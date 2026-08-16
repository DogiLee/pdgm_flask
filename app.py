"""
PDGM · Baskı Dizgi Atölyesi İş Takip Sistemi
--------------------------------------------
Flask + Excel tabanlı, intranet için tek-process sunucu sürümü.

Önemli mimari not:
- Gerçek uygulama state'i hâlâ Excel dosyalarında tutulur.
- Aynı process içindeki eşzamanlı istekler depo.py içindeki RLock ile korunur.
- Birden fazla Python/Waitress process'i aynı data klasörünü paylaşmamalıdır.
- Uzun vadede SQL veritabanına geçiş önerilir.
"""

from __future__ import annotations

import json
import logging
import logging.handlers
import os
import secrets
from collections import OrderedDict
from datetime import date, datetime, timedelta
from functools import wraps
from urllib.parse import urlsplit
from uuid import uuid4

from flask import (
    Flask,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename

import depo
import excel_araclari as ex

KOK = os.path.dirname(os.path.abspath(__file__))
VERI_KLASORU = os.path.join(KOK, "data")
YUKLEME_KLASORU = os.path.join(VERI_KLASORU, "yuklenen_exceller")
KULLANICI_DOSYASI = os.path.join(VERI_KLASORU, "kullanicilar.json")
LOG_DOSYASI = os.path.join(VERI_KLASORU, "uygulama.log")
# macOS AirPlay sıkça 5000'i tuttuğu için varsayılan 5001.
SUNUCU_PORTU = int(os.environ.get("PDGM_PORT", "5001"))

app = Flask(__name__)


# ---------------------------------------------------------------- güvenli config

def _anahtar() -> str:
    """Session anahtarını diskte saklar; ilk çalıştırmada rastgele üretir."""
    yol = os.path.join(VERI_KLASORU, "gizli.key")
    os.makedirs(os.path.dirname(yol), exist_ok=True)
    if not os.path.exists(yol):
        with open(yol, "w", encoding="utf-8") as f:
            f.write(secrets.token_hex(32))
        try:
            os.chmod(yol, 0o600)
        except OSError:
            pass
    with open(yol, encoding="utf-8") as f:
        anahtar = f.read().strip()
    if len(anahtar) < 32:
        raise RuntimeError("data/gizli.key geçersiz veya çok kısa.")
    return anahtar


app.secret_key = _anahtar()
app.config.update(
    MAX_CONTENT_LENGTH=25 * 1024 * 1024,
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=os.environ.get("PDGM_HTTPS", "0") == "1",
    PERMANENT_SESSION_LIFETIME=timedelta(hours=12),
)


def _gunluk_dosya_logu_kur():
    os.makedirs(VERI_KLASORU, exist_ok=True)
    if any(isinstance(h, logging.FileHandler) for h in app.logger.handlers):
        return
    handler = logging.handlers.RotatingFileHandler(
        LOG_DOSYASI,
        maxBytes=2_000_000,
        backupCount=5,
        encoding="utf-8",
    )
    handler.setFormatter(
        logging.Formatter("%(asctime)s %(levelname)s %(message)s")
    )
    app.logger.setLevel(logging.INFO)
    app.logger.addHandler(handler)
    # Werkzeug access log ayrı; uygulama hataları burada kalsın.
    logging.getLogger("waitress").setLevel(logging.INFO)


def _kullanicilari_yukle() -> dict[str, dict]:
    """
    Kullanıcıları JSON dosyasından yükler.

    Dosya yoksa ilk kullanıcıları oluşturur ve SADECE ilk çalıştırmada üretilen
    parolaları konsola yazar. JSON içinde yalnız password hash saklanır.

    İstenirse ilk kurulum parolaları environment variable ile verilebilir:
      PDGM_ADMIN_PASSWORD
      PDGM_OPERATOR_PASSWORD
      PDGM_VIEWER_PASSWORD
    """
    os.makedirs(VERI_KLASORU, exist_ok=True)

    if os.path.exists(KULLANICI_DOSYASI):
        with open(KULLANICI_DOSYASI, encoding="utf-8") as f:
            veri = json.load(f)
        if not isinstance(veri, dict) or not veri:
            raise RuntimeError("data/kullanicilar.json geçersiz.")
        return veri

    baslangic = {
        "admin": {
            "rol": "admin",
            "ad": "Sistem Yöneticisi",
            "env": "PDGM_ADMIN_PASSWORD",
        },
        "operator": {
            "rol": "operator",
            "ad": "Dizgi Operatörü",
            "env": "PDGM_OPERATOR_PASSWORD",
        },
        "izleyici": {
            "rol": "izleyici",
            "ad": "İzleyici",
            "env": "PDGM_VIEWER_PASSWORD",
        },
    }

    sonuc: dict[str, dict] = {}
    ilk_parolalar: dict[str, str] = {}
    for kullanici, bilgi in baslangic.items():
        parola = os.environ.get(bilgi["env"]) or secrets.token_urlsafe(12)
        ilk_parolalar[kullanici] = parola
        sonuc[kullanici] = {
            "sifre_hash": generate_password_hash(parola),
            "rol": bilgi["rol"],
            "ad": bilgi["ad"],
            "aktif": True,
        }

    gecici = KULLANICI_DOSYASI + ".yeni"
    with open(gecici, "w", encoding="utf-8") as f:
        json.dump(sonuc, f, ensure_ascii=False, indent=2)
    os.replace(gecici, KULLANICI_DOSYASI)

    print("\n  İLK KURULUM KULLANICI PAROLALARI")
    print("  --------------------------------")
    for kullanici, parola in ilk_parolalar.items():
        print(f"  {kullanici:10}: {parola}")
    print("  Bu parolaları güvenli bir yerde saklayın. JSON dosyasında yalnız hash vardır.\n")

    return sonuc


_kullanici_mtime: float | None = None
_kullanici_onbellek: dict[str, dict] | None = None


def _kullanicilari_al() -> dict[str, dict]:
    """Diskteki kullanicilar.json değişince yeniden okur (restart gerekmez)."""
    global _kullanici_mtime, _kullanici_onbellek
    try:
        mtime = os.path.getmtime(KULLANICI_DOSYASI)
    except OSError:
        if _kullanici_onbellek is None:
            _kullanici_onbellek = _kullanicilari_yukle()
            try:
                _kullanici_mtime = os.path.getmtime(KULLANICI_DOSYASI)
            except OSError:
                _kullanici_mtime = None
        return _kullanici_onbellek

    if _kullanici_onbellek is None or _kullanici_mtime != mtime:
        with open(KULLANICI_DOSYASI, encoding="utf-8") as f:
            veri = json.load(f)
        if not isinstance(veri, dict) or not veri:
            raise RuntimeError("data/kullanicilar.json geçersiz.")
        _kullanici_onbellek = veri
        _kullanici_mtime = mtime
    return _kullanici_onbellek


_kullanicilari_yukle()
_gunluk_dosya_logu_kur()

# Kayıt dosyalarını yükle / yoksa oluştur.
depo.kur()


# ------------------------------------------------------------------- yetkiler

@app.before_request
def _oturum_kullanici_kontrol():
    """Pasifleştirilmiş veya silinmiş kullanıcıyı oturumda tutma."""
    ad = session.get("kullanici")
    if not ad:
        return None
    kayit = _kullanicilari_al().get(ad)
    if not kayit or not kayit.get("aktif", True):
        session.clear()
        if request.path.startswith("/api/"):
            return jsonify(hata="Oturum sonlandırıldı. Tekrar giriş yapın."), 401
        flash("Hesabınız pasif veya bulunamadı. Tekrar giriş yapın.", "hata")
        return redirect(url_for("giris"))
    # Rol/ad diskten tazelensin (CLI ile değişmiş olabilir).
    session["rol"] = kayit.get("rol") or session.get("rol")
    session["ad"] = kayit.get("ad") or session.get("ad")
    return None

def yetki(*roller):
    def sarmalayici(fn):
        @wraps(fn)
        def ic(*a, **kw):
            if "kullanici" not in session:
                return redirect(url_for("giris", devam=request.path))
            if roller and session.get("rol") not in roller:
                return render_template("yetkisiz.html"), 403
            return fn(*a, **kw)

        return ic

    return sarmalayici


def _csrf_token_uret() -> str:
    token = session.get("csrf_token")
    if not token:
        token = secrets.token_urlsafe(32)
        session["csrf_token"] = token
    return token


def csrf_koru(fn):
    """Cookie-session kullanan state-changing endpoint'leri CSRF'ye karşı korur."""

    @wraps(fn)
    def ic(*a, **kw):
        beklenen = session.get("csrf_token")
        gelen = request.headers.get("X-CSRF-Token") or request.form.get("_csrf_token")
        if not beklenen or not gelen or not secrets.compare_digest(beklenen, gelen):
            if request.path.startswith("/api/"):
                return jsonify(hata="Geçersiz veya eksik CSRF token."), 403
            flash("Oturum doğrulaması başarısız. Sayfayı yenileyip tekrar deneyin.", "hata")
            return redirect(request.referrer or url_for("ana"))
        return fn(*a, **kw)

    return ic


def _guvenli_devam_hedefi(hedef: str | None) -> bool:
    if not hedef:
        return False
    parca = urlsplit(hedef)
    return not parca.scheme and not parca.netloc and hedef.startswith("/") and not hedef.startswith("//")


@app.context_processor
def genel_degiskenler():
    return {
        "oturum_ad": session.get("ad"),
        "oturum_rol": session.get("rol"),
        "oturum_kullanici": session.get("kullanici"),
        "bugun": date.today().strftime("%d.%m.%Y"),
        "csrf_token": _csrf_token_uret() if "kullanici" in session else "",
    }


@app.template_filter("gun")
def gun_filtresi(deger):
    """2026-08-04 09:15:00 -> 04.08.2026"""
    if not deger:
        return "—"
    p = str(deger)[:10].split("-")
    return f"{p[2]}.{p[1]}.{p[0]}" if len(p) == 3 else str(deger)


# --------------------------------------------------------------------- giriş

@app.route("/giris", methods=["GET", "POST"])
def giris():
    if request.method == "POST":
        ad = (request.form.get("kullanici") or "").strip()
        sifre = request.form.get("sifre") or ""
        kayit = _kullanicilari_al().get(ad)

        if (
            kayit
            and kayit.get("aktif", True)
            and check_password_hash(kayit.get("sifre_hash", ""), sifre)
        ):
            session.clear()
            session.permanent = True
            session.update(kullanici=ad, rol=kayit["rol"], ad=kayit["ad"])
            _csrf_token_uret()
            depo.log_ekle(ad, kayit["rol"], "GİRİŞ YAPILDI")
            hedef = request.args.get("devam")
            return redirect(hedef if _guvenli_devam_hedefi(hedef) else url_for("ana"))

        # Kullanıcı adını loglamak audit için faydalı; parola hiçbir zaman loglanmaz.
        depo.log_ekle(ad or "-", "-", "HATALI GİRİŞ")
        return render_template("giris.html", hata="Kullanıcı adı veya şifre hatalı."), 401

    return render_template("giris.html")


@app.route("/cikis", methods=["POST"])
@yetki("admin", "operator", "izleyici")
@csrf_koru
def cikis():
    depo.log_ekle(session["kullanici"], session.get("rol", ""), "ÇIKIŞ YAPILDI")
    session.clear()
    return redirect(url_for("giris"))


@app.route("/")
def ana():
    rol = session.get("rol")
    if rol == "admin":
        return redirect(url_for("yonetim"))
    if rol == "operator":
        return redirect(url_for("operator"))
    if rol == "izleyici":
        return redirect(url_for("panel"))
    return redirect(url_for("giris"))


# ------------------------------------------------------------------- ekranlar

def _pano_verisi():
    kartlar = depo.kartlari_getir()
    sayac = {
        "devam": sum(1 for k in kartlar if k["durum"] == depo.DEVAM),
        "sirada": sum(1 for k in kartlar if k["durum"] == depo.SIRADA),
        "tamam": sum(1 for k in kartlar if k["durum"] == depo.TAMAM),
        "gecikme": sum(
            1 for k in kartlar if k["renk"] == "kotu" and k["durum"] != depo.TAMAM
        ),
    }
    return {
        "kartlar": kartlar,
        "sayac": sayac,
        "guncelleme": datetime.now().strftime("%H:%M:%S"),
    }


@app.route("/panel")
@yetki("admin", "operator", "izleyici")
def panel():
    veri = _pano_verisi()
    kartlar = veri["kartlar"]
    tamamlanan = sorted(
        [k for k in kartlar if k["durum"] == depo.TAMAM],
        key=lambda k: k.get("bitis_zamani") or k.get("gerceklesen_teslim") or "",
        reverse=True,
    )[:5]
    return render_template(
        "panel.html",
        sayac=veri["sayac"],
        guncelleme=veri["guncelleme"],
        devam=[k for k in kartlar if k["durum"] == depo.DEVAM],
        sirada=[k for k in kartlar if k["durum"] == depo.SIRADA][:12],
        tamamlanan=tamamlanan,
    )


@app.route("/operator")
@yetki("admin", "operator")
def operator():
    veri = _pano_verisi()
    return render_template("operator.html", kartlar=veri["kartlar"], sayac=veri["sayac"])


@app.route("/yonetim")
@yetki("admin")
def yonetim():
    veri = _pano_verisi()
    kartlar = veri["kartlar"]
    kaynakta_olmayan = [
        k
        for k in kartlar
        if k.get("kaynakta_yok") and k.get("durum") != depo.TAMAM
    ]
    return render_template(
        "yonetim.html",
        kartlar=kartlar,
        kaynakta_olmayan=kaynakta_olmayan,
        yuklemeler=depo.yuklemeleri_getir(8),
        loglar=depo.loglari_getir(25),
    )


@app.route("/ozet")
@yetki("admin", "izleyici", "operator")
def ozet():
    return render_template("ozet.html", **ozet_hesapla())


@app.route("/api/veriler")
@yetki("admin", "operator", "izleyici")
def api_veriler():
    return jsonify(_pano_verisi())


# ------------------------------------------------------------- operator işlemleri

@app.route("/api/basla", methods=["POST"])
@yetki("admin", "operator")
@csrf_koru
def api_basla():
    veri = request.get_json(silent=True) or {}
    try:
        kart = depo.kart_baslat(
            kart_id=veri.get("kart_id"),
            adet=veri.get("adet"),
            kullanici=session["kullanici"],
            rol=session["rol"],
            aciklama=(veri.get("not") or "").strip(),
        )
    except depo.KartBulunamadi as hata:
        return jsonify(hata=str(hata)), 404
    except depo.IsKuralHatasi as hata:
        return jsonify(hata=str(hata)), 409
    except (TypeError, ValueError) as hata:
        return jsonify(hata=str(hata)), 400

    return jsonify(tamam=True, kart=kart)


@app.route("/api/bitir", methods=["POST"])
@yetki("admin", "operator")
@csrf_koru
def api_bitir():
    veri = request.get_json(silent=True) or {}
    try:
        kart, bitti, mesaj = depo.kart_bitir(
            kart_id=veri.get("kart_id"),
            adet=veri.get("adet"),
            kullanici=session["kullanici"],
            rol=session["rol"],
            aciklama=(veri.get("not") or "").strip(),
        )
    except depo.KartBulunamadi as hata:
        return jsonify(hata=str(hata)), 404
    except depo.IsKuralHatasi as hata:
        return jsonify(hata=str(hata)), 409
    except (TypeError, ValueError) as hata:
        return jsonify(hata=str(hata)), 400

    return jsonify(tamam=True, bitti=bitti, mesaj=mesaj, kart=kart)


@app.route("/api/not", methods=["POST"])
@yetki("admin", "operator")
@csrf_koru
def api_not():
    veri = request.get_json(silent=True) or {}
    try:
        kart = depo.kart_not_guncelle(
            kart_id=veri.get("kart_id"),
            aciklama=(veri.get("not") or "").strip(),
            kullanici=session["kullanici"],
            rol=session["rol"],
        )
    except depo.KartBulunamadi as hata:
        return jsonify(hata=str(hata)), 404
    return jsonify(tamam=True, kart=kart)


# ---------------------------------------------------------------- admin işlemleri

@app.route("/api/admin/duzenle", methods=["POST"])
@yetki("admin")
@csrf_koru
def api_duzenle():
    veri = request.get_json(silent=True) or {}
    try:
        kart = depo.admin_kart_duzenle(
            kart_id=veri.get("kart_id"),
            durum=veri.get("durum"),
            tamamlanan_adet=veri.get("tamamlanan_adet"),
            toplam_adet=veri.get("toplam_adet"),
            aciklama=veri.get("not"),
            kullanici=session["kullanici"],
        )
    except depo.KartBulunamadi as hata:
        return jsonify(hata=str(hata)), 404
    except (depo.IsKuralHatasi, TypeError, ValueError) as hata:
        return jsonify(hata=str(hata)), 400
    return jsonify(tamam=True, kart=kart)


@app.route("/api/admin/kart-sil", methods=["POST"])
@yetki("admin")
@csrf_koru
def api_kart_sil():
    veri = request.get_json(silent=True) or {}
    try:
        depo.admin_kart_gizle(veri.get("kart_id"), session["kullanici"])
    except depo.KartBulunamadi as hata:
        return jsonify(hata=str(hata)), 404
    return jsonify(tamam=True)


# ------------------------------------------------------------ excel yükle/indir

@app.route("/yonetim/yukle", methods=["POST"])
@yetki("admin")
@csrf_koru
def yukle():
    dosya = request.files.get("dosya")
    if not dosya or not dosya.filename:
        flash("Dosya seçilmedi.", "hata")
        return redirect(url_for("yonetim"))

    guvenli_ad = secure_filename(dosya.filename)
    if not guvenli_ad.lower().endswith((".xlsx", ".xlsm")):
        flash("Sadece .xlsx veya .xlsm dosyası yükleyin.", "hata")
        return redirect(url_for("yonetim"))

    os.makedirs(YUKLEME_KLASORU, exist_ok=True)
    ad = f"{datetime.now():%Y%m%d_%H%M%S}_{uuid4().hex[:8]}_{guvenli_ad}"
    yol = os.path.join(YUKLEME_KLASORU, ad)
    dosya.save(yol)

    try:
        sonuc = ex.excelden_aktar(yol, session["kullanici"])
        mesaj = (
            f"Excel aktarıldı · {sonuc['satir']} satır · "
            f"{sonuc['yeni']} yeni kart · {sonuc['guncellenen']} kartın planı güncellendi. "
            f"Mevcut iş durumu/tamamlanan adet korundu ({sonuc.get('workflow_korundu', 0)} kart). "
            f"{sonuc['pasife_alinan']} açık kart son Excel'de yok (ekranda 'Son Excelde yok' olarak kalır)."
        )
        if sonuc.get("pasife_listesi"):
            mesaj += " Örnek: " + ", ".join(sonuc["pasife_listesi"][:5])
        if sonuc.get("uyari"):
            mesaj += f" · {sonuc['uyari']} uyarı"
        if sonuc.get("yedek"):
            mesaj += f" · Yedek: {os.path.basename(sonuc['yedek'])}"
        flash(mesaj, "basari")
        app.logger.info("Excel import OK: %s", mesaj)
    except (ex.ExcelAktarimHatasi, depo.IsKuralHatasi) as hata:
        flash(f"Excel içeriği kabul edilmedi: {hata}", "hata")
        app.logger.warning("Excel import reddedildi: %s", hata)
    except Exception:  # noqa: BLE001
        app.logger.exception("Excel import sırasında beklenmeyen hata")
        flash("Excel okunurken beklenmeyen bir hata oluştu. Sistem yöneticisine başvurun.", "hata")

    return redirect(url_for("yonetim"))


@app.route("/yonetim/yeniden-oku", methods=["POST"])
@yetki("admin")
@csrf_koru
def yeniden_oku():
    """Sadece kartlar.xlsx dosyasını schema + invariant kontrolüyle yeniden okur."""
    try:
        adet = depo.kartlari_diskten_yeniden_yukle()
    except depo.VeriDogrulamaHatasi as hata:
        flash(f"Kayıt dosyası yeniden okunamadı: {hata}", "hata")
        return redirect(url_for("yonetim"))

    depo.log_ekle(
        session["kullanici"],
        "admin",
        "KART DOSYASI YENİDEN OKUNDU",
        detay=f"{adet} kart",
    )
    flash(f"Kayıt dosyası yeniden okundu · {adet} kart.", "basari")
    return redirect(url_for("yonetim"))


@app.route("/yonetim/kayit-dosyasi/<hangi>")
@yetki("admin")
def kayit_dosyasi(hangi):
    dosyalar = {
        "kartlar": depo.KARTLAR_DOSYA,
        "log": depo.LOG_DOSYA,
        "yuklemeler": depo.YUKLEME_DOSYA,
    }
    yol = dosyalar.get(hangi)
    if not yol or not os.path.exists(yol):
        flash("Dosya henüz oluşmamış.", "hata")
        return redirect(url_for("yonetim"))
    depo.log_ekle(session["kullanici"], "admin", "KAYIT DOSYASI İNDİRİLDİ", detay=hangi)
    return send_file(yol, as_attachment=True, download_name=os.path.basename(yol))


@app.route("/yonetim/rapor")
@yetki("admin")
def rapor_indir():
    o = ozet_hesapla()
    ozet_satirlari = [
        ["Rapor tarihi", datetime.now().strftime("%d.%m.%Y %H:%M")],
        ["Toplam kart", o["genel"]["toplam"]],
        ["Sırada", o["genel"]["sirada"]],
        ["Devam eden", o["genel"]["devam"]],
        ["Tamamlanan", o["genel"]["tamam"]],
        ["Süresi aşan (açık işler)", o["genel"]["gecikme"]],
        ["Bu hafta tamamlanan kart", o["donemler"]["Bu hafta"]["kart"]],
        ["Bu ay tamamlanan kart", o["donemler"]["Bu ay"]["kart"]],
        ["Bu yıl tamamlanan kart", o["donemler"]["Bu yıl"]["kart"]],
        ["Bu ay zamanında biten (%)", o["donemler"]["Bu ay"]["zamaninda_yuzde"]],
        ["Bu ay ortalama sapma (gün)", o["donemler"]["Bu ay"]["ort_sapma"]],
    ]
    wb = ex.calisma_kitabi_uret(
        depo.kartlari_getir(sadece_gorunen=False),
        depo.loglari_getir(),
        ozet_satirlari,
    )
    depo.log_ekle(session["kullanici"], "admin", "RAPOR İNDİRİLDİ")
    return send_file(
        ex.kitap_baytlari(wb),
        as_attachment=True,
        download_name=ex.dosya_adi("PDGM_Rapor"),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ------------------------------------------------------------------ özet hesap

def _tamamlanma_tarihi(kart) -> str | None:
    deger = kart.get("bitis_zamani") or kart.get("gerceklesen_teslim")
    return str(deger)[:10] if deger else None


def _donem_ozeti(kartlar, baslangic, bitis):
    alt = baslangic.strftime("%Y-%m-%d")
    ust = bitis.strftime("%Y-%m-%d")
    secilen = []

    for k in kartlar:
        if k["durum"] != depo.TAMAM:
            continue
        tamamlanma = _tamamlanma_tarihi(k)
        if tamamlanma and alt <= tamamlanma <= ust:
            secilen.append(k)

    sapmalar = [k["sapma"] for k in secilen if k["sapma"] is not None]
    zamaninda = sum(1 for s in sapmalar if s <= 0)
    return {
        "kart": len(secilen),
        "adet": sum(k["tamamlanan_adet"] for k in secilen),
        "zamaninda": zamaninda,
        "gecikmeli": len(sapmalar) - zamaninda,
        "zamaninda_yuzde": round(zamaninda / len(sapmalar) * 100) if sapmalar else 0,
        "ort_sapma": round(sum(sapmalar) / len(sapmalar), 1) if sapmalar else 0,
    }


def ozet_hesapla():
    # Özet ve tarihsel rapor için admin tarafından gizlenenler hariç tüm tarihçe gerekir.
    kartlar = depo.kartlari_getir(sadece_gorunen=False)
    bugun = date.today()

    donemler = OrderedDict()
    donemler["Bu hafta"] = _donem_ozeti(
        kartlar, bugun - timedelta(days=bugun.weekday()), bugun
    )
    donemler["Bu ay"] = _donem_ozeti(kartlar, bugun.replace(day=1), bugun)
    donemler["Bu yıl"] = _donem_ozeti(kartlar, bugun.replace(month=1, day=1), bugun)

    haftalar = []
    for geri in range(7, -1, -1):
        bas = bugun - timedelta(days=bugun.weekday() + geri * 7)
        son = min(bas + timedelta(days=6), bugun)
        b, s = bas.strftime("%Y-%m-%d"), son.strftime("%Y-%m-%d")

        planlanan = sum(
            1 for k in kartlar if k.get("plan_teslim") and b <= k["plan_teslim"] <= s
        )
        tamamlanan = sum(
            1
            for k in kartlar
            if k["durum"] == depo.TAMAM
            and (tamamlanma := _tamamlanma_tarihi(k))
            and b <= tamamlanma <= s
        )
        haftalar.append(
            {
                "etiket": bas.strftime("%d.%m"),
                "hafta_no": bas.isocalendar()[1],
                "planlanan": planlanan,
                "tamamlanan": tamamlanan,
                "sapma": tamamlanan - planlanan,
            }
        )

    en_yuksek = max(
        [h["planlanan"] for h in haftalar]
        + [h["tamamlanan"] for h in haftalar]
        + [1]
    )

    genel = {
        "toplam": len(kartlar),
        "sirada": sum(1 for k in kartlar if k["durum"] == depo.SIRADA and k["gorunur"]),
        "devam": sum(1 for k in kartlar if k["durum"] == depo.DEVAM and k["gorunur"]),
        "tamam": sum(1 for k in kartlar if k["durum"] == depo.TAMAM),
        "gecikme": sum(
            1
            for k in kartlar
            if k["gorunur"] and k["renk"] == "kotu" and k["durum"] != depo.TAMAM
        ),
    }
    geciken = [
        k
        for k in kartlar
        if k["gorunur"] and k["renk"] == "kotu" and k["durum"] != depo.TAMAM
    ]

    return {
        "genel": genel,
        "donemler": donemler,
        "haftalar": haftalar,
        "en_yuksek": en_yuksek,
        "geciken_kartlar": geciken,
    }


# ---------------------------------------------------------------------- çalıştır

def calistir():
    depo.process_kilidi_al()
    depo.kur()
    _gunluk_dosya_logu_kur()

    print("\n  PDGM İş Takip Sistemi çalışıyor")
    print(f"  Bu bilgisayarda : http://127.0.0.1:{SUNUCU_PORTU}")
    print(f"  Ağdaki diğer PC : http://<sunucunun-ip-adresi>:{SUNUCU_PORTU}")
    print(f"  Kayıtlar        : {VERI_KLASORU}")
    print("  Sunucu modeli   : tek process + çok thread (ikinci process engelli)")
    print(f"  Uygulama logu   : {LOG_DOSYASI}")
    print("  Durdurmak için  : Ctrl + C\n")
    app.logger.info("Sunucu başladı port=%s", SUNUCU_PORTU)

    try:
        from waitress import serve

        serve(app, host="0.0.0.0", port=SUNUCU_PORTU, threads=8)
    except ImportError:
        app.run(host="0.0.0.0", port=SUNUCU_PORTU, debug=False, threaded=True)


if __name__ == "__main__":
    calistir()
