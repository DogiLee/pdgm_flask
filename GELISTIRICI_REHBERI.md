# PDGM — Geliştirici Rehberi

Bu doküman **koda özellik eklemek**, mevcut ekranları genişletmek veya bu sistemi bir **portal / API katmanının** parçası yapmak isteyen geliştiriciler içindir.

İş kurallarının (Excel merge, durumlar, yedek) kullanıcı diliyle anlatımı için: [`SISTEM_NASIL_CALISIR.md`](./SISTEM_NASIL_CALISIR.md)

---

## 1. Proje haritası

```text
pdgm_flask/
├── app.py                 # HTTP katmanı: route, oturum, CSRF, şablon bağlama
├── depo.py                # Domain + persistence (Excel state, kilit, iş kuralları)
├── excel_araclari.py      # Haftalık Excel parse/validate + rapor workbook
├── kullanici_yonet.py     # CLI kullanıcı yönetimi
├── templates/             # Jinja2 ekranlar
│   ├── base.html          # Layout, CSRF meta, pdgmFetch(), toast
│   ├── giris.html
│   ├── panel.html
│   ├── operator.html
│   ├── yonetim.html
│   ├── ozet.html
│   └── yetkisiz.html
├── static/stil.css
├── data/                  # Runtime veri (gitignore'lu parçalar)
├── ornek_exceller/        # Test Excel'leri
├── SISTEM_NASIL_CALISIR.md
└── GELISTIRICI_REHBERI.md  # bu dosya
```

### Katman kuralı (kritik)

| Katman | Sorumluluk | Yapmaması gereken |
|--------|------------|-------------------|
| `app.py` | Auth, HTTP, flash, JSON cevap, template context | İş kuralı / Excel yazma |
| `depo.py` | Kart state, invariant, lock, dosya yazma | Flask `request` / `session` bilmek |
| `excel_araclari.py` | Dış Excel'i parse etmek, `depo.excel_import_uygula` çağırmak | Doğrudan global state mutasyonu |
| `templates/` | Görünüm + ince JS | İş kuralını JS’te yeniden yazmak |

**Özellik eklerken önce `depo.py` fonksiyonu yaz; sonra `app.py` route; sonra template.**  
Bu sıra hem Flask UI hem ileride portal/API için aynı kalır.

---

## 2. Bellek modeli ve kilit

`depo.py` açılışta Excel’leri okur, RAM’de tutar:

```text
_kartlar: list[dict]
_loglar: list[dict]
_yuklemeler: list[dict]
_kilit: threading.RLock
```

- Okuma/yazma **`with _kilit:`** altında yapılır.
- Dışarıya **kopya** verilir (`kart_gorunumu` / `deepcopy`). Global dict referansı sızdırma.
- Kart + log birlikte değişecekse: önce snapshot al → değiştir → `_kart_log_commit` / `_coklu_yaz` → hata olursa snapshot’a dön.

Yeni state değiştiren fonksiyon iskeleti:

```python
def ornek_islem(kart_id, kullanici, rol, ...):
    global _kartlar, _loglar
    with _kilit:
        eski_kartlar = copy.deepcopy(_kartlar)
        eski_loglar = copy.deepcopy(_loglar)
        try:
            kart = _kart_ref(kart_id)
            if not kart:
                raise KartBulunamadi(...)
            # ... kurallar ...
            _loglar.append(_log_kaydi(...))
            _kart_log_commit(eski_kartlar, eski_loglar)
            return kart_gorunumu(kart)
        except Exception:
            _kartlar = eski_kartlar
            _loglar = eski_loglar
            raise
```

`kart_guncelle` / `toplu_kaydet` **legacy**; yeni kodda kullanma.

---

## 3. Kart şeması

Excel sütun eşlemesi `KART_ALANLARI` ile yapılır: `(ExcelBaşlık, python_key)`.

Önemli alanlar:

| Key | Anlam |
|-----|--------|
| `id` | Dahili sayısal kimlik |
| `anahtar` | `f"{talep_no}|{stok_no}"` — Excel merge anahtarı |
| `durum` | `SIRADA` \| `DEVAM EDİYOR` \| `TAMAMLANDI` |
| `toplam_adet` / `tamamlanan_adet` | Üretim miktarı |
| `excel_durum` | Kaynak Excel’deki ham durum metni |
| `source_active` | Son import batch’inde var mı (0/1) |
| `admin_gizli` | Admin gizledi (import geri açmaz) |
| `aktif` | Legacy yaşam döngüsü |
| `kaynak` | Genelde `"EXCEL"` |

Yeni **kalıcı alan** eklemek:

1. `KART_ALANLARI` sonuna `( "Başlık", "python_key" )` ekle.
2. Sayısalysa `SAYISAL` set’ine ekle.
3. Gerekirse `_kart_normalize` / `_kart_dogrula` güncelle.
4. Eski `kartlar.xlsx` dosyalarında sütun yoksa `_oku` boş/None getirir — default ver.
5. UI / `kart_gorunumu` / rapor (`excel_araclari.calisma_kitabi_uret`) güncelle.

> Sadece ekranda hesaplanan alanlar (`kalan_adet`, `rozet`, `kaynakta_yok`) Excel’e yazılmaz; `kart_gorunumu` içinde üretilir.

---

## 4. Hazır domain API (`depo.py`)

Özellik veya portal yazarken bunları tercih et:

| Fonksiyon | Ne yapar |
|-----------|----------|
| `kur()` | Dosyaları yükle / yoksa oluştur |
| `kartlari_getir(sadece_gorunen=True)` | Liste + görünüm alanları |
| `kart_getir(id)` | Tek kart |
| `kart_baslat(...)` | SIRADA → DEVAM |
| `kart_bitir(...)` | Adet tamamla |
| `kart_not_guncelle(...)` | Not |
| `admin_kart_duzenle(...)` | Admin müdahale |
| `admin_kart_gizle(...)` | Gizle |
| `excel_import_uygula(...)` | Validate edilmiş satırları merge et |
| `log_ekle` / `loglari_getir` | Audit |
| `anlik_yedek(etiket)` | Timestamp’li yedek klasörü |
| `process_kilidi_al()` | İkinci process engeli |

Exception’lar:

- `KartBulunamadi` → HTTP 404
- `IsKuralHatasi` → HTTP 409 (iş kuralı)
- `VeriDogrulamaHatasi` → dosya/schema
- `excel_araclari.ExcelAktarimHatasi` → import parse

`app.py` içinde bu ayrımı koru; generic `Exception`’da path sızdırma.

---

## 5. HTTP / auth yüzeyi (`app.py`)

### Roller

```text
admin | operator | izleyici
```

Koruma:

```python
@app.route("/ornek")
@yetki("admin", "operator")
def ornek():
    ...
```

State değiştiren endpoint:

```python
@app.route("/api/ornek", methods=["POST"])
@yetki("admin", "operator")
@csrf_koru
def api_ornek():
    ...
```

### CSRF

- Form: `<input type="hidden" name="_csrf_token" value="{{ csrf_token }}">`
- `fetch`: `base.html` içindeki `pdgmFetch()` otomatik `X-CSRF-Token` ekler.
- Login (`/giris`) CSRF’siz (oturum yok); sonrası zorunlu.

### Kullanıcılar

- Dosya: `data/kullanicilar.json`
- Runtime: `_kullanicilari_al()` — mtime ile hot-reload
- CLI: `kullanici_yonet.py`
- `@app.before_request`: pasif kullanıcıyı düşürür, rol/ad tazeler

### Mevcut route özeti

| Route | Metod | Rol | Not |
|-------|-------|-----|-----|
| `/giris` | GET/POST | — | Login |
| `/cikis` | POST | hepsi | CSRF |
| `/panel` | GET | hepsi | Pano |
| `/operator` | GET | admin, operator | |
| `/yonetim` | GET | admin | |
| `/ozet` | GET | hepsi | |
| `/api/veriler` | GET | hepsi | JSON pano |
| `/api/basla` | POST | admin, operator | JSON + CSRF |
| `/api/bitir` | POST | admin, operator | |
| `/api/not` | POST | admin, operator | |
| `/api/admin/duzenle` | POST | admin | |
| `/api/admin/kart-sil` | POST | admin | Gizle |
| `/yonetim/yukle` | POST | admin | Multipart Excel |
| `/yonetim/yeniden-oku` | POST | admin | |
| `/yonetim/kayit-dosyasi/<hangi>` | GET | admin | |
| `/yonetim/rapor` | GET | admin | |

---

## 6. Frontend kalıpları

### Yeni sayfa

1. `templates/ornek.html` → `{% extends "base.html" %}`
2. Nav linki `base.html` içinde role göre
3. `app.py` route + `render_template`
4. Stiller: mümkünse mevcut sınıflar (`panel-kutu`, `buton`, `durum-rozet`, …) — `static/stil.css`

### JSON işlem (operatör tarzı)

```javascript
const data = await pdgmFetch("/api/basla", {
  method: "POST",
  body: JSON.stringify({ kart_id: 1, adet: 10, not: "" })
});
toast(data.mesaj || "Tamam");
```

Hata: `pdgmFetch` `throw`; `hataMesaji(err)` toast gösterir.

### Dialog

Native `<dialog>`; kapatma: `[data-dialog-kapat]` (base.js dinleyicisi).

### Operatör filtre state

`sessionStorage` anahtarları: `pdgm-op-filtre`, `pdgm-op-arama`, `pdgm-op-scroll` — reload sonrası koruma.

---

## 7. Özellik ekleme tarifleri

### A) Yeni operatör aksiyonu (ör. “Duraklat”)

1. `depo.py`: `kart_duraklat(...)` + invariant + log + commit  
2. `app.py`: `/api/duraklat` + `@yetki` + `@csrf_koru`  
3. `operator.html`: buton + dialog + `pdgmFetch`  
4. Gerekirse `durum` değerlerini `GECERLI_DURUMLAR` / `_kart_dogrula` genişlet  
5. `SISTEM_NASIL_CALISIR.md` güncelle

### B) Yeni rapor alanı / özet metriği

1. Hesabı `app.py` `ozet_hesapla` veya `depo` helper’ında tut  
2. `ozet.html` göster  
3. İndirilebilir rapor için `excel_araclari.calisma_kitabi_uret` sheet’ine ekle

### C) Yeni Excel sütunu (import)

1. `excel_araclari.BASLIK_ESLESME` — normalize başlık → key  
2. Parse döngüsünde `plan` dict’ine koy  
3. `depo.KART_ALANLARI` + merge’de `mevcut.update(plan)` zaten alır (workflow alanlarına dikkat)  
4. Mevcut kartta **ezilmemesi** gereken alanları `excel_import_uygula` içinde açıkça koru (durum/tamamlanan gibi)

### D) Yeni rol

1. `kullanici_yonet.ROLLER`  
2. İlk kurulum / `_kullanicilari_yukle` (gerekirse)  
3. `@yetki(...)` ve `base.html` nav koşulları  
4. Doküman

### E) Kart’a yeni kalıcı alan (UI + Excel)

Bkz. §3. Sonra yönetim tablosu / operatör kartı / filtre `data-arama` string’i.

---

## 8. Portal / dış sistem entegrasyonu

Bu repo şu an **server-rendered intranet uygulaması**. Portal kurmak isteyenler için önerilen yaklaşımlar:

### Seçenek 1 — Aynı process içinde API genişlet (en ucuz)

Mevcut JSON endpoint’ler zaten var (`/api/basla`, `/api/veriler`, …).

Portal veya mobil istemci için:

1. Yeni `/api/...` route’ları ekle; işi yalnız `depo.*` yapsın.  
2. Auth: kısa vadede session cookie (aynı origin) veya ileride token.  
3. CSRF: cookie session kullanıyorsan state-changing API’de CSRF veya SameSite stratejisini bilinçli seç.  
4. Cevapları stabil JSON şemasına bağla (aşağıdaki örnek).

Örnek cevap sözleşmesi:

```json
{
  "tamam": true,
  "kart": { "id": 1, "durum": "DEVAM EDİYOR", "tamamlanan_adet": 10, "...": "..." }
}
```

Hata:

```json
{ "hata": "Kalan adetten fazla tamamlanamaz." }
```

### Seçenek 2 — `depo.py`’yi kütüphane gibi kullan (ayrı portal servisi)

```text
[Portal FastAPI/Flask] ---> import depo / paylaşılan paket
[Mevcut PDGM UI]      ---> aynı depo
```

**Dikkat:** Excel persistence + `process_kilidi` nedeniyle **iki process aynı `data/` klasörünü paylaşamaz.**  
Portal ayrı process olacaksa önce **SQLite (veya tek writer servis)** şart. Excel haliyle “iki app bir data” = veri bozulması.

Pratik kural:

| Senaryo | Uygun mu? |
|---------|-----------|
| Tek Waitress process; UI + API birlikte | Evet |
| İkinci process aynı `data/` Excel’e yazıyor | Hayır |
| Read-only rapor process (dosyayı kopyalayarak) | Dikkatli / mümkün |
| Ortak DB (SQLite/Postgres) | Evet — doğru uzun vadeli portal temeli |

### Seçenek 3 — Reverse proxy ile “portal kabuğu”

```text
https://portal.firma.local/
  /is-takip/   → PDGM (bu uygulama)
  /ik/         → başka sistem
  /mesaj/      → başka sistem
```

Bu uygulama path-prefix ile çalışacak şekilde `APPLICATION_ROOT` / proxy header ayarı gerekebilir (şu an kök `/` varsayar). Portal sadece SSO + menü ise PDGM’i iframe veya alt path olarak bağlamak yeterli olabilir; kod değişmeden SSO için dış auth gateway gerekir.

### Portal için dokunulmaması gerekenler

- Import merge politikasını portal UI’ında yeniden yazma — `excel_import_uygula` tek kaynak.  
- Kart anahtarını (`talep_no|stok_no`) portalda farklı üretme.  
- `kartlar.xlsx`’i portalın doğrudan openpyxl ile güncellemesi — her zaman `depo` üzerinden.

### İleride SQLite’a geçerken portal dostu sınır

Hedef arayüz (kavramsal):

```text
depo.kartlari_getir()
depo.kart_baslat()
depo.excel_import_uygula()
```

Persistence Excel’den SQLite’a geçse bile **app/portal aynı fonksiyonları çağırmaya devam etsin**.  
Bu yüzden bugün bile iş kuralını `app.py` veya template’e gömme.

---

## 9. Excel import’a dokunurken

Akış:

```text
dosya kaydet
  → excel_araclari.excelden_aktar
      → parse + validate (state yok)
      → depo.excel_import_uygula
          → anlik_yedek
          → merge
          → coklu_yaz
```

Kurallar (özet):

- Mevcut kart: plan güncellenir, **workflow ezilmez**.  
- Yeni kart: Excel durumundan ilk `durum`.  
- Eksik açık kart: `source_active=0`, silinmez, UI’da kalır.  
- Çelişen adet: **tüm import abort**.

Test için: `ornek_exceller/01_…` … `04_…`

---

## 10. UI bileşen sözlüğü

| Sınıf / parça | Kullanım |
|---------------|----------|
| `panel-kutu` | Bölüm kartı |
| `istatistik-grid` / `istatistik` | Sayı kutuları |
| `durum-rozet` + `iyi`/`uyari`/`kotu` | Durum etiketi |
| `buton buton-ana` / `buton-basari` / `buton-hayalet` | Aksiyon |
| `arac-cubugu` + `filtre` | Operatör filtre |
| `modal` (`<dialog>`) | Onay / form |
| `toast` / `bildirim` | Geri bildirim |
| `uyari-kutu` | Yönetim uyarı bandı |

CSS değişkenleri `:root` içinde (`--ana`, `--iyi`, …).

---

## 11. Çalıştırma, ortam, güvenlik checklist (dev)

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py                    # :5001
PDGM_PORT=8080 python app.py
PDGM_HTTPS=1 python app.py       # Secure cookie
```

İlk kullanıcı parolaları: `PDGM_ADMIN_PASSWORD` vb. veya konsol çıktısı.

Geliştirirken:

- [ ] State değişimi `@csrf_koru` + doğru HTTP metodu  
- [ ] İş kuralı `depo` içinde  
- [ ] Lock + snapshot + commit  
- [ ] Audit `log_ekle`  
- [ ] Yeni alan varsa `KART_ALANLARI` + rapor  
- [ ] İkinci process ile test etme (lock bilerek)  
- [ ] Excel özelliği ise `ornek_exceller` ile doğrula  

---

## 12. Bilinçli teknik borç (özellik planlarken)

1. Persistence = Excel → multi-writer portal yok.  
2. Kullanıcı store = JSON (yeterli; karmaşık IAM için dış auth düşün).  
3. Özet hesapları request anında RAM’de — kart sayısı çok artarsa cache/SQL.  
4. Otomatik test yok — kritik merge/baslat/bitir için test eklemek yüksek değer.  
5. FastAPI’ye geçiş **zorunlu değil**; portal ihtiyacı doğunca önce stabil `depo` API + gerekirse SQLite.

---

## 13. Hızlı “nereden bakayım?” indeksi

| İstediğin şey | Dosya |
|---------------|--------|
| Yeni ekran / route | `app.py` + `templates/` |
| İş kuralı / kaydetme | `depo.py` |
| Haftalık Excel formatı | `excel_araclari.py` |
| Nav / CSRF / fetch helper | `templates/base.html` |
| Görsel stil | `static/stil.css` |
| Kullanıcı CLI | `kullanici_yonet.py` |
| Operasyon dokümanı | `SISTEM_NASIL_CALISIR.md` |
| Bu rehber | `GELISTIRICI_REHBERI.md` |

---

*Bu rehber, mevcut Flask + Excel mimarisinde güvenli genişleme ve ileride portal/API ayrımı için yazılmıştır. Domain gerçeğinin tek sahibi `depo.py` olmalıdır.*
