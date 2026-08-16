# PDGM İş Takip Sistemi — Nasıl Çalışır?

Bu doküman, **Excel kalıcı katmanı koruyan** güncel sürümün işleyişini anlatır. SQL geçişi bu dokümanın kapsamı dışındadır.

Kod yapısı, özellik ekleme ve portal entegrasyonu için: [`GELISTIRICI_REHBERI.md`](./GELISTIRICI_REHBERI.md)

---

## 1. Sistem ne işe yarar?

PDGM Baskı Dizgi Atölyesi için:

1. Haftalık iş planı **Excel** ile sisteme alınır.
2. Operatörler web üzerinden kartları **başlatır / adet bitirir / not ekler**.
3. Pano ve özet ekranları canlı durumu gösterir.
4. Gerçek durum her zaman sunucudaki **`data/kartlar.xlsx`** dosyasında tutulur.

Kısaca:

| Kaynak | Rolü |
|--------|------|
| Haftalık Excel (yüklenen dosya) | Planlama / kaynak listesi |
| `data/kartlar.xlsx` | Canlı üretim gerçeği |
| `data/islem_logu.xlsx` | Kim ne yaptı (audit) |
| `data/yuklemeler.xlsx` | Excel yükleme geçmişi |
| `data/kullanicilar.json` | Kullanıcılar (yalnız parola hash) |

---

## 2. Mimari özet

```text
[Haftalık Excel] --yükle--> [excel_araclari.py: parse+validate]
                                    |
                                    v
                            [depo.py: merge + yaz]
                                    |
                                    v
                            data/kartlar.xlsx  <---+
                                                    |
[Operatör web] --API--> [depo.py: baslat/bitir/not]-+
                                                    |
                                            data/islem_logu.xlsx
```

- **Tek Python process** + çok thread (Waitress).
- Aynı `data/` klasörünü **ikinci process açamaz** (`data/sunucu.lock`).
- Excel gerçek veritabanı değildir; bu yüzden tek sunucu zorunludur.

### Ana dosyalar

| Dosya | Görev |
|-------|--------|
| `app.py` | HTTP, oturum, roller, CSRF, ekranlar |
| `depo.py` | Kart state, kilit, yedek, Excel merge |
| `excel_araclari.py` | Haftalık Excel parse / validate / rapor |
| `kullanici_yonet.py` | CLI ile kullanıcı ekleme / pasifleştirme |
| `templates/` + `static/` | Arayüz |

---

## 3. Çalıştırma

```bash
cd /path/to/pdgm_flask
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

Varsayılan adres: **http://127.0.0.1:5001**  
(macOS’ta 5000 sıkça AirPlay tarafından tutulduğu için varsayılan 5001’dir.)

Port değiştirmek için:

```bash
PDGM_PORT=8080 python app.py
```

HTTPS reverse proxy arkasında cookie `Secure` olsun diye:

```bash
PDGM_HTTPS=1 python app.py
```

Durdurmak: `Ctrl + C` (lock dosyası temizlenir).

Eski process çöktüyse ve `data/sunucu.lock` kaldıysa, process gerçekten kapalıysa dosyayı silip tekrar başlatın.

---

## 4. Roller ve ekranlar

| Rol | Ne yapar? |
|-----|-----------|
| **admin** | Excel yükleme, kart düzenleme/gizleme, rapor, tüm ekranlar |
| **operator** | Pano, operatör işlemleri, özet |
| **izleyici** | Pano + özet (işlem yok) |

| Ekran | Açıklama |
|-------|----------|
| **Pano** | Devam / sırada / son tamamlananlar (≈30 sn yenilenir) |
| **Operatör** | İşe başla, adet bitir, not; varsayılan filtre **Aktif** |
| **Özet** | Dönem metrikleri ve haftalık grafik |
| **Yönetim** | Excel aktarımı, kart yönetimi, log, yükleme geçmişi |

İlk çalıştırmada `admin` / `operator` / `izleyici` oluşturulur; parolalar **yalnız bir kez** konsola yazılır. Sonra yalnız hash `data/kullanicilar.json` içinde kalır.

Kullanıcı ekleme (restart gerekmez; dosya değişince okunur):

```bash
python kullanici_yonet.py ekle ahmet operator "Ahmet Yılmaz"
python kullanici_yonet.py listele
python kullanici_yonet.py pasif operator
```

---

## 5. Kart nedir? Anahtar nedir?

Her kart şu anahtarla tanınır:

```text
Talep NO + "|" + Kart Stok No
örnek: TLP-2026-0141|PCB-4410-A
```

Bu anahtar Excel satırını sistemdeki kartla eşleştirir.

**Önemli:** Kaynak Excel’de Talep NO veya Stok No düzeltilirse sistem bunu **yeni kart** sanabilir; eski kart “Son Excelde yok” kalır. Yönetim ekranında bu durum için uyarı kutusu vardır.

### İki tür “durum”

| Alan | Anlamı | Kim yönetir? |
|------|--------|--------------|
| **İş durumu** (`durum`) | `SIRADA` / `DEVAM EDİYOR` / `TAMAMLANDI` | Web (operatör / admin) |
| **Excel durumu** (`excel_durum`) | Kaynak dosyadaki metin (`DIZGI`, `TESLIM`, …) | Excel import |

Operatör ekranında ikisi ayrı gösterilir: **İş:** … · **Excel:** …

---

## 6. Operatör işlemleri neyi değiştirir?

Tüm ilerleme **`data/kartlar.xlsx`**’e yazılır; ayrıca satır **`islem_logu.xlsx`**’e düşer.

### İşe Başla (`SIRADA` → `DEVAM EDİYOR`)

- Başlatılacak adet girilir (kalan adetten fazla olamaz).
- Kart devam durumuna geçer.
- Operatör adı ve başlama zamanı kaydolur.

### Adet Bitir

- Tamamlanan adet eklenir.
- Toplam adede ulaşılırsa → `TAMAMLANDI` + bitiş zamanı.
- Ulaşılmazsa → kart `DEVAM EDİYOR` kalır.

### Not

- Kart notu güncellenir; durum değişmez.

### Admin düzenle / gizle

- Admin iş durumunu ve adetleri elle düzeltebilir.
- **Gizle:** kart operatör/panodan düşer; sonraki Excel import’u gizlemeyi otomatik geri açmaz.

---

## 7. Excel yükleme — adım adım

Yönetim → **Haftalık Excel Yükle**.

### 7.1 Dosya kabulü

- Uzantı: `.xlsx` veya `.xlsm`
- Kopya saklanır: `data/yuklenen_exceller/…`
- Boyut üst sınırı: 25 MB

### 7.2 Parse + doğrulama (henüz state değişmez)

`excel_araclari.py` dosyayı tamamen okur. Hata varsa **hiçbir kart değişmez**.

Beklenen sütunlar (başlık satırı ilk ~10 satırda aranır):

| Sütun (örnek ad) | Zorunluluk |
|------------------|------------|
| Talep NO ve/veya Kart Stok No | En az biri |
| Kart Üretim Adet / Üretim Adet / Adet | Zorunlu |
| Sıra, Talep Sahibi, plan tarihleri, DURUM, PCB | Opsiyonel ama kullanılır |

Reddedilen örnekler:

- Aynı Excel içinde çift `Talep NO + Stok No`
- Okunamayan / &lt; 1 adet
- Plan başlangıç &gt; plan teslim
- Okunamayan tarih alanları (doluysa)

### 7.3 Kaynak DURUM → ilk iş durumu (yalnız **yeni** kartlar)

Normalize edilmiş exact eşleşme (substring yok):

| Excel DURUM (örnekler) | Yeni kartın ilk iş durumu |
|------------------------|---------------------------|
| `TESLIM`, `TESLIM EDILDI`, `TAMAMLANDI`, … | `TAMAMLANDI` |
| `DIZGI`, `DIZGIDE`, `DEVAM EDIYOR`, … | `DEVAM EDİYOR` |
| boş, `SIRADA`, `BEKLIYOR`, **`TESLIM EDILMEDI`** | `SIRADA` |
| Bilinmeyen metin | `SIRADA` + uyarı sayacı |

> `TESLIM EDILMEDI` bilinçli olarak **tamamlanmış sayılmaz**.

### 7.4 Anlık yedek

Import uygulanmadan hemen önce:

```text
data/yedekler/YYYYMMDD_HHMMSS_import_oncesi/
  ├── kartlar.xlsx
  ├── islem_logu.xlsx
  └── yuklemeler.xlsx
```

Sorun olursa bu klasörden geri kopyalanabilir.

### 7.5 Merge (asıl politika)

Her satır için:

#### A) Sistemde **yok** → yeni kart

- Excel durumundan ilk iş durumu atanır.
- Tamamlandıysa ve gerçekleşen teslim tarihi varsa bitiş zamanı yazılır.
- Plan tarihi, gerçekleşen teslim yoksa **uydurma bitiş saati olarak kullanılmaz**.

#### B) Sistemde **var** (aynı anahtar) → plan güncelle, workflow koru

| Excel’den gelen | Ne olur? |
|-----------------|----------|
| Plan alanları (adet metni, tarihler, talep sahibi, PCB, excel_durum metni, toplam adet*) | **Güncellenir** |
| İş durumu (`SIRADA` / `DEVAM` / `TAMAM`) | **Dokunulmaz** |
| `tamamlanan_adet`, operatör, başlama/bitiş, not | **Dokunulmaz** |
| Admin gizlemiş mi? | **Gizli kalır** |
| `source_active` | Yeniden **1** yapılır (kaynakta tekrar görüldü) |

\* Toplam adet güncellemesi kurallara bağlıdır; çelişirse **tüm import iptal** edilir (aşağıda).

#### C) Bu Excel’de **olmayan** eski Excel kaynaklı kartlar

- Silinmez.
- `source_active = 0` olur (“Son Excelde yok”).
- **Açık işler** (`SIRADA` / `DEVAM`) operatör ve panoda **görünmeye devam eder**.
- Tamamlanmışlar rapor/tarihçe için kalır.

### 7.6 Import’un tamamen reddedildiği durumlar

Tek satır bile kuralı bozarsa **hiçbir değişiklik uygulanmaz** (bellek geri alınır):

1. Kart sistemde `TAMAMLANDI` iken Excel toplam adedi farklı.
2. Kart `DEVAM` iken Excel toplamı ≤ tamamlanan adet.
3. Excel toplamı &lt; sistemdeki tamamlanan adet.
4. Parse/validate hataları (çift anahtar, bozuk tarih, vb.).

### 7.7 Başarılı yükleme sonrası mesaj

Örnek anlam:

- X satır okundu
- Y yeni kart
- Z kartın **planı** güncellendi
- İş durumu / tamamlanan adet **korundu**
- N açık kart son Excel’de yok
- Yedek klasör adı

Ayrıca `yuklemeler.xlsx` ve işlem loguna kayıt düşer.

---

## 8. “Excel ilerlemeyi ezer mi?”

**Hayır.** Mevcut kartlarda Excel:

- planı günceller,
- **operatör ilerlemesini ezmez.**

Yani operatör 50/200 bitirmişken yeni haftalık Excel gelse bile kart `DEVAM EDİYOR` ve `50` tamamlanmış kalır; yalnızca plan alanları yenilenir.

Excel’in iş durumunu belirlediği tek an: **kart ilk kez sisteme girdiğinde**.

---

## 9. Görünürlük kuralları

Kart panoda / operatörde görünmez eğer:

- `aktif != 1`, veya
- admin gizlemiş (`admin_gizli = 1`)

Açık kartlar son Excel’de olmasa bile **görünür** kalır (üretim kaybolmasın diye).  
Bunlar **“Son Excelde yok”** rozeti taşır; yönetim ekranında ayrıca listelenir.

---

## 10. Örnek senaryo

1. Admin `01_haftalik_plan.xlsx` yükler → 12 kart oluşur.
2. Operatör birkaç kartı başlatır, adet bitirir → `kartlar.xlsx` güncellenir.
3. Admin `02_haftalik_guncelleme.xlsx` yükler:
   - Ortak anahtarlı kartların **planı** yenilenir, **ilerleme korunur**.
   - Yeni satırlar eklenir.
   - Bu dosyada olmayan açık kartlar “Son Excelde yok” olur ama ekranda kalır.
4. Hatalı dosya (`03_gecersiz…` / `04_duplicate…`) yüklenirse → kırmızı hata, state aynı kalır.

Örnek dosyalar: `ornek_exceller/`

---

## 11. Yedekler ve loglar

| Konum | Ne zaman |
|-------|----------|
| `data/yedekler/YYYYMMDD_kartlar.xlsx` | Günlük ilk yazımda (eski davranış) |
| `data/yedekler/…_import_oncesi/` | Her Excel import öncesi anlık paket |
| `data/uygulama.log` | Uygulama logu (dönen dosya) |
| `data/yuklenen_exceller/` | Yüklenen ham Excel kopyaları |

Geri dönüş (manuel): ilgili yedek klasöründeki dosyaları `data/` altına kopyalayıp uygulamayı yeniden başlatın veya Yönetim → **Kart Dosyasını Yeniden Oku**.

---

## 12. Güvenlik (intranet varsayımı)

- Session: HttpOnly, SameSite=Lax, isteğe bağlı Secure
- CSRF: form + `fetch` (`X-CSRF-Token`)
- Parolalar: Werkzeug hash; kaynak kodda yok
- State değiştiren `/cikis` ve `/yonetim/yeniden-oku`: POST
- Login redirect: `//evil…` gibi açık yönlendirmelere kapalı
- Beklenmeyen import hataları kullanıcıya path dump etmez

Bu sistem **intranet / güvenilen ağ** içindir; internete açık bırakmak için ek sertleştirme gerekir.

---

## 13. Bilinçli sınırlar (SQL’siz)

1. Tek process + tek `data/` klasörü.
2. Excel transactional değildir; çökme anında teorik risk (temp + rollback ile azaltıldı).
3. Kart kimliği kaynak Excel anahtarına bağlıdır.
4. Uzun vadede doğru çözüm SQL’e geçiştir; depo API’si (`kart_baslat` / `kart_bitir` / `excel_import_uygula`) buna uygun ayrılmıştır.

---

## 14. Hızlı kontrol listesi (prod)

- [ ] Tek sunucuda çalışıyor (`sunucu.lock` uyarısı yok)
- [ ] `PDGM_PORT` / firewall uygun
- [ ] İlk admin parolası güvenli yerde
- [ ] Ortak `operator` hesabı production’da pasifleştirildi; kişi bazlı kullanıcılar açıldı
- [ ] Haftalık Excel şablonu sütun adları sistemle uyumlu
- [ ] `data/yedekler/` periyodik olarak başka diske kopyalanıyor
- [ ] HTTPS varsa `PDGM_HTTPS=1`

---

*Son güncelleme: Excel merge + workflow koruma + “Son Excelde yok” görünürlüğü + import öncesi anlık yedek + process kilidi + kullanıcı hot-reload sürümü.*
