"""PDGM kullanıcı dosyasına plaintext parola saklamadan kullanıcı ekleme/pasifleştirme aracı."""

from __future__ import annotations

import getpass
import json
import os
import sys

from werkzeug.security import generate_password_hash

KOK = os.path.dirname(os.path.abspath(__file__))
DOSYA = os.path.join(KOK, "data", "kullanicilar.json")
ROLLER = {"admin", "operator", "izleyici"}


def oku():
    if not os.path.exists(DOSYA):
        print("Önce uygulamayı bir kez çalıştırın; kullanicilar.json oluşturulsun.")
        raise SystemExit(1)
    with open(DOSYA, encoding="utf-8") as f:
        return json.load(f)


def yaz(veri):
    gecici = DOSYA + ".yeni"
    with open(gecici, "w", encoding="utf-8") as f:
        json.dump(veri, f, ensure_ascii=False, indent=2)
    os.replace(gecici, DOSYA)


def ekle(kullanici, rol, ad):
    if rol not in ROLLER:
        raise SystemExit(f"Rol {sorted(ROLLER)} değerlerinden biri olmalı.")
    veri = oku()
    if kullanici in veri:
        raise SystemExit("Bu kullanıcı adı zaten var.")
    p1 = getpass.getpass("Parola: ")
    p2 = getpass.getpass("Parola tekrar: ")
    if p1 != p2 or len(p1) < 8:
        raise SystemExit("Parolalar eşleşmeli ve en az 8 karakter olmalı.")
    veri[kullanici] = {
        "sifre_hash": generate_password_hash(p1),
        "rol": rol,
        "ad": ad,
        "aktif": True,
    }
    yaz(veri)
    print(f"Kullanıcı eklendi: {kullanici} ({rol})")


def aktiflik(kullanici, aktif):
    veri = oku()
    if kullanici not in veri:
        raise SystemExit("Kullanıcı bulunamadı.")
    veri[kullanici]["aktif"] = aktif
    yaz(veri)
    print(f"{kullanici}: {'aktif' if aktif else 'pasif'}")


def listele():
    for kullanici, bilgi in oku().items():
        print(f"{kullanici:20} {bilgi.get('rol','-'):10} {bilgi.get('ad','-')} aktif={bilgi.get('aktif', True)}")


def main():
    if len(sys.argv) < 2:
        raise SystemExit(
            "Kullanım:\n"
            "  python kullanici_yonet.py listele\n"
            "  python kullanici_yonet.py ekle <kullanici> <admin|operator|izleyici> <Gorunen Ad>\n"
            "  python kullanici_yonet.py pasif <kullanici>\n"
            "  python kullanici_yonet.py aktif <kullanici>"
        )
    komut = sys.argv[1].lower()
    if komut == "listele":
        listele()
    elif komut == "ekle" and len(sys.argv) >= 5:
        ekle(sys.argv[2], sys.argv[3], " ".join(sys.argv[4:]))
    elif komut == "pasif" and len(sys.argv) == 3:
        aktiflik(sys.argv[2], False)
    elif komut == "aktif" and len(sys.argv) == 3:
        aktiflik(sys.argv[2], True)
    else:
        raise SystemExit("Geçersiz komut veya eksik argüman.")


if __name__ == "__main__":
    main()
