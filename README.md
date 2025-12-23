# O365-AutoReply-Manager (v1.0) by NetronIC
GUI tool for managing Exchange Online Out-of-Office settings

Az **O365 AutoReply Manager** egy PowerShell-alapú GUI alkalmazás, amely megkönnyíti a Microsoft 365 adminisztrátorok számára a kimenő automatikus válaszok (OOF) kezelését. Nincs szükség parancssorokra; kezelje szervezete postaládáit egy modern, intuitív felületen.

## 🚀 Főbb funkciók
- **Gyors keresés:** Több száz felhasználó közötti hatékony szűrés.
- **Részletes üzenetkezelés:** Külön belső és külső válaszüzenetek szerkesztése.
- **Batch Mode:** Akár több száz felhasználó válaszüzenetének egyidejű beállítása.
- **CSV Import/Export:** Teljes szervezet szintű mentés vagy tömeges beállítás fájlból.
- **Modern hitelesítés:** Támogatja az MFA (többfaktoros) és a Device Code bejelentkezést.

## 🛠️ Telepítés
1. Töltse le a legfrissebb `O365Manager.exe` vagy script csomagot.
2. Győződjön meg róla, hogy az **ExchangeOnlineManagement v3.0+** modul telepítve van:
powershell Install-Module -Name ExchangeOnlineManagement -Force
