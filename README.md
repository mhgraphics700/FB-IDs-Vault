# FB IDs Vault

Android app (WebView) + Google Apps Script backend for keeping FB IDs in a Google Sheet and handing out a customer's quantity.

## Build the APK
1. Open this folder in Android Studio (File > Open), let Gradle sync.
2. Build > Build Bundle(s) / APK(s) > Build APK(s)   (or Generate Signed Bundle for Play Store).
   Package name: `com.fbidsvault.app`, version 2.0.

## Google Sheet backend
`backend/Code.gs` is the only script file you need. Paste it into Extensions > Apps Script, Deploy > Web app
(Execute as: Me, Who has access: Anyone), copy the `/exec` link into the app (Settings > Manage Google Sheets).
If you already use the previous script version 6, you do not have to redeploy for the app to keep working.

If you change `app/src/main/assets/index.html` and also use the web-page version, run `python3 tools/build_backend.py`.

## Icons
`store/make_icons.py` redraws the old-phone icons; `store/ic_launcher_512.png` is the Play Store icon.

## GitHub: publish and update
1. Create a **public** repository on github.com (for example `FB-IDs-Vault`) and upload this folder:
   `git init && git add . && git commit -m "FB IDs Vault" && git branch -M main && git remote add origin https://github.com/USERNAME/FB-IDs-Vault.git && git push -u origin main`
2. In the app: Settings > App updates > type `USERNAME/FB-IDs-Vault` (no code change needed).
3. Create the signing key once (`bash tools/make_keystore.sh`, or use the one delivered separately), then add the 4 repository secrets it prints
   (Settings > Secrets and variables > Actions): KEYSTORE_BASE64, KEYSTORE_PASSWORD, KEY_ALIAS, KEY_PASSWORD.
   Back up the `.jks` file; do not commit it.
4. Release a version: `git tag v2.1 && git push origin v2.1`. GitHub Actions builds the APK and puts it under Releases.
   The version number comes from the tag.
5. Installed apps check the latest Release (every 12 hours, or Settings > App updates > Check for updates) and show an
   **Update** button that downloads the APK. The first install must also use the same signed APK from Releases.

The app keeps loading everything in the background: after unlock it prepares Orders and Customers, and it syncs with the sheet every 90 seconds while it is open.
