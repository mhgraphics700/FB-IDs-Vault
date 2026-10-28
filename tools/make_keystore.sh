#!/usr/bin/env bash
# Creates the signing key ONCE. Keep the .jks file and passwords safe (a password manager / backup).
# If you lose this key you can never publish an update that installs over the old app.
set -e
ALIAS=fbidsvault
keytool -genkeypair -v -keystore fbidsvault-release.jks -alias "$ALIAS" -keyalg RSA -keysize 2048 -validity 12000
base64 -w0 fbidsvault-release.jks > keystore.b64 2>/dev/null || base64 fbidsvault-release.jks | tr -d '\n' > keystore.b64
echo
echo "Now add these 4 secrets in GitHub: repo > Settings > Secrets and variables > Actions > New repository secret"
echo "  KEYSTORE_BASE64   = contents of keystore.b64"
echo "  KEYSTORE_PASSWORD = the keystore password you just typed"
echo "  KEY_ALIAS         = $ALIAS"
echo "  KEY_PASSWORD      = the key password you just typed"
echo "Then delete keystore.b64 (keep fbidsvault-release.jks backed up, never commit it)."
