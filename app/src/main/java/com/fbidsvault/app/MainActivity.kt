package com.fbidsvault.app

import android.annotation.SuppressLint
import android.app.DownloadManager
import android.os.Build
import android.os.Handler
import android.os.Looper
import android.provider.Settings
import androidx.core.content.FileProvider
import java.io.File
import android.content.ContentValues
import android.content.Intent
import android.content.res.Configuration
import android.net.Uri
import android.os.Bundle
import android.os.Environment
import android.provider.MediaStore
import android.view.WindowManager
import android.webkit.JavascriptInterface
import android.webkit.JsResult
import android.webkit.RenderProcessGoneDetail
import android.webkit.ValueCallback
import android.webkit.WebChromeClient
import android.webkit.WebResourceRequest
import android.webkit.WebResourceResponse
import android.webkit.WebSettings
import android.webkit.WebView
import android.webkit.WebViewClient
import android.widget.Toast
import androidx.activity.OnBackPressedCallback
import androidx.activity.result.contract.ActivityResultContracts
import androidx.appcompat.app.AppCompatActivity
import androidx.biometric.BiometricManager
import androidx.biometric.BiometricPrompt
import androidx.core.content.ContextCompat
import androidx.core.view.ViewCompat
import androidx.core.view.WindowCompat
import androidx.core.view.WindowInsetsCompat
import androidx.webkit.WebViewAssetLoader
import com.google.android.material.dialog.MaterialAlertDialogBuilder

class MainActivity : AppCompatActivity() {

    // The page is bundled inside the APK (app/src/main/assets/index.html) and served from a private
    // https address, so it opens without internet and does not need file:// access.
    private val host = "appassets.androidplatform.net"
    private val appUrl = "https://$host/assets/index.html"

    private lateinit var web: WebView
    private lateinit var loader: WebViewAssetLoader
    private var fileCallback: ValueCallback<Array<Uri>>? = null
    private var lastBack = 0L

    // ---- background update download (APK from GitHub Releases) ----
    private var dlId = -1L
    private val ui = Handler(Looper.getMainLooper())
    private fun apkFile() = File(getExternalFilesDir(null), "update.apk")
    private fun js(code: String) { runOnUiThread { web.evaluateJavascript(code, null) } }

    private fun finishDownload() {
        dlId = -1L
        val f = apkFile()
        val info = packageManager.getPackageArchiveInfo(f.path, 0)
        if (info == null || info.packageName != packageName) {
            f.delete()
            js("window.onUpdFail&&window.onUpdFail()")
        } else js("window.onUpdReady&&window.onUpdReady()")
    }

    // polls the system download every 0.7 s and tells the page the percentage
    private val poll = object : Runnable {
        override fun run() {
            if (dlId == -1L) return
            val dm = getSystemService(DOWNLOAD_SERVICE) as DownloadManager
            val c = dm.query(DownloadManager.Query().setFilterById(dlId))
            var again = true
            if (c != null) {
                if (c.moveToFirst()) {
                    val st = c.getInt(c.getColumnIndexOrThrow(DownloadManager.COLUMN_STATUS))
                    val got = c.getLong(c.getColumnIndexOrThrow(DownloadManager.COLUMN_BYTES_DOWNLOADED_SO_FAR))
                    val tot = c.getLong(c.getColumnIndexOrThrow(DownloadManager.COLUMN_TOTAL_SIZE_BYTES))
                    if (st == DownloadManager.STATUS_SUCCESSFUL) { again = false; finishDownload() }
                    else if (st == DownloadManager.STATUS_FAILED) { again = false; dlId = -1L; js("window.onUpdFail&&window.onUpdFail()") }
                    else { val p = if (tot > 0) (got * 100 / tot).toInt() else 0; js("window.onUpdProgress&&window.onUpdProgress($p)") }
                } else { again = false; dlId = -1L; js("window.onUpdFail&&window.onUpdFail()") }
                c.close()
            }
            if (again) ui.postDelayed(this, 700)
        }
    }

    private val picker = registerForActivityResult(ActivityResultContracts.GetContent()) { uri ->
        fileCallback?.onReceiveValue(if (uri != null) arrayOf(uri) else null)
        fileCallback = null
    }

    @SuppressLint("SetJavaScriptEnabled")
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)

        applySecure(prefs().getBoolean("secure", true))

        loader = WebViewAssetLoader.Builder()
            .setDomain(host)
            .addPathHandler("/assets/", WebViewAssetLoader.AssetsPathHandler(this))
            .build()

        web = WebView(this)
        setContentView(web)
        web.setBackgroundColor(android.graphics.Color.TRANSPARENT)

        // status bar and navigation bar icons follow the light / dark theme
        val night = (resources.configuration.uiMode and Configuration.UI_MODE_NIGHT_MASK) == Configuration.UI_MODE_NIGHT_YES
        WindowCompat.getInsetsController(window, web).apply {
            isAppearanceLightStatusBars = !night
            isAppearanceLightNavigationBars = !night
        }

        // keep the page clear of the status bar and navigation bar
        ViewCompat.setOnApplyWindowInsetsListener(web) { v, insets ->
            val b = insets.getInsets(WindowInsetsCompat.Type.systemBars())
            v.setPadding(b.left, b.top, b.right, b.bottom)
            insets
        }

        web.settings.apply {
            javaScriptEnabled = true
            domStorageEnabled = true      // saved PIN, connections and offline data (localStorage)
            allowFileAccess = false
            allowContentAccess = false
            mixedContentMode = WebSettings.MIXED_CONTENT_NEVER_ALLOW
            setSupportZoom(false)
        }

        web.webViewClient = object : WebViewClient() {
            override fun shouldInterceptRequest(view: WebView, request: WebResourceRequest): WebResourceResponse? =
                loader.shouldInterceptRequest(request.url)

            override fun shouldOverrideUrlLoading(view: WebView, request: WebResourceRequest): Boolean {
                if (request.url.host == host) return false
                try { startActivity(Intent(Intent.ACTION_VIEW, request.url)) } catch (_: Exception) {}
                return true
            }

            override fun onRenderProcessGone(view: WebView, detail: RenderProcessGoneDetail): Boolean {
                recreate()   // the page crashed or was killed to free memory: start fresh instead of closing
                return true
            }
        }

        web.webChromeClient = object : WebChromeClient() {
            // themed dialogs instead of the plain browser ones (Undo, Replace, Logout, ...)
            override fun onJsConfirm(view: WebView, url: String?, message: String?, result: JsResult): Boolean {
                MaterialAlertDialogBuilder(this@MainActivity)
                    .setTitle(R.string.app_name)
                    .setMessage(message)
                    .setPositiveButton("OK") { _, _ -> result.confirm() }
                    .setNegativeButton("Cancel") { _, _ -> result.cancel() }
                    .setOnCancelListener { result.cancel() }
                    .show()
                return true
            }

            override fun onJsAlert(view: WebView, url: String?, message: String?, result: JsResult): Boolean {
                MaterialAlertDialogBuilder(this@MainActivity)
                    .setTitle(R.string.app_name)
                    .setMessage(message)
                    .setPositiveButton("OK") { _, _ -> result.confirm() }
                    .setOnCancelListener { result.confirm() }
                    .show()
                return true
            }

            // lets the page open the file picker (backup restore)
            override fun onShowFileChooser(
                view: WebView, callback: ValueCallback<Array<Uri>>, params: FileChooserParams
            ): Boolean {
                fileCallback?.onReceiveValue(null)
                fileCallback = callback
                try { picker.launch("*/*") } catch (e: Exception) { fileCallback = null; callback.onReceiveValue(null); return false }
                return true
            }
        }
        web.addJavascriptInterface(Bridge(), "Android")

        if (savedInstanceState == null) web.loadUrl(appUrl) else web.restoreState(savedInstanceState)

        onBackPressedDispatcher.addCallback(this, object : OnBackPressedCallback(true) {
            override fun handleOnBackPressed() {
                // the page closes a popup, the menu, or goes back to the Sell tab; otherwise press twice to exit
                web.evaluateJavascript("(function(){try{return window.fbBack?window.fbBack():false}catch(e){return false}})()") { r ->
                    if (r == "true") return@evaluateJavascript
                    val now = System.currentTimeMillis()
                    if (now - lastBack < 2000) finish()
                    else {
                        lastBack = now
                        Toast.makeText(this@MainActivity, "Press back again to exit", Toast.LENGTH_SHORT).show()
                    }
                }
            }
        })
    }

    override fun onSaveInstanceState(outState: Bundle) {
        super.onSaveInstanceState(outState)
        web.saveState(outState)
    }

    override fun onDestroy() {
        ui.removeCallbacksAndMessages(null)
        fileCallback?.onReceiveValue(null)
        fileCallback = null
        super.onDestroy()
    }

    private fun prefs() = getSharedPreferences("fbidsvault", MODE_PRIVATE)

    // Blocks screenshots and the recent-apps preview (the IDs are valuable)
    private fun applySecure(on: Boolean) {
        if (on) window.setFlags(WindowManager.LayoutParams.FLAG_SECURE, WindowManager.LayoutParams.FLAG_SECURE)
        else window.clearFlags(WindowManager.LayoutParams.FLAG_SECURE)
    }

    // Called from the page: Android.saveFile(name, text, mime) and friends
    inner class Bridge {
        @JavascriptInterface
        fun saveFile(name: String, text: String, mime: String) {
            try {
                val values = ContentValues().apply {
                    put(MediaStore.MediaColumns.DISPLAY_NAME, name)
                    put(MediaStore.MediaColumns.MIME_TYPE, mime)
                    put(MediaStore.MediaColumns.RELATIVE_PATH, Environment.DIRECTORY_DOWNLOADS)
                }
                val uri = contentResolver.insert(MediaStore.Downloads.EXTERNAL_CONTENT_URI, values)!!
                contentResolver.openOutputStream(uri)!!.use { it.write(text.toByteArray(Charsets.UTF_8)) }
                runOnUiThread {
                    Toast.makeText(this@MainActivity, "Saved to Downloads: $name", Toast.LENGTH_LONG).show()
                }
            } catch (e: Exception) {
                runOnUiThread {
                    Toast.makeText(this@MainActivity, "Save failed", Toast.LENGTH_LONG).show()
                }
            }
        }

        @JavascriptInterface
        fun shareText(text: String) {
            runOnUiThread {
                val i = Intent(Intent.ACTION_SEND).apply {
                    type = "text/plain"
                    putExtra(Intent.EXTRA_TEXT, text)
                }
                startActivity(Intent.createChooser(i, "Share"))
            }
        }

        @JavascriptInterface
        fun appVersion(): String = try {
            packageManager.getPackageInfo(packageName, 0).versionName ?: ""
        } catch (e: Exception) { "" }

        // Only GitHub links may be opened from the page (used by the update check)
        @JavascriptInterface
        fun openUrl(url: String) {
            if (!url.startsWith("https://github.com/")) return
            runOnUiThread {
                try { startActivity(Intent(Intent.ACTION_VIEW, Uri.parse(url))) } catch (_: Exception) {}
            }
        }

        // Starts downloading the new APK in the background (GitHub links only)
        @JavascriptInterface
        fun downloadUpdate(url: String) {
            if (!url.startsWith("https://github.com/") || !url.endsWith(".apk", true)) return
            runOnUiThread {
                if (dlId != -1L) return@runOnUiThread
                try {
                    val f = apkFile()
                    if (f.exists()) f.delete()
                    val req = DownloadManager.Request(Uri.parse(url))
                        .setTitle("FB IDs Vault update")
                        .setMimeType("application/vnd.android.package-archive")
                        .setNotificationVisibility(DownloadManager.Request.VISIBILITY_VISIBLE)
                        .setDestinationInExternalFilesDir(this@MainActivity, null, "update.apk")
                    val dm = getSystemService(DOWNLOAD_SERVICE) as DownloadManager
                    dlId = dm.enqueue(req)
                    ui.post(poll)
                } catch (e: Exception) {
                    dlId = -1L
                    js("window.onUpdFail&&window.onUpdFail()")
                }
            }
        }

        // true when a newer APK is already downloaded and waiting to be installed
        @JavascriptInterface
        fun updateReady(): Boolean = try {
            val f = apkFile()
            if (dlId != -1L || !f.exists()) false
            else {
                val info = packageManager.getPackageArchiveInfo(f.path, 0)
                val cur = packageManager.getPackageInfo(packageName, 0)
                if (info != null && info.packageName == packageName && info.versionName != cur.versionName) true
                else { f.delete(); false }
            }
        } catch (e: Exception) { false }

        @JavascriptInterface
        fun installUpdate() {
            runOnUiThread {
                val f = apkFile()
                if (!f.exists()) return@runOnUiThread
                if (Build.VERSION.SDK_INT >= 26 && !packageManager.canRequestPackageInstalls()) {
                    Toast.makeText(this@MainActivity, "Allow installs from this app, then tap Install again", Toast.LENGTH_LONG).show()
                    try { startActivity(Intent(Settings.ACTION_MANAGE_UNKNOWN_APP_SOURCES, Uri.parse("package:$packageName"))) } catch (_: Exception) {}
                    return@runOnUiThread
                }
                try {
                    val uri = FileProvider.getUriForFile(this@MainActivity, "$packageName.fileprovider", f)
                    val i = Intent(Intent.ACTION_VIEW).apply {
                        setDataAndType(uri, "application/vnd.android.package-archive")
                        addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION or Intent.FLAG_ACTIVITY_NEW_TASK)
                    }
                    startActivity(i)
                } catch (e: Exception) {
                    Toast.makeText(this@MainActivity, "Could not open the installer", Toast.LENGTH_LONG).show()
                }
            }
        }

        @JavascriptInterface
        fun isSecure(): Boolean = prefs().getBoolean("secure", true)

        @JavascriptInterface
        fun setSecure(on: Boolean) {
            prefs().edit().putBoolean("secure", on).apply()
            runOnUiThread { applySecure(on) }
        }

        @JavascriptInterface
        fun canBio(): Boolean =
            BiometricManager.from(this@MainActivity)
                .canAuthenticate(BiometricManager.Authenticators.BIOMETRIC_WEAK) == BiometricManager.BIOMETRIC_SUCCESS

        @JavascriptInterface
        fun bio() {
            runOnUiThread {
                val prompt = BiometricPrompt(
                    this@MainActivity,
                    ContextCompat.getMainExecutor(this@MainActivity),
                    object : BiometricPrompt.AuthenticationCallback() {
                        override fun onAuthenticationSucceeded(result: BiometricPrompt.AuthenticationResult) {
                            web.evaluateJavascript("window.onBio && window.onBio(true)", null)
                        }
                        override fun onAuthenticationError(errorCode: Int, errString: CharSequence) {
                            web.evaluateJavascript("window.onBio && window.onBio(false)", null)
                        }
                    }
                )
                val info = BiometricPrompt.PromptInfo.Builder()
                    .setTitle("FB IDs Vault")
                    .setSubtitle("Unlock with fingerprint")
                    .setNegativeButtonText("Use PIN")
                    .setAllowedAuthenticators(BiometricManager.Authenticators.BIOMETRIC_WEAK)
                    .build()
                prompt.authenticate(info)
            }
        }
    }
}
