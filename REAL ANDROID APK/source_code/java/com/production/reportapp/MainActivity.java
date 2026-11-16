package com.production.reportapp;

import android.annotation.SuppressLint;
import android.app.AlertDialog;
import android.app.DownloadManager;
import android.content.BroadcastReceiver;
import android.content.Context;
import android.content.Intent;
import android.content.IntentFilter;
import android.content.SharedPreferences;
import android.graphics.Bitmap;
import android.net.Uri;
import android.os.AsyncTask;
import android.os.Build;
import android.os.Bundle;
import android.os.Environment;
import android.view.LayoutInflater;
import android.view.Menu;
import android.view.MenuItem;
import android.view.View;
import android.webkit.CookieManager;
import android.webkit.URLUtil;
import android.webkit.ValueCallback;
import android.webkit.WebChromeClient;
import android.webkit.WebResourceError;
import android.webkit.WebResourceRequest;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.webkit.WebViewClient;
import android.widget.Button;
import android.widget.EditText;
import android.widget.ProgressBar;
import android.widget.TextView;
import android.widget.Toast;

import androidx.annotation.Nullable;
import androidx.appcompat.app.AppCompatActivity;
import androidx.core.content.FileProvider;

import org.json.JSONObject;

import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.InetSocketAddress;
import java.net.Socket;
import java.net.URL;

public class MainActivity extends AppCompatActivity {

    private static final String PREFS_NAME = "ReportAppPrefs";
    private static final String KEY_SERVER_URL = "server_url";
    private static final String DEFAULT_HOST = "192.168.99.116";
    private static final String DEFAULT_PORT = "5005";
    private static final int FILE_CHOOSER_REQUEST_CODE = 1001;

    // Public GitHub Cloud Registry endpoints for zero-friction auto-discovery
    private static final String CLOUD_DISCOVERY_PRIMARY = "https://raw.githubusercontent.com/abellrajuu/IPA-PRODUCTION-TEST-REPORT/gh-pages/live_url.json";
    private static final String CLOUD_DISCOVERY_SECONDARY = "https://abellrajuu.github.io/IPA-PRODUCTION-TEST-REPORT/live_url.json";

    private WebView webView;
    private ProgressBar progressBar;
    private View errorLayout;
    private TextView errorTextView;
    private SharedPreferences prefs;
    private ValueCallback<Uri[]> uploadMessage;
    private boolean hasCheckedUpdateThisSession = false;
    private boolean isPageError = false;

    @SuppressLint("SetJavaScriptEnabled")
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        prefs = getSharedPreferences(PREFS_NAME, Context.MODE_PRIVATE);

        webView = findViewById(R.id.webView);
        progressBar = findViewById(R.id.progressBar);
        errorLayout = findViewById(R.id.errorLayout);
        errorTextView = findViewById(R.id.errorTextView);

        Button btnSettings = findViewById(R.id.btnChangeServer);
        if (btnSettings != null) {
            btnSettings.setOnClickListener(v -> showServerConfigDialog());
        }

        Button btnRetry = findViewById(R.id.btnRetry);
        if (btnRetry != null) {
            btnRetry.setText("Connect / Retry");
            btnRetry.setOnClickListener(v -> connectToProductionServer(true));
        }

        configureWebView();

        // Connect immediately on startup
        connectToProductionServer(false);
    }

    @Override
    protected void onResume() {
        super.onResume();
        if (errorLayout != null && errorLayout.getVisibility() == View.VISIBLE) {
            connectToProductionServer(false);
        }
    }

    @Override
    protected void onPause() {
        super.onPause();
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.LOLLIPOP) {
            CookieManager.getInstance().flush();
        }
    }

    @SuppressLint("SetJavaScriptEnabled")
    private void configureWebView() {
        WebSettings settings = webView.getSettings();
        settings.setJavaScriptEnabled(true);
        settings.setDomStorageEnabled(true);
        settings.setDatabaseEnabled(true);
        settings.setAllowFileAccess(true);
        settings.setAllowContentAccess(true);
        settings.setUseWideViewPort(true);
        settings.setLoadWithOverviewMode(true);
        settings.setBuiltInZoomControls(true);
        settings.setDisplayZoomControls(false);
        settings.setSupportZoom(true);

        String baseUa = settings.getUserAgentString();
        settings.setUserAgentString(baseUa + " ProductionReportApp/3.2.0 (Android; NativeApp; OnlineMobileData)");

        CookieManager cookieManager = CookieManager.getInstance();
        cookieManager.setAcceptCookie(true);
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.LOLLIPOP) {
            cookieManager.setAcceptThirdPartyCookies(webView, true);
        }

        webView.setWebViewClient(new WebViewClient() {
            @Override
            public void onPageStarted(WebView view, String url, Bitmap favicon) {
                if (!"about:blank".equals(url)) {
                    isPageError = false;
                    progressBar.setVisibility(View.VISIBLE);
                }
            }

            @Override
            public void onPageFinished(WebView view, String url) {
                progressBar.setVisibility(View.GONE);
                if (isPageError || "about:blank".equals(url)) {
                    webView.setVisibility(View.GONE);
                    errorLayout.setVisibility(View.VISIBLE);
                    return;
                }
                errorLayout.setVisibility(View.GONE);
                webView.setVisibility(View.VISIBLE);
                if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.LOLLIPOP) {
                    CookieManager.getInstance().flush();
                }
                checkForAppUpdate(false);
            }

            @Override
            public void onReceivedError(WebView view, WebResourceRequest request, WebResourceError error) {
                if (request.isForMainFrame()) {
                    isPageError = true;
                    progressBar.setVisibility(View.GONE);
                    view.loadUrl("about:blank");
                    webView.setVisibility(View.GONE);
                    errorLayout.setVisibility(View.VISIBLE);
                    errorTextView.setText("Could not reach Production System.\n\nPlease check your mobile data (4G/5G) or Wi-Fi, then tap 'Connect / Retry'.");
                }
            }

            @Override
            public void onReceivedError(WebView view, int errorCode, String description, String failingUrl) {
                isPageError = true;
                progressBar.setVisibility(View.GONE);
                view.loadUrl("about:blank");
                webView.setVisibility(View.GONE);
                errorLayout.setVisibility(View.VISIBLE);
                errorTextView.setText("Could not reach Production System.\n\nPlease check your mobile data (4G/5G) or Wi-Fi, then tap 'Connect / Retry'.");
            }
        });

        webView.setWebChromeClient(new WebChromeClient() {
            @Override
            public void onProgressChanged(WebView view, int newProgress) {
                if (newProgress < 100) {
                    progressBar.setVisibility(View.VISIBLE);
                    progressBar.setProgress(newProgress);
                } else {
                    progressBar.setVisibility(View.GONE);
                }
            }

            @Override
            public boolean onShowFileChooser(WebView webView, ValueCallback<Uri[]> filePathCallback, FileChooserParams fileChooserParams) {
                if (uploadMessage != null) {
                    uploadMessage.onReceiveValue(null);
                }
                uploadMessage = filePathCallback;

                Intent intent = new Intent(Intent.ACTION_GET_CONTENT);
                intent.addCategory(Intent.CATEGORY_OPENABLE);
                intent.setType("*/*");
                startActivityForResult(Intent.createChooser(intent, "Select Job / CAD File"), FILE_CHOOSER_REQUEST_CODE);
                return true;
            }
        });

        webView.setDownloadListener((url, userAgent, contentDisposition, mimetype, contentLength) -> {
            try {
                String filename = URLUtil.guessFileName(url, contentDisposition, mimetype);
                DownloadManager.Request request = new DownloadManager.Request(Uri.parse(url));
                request.setMimeType(mimetype);
                String cookies = CookieManager.getInstance().getCookie(url);
                request.addRequestHeader("cookie", cookies);
                request.addRequestHeader("User-Agent", userAgent);
                request.setDescription("Downloading report from Production System...");
                request.setTitle(filename);
                request.setNotificationVisibility(DownloadManager.Request.VISIBILITY_VISIBLE_NOTIFY_COMPLETED);
                request.setDestinationInExternalPublicDir(Environment.DIRECTORY_DOWNLOADS, filename);

                DownloadManager dm = (DownloadManager) getSystemService(Context.DOWNLOAD_SERVICE);
                if (dm != null) {
                    dm.enqueue(request);
                    Toast.makeText(MainActivity.this, "Downloading " + filename + " to Downloads folder...", Toast.LENGTH_SHORT).show();
                }
            } catch (Exception e) {
                Toast.makeText(MainActivity.this, "Download failed: " + e.getMessage(), Toast.LENGTH_LONG).show();
            }
        });

        // Direct Javascript Interface for seamless 1-click & auto downloads and updates
        webView.addJavascriptInterface(new Object() {
            @android.webkit.JavascriptInterface
            public void downloadFile(String url, String filename, String mimetype) {
                runOnUiThread(() -> {
                    try {
                        String downloadUrl = url;
                        if (!downloadUrl.startsWith("http://") && !downloadUrl.startsWith("https://")) {
                            String serverUrl = getServerUrl();
                            if (serverUrl != null && !serverUrl.isEmpty()) {
                                Uri sUri = Uri.parse(serverUrl);
                                downloadUrl = sUri.getScheme() + "://" + sUri.getAuthority() + (url.startsWith("/") ? "" : "/") + url;
                            }
                        }
                        String finalName = (filename != null && !filename.isEmpty()) ? filename : "Report.pdf";
                        DownloadManager.Request request = new DownloadManager.Request(Uri.parse(downloadUrl));
                        if (mimetype != null && !mimetype.isEmpty()) {
                            request.setMimeType(mimetype);
                        }
                        String cookies = CookieManager.getInstance().getCookie(downloadUrl);
                        if (cookies != null && !cookies.isEmpty()) {
                            request.addRequestHeader("cookie", cookies);
                        }
                        request.addRequestHeader("User-Agent", webView.getSettings().getUserAgentString());
                        request.setDescription("Production Test Report download");
                        request.setTitle(finalName);
                        request.setNotificationVisibility(DownloadManager.Request.VISIBILITY_VISIBLE_NOTIFY_COMPLETED);
                        request.setDestinationInExternalPublicDir(Environment.DIRECTORY_DOWNLOADS, finalName);

                        DownloadManager dm = (DownloadManager) getSystemService(Context.DOWNLOAD_SERVICE);
                        if (dm != null) {
                            dm.enqueue(request);
                            Toast.makeText(MainActivity.this, "Downloading " + finalName + " to phone Downloads...", Toast.LENGTH_SHORT).show();
                        }
                    } catch (Exception e) {
                        Toast.makeText(MainActivity.this, "Download error: " + e.getMessage(), Toast.LENGTH_LONG).show();
                    }
                });
            }

            @android.webkit.JavascriptInterface
            public void checkForUpdates() {
                runOnUiThread(() -> checkForAppUpdate(true));
            }
        }, "AndroidNative");
    }

    private String getServerUrl() {
        return prefs.getString(KEY_SERVER_URL, "");
    }

    private void saveServerUrl(String url) {
        prefs.edit().putString(KEY_SERVER_URL, url).apply();
    }

    private void loadUrlDirectly(String url) {
        if (!url.contains("is_app=1")) {
            url += (url.contains("?") ? "&" : "?") + "is_app=1";
        }
        isPageError = false;
        progressBar.setVisibility(View.VISIBLE);
        errorLayout.setVisibility(View.GONE);
        webView.setVisibility(View.VISIBLE);
        webView.loadUrl(url);
    }

    // --- Smart Multi-Path Connection Logic (Optimized for Mobile Data) ---
    public void connectToProductionServer(final boolean manual) {
        progressBar.setVisibility(View.VISIBLE);
        if (manual) {
            Toast.makeText(this, "Connecting to Production Server...", Toast.LENGTH_SHORT).show();
        }

        new Thread(() -> {
            // Path 1: USB Direct (if connected via USB to laptop)
            if (isPortOpen("127.0.0.1", 5005, 200)) {
                String localUsb = "http://127.0.0.1:5005/portal?is_app=1";
                runOnUiThread(() -> {
                    saveServerUrl(localUsb);
                    loadUrlDirectly(localUsb);
                    Toast.makeText(MainActivity.this, "Connected via USB", Toast.LENGTH_SHORT).show();
                });
                return;
            }

            // Path 2: Office Wi-Fi (if connected to office Wi-Fi network)
            if (isPortOpen("192.168.99.116", 5005, 300)) {
                String localWifi = "http://192.168.99.116:5005/portal?is_app=1";
                runOnUiThread(() -> {
                    saveServerUrl(localWifi);
                    loadUrlDirectly(localWifi);
                    Toast.makeText(MainActivity.this, "Connected via Office Wi-Fi", Toast.LENGTH_SHORT).show();
                });
                return;
            }

            // Path 3: Live Cloud Discovery via GitHub (Fast Mobile Data Path)
            String cloudUrl = fetchCloudServerUrl();
            if (cloudUrl != null && !cloudUrl.isEmpty()) {
                runOnUiThread(() -> {
                    saveServerUrl(cloudUrl);
                    loadUrlDirectly(cloudUrl);
                    Toast.makeText(MainActivity.this, "Connected Online via Mobile Data", Toast.LENGTH_SHORT).show();
                });
                return;
            }

            // Path 4: Previously saved cloud URL (fallback)
            String savedUrl = getServerUrl();
            if (savedUrl != null && !savedUrl.isEmpty() && !savedUrl.contains("127.0.0.1") && !savedUrl.contains("192.168.")) {
                runOnUiThread(() -> {
                    loadUrlDirectly(savedUrl);
                    Toast.makeText(MainActivity.this, "Connecting to saved link...", Toast.LENGTH_SHORT).show();
                });
                return;
            }

            // If ALL fail: DO NOT load failed URL into WebView (prevents Chromium error screen)
            runOnUiThread(() -> {
                progressBar.setVisibility(View.GONE);
                webView.setVisibility(View.GONE);
                errorLayout.setVisibility(View.VISIBLE);
                errorTextView.setText("Could not reach Production System.\n\nPlease check your mobile data (4G/5G) connection, then tap 'Connect / Retry'.");
                if (manual) {
                    Toast.makeText(MainActivity.this, "No internet or server connection detected.", Toast.LENGTH_LONG).show();
                }
            });
        }).start();
    }

    private String fetchCloudServerUrl() {
        String[] endpoints = new String[]{
                CLOUD_DISCOVERY_PRIMARY + "?t=" + System.currentTimeMillis(),
                CLOUD_DISCOVERY_SECONDARY + "?t=" + System.currentTimeMillis()
        };
        for (String endpoint : endpoints) {
            HttpURLConnection conn = null;
            try {
                URL u = new URL(endpoint);
                conn = (HttpURLConnection) u.openConnection();
                conn.setConnectTimeout(4000);
                conn.setReadTimeout(4000);
                conn.setRequestMethod("GET");
                conn.setRequestProperty("User-Agent", "ProductionReportApp Native");
                if (conn.getResponseCode() == 200) {
                    BufferedReader reader = new BufferedReader(new InputStreamReader(conn.getInputStream()));
                    StringBuilder sb = new StringBuilder();
                    String line;
                    while ((line = reader.readLine()) != null) {
                        sb.append(line);
                    }
                    reader.close();
                    JSONObject json = new JSONObject(sb.toString());
                    String url = json.optString("url", "").trim();
                    if (!url.isEmpty()) {
                        return url;
                    }
                }
            } catch (Exception ignored) {
            } finally {
                if (conn != null) conn.disconnect();
            }
        }
        return null;
    }

    public void showServerConfigDialog() {
        AlertDialog.Builder builder = new AlertDialog.Builder(this);
        View dialogView = LayoutInflater.from(this).inflate(R.layout.dialog_server_config, null);
        builder.setView(dialogView);

        EditText etHost = dialogView.findViewById(R.id.etHost);
        EditText etPort = dialogView.findViewById(R.id.etPort);
        Button btnMdns = dialogView.findViewById(R.id.btnUseMdns);
        Button btnAutoScan = dialogView.findViewById(R.id.btnAutoScan);

        String curUrl = getServerUrl();
        try {
            Uri uri = Uri.parse(curUrl);
            if (uri.getHost() != null) etHost.setText(uri.getHost());
            if (uri.getPort() != -1) etPort.setText(String.valueOf(uri.getPort()));
            else etPort.setText(DEFAULT_PORT);
        } catch (Exception e) {
            etHost.setText(DEFAULT_HOST);
            etPort.setText(DEFAULT_PORT);
        }

        AlertDialog dialog = builder.create();

        btnMdns.setOnClickListener(v -> {
            etHost.setText(DEFAULT_HOST);
            etPort.setText(DEFAULT_PORT);
        });

        btnAutoScan.setOnClickListener(v -> {
            btnAutoScan.setEnabled(false);
            btnAutoScan.setText("Scanning Wi-Fi...");
            new SubnetScannerTask(etHost, btnAutoScan).execute();
        });

        dialogView.findViewById(R.id.btnSaveConnect).setOnClickListener(v -> {
            String host = etHost.getText().toString().trim();
            String port = etPort.getText().toString().trim();

            if (host.isEmpty()) {
                etHost.setError("Please enter Hostname or IP");
                return;
            }

            String newUrl;
            if (host.startsWith("http://") || host.startsWith("https://")) {
                newUrl = host.endsWith("/") ? host + "portal" : (host.contains("/portal") ? host : host + "/portal");
            } else if (port.equals("443")) {
                newUrl = "https://" + host + "/portal";
            } else {
                if (port.isEmpty()) port = DEFAULT_PORT;
                newUrl = "http://" + host + ":" + port + "/portal";
            }
            saveServerUrl(newUrl);
            dialog.dismiss();
            hasCheckedUpdateThisSession = false;
            loadUrlDirectly(newUrl);
        });

        dialog.show();
    }

    public void checkForAppUpdate(final boolean manual) {
        if (!manual && hasCheckedUpdateThisSession) return;
        hasCheckedUpdateThisSession = true;

        new Thread(() -> {
            HttpURLConnection conn = null;
            try {
                String serverUrl = getServerUrl();
                if (serverUrl == null || serverUrl.isEmpty()) return;
                Uri sUri = Uri.parse(serverUrl);
                String versionApiUrl = sUri.getScheme() + "://" + sUri.getAuthority() + "/api/version";

                URL u = new URL(versionApiUrl);
                conn = (HttpURLConnection) u.openConnection();
                conn.setConnectTimeout(4000);
                conn.setReadTimeout(4000);
                conn.setRequestMethod("GET");
                conn.setRequestProperty("User-Agent", "ProductionReportApp Native");

                final int responseCode = conn.getResponseCode();
                if (responseCode == 200) {
                    BufferedReader reader = new BufferedReader(new InputStreamReader(conn.getInputStream()));
                    StringBuilder sb = new StringBuilder();
                    String line;
                    while ((line = reader.readLine()) != null) {
                        sb.append(line);
                    }
                    reader.close();

                    JSONObject json = new JSONObject(sb.toString());
                    int serverVersionCode = json.optInt("version_code", 0);
                    String serverVersion = json.optString("version", "Latest");
                    String downloadUrl = json.optString("download_url", "/download-app");
                    String releaseNotes = json.optString("release_notes", "Bug fixes and performance improvements.");

                    int myVersionCode = 1;
                    String myVersionName = "3.2.0";
                    try {
                        android.content.pm.PackageInfo pInfo = getPackageManager().getPackageInfo(getPackageName(), 0);
                        myVersionCode = pInfo.versionCode;
                        myVersionName = pInfo.versionName;
                    } catch (Exception ignored) {}

                    if (serverVersionCode > myVersionCode) {
                        runOnUiThread(() -> showUpdateDialog(serverVersion, downloadUrl, releaseNotes));
                    } else if (manual) {
                        final String vName = myVersionName;
                        runOnUiThread(() -> Toast.makeText(MainActivity.this, "App is up to date (v" + vName + ")", Toast.LENGTH_SHORT).show());
                    }
                } else if (manual) {
                    runOnUiThread(() -> Toast.makeText(MainActivity.this, "Could not check updates (Server code: " + responseCode + ")", Toast.LENGTH_SHORT).show());
                }
            } catch (Exception e) {
                if (manual) {
                    runOnUiThread(() -> Toast.makeText(MainActivity.this, "Update check: " + e.getMessage(), Toast.LENGTH_SHORT).show());
                }
            } finally {
                if (conn != null) conn.disconnect();
            }
        }).start();
    }

    private void showUpdateDialog(String newVersion, String downloadUrl, String releaseNotes) {
        new AlertDialog.Builder(this)
                .setTitle("New App Update Available!")
                .setMessage("A new version (" + newVersion + ") of Production Test Report is available on the server.\n\n" + releaseNotes + "\n\nWould you like to download and install it now?")
                .setCancelable(true)
                .setPositiveButton("Update Now", (d, w) -> downloadAndInstallApk(downloadUrl, newVersion))
                .setNegativeButton("Later", null)
                .show();
    }

    private void downloadAndInstallApk(String url, String version) {
        try {
            String fullDownloadUrl = url;
            if (!fullDownloadUrl.startsWith("http://") && !fullDownloadUrl.startsWith("https://")) {
                String serverUrl = getServerUrl();
                if (serverUrl != null && !serverUrl.isEmpty()) {
                    Uri sUri = Uri.parse(serverUrl);
                    fullDownloadUrl = sUri.getScheme() + "://" + sUri.getAuthority() + (url.startsWith("/") ? "" : "/") + url;
                }
            }

            final String apkFileName = "ProductionReportApp_v" + version + ".apk";
            DownloadManager.Request request = new DownloadManager.Request(Uri.parse(fullDownloadUrl));
            request.setTitle("Updating Production Test Report");
            request.setDescription("Downloading v" + version + " from server...");
            request.setMimeType("application/vnd.android.package-archive");
            request.setNotificationVisibility(DownloadManager.Request.VISIBILITY_VISIBLE_NOTIFY_COMPLETED);
            request.setDestinationInExternalPublicDir(Environment.DIRECTORY_DOWNLOADS, apkFileName);

            final DownloadManager dm = (DownloadManager) getSystemService(Context.DOWNLOAD_SERVICE);
            if (dm == null) return;
            final long downloadId = dm.enqueue(request);
            Toast.makeText(this, "Downloading update... You will be prompted to install once finished.", Toast.LENGTH_LONG).show();

            BroadcastReceiver onComplete = new BroadcastReceiver() {
                @Override
                public void onReceive(Context context, Intent intent) {
                    long id = intent.getLongExtra(DownloadManager.EXTRA_DOWNLOAD_ID, -1);
                    if (id == downloadId) {
                        try {
                            context.unregisterReceiver(this);
                        } catch (Exception ignored) {}

                        File file = new File(Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS), apkFileName);
                        if (file.exists()) {
                            Intent installIntent = new Intent(Intent.ACTION_VIEW);
                            installIntent.setFlags(Intent.FLAG_ACTIVITY_NEW_TASK);
                            if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.N) {
                                Uri apkUri = FileProvider.getUriForFile(MainActivity.this, getPackageName() + ".fileprovider", file);
                                installIntent.setDataAndType(apkUri, "application/vnd.android.package-archive");
                                installIntent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);
                            } else {
                                installIntent.setDataAndType(Uri.fromFile(file), "application/vnd.android.package-archive");
                            }
                            startActivity(installIntent);
                        }
                    }
                }
            };
            if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.TIRAMISU) {
                registerReceiver(onComplete, new IntentFilter(DownloadManager.ACTION_DOWNLOAD_COMPLETE), Context.RECEIVER_EXPORTED);
            } else {
                registerReceiver(onComplete, new IntentFilter(DownloadManager.ACTION_DOWNLOAD_COMPLETE));
            }
        } catch (Exception e) {
            Toast.makeText(this, "Update error: " + e.getMessage(), Toast.LENGTH_LONG).show();
        }
    }

    @Override
    protected void onActivityResult(int requestCode, int resultCode, @Nullable Intent data) {
        super.onActivityResult(requestCode, resultCode, data);
        if (requestCode == FILE_CHOOSER_REQUEST_CODE) {
            if (uploadMessage == null) return;
            uploadMessage.onReceiveValue(WebChromeClient.FileChooserParams.parseResult(resultCode, data));
            uploadMessage = null;
        }
    }

    @Override
    public boolean onCreateOptionsMenu(Menu menu) {
        getMenuInflater().inflate(R.menu.main_menu, menu);
        return true;
    }

    @Override
    public boolean onOptionsItemSelected(MenuItem item) {
        if (item.getItemId() == R.id.action_server_settings) {
            showServerConfigDialog();
            return true;
        } else if (item.getItemId() == R.id.action_check_updates) {
            checkForAppUpdate(true);
            return true;
        } else if (item.getItemId() == R.id.action_reload) {
            connectToProductionServer(true);
            return true;
        }
        return super.onOptionsItemSelected(item);
    }

    @Override
    public void onBackPressed() {
        if (webView != null && webView.canGoBack()) {
            webView.goBack();
        } else {
            super.onBackPressed();
        }
    }

    public static boolean isPortOpen(String host, int port, int timeoutMs) {
        return SubnetScannerTask.isPortOpen(host, port, timeoutMs);
    }

    private static class SubnetScannerTask extends AsyncTask<Void, Void, String> {
        private final EditText etHost;
        private final Button btnScan;

        SubnetScannerTask(EditText etHost, Button btnScan) {
            this.etHost = etHost;
            this.btnScan = btnScan;
        }

        @Override
        protected String doInBackground(Void... voids) {
            int port = Integer.parseInt(DEFAULT_PORT);
            if (isPortOpen("127.0.0.1", port, 200)) {
                return "127.0.0.1";
            }
            if (isPortOpen(DEFAULT_HOST, port, 400)) {
                return DEFAULT_HOST;
            }
            String[] subnets = {"192.168.99.", "192.168.1.", "192.168.0.", "192.168.43."};
            for (String sub : subnets) {
                for (int i = 2; i <= 254; i++) {
                    String candidate = sub + i;
                    if (isPortOpen(candidate, port, 40)) {
                        return candidate;
                    }
                }
            }
            return null;
        }

        public static boolean isPortOpen(String host, int port, int timeoutMs) {
            try (Socket socket = new Socket()) {
                socket.connect(new InetSocketAddress(host, port), timeoutMs);
                return true;
            } catch (IOException e) {
                return false;
            }
        }

        @Override
        protected void onPostExecute(String foundHost) {
            btnScan.setEnabled(true);
            btnScan.setText("Auto-Scan Wi-Fi");
            if (foundHost != null) {
                etHost.setText(foundHost);
                Toast.makeText(etHost.getContext(), "Found Server at: " + foundHost, Toast.LENGTH_SHORT).show();
            } else {
                Toast.makeText(etHost.getContext(), "Server not found on Wi-Fi scan.", Toast.LENGTH_SHORT).show();
            }
        }
    }
}
