package com.secretovens.samanlist;

import android.app.Activity;
import android.content.Intent;
import android.net.Uri;
import android.os.Bundle;
import android.webkit.ValueCallback;
import android.webkit.WebChromeClient;
import android.webkit.WebResourceError;
import android.webkit.WebResourceRequest;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.webkit.WebViewClient;
import android.widget.Toast;

public class MainActivity extends Activity {
    // The app screens live online, so every update reaches all phones automatically.
    static final String APP_URL = "https://secretovens.github.io/secret-ovens-app/";
    static final String APP_HOST = "secretovens.github.io";
    private WebView web;
    private ValueCallback<Uri[]> filePick;
    static final int PICK = 41;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        web = new WebView(this);
        WebSettings s = web.getSettings();
        s.setJavaScriptEnabled(true);
        s.setDomStorageEnabled(true);
        s.setCacheMode(WebSettings.LOAD_DEFAULT);
        web.clearCache(true); // always fetch the newest version on start
        web.setWebViewClient(new WebViewClient() {
            @Override
            public boolean shouldOverrideUrlLoading(WebView view, WebResourceRequest req) {
                return openOutside(req.getUrl());
            }

            @SuppressWarnings("deprecation")
            @Override
            public boolean shouldOverrideUrlLoading(WebView view, String url) {
                return openOutside(Uri.parse(url));
            }

            @Override
            public void onReceivedError(WebView view, WebResourceRequest req, WebResourceError err) {
                if (req.isForMainFrame()) view.loadUrl("file:///android_asset/offline.html");
            }

            @SuppressWarnings("deprecation")
            @Override
            public void onReceivedError(WebView view, int code, String desc, String url) {
                if (url != null && url.startsWith(APP_URL)) view.loadUrl("file:///android_asset/offline.html");
            }
        });
        // बिल और सामान की फ़ोटो के लिए: कैमरा/गैलरी से फ़ोटो चुनना
        web.setWebChromeClient(new WebChromeClient() {
            @Override
            public boolean onShowFileChooser(WebView view, ValueCallback<Uri[]> cb, FileChooserParams params) {
                if (filePick != null) filePick.onReceiveValue(null);
                filePick = cb;
                try {
                    Intent pick = new Intent(Intent.ACTION_GET_CONTENT);
                    pick.addCategory(Intent.CATEGORY_OPENABLE);
                    pick.setType("image/*");
                    startActivityForResult(Intent.createChooser(pick, "फ़ोटो चुनें"), PICK);
                } catch (Exception e) {
                    filePick = null;
                    Toast.makeText(MainActivity.this, "फ़ोटो नहीं खुल पाई", Toast.LENGTH_LONG).show();
                    return false;
                }
                return true;
            }
        });
        setContentView(web);
        web.loadUrl(APP_URL);
    }

    // WhatsApp and other links open in their own app
    private boolean openOutside(Uri uri) {
        String scheme = uri.getScheme();
        if ("file".equals(scheme)) return false;
        if (APP_HOST.equals(uri.getHost())) return false;
        try {
            startActivity(new Intent(Intent.ACTION_VIEW, uri));
        } catch (Exception e) {
            Toast.makeText(this, "यह लिंक नहीं खुल पाया", Toast.LENGTH_LONG).show();
        }
        return true;
    }

    @Override
    protected void onActivityResult(int req, int res, Intent data) {
        super.onActivityResult(req, res, data);
        if (req != PICK || filePick == null) return;
        Uri[] out = null;
        if (res == RESULT_OK && data != null && data.getData() != null) out = new Uri[]{ data.getData() };
        filePick.onReceiveValue(out);
        filePick = null;
    }

    @Override
    public void onBackPressed() {
        web.evaluateJavascript(
            "(window.appBack && window.appBack()) ? 'closed' : 'none'",
            r -> { if (r == null || !r.contains("closed")) MainActivity.super.onBackPressed(); });
    }
}
