# PRODUCTION TEST REPORT - REAL ANDROID APK

## Official System Identity
* **Application Title**: Production Test Report
* **Company**: IPA Weighing & Automation (IPA Private Limited)
* **Package Name**: `com.production.reportapp`
* **Version**: 2.0 (Target API 29)

---

## What is in this Folder:
* `ProductionReportApp.apk`: The real, compiled, signed, and ready-to-install Android APK (738 KB).
* `INSTALL_REAL_APK.bat`: 1-Click installer for any USB-connected Android phone with auto port-reversing.
* `source_code/`:
  * `java/com/production/reportapp/MainActivity.java`: Complete Java source code with custom User-Agent, Android DownloadManager listener, camera/file upload handler, and local network scanner.
  * `res/`: Complete Android UI resources (layouts, strings, colors, drawables, and IPA logo launcher icons).
  * `AndroidManifest.xml`: Official Android Manifest with camera, storage, and internet permissions.
  * `build.gradle` & `settings.gradle`: Android project configuration.

---

## Key Guarantees & Features:
1. **Zero Left-to-Right Scrolling**:
   - The topbar and main containers are strictly confined to 100% width on phone screens (~210px header).
   - No horizontal overflow, no screen wobbling, and no cutoff buttons.

2. **Perfect Data & Heading Alignment**:
   - Data table headings and input rows never intersect, overlap, or collide.
   - All multi-column tables are wrapped in touch-scrollable `.table-responsive` cards.

3. **No "Install APK" Prompts Inside the App**:
   - The APK identifies itself cleanly as `is_app`.
   - All website install banners and topbar download links are completely hidden inside the app.

4. **All 15 Systems & Auto-Calculations Supported**:
   - **Belt Scale**: RPM, speed, design load (kg/m), pulses/second, multi-point calibration.
   - **Weigh Feeder**: Belt loading, feed rate, tacho auto-calc.
   - **Digital Indicator**: 8-channel load cell matrix, tare/span, error % calculation.
   - **Batching System**: Setpoint, batch tolerances, cumulative totals.
   - **Crane Scale, Signal Conditioner, Trip Safe, Vibration Switch, Accelerometers, Vibration Meter, Charge Amplifier, ODD & DD Systems, Loss-in-Weigh Feeder, and Misc Reports**.

5. **100% Direct Localhost Speed**:
   - Zero internet tunnels, zero third-party proxies.
   - Directly connects to your laptop over Wi-Fi (`http://192.168.1.20:5000`) or USB (`http://127.0.0.1:5000`).

---

## How to Install on Your Phone:
* **Option 1 (USB Cable)**:
  Connect your phone to this laptop with a USB cable (USB Debugging ON), then double-click `INSTALL_REAL_APK.bat`.
* **Option 2 (Phone Browser)**:
  Connect your phone to the same Wi-Fi as your laptop, open `http://192.168.1.20:5000/download-app` in Chrome, and tap install.
