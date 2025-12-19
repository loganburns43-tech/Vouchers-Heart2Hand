# Heart to Hand & Vouchers – offline switching

## How staff should use offline mode
1. **Preload both apps while online:** open each app and tap **“Preload both apps for offline”** so both tabs are created.
2. **Keep both tabs open:** leave the Heart to Hand and Vouchers tabs open in the kiosk browser.
3. **If Wi-Fi drops:** continue working in the open tabs; switching works only if the other tab is already open.
4. **When back online:** queued changes will sync automatically.

## Manual test checklist
- From Heart to Hand, click **Switch to Vouchers** → Vouchers opens/reuses a tab; if popups are blocked, the in-app message shows the URL to open.
- From Vouchers, click **Switch to Heart to Hand** → Heart to Hand opens/reuses a tab; blocked popup message appears with the URL if needed.
- Click **Preload both apps for offline** in each app → creates the other tab, shows a “Preloaded” pill, and refocuses the current tab.
- While offline, switching shows the offline notice and attempts to focus the existing named tab without errors.
- The preloaded indicator persists after reload and shows a timestamp.
