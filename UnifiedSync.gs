// ==============================
// Unified Offline Sync Helpers
// ==============================
// Exposes read-only export and replay-style import helpers that combine
// the Heart to Hand and Vouchers apps without bypassing their integrity
// checks. All writes are delegated to the existing, validated entry
// points (saveHHEntry + addToFormResponses1) so headers, locks, rebuilds,
// and signature/phone handling stay intact.

/**
 * Single entry point router.
 * Usage:
 *   no app param   -> renders combined container page with two iframes/tabs
 *   ?app=vouchers  -> renders voucher lookup UI (keeps voucher math/offline logic intact)
 *   ?app=heart     -> renders Heart to Hand UI
 */
function doGet(e) {
  const app = (e && e.parameter && e.parameter.app) || '';
  if (String(app).toLowerCase() === 'heart') {
    return renderHeartToHand();
  }
  if (String(app).toLowerCase() === 'vouchers') {
    return renderVoucherLookup();
  }
  try {
    const t = HtmlService.createTemplateFromFile('Combined');
    t.baseUrl = ScriptApp.getService().getUrl();
    return t.evaluate().setTitle('Heart to Hand & Vouchers');
  } catch (_err) {
    // Fallback: if Combined.html is missing in this deployment, serve Heart to Hand.
    return renderHeartToHand();
  }
}

/**
 * Build a unified export payload combining both apps.
 * @param {Object} [options]
 * @param {Array<{lastName:string, firstName:string}>} [options.hhHistoryNames]
 *        Optional list of client names to include detailed Heart to Hand
 *        history for. If omitted, only the client list is returned.
 */
function exportUnifiedData(options) {
  const opts = options || {};
  const hhHistoryNames = Array.isArray(opts.hhHistoryNames) ? opts.hhHistoryNames : [];

  const heartToHand = {
    clients: getAllHHClients()
  };

  if (hhHistoryNames.length) {
    const historyByClient = {};
    hhHistoryNames.forEach((name) => {
      if (!name || !name.lastName || !name.firstName) return;
      try {
        const key = `${String(name.lastName).trim()}|${String(name.firstName).trim()}`;
        historyByClient[key] = getHHHistory(name.lastName, name.firstName);
      } catch (_err) {
        // If a specific client lookup fails, skip it without breaking the export.
      }
    });
    heartToHand.historyByClient = historyByClient;
  }

  const vouchers = {
    balances: getAllClientBalances(),
    ledger: fetchVoucherJson()
  };

  return { vouchers, heartToHand };
}

/**
 * Replay pending offline writes into both apps through the validated
 * entry points. Any failures are captured per-item so a partial failure
 * does not hide other successes.
 *
 * Payload shape:
 * {
 *   vouchers: {
 *     responses: [ ...objects consumable by addToFormResponses1 ],
 *   },
 *   heartToHand: {
 *     entries: [ ...objects consumable by saveHHEntry ],
 *   }
 * }
 */
function importUnifiedData(payload) {
  const body = payload || {};

  const results = {
    vouchers: { responses: [] },
    heartToHand: { entries: [] },
  };

  // --- Heart to Hand ---
  const hhEntries = Array.isArray(body.heartToHand && body.heartToHand.entries)
    ? body.heartToHand.entries
    : [];

  hhEntries.forEach((entry) => {
    try {
      const res = saveHHEntry(entry);
      results.heartToHand.entries.push({ ok: true, result: res });
    } catch (err) {
      results.heartToHand.entries.push({ ok: false, error: String(err && err.message ? err.message : err) });
    }
  });

  // --- Vouchers (Form responses path) ---
  const voucherResponses = Array.isArray(body.vouchers && body.vouchers.responses)
    ? body.vouchers.responses
    : [];

  voucherResponses.forEach((resp) => {
    try {
      const res = addToFormResponses1(resp);
      results.vouchers.responses.push({ ok: true, result: res });
    } catch (err) {
      results.vouchers.responses.push({ ok: false, error: String(err && err.message ? err.message : err) });
    }
  });

  return results;
}

/**
 * Optional HTTPS entry point for export/import.
 * POST body JSON: { action: "exportUnified"|"importUnified", data?: {...}, options?: {...} }
 */
function doPost(e) {
  const raw = e && e.postData && e.postData.contents;
  let parsed = {};
  try {
    parsed = raw ? JSON.parse(raw) : {};
  } catch (_err) {
    parsed = {};
  }

  const action = (parsed && parsed.action) || (e && e.parameter && e.parameter.action);

  let output;
  try {
    if (action === 'exportUnified') {
      output = exportUnifiedData(parsed.options || {});
    } else if (action === 'importUnified') {
      output = importUnifiedData(parsed.data || parsed.payload || {});
    } else {
      output = { error: 'Unsupported action', action };
    }
  } catch (err) {
    output = { error: String(err && err.message ? err.message : err), action };
  }

  return ContentService
    .createTextOutput(JSON.stringify(output))
    .setMimeType(ContentService.MimeType.JSON);
}
