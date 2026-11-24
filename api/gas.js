/**
 * ============================================================================
 * VERCEL SERVERLESS API PROXY UNTUK GOOGLE APPS SCRIPT (100% BEBAS CORS)
 * ============================================================================
 * File ini berjalan di server Vercel (Node.js) sehingga request dari browser
 * dikirim ke domain Vercel yang sama (/api/gas), kemudian server Vercel yang
 * meneruskannya ke Google Apps Script dan mengikuti redirect 302 Google.
 * Hasilnya: TIDAK AKAN PERNAH terkena "Failed to fetch" atau blokir CORS/AdBlocker!
 */

const DEFAULT_GAS_URL =
  process.env.GAS_WEB_APP_URL ||
  'https://script.google.com/macros/s/AKfycbyfOKlaGhb8cKmGo2PCIEGI5Jp2gF_4_gh4I3Q0-pTCMtMc3Rzi7A6CDBuQNsgoNR7p/exec';

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  try {
    let payload = req.body || {};
    if (typeof payload === 'string') {
      try {
        payload = JSON.parse(payload);
      } catch (e) {
        payload = {};
      }
    }

    const action = String(payload.action || (req.query && req.query.action) || '').trim();
    const args = Array.isArray(payload.args)
      ? payload.args
      : (req.query && req.query.args ? JSON.parse(req.query.args) : []);
    const spreadsheetId = String(
      payload.spreadsheetId ||
      (req.query && req.query.spreadsheetId) ||
      process.env.CLIENT_SPREADSHEET_ID ||
      ''
    ).trim();
    const clientHost = String(
      payload.clientHost ||
      (req.headers && (req.headers['x-forwarded-host'] || req.headers.host)) ||
      ''
    ).trim();
    const targetGasUrl = String(payload.gasUrl || DEFAULT_GAS_URL).trim();

    if (!targetGasUrl || !/^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec/i.test(targetGasUrl)) {
      return res.status(400).json({
        status: 'error',
        message: 'URL Web App Google Apps Script (/exec) belum dikonfigurasi dengan benar.'
      });
    }

    if (!action) {
      return res.status(400).json({
        status: 'error',
        message: 'Parameter action API tidak ditemukan.'
      });
    }

    // 1. Coba kirim via POST ke doPost(e) di Google Apps Script (Sertakan spreadsheetId & clientHost pembeli)
    let gasResponse = await fetch(targetGasUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'text/plain;charset=utf-8',
        'Accept': 'application/json, text/plain, */*'
      },
      body: JSON.stringify({ action, args, spreadsheetId, clientHost }),
      redirect: 'follow'
    });

    let rawText = await gasResponse.text();

    // 2. Fallback pintar: Jika deployment Apps Script lama belum memiliki doPost(e)
    //    atau mengembalikan halaman HTML, coba panggil via GET ?action=...&args=...&spreadsheetId=...&clientHost=...
    if (rawText.trim().startsWith('<') || !rawText.trim().startsWith('{')) {
      const urlWithQuery =
        targetGasUrl +
        (targetGasUrl.includes('?') ? '&' : '?') +
        'action=' + encodeURIComponent(action) +
        '&args=' + encodeURIComponent(JSON.stringify(args)) +
        (spreadsheetId ? '&spreadsheetId=' + encodeURIComponent(spreadsheetId) : '') +
        (clientHost ? '&clientHost=' + encodeURIComponent(clientHost) : '');

      const getResp = await fetch(urlWithQuery, {
        method: 'GET',
        headers: { 'Accept': 'application/json, text/plain, */*' },
        redirect: 'follow'
      });
      const getRawText = await getResp.text();
      if (getRawText.trim().startsWith('{')) {
        rawText = getRawText;
      }
    }

    // 3. Jika response masih berupa HTML dari Google, berikan pesan diagnosis yang sangat jelas
    if (rawText.trim().startsWith('<')) {
      if (rawText.includes('accounts.google.com') || rawText.includes('ServiceLogin')) {
        return res.status(200).json({
          status: 'error',
          message:
            'Akses ditolak oleh Google Apps Script (Meminta Login Google). Pastikan saat Deploy di Apps Script, bagian "Who has access / Siapa yang memiliki akses" dipilih "Anyone / Siapa saja".'
        });
      }
      if (rawText.includes('Script function not found: doPost')) {
        return res.status(200).json({
          status: 'error',
          message:
            'Fungsi doPost belum aktif di link Apps Script Anda! Silakan copy-paste kode code.gs terbaru ke Editor Apps Script, lalu klik Deploy -> Manage Deployments -> Edit -> pilih "New version / Versi baru" -> Deploy.'
        });
      }
      return res.status(200).json({
        status: 'error',
        message:
          'Deployment Google Apps Script Anda masih menjalankan versi kode lama (mengembalikan HTML, bukan JSON). Silakan salin code.gs terbaru ke Apps Script, lalu klik Deploy -> New deployment (atau Manage deployments -> Edit -> New version).'
      });
    }

    const data = JSON.parse(rawText);
    return res.status(200).json(data);
  } catch (err) {
    return res.status(200).json({
      status: 'error',
      message: 'Gagal menghubungi Google Apps Script: ' + (err && err.message ? err.message : String(err))
    });
  }
};
