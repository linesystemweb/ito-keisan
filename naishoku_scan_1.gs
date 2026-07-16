/**
 * 内職 出荷・入荷管理 - 共有バックエンド v7 (Google Apps Script)
 *
 * v7の変更点:
 * - 出荷登録時に「出荷日（依頼日）」を選択できるようになったため、
 *   クライアントから送られる ts（任意）をそのまま採用するようにしました
 *
 * v6までの変更点:
 * - 1回の出荷・入荷で複数商品（最大10点）を1つの伝票番号にまとめられます
 * - 入荷登録時に「残りを打ち切って完結する」を選べます
 *
 * セットアップは同じ手順です:
 * 拡張機能 → Apps Script → このコードを貼り付けて保存 → デプロイ（新しいバージョン）
 */

const SHIPMENTS_SHEET = '依頼';
const RECEIPTS_SHEET = '受領';
const TZ = 'Asia/Tokyo';
const MAX_LINES = 10;

function doGet(e) {
  var action = e && e.parameter ? e.parameter.action : null;
  if (action === 'list') {
    return jsonOut_({
      shipments: listShipments_(),
      receipts: listReceipts_()
    });
  }
  return jsonOut_({ error: 'unknown action' });
}

function doPost(e) {
  var body = {};
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut_({ error: 'invalid body' });
  }

  if (body.action === 'add_shipment_batch') {
    var r = addShipmentBatch_(body.entries, body.worker, body.ts);
    if (r.error) return jsonOut_({ error: r.error });
    return jsonOut_({ ok: true, record: r });
  }
  if (body.action === 'add_receipt_batch') {
    var r2 = addReceiptBatch_(body.items, body.worker);
    if (r2.error) return jsonOut_({ error: r2.error });
    return jsonOut_({ ok: true, record: r2 });
  }
  if (body.action === 'delete_shipment') {
    var r3 = deleteShipmentLine_(body.voucherNo, body.line);
    if (r3 && r3.error) return jsonOut_({ error: r3.error });
    return jsonOut_({ ok: true });
  }
  if (body.action === 'delete_receipt') {
    deleteReceiptLine_(body.voucherNo, body.line);
    return jsonOut_({ ok: true });
  }
  return jsonOut_({ error: 'unknown action' });
}

/* ====================== SHEET SETUP ====================== */

function getSpreadsheet_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function getShipmentsSheet_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(SHIPMENTS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SHIPMENTS_SHEET);
    sheet.appendRow(['伝票番号', '行', '日付', '時刻', 'ts', '内職者コード', '内職さん',
      '商品コード', '商品名', 'JAN', '出荷数', '回収済数', '残数', '備考', '状態', '完結区分']);
  }
  return sheet;
}

function getReceiptsSheet_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(RECEIPTS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(RECEIPTS_SHEET);
    sheet.appendRow(['伝票番号', '行', '日付', '時刻', 'ts', '出荷伝票番号', '出荷行', '委託日',
      '内職者コード', '内職さん', '商品コード', '商品名', '受領数', '備考']);
  }
  return sheet;
}

/* ====================== VOUCHER NUMBER ====================== */

function getNextVoucherNo_() {
  var max = 0;
  [getShipmentsSheet_(), getReceiptsSheet_()].forEach(function (sheet) {
    var values = sheet.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      var v = String(values[i][0] || '');
      var n = parseInt(v.replace(/[^0-9]/g, ''), 10);
      if (!isNaN(n) && n > max) max = n;
    }
  });
  var padded = ('000000' + (max + 1)).slice(-6);
  return 'W' + padded;
}

/* ====================== SHIPMENTS (出荷/依頼) ====================== */
// 列: 伝票番号|行|日付|時刻|ts|内職者コード|内職さん|商品コード|商品名|JAN|出荷数|回収済数|残数|備考|状態|完結区分

function addShipmentBatch_(entries, worker, tsOverride) {
  if (!entries || !entries.length) return { error: '商品が選択されていません' };
  if (entries.length > MAX_LINES) return { error: `1伝票につき最大${MAX_LINES}点までです` };
  if (!worker || !worker.code) return { error: '内職者情報がありません' };

  var sheet = getShipmentsSheet_();
  var voucherNo = getNextVoucherNo_();
  var ts = tsOverride || Date.now();
  var d = new Date(ts);
  var dateStr = Utilities.formatDate(d, TZ, 'yyyy/MM/dd');
  var timeStr = Utilities.formatDate(d, TZ, 'HH:mm');

  var items = [];
  entries.forEach(function (entry, idx) {
    var line = idx + 1;
    var qty = Number(entry.qty) || 0;
    sheet.appendRow([
      voucherNo, line, dateStr, timeStr, ts,
      worker.code, worker.name,
      entry.productCode || '', entry.productName || '', entry.jan || '',
      qty, 0, qty, entry.note || '', 'open', ''
    ]);
    items.push({
      line: line, productCode: entry.productCode, productName: entry.productName,
      jan: entry.jan || '', qty: qty, note: entry.note || ''
    });
  });

  return { voucherNo: voucherNo, date: dateStr, time: timeStr, workerCode: worker.code, workerName: worker.name, items: items };
}

function listShipments_() {
  var sheet = getShipmentsSheet_();
  var values = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    if (!r[0]) continue;
    out.push({
      voucherNo: r[0], line: r[1], date: r[2], time: r[3], ts: r[4],
      workerCode: r[5], workerName: r[6], productCode: r[7], productName: r[8], jan: r[9],
      qty: r[10], returnedQty: r[11], remaining: r[12], note: r[13], status: r[14], finalizeType: r[15]
    });
  }
  return out;
}

function findShipmentRow_(sheet, voucherNo, line) {
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(voucherNo) && String(values[i][1]) === String(line)) {
      return { row: i + 1, data: values[i] };
    }
  }
  return null;
}

function deleteShipmentLine_(voucherNo, line) {
  var sheet = getShipmentsSheet_();
  var found = findShipmentRow_(sheet, voucherNo, line);
  if (!found) return { error: 'not found' };
  var returnedQty = Number(found.data[11]) || 0;
  if (returnedQty > 0) {
    return { error: 'この出荷にはすでに入荷記録が紐づいているため削除できません。先に該当する入荷記録を削除してください。' };
  }
  sheet.deleteRow(found.row);
  return { ok: true };
}

/* ====================== RECEIPTS (入荷/受領) ====================== */
// items: [{ shipmentVoucherNo, shipmentLine, qty, finalize(bool), note }]
// 列: 伝票番号|行|日付|時刻|ts|出荷伝票番号|出荷行|委託日|内職者コード|内職さん|商品コード|商品名|受領数|備考

function addReceiptBatch_(items, worker) {
  if (!items || !items.length) return { error: '受領する商品が選択されていません' };
  if (items.length > MAX_LINES) return { error: `1伝票につき最大${MAX_LINES}点までです` };

  var shipSheet = getShipmentsSheet_();
  var recSheet = getReceiptsSheet_();
  var voucherNo = getNextVoucherNo_();
  var ts = Date.now();
  var d = new Date(ts);
  var dateStr = Utilities.formatDate(d, TZ, 'yyyy/MM/dd');
  var timeStr = Utilities.formatDate(d, TZ, 'HH:mm');

  var outItems = [];

  for (var idx = 0; idx < items.length; idx++) {
    var it = items[idx];
    var found = findShipmentRow_(shipSheet, it.shipmentVoucherNo, it.shipmentLine);
    if (!found) {
      return { error: `伝票 ${it.shipmentVoucherNo}-${it.shipmentLine} が見つかりません` };
    }
    var row = found.row;
    var data = found.data;
    var qtyShipped = Number(data[10]) || 0;
    var returnedQty = Number(data[11]) || 0;
    var qtyThis = Number(it.qty) || 0;

    var newReturned = returnedQty + qtyThis;
    var newRemaining = qtyShipped - newReturned;
    var finalizeType = '';
    var newStatus;

    if (it.finalize) {
      // 残りを打ち切って完結する
      newRemaining = 0;
      newStatus = 'closed';
      finalizeType = newReturned >= qtyShipped ? '' : '打切り';
    } else {
      newStatus = newRemaining <= 0 ? 'closed' : 'open';
      finalizeType = newRemaining <= 0 ? '' : '';
    }

    shipSheet.getRange(row, 12, 1, 3).setValues([[newReturned, newRemaining, newStatus]]); // 回収済数,残数,状態
    if (finalizeType) shipSheet.getRange(row, 16).setValue(finalizeType); // 完結区分

    var line = idx + 1;
    recSheet.appendRow([
      voucherNo, line, dateStr, timeStr, ts,
      data[0], data[1], data[2], // 出荷伝票番号, 出荷行, 委託日
      data[5], data[6], data[7], data[8], // 内職者コード,内職さん,商品コード,商品名
      qtyThis, it.note || ''
    ]);

    outItems.push({
      line: line, shipmentVoucherNo: data[0], shipmentLine: data[1], shipmentDate: data[2],
      productCode: data[7], productName: data[8],
      qty: qtyThis, note: it.note || '',
      shipmentRemaining: newRemaining, shipmentStatus: newStatus, finalizeType: finalizeType
    });
  }

  return {
    voucherNo: voucherNo, date: dateStr, time: timeStr,
    workerCode: worker ? worker.code : '', workerName: worker ? worker.name : '',
    items: outItems
  };
}

function listReceipts_() {
  var sheet = getReceiptsSheet_();
  var values = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    if (!r[0]) continue;
    out.push({
      voucherNo: r[0], line: r[1], date: r[2], time: r[3], ts: r[4],
      shipmentVoucherNo: r[5], shipmentLine: r[6], shipmentDate: r[7],
      workerCode: r[8], workerName: r[9], productCode: r[10], productName: r[11],
      qty: r[12], note: r[13]
    });
  }
  return out;
}

function deleteReceiptLine_(voucherNo, line) {
  var recSheet = getReceiptsSheet_();
  var values = recSheet.getDataRange().getValues();
  var rowIdx = -1, shipVoucherNo = null, shipLine = null, qty = 0;
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(voucherNo) && String(values[i][1]) === String(line)) {
      rowIdx = i + 1;
      shipVoucherNo = values[i][5];
      shipLine = values[i][6];
      qty = Number(values[i][12]) || 0;
      break;
    }
  }
  if (rowIdx === -1) return;
  recSheet.deleteRow(rowIdx);

  var shipSheet = getShipmentsSheet_();
  var found = findShipmentRow_(shipSheet, shipVoucherNo, shipLine);
  if (found) {
    var qtyShipped = Number(found.data[10]) || 0;
    var returnedQty = Number(found.data[11]) || 0;
    var newReturned = Math.max(0, returnedQty - qty);
    var newRemaining = qtyShipped - newReturned;
    var newStatus = newRemaining <= 0 ? 'closed' : 'open';
    shipSheet.getRange(found.row, 12, 1, 3).setValues([[newReturned, newRemaining, newStatus]]);
    shipSheet.getRange(found.row, 16).setValue(''); // 完結区分クリア（再オープン扱い）
  }
}

/* ====================== UTIL ====================== */

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
