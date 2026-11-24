/**
 * 口コミページのアンケート回答を、このスプレッドシートに記録する Apps Script
 *
 * 使い方（README「スプレッドシート連携」も参照）
 *   1. 記録用のスプレッドシートを開き、[拡張機能] → [Apps Script]
 *   2. このファイルの中身をすべて貼り付けて保存
 *   3. [デプロイ] → [新しいデプロイ] → 種類「ウェブアプリ」
 *        次のユーザーとして実行：自分 ／ アクセスできるユーザー：全員
 *   4. 表示された「ウェブアプリのURL」を index.html の CONFIG.sheetUrl に貼る
 *
 * ページから届くデータ
 *   type: "survey" … アンケートに答えた → 1行追加
 *   type: "google" … 「Googleで口コミを書く」を押した → 同じIDの行を「押した」にする
 *
 * ★ このファイルを直したら：[デプロイ] → [デプロイを管理] → 鉛筆 → バージョン「新バージョン」→ [デプロイ]
 *   （保存しただけでは反映されない。URLは変わらない）
 */

const SHEET_NAME = 'アンケート';
const HEADERS = ['日時', 'ID', '答え方', 'ご利用', '来店経路', '品物', '印象に残ったこと',
                 'あなたの言葉', '下書きの文章', 'Google口コミボタン', 'ボタンを押した日時', 'コピーした文章'];
const COL_OWN     = 8;    // 「あなたの言葉」（H列）
const COL_GOOGLE  = 10;   // 「Google口コミボタン」押した／押していない（J列）
const COL_PRESSED = 11;   // 「ボタンを押した日時」（K列）
const COL_FINAL   = 12;   // 「コピーした文章」（L列）

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);   // 同時に届いても行が混ざらないように
  try {
    const d = JSON.parse(e.postData.contents);
    const id = String(d.id || '');
    if (!/^[0-9a-f]{16}$/.test(id)) return reply_('bad id');   // ページ以外からのいたずら対策（簡易）

    const sh = getSheet_();
    if (d.type === 'survey') {
      sh.appendRow([
        new Date(), id,
        safe_(d.mode, 10), safe_(d.visit, 20), safe_(d.route, 200), safe_(d.items, 300),
        safe_(d.points, 400), safe_(d.own, 200), safe_(d.text, 1000), '押していない', '', '',
      ]);
    } else if (d.type === 'google') {
      const row = findRow_(sh, id);
      if (row) {
        sh.getRange(row, COL_GOOGLE).setValue('押した');
        sh.getRange(row, COL_PRESSED).setValue(new Date());
        sh.getRange(row, COL_FINAL).setValue(safe_(d.text, 1000));
        if (d.own) sh.getRange(row, COL_OWN).setValue(safe_(d.own, 200));   // 口コミ画面で書き足した一言
      } else {
        // アンケートをとばして口コミへ進んだ人
        sh.appendRow([new Date(), id, '（アンケートなし）', '', '', '', '', safe_(d.own, 200), '',
                      '押した', new Date(), safe_(d.text, 1000)]);
      }
    }
    return reply_('ok');
  } catch (err) {
    return reply_('error');
  } finally {
    lock.releaseLock();
  }
}

// シートがなければ作り、1行目に見出しを入れる
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  // 見出し行がない・古い場合は書き直す（列を追加したときも自動で更新される）
  const head = sh.getRange(1, 1, 1, HEADERS.length);
  if (head.getValues()[0].join('|') !== HEADERS.join('|')) {
    head.setValues([HEADERS]).setFontWeight('bold').setBackground('#d2d5ec');
    sh.setFrozenRows(1);
  }
  return sh;
}

// 同じIDの行を下から探す（最新の回答）
function findRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const ids = sh.getRange(2, 2, last - 1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) if (ids[i][0] === id) return i + 2;
  return 0;
}

// 長さを制限し、「=」などで始まる文字は数式として実行されないよう先頭に ' をつける
function safe_(v, max) {
  let s = String(v == null ? '' : v).slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function reply_(msg) {
  return ContentService.createTextOutput(msg).setMimeType(ContentService.MimeType.TEXT);
}

// 動作確認用：Apps Script の画面でこの関数を選んで [実行] すると、テスト行が1行入る
function testInsert() {
  doPost({ postData: { contents: JSON.stringify({
    id: '0123456789abcdef', type: 'survey', mode: 'かんたん', visit: '初めて',
    route: 'Googleマップ', items: '時計', points: '説明が丁寧', own: '', text: 'テストです',
  }) } });
}
