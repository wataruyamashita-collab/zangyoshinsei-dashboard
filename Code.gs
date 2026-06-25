/**
TeamSpirit 残業申請・承認 取込管理
月次・週次対応版：CSVファイル選択 → 直近原本 → 取込データ更新 → 月次・週次集計 → HTMLダッシュボード更新

【運用方針】
- CSV貼付シートへの手動貼付は不要
- メニュー「TeamSpirit取込」→「ローカルCSVファイルを選択して取り込み」からCドライブ等のCSVを選択
- ブラウザ側でCSVを読み取り、Apps Script側にCSVテキストを渡して集計
- Apps ScriptはユーザーPCのCドライブを直接読むのではなく、ユーザーが選択したファイルのみを処理する
- 取込データ_残業申請に取込キー単位で追加・更新し、月次・週次の両方に対応する
*/

const TS_CONFIG = {
  TIMEZONE: 'Asia/Tokyo',
  MAX_ROWS: 30000,
  DEFAULT_CLOSING_TIME: '17:30',
  STAFF_CATEGORY: 'スタッフ部門',
  SALES_CATEGORY: '営業拠点',
  UNCLASSIFIED_CATEGORY: '未分類',
  LOCK_TIMEOUT_MS: 30000,
  SHEETS: {
    RAW: '原本_残業申請',
    ACCUM: '取込データ_残業申請',
    SETTINGS: '設定',
    DEPT_MASTER: '部署マスタ',
    IMPORT_LOG: '取込ログ',
    ERRORS: 'エラー一覧',
    SUMMARY_DEPT: '集計_部署別',
    SUMMARY_PERSON: '集計_個人別',
    SUMMARY_APPROVER: '集計_承認者別',
    SUMMARY_WEEK_DEPT: '集計_週次_部署別',
    SUMMARY_WEEK_PERSON: '集計_週次_個人別',
    SUMMARY_WEEK_APPROVER: '集計_週次_承認者別',
    DASHBOARD_STAFF: 'ダッシュボード_スタッフ部門',
    DASHBOARD_SALES: 'ダッシュボード_営業部門参考',
    DASHBOARD_WEEK_STAFF: 'ダッシュボード_週次_スタッフ部門',
    DASHBOARD_WEEK_SALES: 'ダッシュボード_週次_営業部門参考'
  },
  REQUIRED_HEADERS: [
    '部署名',
    '社員コード',
    '社員名',
    '日付',
    '曜日',
    '残業申請:申請対象社員コード',
    '残業申請:申請対象社員名',
    '残業申請:申請種類',
    '残業申請:ステータス',
    '残業申請:申請日時',
    '残業申請:申請者',
    '残業申請:承認日時',
    '残業申請:承認者',
    '残業申請:承認却下コメント',
    '残業申請:内容',
    '残業申請:備考'
  ],
  HELPER_HEADERS: [
    '対象年月',
    '対象週',
    '週開始日',
    '週終了日',
    '集計対象',
    '定時基準',
    '申請日時_DT',
    '承認日時_DT',
    '定時前申請',
    '定時前承認',
    '翌日承認',
    '部門区分',
    '取込キー'
  ]
};

/**
HTML_UPLOAD_DIALOG is embedded to avoid HtmlService file lookups.
*/
const HTML_UPLOAD_DIALOG = String.raw`<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <meta charset="UTF-8">
  <style>
    body {
      font-family: Arial, "Yu Gothic", "Meiryo", sans-serif;
      font-size: 14px;
      color: #202124;
      padding: 18px;
      line-height: 1.6;
    }
    h2 {
      font-size: 18px;
      margin: 0 0 12px;
      color: #17466f;
    }
    .box {
      border: 1px solid #dadce0;
      border-radius: 8px;
      padding: 14px;
      margin-top: 12px;
      background: #fff;
    }
    .note {
      color: #5f6368;
      font-size: 13px;
      margin-top: 8px;
    }
    label {
      display: block;
      font-weight: bold;
      margin-top: 12px;
      margin-bottom: 4px;
    }
    input[type="file"], select {
      width: 100%;
      box-sizing: border-box;
      padding: 8px;
      border: 1px solid #dadce0;
      border-radius: 4px;
      background: #fff;
    }
    button {
      margin-top: 16px;
      padding: 9px 16px;
      border: none;
      border-radius: 4px;
      cursor: pointer;
      font-weight: bold;
    }
    .primary {
      background: #17466f;
      color: #fff;
    }
    .secondary {
      background: #f1f3f4;
      color: #202124;
      margin-left: 8px;
    }
    .status {
      margin-top: 14px;
      padding: 10px;
      border-radius: 6px;
      background: #f8f9fa;
      white-space: pre-wrap;
      min-height: 42px;
    }
    .error {
      background: #fce8e6;
      color: #b3261e;
    }
    .success {
      background: #e6f4ea;
      color: #137333;
    }
  </style>
</head>
<body>
  <h2>TeamSpirit CSVファイル取り込み</h2>

  <div class="box">
    <div>
      TeamSpiritから出力したCSVファイルを選択してください。<br>
      取り込んだデータは申請単位で保持・更新され、月次・週次の両方の集計に反映されます。
    </div>

    <label for="csvFile">CSVファイル</label>
    <input type="file" id="csvFile" accept=".csv,text/csv">

    <label for="encoding">文字コード</label>
    <select id="encoding">
      <option value="auto" selected>自動判定</option>
      <option value="utf-8">UTF-8</option>
      <option value="shift_jis">Shift_JIS / CP932</option>
    </select>

    <div class="note">
      ※Apps ScriptがCドライブを直接読むのではなく、選択したファイルの内容だけをブラウザ経由で取り込みます。<br>
      ※同じ申請は申請キーで更新されるため、毎日取り込んでも二重計上されにくい設計です。
    </div>

    <button class="primary" onclick="startImport()">取り込み実行</button>
    <button class="secondary" onclick="google.script.host.close()">閉じる</button>

    <div id="status" class="status">CSVファイルを選択して「取り込み実行」を押してください。</div>
  </div>

  <script>

    const REQUIRED_HEADERS = [
      '部署名',
      '社員コード',
      '社員名',
      '日付',
      '曜日',
      '残業申請:申請対象社員コード',
      '残業申請:申請対象社員名',
      '残業申請:申請種類',
      '残業申請:ステータス',
      '残業申請:申請日時',
      '残業申請:申請者',
      '残業申請:承認日時',
      '残業申請:承認者',
      '残業申請:承認却下コメント',
      '残業申請:内容',
      '残業申請:備考'
    ];
    function startImport() {
      const fileInput = document.getElementById('csvFile');
      const encoding = document.getElementById('encoding').value;
      const status = document.getElementById('status');

      status.className = 'status';
      status.textContent = '';

      if (!fileInput.files || fileInput.files.length === 0) {
        status.className = 'status error';
        status.textContent = 'CSVファイルを選択してください。';
        return;
      }

      const file = fileInput.files[0];

      if (!/\.csv$/i.test(file.name)) {
        status.className = 'status error';
        status.textContent = 'CSVファイルを選択してください。';
        return;
      }

      status.textContent = 'CSVファイルを読み込んでいます...';

      const reader = new FileReader();

      reader.onload = function(event) {
        try {
          const buffer = event.target.result;
          const decoded = decodeCsvBuffer(buffer, encoding);

          if (!decoded.text || decoded.text.trim() === '') {
            throw new Error('CSVファイルの内容が空です。');
          }

          status.textContent = '取り込み・データ更新・月次週次集計を実行しています...';

          google.script.run
            .withSuccessHandler(function(result) {
              status.className = 'status success';
              status.textContent =
                '取り込みが完了しました。\n\n' +
                'ファイル名：' + result.fileName + '\n' +
                '対象年月：' + result.targetMonth + '\n' +
                '対象週：' + result.targetWeek + '（' + result.weekStart + '〜' + result.weekEnd + '）\n\n' +
                '新規追加：' + result.added + '件\n' +
                '更新：' + result.updated + '件\n' +
                '取込対象外：' + result.skipped + '件\n\n' +
                '月次対象件数：' + result.monthlyImportCount + '件\n' +
                '月次スタッフ部門：' + result.monthlyStaffCount + '件\n' +
                '月次営業参考：' + result.monthlySalesCount + '件\n\n' +
                '週次対象件数：' + result.weeklyImportCount + '件\n' +
                '週次スタッフ部門：' + result.weeklyStaffCount + '件\n' +
                '週次営業参考：' + result.weeklySalesCount + '件\n\n' +
                'エラー・確認件数：' + result.errorCount + '件\n' +
                '文字コード：' + result.encoding + '\n\n' +
                'メニュー「TeamSpirit取込」→「HTMLダッシュボードを開く」から、月次・週次の結果を確認してください。';
            })
            .withFailureHandler(function(error) {
              status.className = 'status error';
              status.textContent = '取り込みでエラーが発生しました。\n\n' + (error && error.message ? error.message : error);
            })
            .importLocalCsvText({
              fileName: file.name,
              fileSize: file.size,
              encoding: decoded.encoding,
              csvText: decoded.text
            });
        } catch (error) {
          status.className = 'status error';
          status.textContent = 'CSV読込でエラーが発生しました。\n\n' + error.message;
        }
      };

      reader.onerror = function() {
        status.className = 'status error';
        status.textContent = 'ファイルを読み込めませんでした。';
      };

      reader.readAsArrayBuffer(file);
    }

    function decodeCsvBuffer(buffer, encoding) {
      if (encoding === 'utf-8') {
        return {
          text: decodeWith('utf-8', buffer),
          encoding: 'UTF-8'
        };
      }

      if (encoding === 'shift_jis') {
        return {
          text: decodeWith('shift_jis', buffer),
          encoding: 'Shift_JIS'
        };
      }

      const candidates = [
        { text: decodeWith('utf-8', buffer), encoding: 'UTF-8(auto)' },
        { text: decodeWith('shift_jis', buffer), encoding: 'Shift_JIS(auto)' }
      ].map(function(candidate) {
        return Object.assign(candidate, { quality: inspectCsvDecodeQuality(candidate.text) });
      });

      candidates.sort(compareDecodeCandidates);

      return {
        text: candidates[0].text,
        encoding: candidates[0].quality.requiredHeaderMatches >= REQUIRED_HEADERS.length
          ? candidates[0].encoding
          : candidates[0].encoding + ' fallback'
      };
    }

    function decodeWith(label, buffer) {
      try {
        return new TextDecoder(label, { fatal: false }).decode(buffer);
      } catch (error) {
        if (label === 'shift_jis') {
          return new TextDecoder('utf-8', { fatal: false }).decode(buffer);
        }
        throw error;
      }
    }

    function inspectCsvDecodeQuality(text) {
      const headers = parseCsvHeaderLine(text).map(normalizeHeader);
      return {
        requiredHeaderMatches: REQUIRED_HEADERS
          .map(normalizeHeader)
          .filter(function(required) { return headers.indexOf(required) !== -1; })
          .length,
        looksLikeTeamSpiritCsv: looksLikeTeamSpiritCsv(text),
        replacementChars: countReplacementChars(text)
      };
    }

    function compareDecodeCandidates(a, b) {
      if (a.quality.requiredHeaderMatches !== b.quality.requiredHeaderMatches) {
        return b.quality.requiredHeaderMatches - a.quality.requiredHeaderMatches;
      }
      if (a.quality.looksLikeTeamSpiritCsv !== b.quality.looksLikeTeamSpiritCsv) {
        return a.quality.looksLikeTeamSpiritCsv ? -1 : 1;
      }
      return a.quality.replacementChars - b.quality.replacementChars;
    }

    function parseCsvHeaderLine(text) {
      const firstLine = String(text || '').replace(/^\uFEFF/, '').split(/\r\n|\n|\r/)[0] || '';
      const cells = [];
      let current = '';
      let inQuotes = false;

      for (let i = 0; i < firstLine.length; i++) {
        const ch = firstLine.charAt(i);
        const next = firstLine.charAt(i + 1);

        if (ch === '"' && inQuotes && next === '"') {
          current += '"';
          i++;
        } else if (ch === '"') {
          inQuotes = !inQuotes;
        } else if (ch === ',' && !inQuotes) {
          cells.push(current);
          current = '';
        } else {
          current += ch;
        }
      }

      cells.push(current);
      return cells;
    }

    function normalizeHeader(header) {
      return String(header || '')
        .replace(/\s/g, '')
        .replace(/：/g, ':')
        .trim();
    }

    function looksLikeTeamSpiritCsv(text) {
      const sample = String(text || '').slice(0, 3000);
      return sample.indexOf('部署名') !== -1 &&
             sample.indexOf('残業申請') !== -1 &&
             sample.indexOf('申請日時') !== -1;
    }

    function countReplacementChars(text) {
      const matched = String(text || '').match(/\uFFFD/g);
      return matched ? matched.length : 0;
    }
  </script>
</body>
</html>
`;

/**
HTML_DASHBOARD is embedded to avoid HtmlService file lookups.
*/
const HTML_DASHBOARD = String.raw`<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <meta charset="UTF-8">
  <style>
    :root {
      --blue: #17466f;
      --blue-dark: #0f3554;
      --bg: #f5f7fb;
      --surface: #ffffff;
      --surface-soft: #f9fbfd;
      --line: #d7dee8;
      --line-soft: #edf1f7;
      --text: #1f2937;
      --sub: #667085;
      --good: #137333;
      --bad: #b42318;
      --warn: #b26a00;
      --neutral: #64748b;
      --chip: #eaf1f8;
    }

    * {
      box-sizing: border-box;
    }

    body {
      margin: 0;
      font-family: Arial, "Yu Gothic", "Meiryo", sans-serif;
      background: var(--bg);
      color: var(--text);
      font-size: 14px;
    }

    .app-header {
      background: linear-gradient(135deg, var(--blue), var(--blue-dark));
      color: white;
      padding: 18px 24px;
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 16px;
    }

    .app-title {
      font-size: 20px;
      font-weight: 800;
      letter-spacing: .01em;
    }

    .app-subtitle {
      font-size: 12px;
      opacity: .9;
      margin-top: 4px;
    }

    .actions {
      display: flex;
      gap: 8px;
      align-items: center;
    }

    .actions button {
      border: 1px solid rgba(255,255,255,.45);
      background: rgba(255,255,255,.12);
      color: white;
      border-radius: 8px;
      padding: 8px 12px;
      cursor: pointer;
      font-weight: 700;
    }

    .actions button:hover {
      background: rgba(255,255,255,.22);
    }

    .wrap {
      padding: 18px 24px 28px;
    }

    .status,
    .error-box {
      background: var(--surface);
      border: 1px solid var(--line);
      border-radius: 14px;
      padding: 16px;
      margin-bottom: 14px;
      color: var(--sub);
      white-space: pre-wrap;
    }

    .error-box {
      background: #fce8e6;
      color: var(--bad);
      border-color: #f3b8b2;
    }

    .meta {
      display: grid;
      grid-template-columns: repeat(4, minmax(140px, 1fr));
      gap: 10px;
      margin-bottom: 14px;
    }

    .meta-card {
      background: var(--surface);
      border: 1px solid var(--line);
      border-radius: 12px;
      padding: 11px 12px;
      font-size: 12px;
      color: var(--sub);
      min-height: 58px;
    }

    .meta-card b {
      display: block;
      color: var(--blue);
      font-size: 13px;
      margin-top: 4px;
      word-break: break-word;
    }

    .tabs {
      display: flex;
      gap: 8px;
      margin-bottom: 0;
      flex-wrap: wrap;
    }

    .tab {
      border: 1px solid var(--line);
      background: var(--chip);
      border-radius: 12px 12px 0 0;
      padding: 11px 22px;
      cursor: pointer;
      font-weight: 800;
      color: var(--blue);
      min-width: 190px;
      text-align: center;
    }

    .tab.active {
      background: var(--surface);
      border-bottom-color: var(--surface);
      box-shadow: 0 -1px 2px rgba(16,24,40,.04);
    }

    .section {
      display: none;
    }

    .section.active {
      display: block;
    }

    .department-shell {
      background: var(--surface);
      border: 1px solid var(--line);
      border-radius: 0 14px 14px 14px;
      padding: 16px;
      box-shadow: 0 1px 2px rgba(16,24,40,.04);
      margin-top: -1px;
    }

    .department-head {
      display: flex;
      justify-content: space-between;
      align-items: baseline;
      gap: 12px;
      margin-bottom: 14px;
    }

    .department-title {
      font-size: 18px;
      font-weight: 900;
      color: var(--blue);
    }

    .department-period {
      font-size: 12px;
      color: var(--sub);
    }

    .period-main {
      border: 1px solid var(--line);
      border-radius: 14px;
      background: #fff;
      padding: 16px;
      margin-bottom: 16px;
    }

    .period-sub {
      border: 1px solid var(--line);
      border-radius: 14px;
      background: var(--surface-soft);
      padding: 14px;
      margin-top: 16px;
    }

    .period-header {
      display: flex;
      justify-content: space-between;
      align-items: baseline;
      gap: 12px;
      margin-bottom: 12px;
    }

    .period-title {
      font-size: 17px;
      font-weight: 900;
      color: var(--blue);
    }

    .period-title small {
      font-size: 12px;
      color: var(--sub);
      font-weight: 700;
      margin-left: 8px;
    }

    .period-info {
      color: var(--sub);
      font-size: 12px;
    }

    .kpi-grid {
      display: grid;
      grid-template-columns: repeat(3, minmax(180px, 1fr));
      gap: 12px;
      margin-bottom: 12px;
    }

    .support-grid {
      display: grid;
      grid-template-columns: repeat(3, minmax(180px, 1fr));
      gap: 12px;
      margin-bottom: 12px;
    }

    .kpi-card {
      background: #fbfdff;
      border: 1px solid #e4ebf3;
      border-radius: 14px;
      padding: 14px;
      min-height: 108px;
    }

    .kpi-card.primary {
      background: linear-gradient(180deg, #ffffff, #f7fbff);
      border-color: #d7e5f5;
    }

    .kpi-label {
      margin: 0 0 8px;
      font-size: 13px;
      color: var(--sub);
      font-weight: 800;
    }

    .kpi-value {
      font-size: 30px;
      font-weight: 900;
      letter-spacing: -.02em;
    }

    .kpi-delta {
      font-size: 12px;
      margin-top: 7px;
      font-weight: 800;
    }

    .good { color: var(--good); }
    .bad { color: var(--bad); }
    .neutral { color: var(--neutral); }

    .attention {
      background: #fff8ed;
      border: 1px solid #ffe0ad;
      border-radius: 14px;
      padding: 12px;
      margin: 12px 0;
    }

    .attention-title {
      font-size: 13px;
      font-weight: 900;
      color: var(--warn);
      margin-bottom: 8px;
    }

    .attention-list {
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
    }

    .attention-chip {
      display: inline-flex;
      gap: 6px;
      align-items: center;
      border-radius: 999px;
      background: #fff;
      border: 1px solid #f5c36d;
      padding: 5px 9px;
      color: #8a4b00;
      font-size: 12px;
      font-weight: 800;
    }

    .table-card {
      border: 1px solid var(--line-soft);
      border-radius: 12px;
      background: #fff;
      overflow: hidden;
      margin-top: 12px;
    }

    .table-head {
      display: flex;
      justify-content: space-between;
      align-items: center;
      padding: 10px 12px;
      border-bottom: 1px solid var(--line-soft);
      background: #fbfdff;
    }

    .table-title {
      font-size: 14px;
      font-weight: 900;
      color: var(--blue);
    }

    .table-caption {
      font-size: 12px;
      color: var(--sub);
    }

    .table-wrap {
      max-height: 420px;
      overflow: auto;
    }

    table {
      width: 100%;
      border-collapse: collapse;
      background: white;
      font-size: 13px;
    }

    th, td {
      border-bottom: 1px solid #edf1f7;
      padding: 9px 10px;
      text-align: right;
      white-space: nowrap;
      vertical-align: middle;
    }

    th {
      background: #eef4fb;
      color: #17324d;
      font-weight: 900;
      position: sticky;
      top: 0;
      z-index: 1;
    }

    th:first-child, td:first-child {
      text-align: left;
      min-width: 220px;
      position: sticky;
      left: 0;
      background: inherit;
      z-index: 0;
    }

    th:first-child {
      background: #eef4fb;
      z-index: 2;
    }

    td:first-child {
      background: white;
      font-weight: 800;
    }

    .metric-main {
      font-weight: 900;
      display: block;
    }

    .metric-sub {
      font-size: 11px;
      font-weight: 800;
      display: block;
      margin-top: 2px;
    }

    .alert {
      display: inline-block;
      border-radius: 999px;
      background: #fff4e5;
      color: var(--warn);
      padding: 3px 8px;
      font-size: 12px;
      font-weight: 900;
    }

    .ok {
      display: inline-block;
      border-radius: 999px;
      background: #e6f4ea;
      color: var(--good);
      padding: 3px 8px;
      font-size: 12px;
      font-weight: 900;
    }

    .note {
      color: var(--sub);
      font-size: 12px;
      line-height: 1.6;
      margin-top: 12px;
      background: white;
      border: 1px solid var(--line);
      border-radius: 12px;
      padding: 12px;
    }

    @media print {
      .actions, .tabs {
        display: none !important;
      }

      .wrap {
        padding: 10px;
      }

      .section {
        display: none !important;
      }

      .section.active {
        display: block !important;
      }

      .department-shell,
      .period-main,
      .period-sub,
      .table-card {
        box-shadow: none;
      }

      .table-wrap {
        max-height: none;
        overflow: visible;
      }

      body {
        background: white;
      }
    }
  </style>
</head>
<body>
  <div class="app-header">
    <div>
      <div class="app-title">TeamSpirit 残業申請・承認ダッシュボード</div>
      <div class="app-subtitle">残業申請・承認の事前運用状況を月次・週次で確認します。</div>
    </div>
    <div class="actions">
      <button onclick="loadData()">再読み込み</button>
      <button onclick="window.print()">表示中のタブを印刷</button>
      <button onclick="google.script.host.close()">閉じる</button>
    </div>
  </div>

  <div class="wrap">
    <div id="message" class="status">ダッシュボードを読み込んでいます...</div>

    <div id="content" style="display:none;">
      <div class="meta">
        <div class="meta-card">対象年月<b id="targetMonth"></b></div>
        <div class="meta-card">対象週<b id="targetWeek"></b></div>
        <div class="meta-card">最終取込日時<b id="lastImportTime"></b></div>
        <div class="meta-card">エラー・確認件数<b id="errorCount"></b></div>
      </div>

      <div class="tabs">
        <button id="tabStaff" class="tab active" onclick="switchTab('staff')">スタッフ部門</button>
        <button id="tabSales" class="tab" onclick="switchTab('sales')">営業部門（参考）</button>
      </div>

      <div id="staff" class="section active"></div>
      <div id="sales" class="section"></div>

      <div class="note" id="dashboardNote"></div>
    </div>
  </div>

  <script>
    const Direction = {
      HIGHER_IS_GOOD: 'higher',
      LOWER_IS_GOOD: 'lower',
      NEUTRAL: 'neutral'
    };

    let dashboardData = null;
    let dashboardLoadTimer = null;

    function loadData() {
      const message = document.getElementById('message');
      const content = document.getElementById('content');

      message.className = 'status';
      message.textContent = 'ダッシュボードを読み込んでいます...';
      message.style.display = 'block';
      content.style.display = 'none';

      if (typeof google === 'undefined' || !google.script || !google.script.run) {
        message.className = 'error-box';
        message.textContent = 'このHTMLはブラウザで直接開けません。スプレッドシートのメニュー「TeamSpirit取込」→「HTMLダッシュボードを開く」から開いてください。';
        return;
      }

      if (dashboardLoadTimer) {
        clearTimeout(dashboardLoadTimer);
      }
      dashboardLoadTimer = setTimeout(function() {
        message.className = 'error-box';
        message.textContent = 'ダッシュボードデータの取得に時間がかかっています。\n\n' +
          '考えられる原因：取込データ件数が多い、シートの読込範囲が広い、またはサーバー側処理が停止しています。\n' +
          'Apps Script画面の「実行数」で getHtmlDashboardData の結果を確認してください。';
      }, 45000);

      google.script.run
        .withSuccessHandler(function(data) {
          if (dashboardLoadTimer) {
            clearTimeout(dashboardLoadTimer);
            dashboardLoadTimer = null;
          }

          try {
            dashboardData = normalizeDashboardPayload(data);

            if (!dashboardData.ok) {
              message.className = 'error-box';
              message.textContent = dashboardData.message || 'ダッシュボードデータを取得できませんでした。';
              return;
            }

            render(dashboardData);
            message.style.display = 'none';
            content.style.display = 'block';
          } catch (error) {
            message.className = 'error-box';
            message.textContent = 'ダッシュボード描画時にエラーが発生しました。\n\n' +
              (error && error.message ? error.message : error) +
              '\n\nサーバーから返ったデータ構造とDashboard.html側の参照が一致していない可能性があります。';
          }
        })
        .withFailureHandler(function(error) {
          if (dashboardLoadTimer) {
            clearTimeout(dashboardLoadTimer);
            dashboardLoadTimer = null;
          }
          message.className = 'error-box';
          message.textContent = 'ダッシュボードの読み込みでエラーが発生しました。\n\n' + (error && error.message ? error.message : error);
        })
        .getHtmlDashboardData();
    }


    function normalizeDashboardPayload(data) {
      if (typeof data === 'string') {
        try {
          data = JSON.parse(data);
        } catch (error) {
          return {
            ok: false,
            message: 'サーバーから返ったデータを解析できませんでした。\n\n' +
              (error && error.message ? error.message : error)
          };
        }
      }

      if (!data || typeof data !== 'object') {
        return { ok: false, message: 'サーバーから有効なデータが返っていません。' };
      }

      data.settings = data.settings || {};
      data.errorCount = Number(data.errorCount || 0);
      data.monthly = data.monthly || {};
      data.weekly = data.weekly || {};
      data.monthly.staff = normalizeCategoryPayload(data.monthly.staff);
      data.monthly.sales = normalizeCategoryPayload(data.monthly.sales);
      data.weekly.staff = normalizeCategoryPayload(data.weekly.staff);
      data.weekly.sales = normalizeCategoryPayload(data.weekly.sales);
      data.monthly.label = data.monthly.label || '';
      data.monthly.previousLabel = data.monthly.previousLabel || '';
      data.weekly.label = data.weekly.label || '';
      data.weekly.previousLabel = data.weekly.previousLabel || '';
      return data;
    }

    function normalizeCategoryPayload(category) {
      const emptyTotal = {
        count: 0,
        beforeApplyRate: 0,
        beforeApproveRate: 0,
        nextDayApproveRate: 0,
        notApprovedCount: 0
      };
      return {
        category: category && category.category ? category.category : '',
        current: Object.assign({}, emptyTotal, category && category.current ? category.current : {}),
        previous: Object.assign({}, emptyTotal, category && category.previous ? category.previous : {}),
        diff: Object.assign({
          count: 0,
          countGrowth: 0,
          beforeApplyRate: 0,
          beforeApproveRate: 0,
          nextDayApproveRate: 0,
          notApprovedCount: 0
        }, category && category.diff ? category.diff : {}),
        details: Array.isArray(category && category.details) ? category.details : []
      };
    }

    function render(data) {
      document.getElementById('targetMonth').textContent = data.settings.targetMonth || '-';
      document.getElementById('targetWeek').textContent = data.settings.targetWeek || '-';
      document.getElementById('lastImportTime').textContent = data.settings.lastImportTime || '-';
      document.getElementById('errorCount').textContent = data.errorCount + '件';
      document.getElementById('dashboardNote').textContent = data.settings.dashboardNote || '';

      document.getElementById('staff').innerHTML = renderDepartment({
        title: 'スタッフ部門',
        monthly: data.monthly.staff,
        weekly: data.weekly.staff,
        monthlyLabel: data.monthly.label,
        previousMonthLabel: data.monthly.previousLabel,
        weeklyLabel: data.weekly.label,
        previousWeekLabel: data.weekly.previousLabel,
        isReference: false
      });

      document.getElementById('sales').innerHTML = renderDepartment({
        title: '営業部門（参考）',
        monthly: data.monthly.sales,
        weekly: data.weekly.sales,
        monthlyLabel: data.monthly.label,
        previousMonthLabel: data.monthly.previousLabel,
        weeklyLabel: data.weekly.label,
        previousWeekLabel: data.weekly.previousLabel,
        isReference: true
      });
    }

    function switchTab(name) {
      document.getElementById('tabStaff').classList.toggle('active', name === 'staff');
      document.getElementById('tabSales').classList.toggle('active', name === 'sales');
      document.getElementById('staff').classList.toggle('active', name === 'staff');
      document.getElementById('sales').classList.toggle('active', name === 'sales');
    }

    function renderDepartment(props) {
      const monthly = renderPeriod({
        importance: 'main',
        title: '月次',
        subtitle: props.isReference ? '営業部門は参考値です。' : '月次確認の中心指標です。',
        category: props.monthly,
        label: props.monthlyLabel,
        previousLabel: props.previousMonthLabel,
        compareLabel: '前月'
      });

      const weekly = renderPeriod({
        importance: 'sub',
        title: '週次',
        subtitle: '直近の運用状況を確認します。',
        category: props.weekly,
        label: props.weeklyLabel,
        previousLabel: props.previousWeekLabel,
        compareLabel: '前週'
      });

      return \`
        <div class="department-shell">
          <div class="department-head">
            <div class="department-title">\${escapeHtml(props.title)}</div>
            <div class="department-period">月次：\${escapeHtml(props.monthlyLabel || '-')} ／ 週次：\${escapeHtml(props.weeklyLabel || '-')}</div>
          </div>
          \${monthly}
          \${weekly}
        </div>
      \`;
    }

    function renderPeriod(props) {
      const className = props.importance === 'main' ? 'period-main' : 'period-sub';

      return \`
        <div class="\${className}">
          <div class="period-header">
            <div class="period-title">\${escapeHtml(props.title)}<small>\${escapeHtml(props.subtitle)}</small></div>
            <div class="period-info">対象：\${escapeHtml(props.label || '-')} ／ \${escapeHtml(props.compareLabel)}：\${escapeHtml(props.previousLabel || '-')}</div>
          </div>
          \${renderKpis(props.category)}
          \${renderAttention(props.category)}
          \${renderDetailTable(props.category, props.title + ' 部署別明細', props.compareLabel)}
        </div>
      \`;
    }

    function renderKpis(category) {
      const c = category.current;
      const d = category.diff;

      return \`
        <div class="kpi-grid">
          \${renderKpi('定時前申請率', pct(c.beforeApplyRate), pt(d.beforeApplyRate), d.beforeApplyRate, Direction.HIGHER_IS_GOOD, true)}
          \${renderKpi('定時前承認率', pct(c.beforeApproveRate), pt(d.beforeApproveRate), d.beforeApproveRate, Direction.HIGHER_IS_GOOD, true)}
          \${renderKpi('翌日承認率', pct(c.nextDayApproveRate), pt(d.nextDayApproveRate), d.nextDayApproveRate, Direction.HIGHER_IS_GOOD, true)}
        </div>
        <div class="support-grid">
          \${renderKpi('申請件数', c.count + '件', signedNum(d.count) + '件', d.count, Direction.NEUTRAL, false)}
          \${renderKpi('未承認件数', c.notApprovedCount + '件', signedNum(d.notApprovedCount) + '件', d.notApprovedCount, Direction.LOWER_IS_GOOD, false)}
          \${renderKpi('確認部署数', attentionDepartments(category).length + '部署', '', 0, Direction.NEUTRAL, false)}
        </div>
      \`;
    }

    function renderKpi(label, value, deltaText, deltaValue, direction, primary) {
      const color = scoreClass(deltaValue, direction);
      return \`
        <div class="kpi-card \${primary ? 'primary' : ''}">
          <div class="kpi-label">\${escapeHtml(label)}</div>
          <div class="kpi-value">\${escapeHtml(value)}</div>
          <div class="kpi-delta \${color}">\${deltaText ? '差分：' + escapeHtml(deltaText) : '—'}</div>
        </div>
      \`;
    }

    function renderAttention(category) {
      const targets = attentionDepartments(category);

      if (targets.length === 0) {
        return \`
          <div class="attention" style="background:#f0f8f2;border-color:#cce8d3;">
            <div class="attention-title" style="color:var(--good);">確認対象部署</div>
            <div class="attention-list">
              <span class="attention-chip" style="border-color:#b8dfc1;color:var(--good);">現時点で要確認部署はありません</span>
            </div>
          </div>
        \`;
      }

      return \`
        <div class="attention">
          <div class="attention-title">確認対象部署</div>
          <div class="attention-list">
            \${targets.map(row => \`
              <span class="attention-chip">\${escapeHtml(row.deptName)}：承認率 \${pct(row.beforeApproveRate)}／未承認 \${num(row.notApprovedCount)}件</span>
            \`).join('')}
          </div>
        </div>
      \`;
    }

    function attentionDepartments(category) {
      return (category.details || [])
        .filter(row => row.alert || Number(row.notApprovedCount || 0) > 0)
        .sort((a, b) => {
          if (Number(b.notApprovedCount || 0) !== Number(a.notApprovedCount || 0)) {
            return Number(b.notApprovedCount || 0) - Number(a.notApprovedCount || 0);
          }
          return Number(a.beforeApproveRate || 0) - Number(b.beforeApproveRate || 0);
        })
        .slice(0, 5);
    }

    function renderDetailTable(category, title, compareLabel) {
      const rows = category.details || [];

      const body = rows.length === 0
        ? \`<tr><td colspan="7">対象データがありません。</td></tr>\`
        : rows.map(row => \`
          <tr>
            <td>\${escapeHtml(row.deptName)}</td>
            <td>\${renderCountCell(row.count, row.countDiff)}</td>
            <td>\${renderRateCell(row.beforeApplyRate, row.beforeApplyRateDiff)}</td>
            <td>\${renderRateCell(row.beforeApproveRate, row.beforeApproveRateDiff)}</td>
            <td>\${renderRateCell(row.nextDayApproveRate, row.nextDayApproveRateDiff)}</td>
            <td>\${renderLowerGoodCountCell(row.notApprovedCount, row.notApprovedCountDiff)}</td>
            <td>\${row.alert ? '<span class="alert">' + escapeHtml(row.alert) + '</span>' : '<span class="ok">良好</span>'}</td>
          </tr>
        \`).join('');

      return \`
        <div class="table-card">
          <div class="table-head">
            <div class="table-title">\${escapeHtml(title)}</div>
            <div class="table-caption">差分は\${escapeHtml(compareLabel)}比</div>
          </div>
          <div class="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>部署名</th>
                  <th>申請件数</th>
                  <th>定時前申請率</th>
                  <th>定時前承認率</th>
                  <th>翌日承認率</th>
                  <th>未承認</th>
                  <th>判定</th>
                </tr>
              </thead>
              <tbody>\${body}</tbody>
            </table>
          </div>
        </div>
      \`;
    }

    function renderCountCell(value, diff) {
      return \`
        <span class="metric-main">\${num(value)}件</span>
        <span class="metric-sub neutral">\${signedNum(diff)}件</span>
      \`;
    }

    function renderRateCell(value, diff) {
      return \`
        <span class="metric-main">\${pct(value)}</span>
        <span class="metric-sub \${scoreClass(diff, Direction.HIGHER_IS_GOOD)}">\${pt(diff)}</span>
      \`;
    }

    function renderLowerGoodCountCell(value, diff) {
      const n = Number(value || 0);
      const klass = n > 0 ? 'bad' : 'good';
      return \`
        <span class="metric-main \${klass}">\${num(n)}件</span>
        <span class="metric-sub \${scoreClass(diff, Direction.LOWER_IS_GOOD)}">\${signedNum(diff)}件</span>
      \`;
    }

    function scoreClass(value, direction) {
      const n = Number(value || 0);

      if (direction === Direction.NEUTRAL || n === 0) {
        return 'neutral';
      }

      if (direction === Direction.HIGHER_IS_GOOD) {
        return n > 0 ? 'good' : 'bad';
      }

      if (direction === Direction.LOWER_IS_GOOD) {
        return n < 0 ? 'good' : 'bad';
      }

      return 'neutral';
    }

    function pct(value) {
      const n = Number(value || 0);
      return (n * 100).toFixed(1) + '%';
    }

    function pt(value) {
      const n = Number(value || 0) * 100;
      const sign = n > 0 ? '+' : '';
      return sign + n.toFixed(1) + 'pt';
    }

    function num(value) {
      return Number(value || 0).toLocaleString('ja-JP');
    }

    function signedNum(value) {
      const n = Number(value || 0);
      const sign = n > 0 ? '+' : '';
      return sign + n.toLocaleString('ja-JP');
    }

    function escapeHtml(value) {
      return String(value == null ? '' : value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
    }

    loadData();
  </script>
</body>
</html>
`;

// 部署判定で使う正規表現は一度だけコンパイルする。
const REGEX_SALES_DEPT = /営業所|出張所|エリア/;
const RAW_TEXT_COLUMN_COUNT = TS_CONFIG.REQUIRED_HEADERS.length;
const DANGEROUS_SHEET_TEXT_PREFIX = /^[=+\-@\t\r]/;

/**
スプレッドシートを開いたときにメニューを追加
*/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('TeamSpirit取込')
    .addItem('ローカルCSVファイルを選択して取り込み', 'showLocalCsvUploadDialog')
    .addItem('HTMLダッシュボードを開く', 'showHtmlDashboard')
    .addItem('月次・週次を再集計', 'rebuildMonthlyAndWeeklyFromAccum')
    .addItem('エラー一覧を更新', 'refreshErrorListFromRaw')
    .addSeparator()
    .addItem('対象月を確定にする', 'markCurrentMonthAsFinal')
    .addSeparator()
    .addItem('初期設定を確認', 'checkInitialSetup')
    .addToUi();
}

/**
ローカルCSV選択ダイアログを表示
*/
function showLocalCsvUploadDialog() {
  ensureSpreadsheetTimeZone_(SpreadsheetApp.getActiveSpreadsheet());
  ensureBaseSheets_(SpreadsheetApp.getActiveSpreadsheet());

  const html = HtmlService
    .createHtmlOutput(HTML_UPLOAD_DIALOG)
    .setWidth(620)
    .setHeight(430);

  SpreadsheetApp.getUi().showModalDialog(html, 'TeamSpirit CSVファイル取り込み');
}


/**
設定値を一括更新する。
Spreadsheet I/Oを減らすため、取込処理ではこちらを優先して使用する。
*/
function setSettingValuesBulk_(settingsObj) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TS_CONFIG.SHEETS.SETTINGS);

  if (!sheet) {
    throw new Error(`シートが見つかりません：${TS_CONFIG.SHEETS.SETTINGS}`);
  }

  const existingRange = sheet.getDataRange();
  let values = existingRange.getValues();

  if (!values || values.length === 0 || values[0].length === 0 || isBlank_(values[0][0])) {
    values = [['設定項目', '設定値', '備考']];
  }

  const columnCount = Math.max(values[0].length, 3);
  values = values.map(row => {
    const copy = row.slice();
    while (copy.length < columnCount) copy.push('');
    return copy;
  });

  const keyToRowIndex = {};
  values.forEach((row, index) => {
    const key = String(row[0] || '').trim();
    if (key) {
      keyToRowIndex[key] = index;
    }
  });

  Object.entries(settingsObj || {}).forEach(([key, value]) => {
    if (keyToRowIndex[key] !== undefined) {
      values[keyToRowIndex[key]][1] = value;
    } else {
      const row = new Array(columnCount).fill('');
      row[0] = key;
      row[1] = value;
      values.push(row);
    }
  });

  ensureSheetSize_(sheet, values.length, columnCount);
  sheet.getRange(1, 1, values.length, columnCount).setValues(values);
}

/**
NaN混入を防ぐ数値パース。
*/
function safeParseNumber_(value) {
  if (value === null || value === undefined || value === '') {
    return 0;
  }

  const n = Number(value);
  return isNaN(n) ? 0 : n;
}


/**
HTMLダイアログから渡されたCSVテキストを取り込む
*/
function importLocalCsvText(payload) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(TS_CONFIG.LOCK_TIMEOUT_MS)) {
    throw new Error('別の取込・再集計処理が実行中です。少し待ってから再実行してください。');
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureSpreadsheetTimeZone_(ss);
    ensureBaseSheets_(ss);

    if (!payload || isBlank_(payload.csvText)) {
      throw new Error('CSVファイルの内容が空です。TeamSpiritから出力したCSVファイルを選択してください。');
    }

    const fileName = payload.fileName || 'ローカルCSV';
    const encoding = payload.encoding || 'auto';
    const csvText = normalizeCsvText_(payload.csvText);

    const csvValues = Utilities.parseCsv(csvText);
    const csvInfo = getCsvDataInfoFromValues_(csvValues);
    validateCsvHeaders_(csvInfo.headers);

    const deptMaster = getDeptMaster_();
    const settings = getSettings_();
    const periods = detectTargetPeriods_(csvInfo.rows, csvInfo.headerMap);

    const targetMonth = periods.targetMonth || settings.targetMonth;
    const targetWeek = periods.targetWeek || settings.targetWeek;
    const now = new Date();

    const settingUpdates = {
      '取込区分': '速報',
      '最終取込日時': now
    };

    if (targetMonth) {
      settingUpdates['対象年月'] = targetMonth;
      settings.targetMonth = targetMonth;
    }
    if (targetWeek) {
      settingUpdates['対象週'] = targetWeek;
      settings.targetWeek = targetWeek;
    }
    if (periods.weekStart) {
      settingUpdates['週開始日'] = periods.weekStart;
      settings.weekStart = periods.weekStart;
    }
    if (periods.weekEnd) {
      settingUpdates['週終了日'] = periods.weekEnd;
      settings.weekEnd = periods.weekEnd;
    }

    setSettingValuesBulk_(settingUpdates);

    settings.importType = '速報';
    settings.lastImportTime = now;

    const processed = buildProcessedRows_(csvInfo.rows, csvInfo.headerMap, deptMaster, targetMonth, settings);
    writeRawSheet_(ss.getSheetByName(TS_CONFIG.SHEETS.RAW), processed);

    const upsertResult = upsertAccumulatedRows_(processed.headers, processed.dataRows);
    const accumInfo = getAccumulatedDataInfo_();

    const errors = buildErrorRows_(csvInfo.rows, csvInfo.headerMap, deptMaster, targetMonth);
    writeErrorSheet_(errors, targetMonth);

    const summaries = buildMonthlyWeeklySummaryBundle_(accumInfo.rows, accumInfo.headers, settings);
    writeAllSummarySheets_(summaries);
    writeAllDashboards_(summaries, settings);

    appendImportLog_({
      importType: 'ローカルCSV取込',
      targetMonth: targetMonth,
      targetWeek: targetWeek,
      fileName: fileName,
      importMethod: 'HTMLファイル選択',
      importCount: summaries.monthly.current.totalImportCount,
      result: errors.length === 0 ? '成功' : '確認あり',
      memo: `文字コード：${encoding}／追加：${upsertResult.added}件／更新：${upsertResult.updated}件／除外：${upsertResult.skipped}件／エラー・確認件数：${errors.length}`
    });

    SpreadsheetApp.flush();

    return {
      ok: true,
      targetMonth: targetMonth,
      targetWeek: targetWeek,
      weekStart: periods.weekStart ? formatDateForDisplay_(periods.weekStart) : '',
      weekEnd: periods.weekEnd ? formatDateForDisplay_(periods.weekEnd) : '',
      monthlyImportCount: summaries.monthly.current.totalImportCount,
      weeklyImportCount: summaries.weekly.current.totalImportCount,
      errorCount: errors.length,
      fileName: fileName,
      encoding: encoding,
      added: upsertResult.added,
      updated: upsertResult.updated,
      skipped: upsertResult.skipped,
      monthlyStaffCount: getCategoryTotalCount_(summaries.monthly.current.deptRows, TS_CONFIG.STAFF_CATEGORY),
      weeklyStaffCount: getCategoryTotalCount_(summaries.weekly.current.deptRows, TS_CONFIG.STAFF_CATEGORY),
      monthlySalesCount: getCategoryTotalCount_(summaries.monthly.current.deptRows, TS_CONFIG.SALES_CATEGORY),
      weeklySalesCount: getCategoryTotalCount_(summaries.weekly.current.deptRows, TS_CONFIG.SALES_CATEGORY)
    };
  } catch (error) {
    appendImportLog_({
      importType: 'ローカルCSV取込',
      targetMonth: getSettingValue_('対象年月'),
      targetWeek: getSettingValue_('対象週'),
      fileName: payload && payload.fileName ? payload.fileName : '',
      importMethod: 'HTMLファイル選択',
      importCount: 0,
      result: '失敗',
      memo: error.message
    });
    throw error;
  } finally {
    lock.releaseLock();
  }
}

/**
原本_残業申請 から再集計する
*/
function rebuildMonthlyAndWeeklyFromAccum() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(TS_CONFIG.LOCK_TIMEOUT_MS)) {
    SpreadsheetApp.getUi().alert('別の取込・再集計処理が実行中です。少し待ってから再実行してください。');
    return;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureSpreadsheetTimeZone_(ss);
    ensureBaseSheets_(ss);

    const accumInfo = getAccumulatedDataInfo_();

    if (accumInfo.rows.length === 0) {
      SpreadsheetApp.getUi().alert('取込データ_残業申請にデータがありません。先にCSVファイルを取り込んでください。');
      return;
    }

    const settings = getSettings_();
    const summaries = buildMonthlyWeeklySummaryBundle_(accumInfo.rows, accumInfo.headers, settings);

    writeAllSummarySheets_(summaries);
    writeAllDashboards_(summaries, settings);

    appendImportLog_({
      importType: '月次・週次再集計',
      targetMonth: settings.targetMonth,
      targetWeek: settings.targetWeek,
      fileName: '取込データ_残業申請',
      importMethod: '取込データ再集計',
      importCount: summaries.monthly.current.totalImportCount,
      result: '成功',
      memo: '取込データ_残業申請から月次・週次を再集計'
    });

    SpreadsheetApp.getUi().alert(
      '月次・週次の再集計が完了しました。\n\n' +
      `対象年月：${settings.targetMonth}\n` +
      `対象週：${settings.targetWeek}\n` +
      `月次対象件数：${summaries.monthly.current.totalImportCount}件\n` +
      `週次対象件数：${summaries.weekly.current.totalImportCount}件`
    );
  } finally {
    lock.releaseLock();
  }
}

function rebuildFromRaw() {
  rebuildMonthlyAndWeeklyFromAccum();
}

/**
原本からエラー一覧を更新する
*/
function refreshErrorListFromRaw() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSpreadsheetTimeZone_(ss);
  ensureBaseSheets_(ss);

  const rawSheet = ss.getSheetByName(TS_CONFIG.SHEETS.RAW);
  const values = rawSheet.getDataRange().getValues();

  if (values.length < 2) {
    SpreadsheetApp.getUi().alert('原本_残業申請にデータがありません。先にCSVファイルを取り込んでください。');
    return;
  }

  const headers = values[0].map(v => String(v).trim());
  const headerMap = buildHeaderMap_(headers);
  const rows = values.slice(1).filter(row => row.some(v => !isBlank_(v)));
  const deptMaster = getDeptMaster_();
  const targetMonth = getSettingValue_('対象年月');

  const errors = buildErrorRows_(rows, headerMap, deptMaster, targetMonth);
  writeErrorSheet_(errors, targetMonth);

  SpreadsheetApp.getUi().alert(`エラー一覧を更新しました。\n\nエラー・確認件数：${errors.length}件`);
}

/**
対象月を確定にする
*/
function markCurrentMonthAsFinal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSpreadsheetTimeZone_(ss);
  ensureBaseSheets_(ss);

  const currentMonth = getSettingValue_('対象年月');
  const now = new Date();

  setSettingValue_('取込区分', '確定');
  setSettingValue_('最終取込日時', now);

  appendImportLog_({
    importType: '月次確定',
    targetMonth: currentMonth,
    fileName: '',
    importMethod: '手動確定',
    importCount: '',
    result: '確定',
    memo: '対象月を確定区分に変更'
  });

  SpreadsheetApp.getUi().alert(`対象月 ${currentMonth} を「確定」に変更しました。`);
}

/**
初期設定確認
*/
function checkInitialSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSpreadsheetTimeZone_(ss);
  ensureBaseSheets_(ss);
  ensureDefaultSettings_();

  const targetMonth = getSettingValue_('対象年月');
  const targetWeek = getSettingValue_('対象週');
  const closingTime = getSettingValue_('定時時刻');
  const importType = getSettingValue_('取込区分');
  const timeZone = ss.getSpreadsheetTimeZone();

  SpreadsheetApp.getUi().alert(
    '初期設定は確認できました。\n\n' +
    `対象年月：${targetMonth}\n` +
    `対象週：${targetWeek}\n` +
    `取込区分：${importType}\n` +
    `定時時刻：${closingTime}\n` +
    `タイムゾーン：${timeZone}`
  );
}

/**
必要シートを作成
*/
function ensureBaseSheets_(ss) {
  Object.keys(TS_CONFIG.SHEETS).forEach(key => {
    const name = TS_CONFIG.SHEETS[key];
    if (!ss.getSheetByName(name)) {
      ss.insertSheet(name);
    }
  });

  ensureDefaultSettings_();
  ensureDeptMasterHeader_();
  ensureAccumHeader_();
}

/**
設定シートの初期値
*/
function ensureDefaultSettings_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TS_CONFIG.SHEETS.SETTINGS);
  const existing = sheet.getDataRange().getValues();
  const keys = existing.map(row => String(row[0] || '').trim());

  const defaults = [
    ['設定項目', '設定値', '備考'],
    ['対象年月', Utilities.formatDate(new Date(), TS_CONFIG.TIMEZONE, 'yyyy-MM'), 'yyyy-MM形式。CSV取込時に自動判定します。'],
    ['対象週', getWeekInfo_(new Date()).weekKey, 'yyyy/mm/dd週形式。CSV取込時に自動判定します。'],
    ['週開始日', getWeekInfo_(new Date()).weekStart, '月曜始まり'],
    ['週終了日', getWeekInfo_(new Date()).weekEnd, '日曜終わり'],
    ['取込区分', '速報', '速報／確定'],
    ['定時時刻', TS_CONFIG.DEFAULT_CLOSING_TIME, '部署マスタに定時時刻がない場合の既定値'],
    ['営業拠点参考表示', true, 'TRUEなら営業部門も参考表示'],
    ['翌日承認判定', '暦日翌日23:59', '対象日の翌日23:59:59まで'],
    ['承認率注意基準', 0.8, '定時前承認率がこの値未満の場合、要確認'],
    ['最終取込日時', '', '取込完了時刻'],
    ['ダッシュボード注記', 'TeamSpirit上の残業申請・承認データをもとに、事前申請・事前承認の運用状況を可視化するものです。\n勤怠締め前のデータは速報値であり、申請・承認状況の更新により数値が変動する場合があります。', '表示用注記']
  ];

  if (existing.length === 1 && isBlank_(existing[0][0])) {
    sheet.clearContents();
  }

  if (sheet.getLastRow() === 0) {
    ensureSheetSize_(sheet, defaults.length, defaults[0].length);
    sheet.getRange(1, 1, defaults.length, defaults[0].length).setValues(defaults);
    formatHeaderRow_(sheet, 1, 3);
    sheet.setColumnWidths(1, 3, 120);
    return;
  }

  defaults.forEach((row, index) => {
    if (index === 0) {
      if (keys[0] !== '設定項目') {
        sheet.getRange(1, 1, 1, 3).setValues([defaults[0]]);
        formatHeaderRow_(sheet, 1, 3);
      }
      return;
    }

    if (!keys.includes(row[0])) {
      sheet.appendRow(row);
    }
  });
}

/**
部署マスタのヘッダー
*/
function ensureDeptMasterHeader_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TS_CONFIG.SHEETS.DEPT_MASTER);

  if (sheet.getLastRow() === 0 || isBlank_(sheet.getRange(1, 1).getValue())) {
    sheet.clearContents();

    const header = [['部署名', '部門区分', '集計区分', '定時時刻', '備考']];
    const defaults = [
      ['【TJ】管理課', 'スタッフ部門', '本集計', '17:30', ''],
      ['【TJ】業務支援課', 'スタッフ部門', '本集計', '17:30', ''],
      ['【TJ】ｸﾗｳﾄﾞ推進課', 'スタッフ部門', '本集計', '17:30', ''],
      ['【TJ】ｶｽﾀﾏｰｻｸｾｽ推進課', 'スタッフ部門', '本集計', '17:30', '']
    ];

    const output = header.concat(defaults);
    ensureSheetSize_(sheet, output.length, output[0].length);
    sheet.getRange(1, 1, output.length, output[0].length).setValues(output);
    formatHeaderRow_(sheet, 1, 5);
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, 5, 120);
  }
}

/**
CSVテキストを正規化
*/
function normalizeCsvText_(text) {
  return String(text || '')
    .replace(/^\uFEFF/, '')
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n');
}

/**
CSV配列からデータ情報を作成
*/
function getCsvDataInfoFromValues_(values) {
  if (!values || values.length < 2) {
    throw new Error('CSVにデータ行がありません。TeamSpiritから出力したCSVファイルを確認してください。');
  }

  if (values.length > TS_CONFIG.MAX_ROWS) {
    throw new Error(`CSVの行数が上限 ${TS_CONFIG.MAX_ROWS} 行を超えています。対象月を絞って出力してください。`);
  }

  const headers = values[0].slice(0, TS_CONFIG.REQUIRED_HEADERS.length).map(v => String(v).trim());
  const rows = values
    .slice(1)
    .filter(row => row.slice(0, TS_CONFIG.REQUIRED_HEADERS.length).some(v => !isBlank_(v)));

  const headerMap = buildHeaderMap_(headers);

  return { values, headers, rows, headerMap };
}

/**
ヘッダーマップ作成
*/
function buildHeaderMap_(headers) {
  const map = {};

  headers.forEach((header, index) => {
    const normalized = normalizeHeader_(header);
    if (normalized) {
      map[normalized] = index;
      map[String(header || '').trim()] = index;
    }
  });

  return map;
}

/**
CSVヘッダー検証
*/
function validateCsvHeaders_(headers) {
  const normalizedActual = headers.map(normalizeHeader_);
  const missing = [];

  TS_CONFIG.REQUIRED_HEADERS.forEach(required => {
    if (!normalizedActual.includes(normalizeHeader_(required))) {
      missing.push(required);
    }
  });

  if (missing.length > 0) {
    throw new Error(
      'CSVの必要列が不足しています。\n\n不足列：\n' +
      missing.join('\n') +
      '\n\nTeamSpiritレポートの列設定を確認してください。'
    );
  }
}

/**
ヘッダー正規化
*/
function normalizeHeader_(header) {
  return String(header || '')
    .replace(/\s/g, '')
    .replace(/：/g, ':')
    .trim();
}

/**
CSVから対象年月を自動判定
*/
function detectTargetMonth_(rows, headerMap) {
  const idxDate = headerMap[normalizeHeader_('日付')];
  const idxTargetCode = headerMap[normalizeHeader_('残業申請:申請対象社員コード')];
  const idxApplyDateTime = headerMap[normalizeHeader_('残業申請:申請日時')];
  const idxStatus = headerMap[normalizeHeader_('残業申請:ステータス')];

  const monthCounts = {};

  rows.forEach(row => {
    if (isExcludedStatus_(row[idxStatus]) || isBlank_(row[idxTargetCode]) || isBlank_(row[idxApplyDateTime])) {
      return;
    }

    const date = parseDate_(row[idxDate]);

    if (!date) {
      return;
    }

    const ym = Utilities.formatDate(date, TS_CONFIG.TIMEZONE, 'yyyy-MM');
    monthCounts[ym] = (monthCounts[ym] || 0) + 1;
  });

  const months = Object.keys(monthCounts);

  if (months.length === 0) {
    return '';
  }

  months.sort((a, b) => monthCounts[b] - monthCounts[a]);
  return months[0];
}

/**
原本用データ作成
*/
function buildProcessedRows_(rows, headerMap, deptMaster, targetMonth, settings) {
  const headers = TS_CONFIG.REQUIRED_HEADERS.concat(TS_CONFIG.HELPER_HEADERS);
  const dataRows = [];
  const idx = name => headerMap[normalizeHeader_(name)];

  rows.forEach(row => {
    const base = TS_CONFIG.REQUIRED_HEADERS.map(header => row[idx(header)]);
    const deptName = row[idx('部署名')];
    const targetCode = row[idx('残業申請:申請対象社員コード')];
    const applyType = row[idx('残業申請:申請種類')];
    const status = String(row[idx('残業申請:ステータス')] || '');
    const dateValue = row[idx('日付')];
    const applyDateTimeValue = row[idx('残業申請:申請日時')];
    const approveDateTimeValue = row[idx('残業申請:承認日時')];

    const targetDate = parseDate_(dateValue);
    const applyDateTime = parseDate_(applyDateTimeValue);
    const approveDateTime = parseDate_(approveDateTimeValue);

    const rowMonth = targetDate ? Utilities.formatDate(targetDate, TS_CONFIG.TIMEZONE, 'yyyy-MM') : '';
    const week = targetDate ? getWeekInfo_(targetDate) : { weekKey: '', weekStart: '', weekEnd: '' };
    const excludedStatus = isExcludedStatus_(status);

    const isApplicationLike =
      !!targetDate &&
      !isBlank_(targetCode) &&
      !isBlank_(applyDateTimeValue) &&
      !!applyDateTime;

    const isTarget = isApplicationLike && !excludedStatus;

    const category = getDeptCategory_(String(deptName || ''), deptMaster);
    const closingTime = getClosingTimeForDept_(String(deptName || ''), deptMaster, settings.closingTime);
    const closingDateTime = targetDate ? combineDateAndTime_(targetDate, closingTime) : null;

    const beforeApply =
      isTarget &&
      applyDateTime &&
      closingDateTime &&
      applyDateTime.getTime() <= closingDateTime.getTime();

    const beforeApprove =
      isTarget &&
      approveDateTime &&
      closingDateTime &&
      approveDateTime.getTime() <= closingDateTime.getTime();

    const nextDayLimit = targetDate
      ? new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate() + 1, 23, 59, 59)
      : null;

    const nextDayApprove =
      isTarget &&
      approveDateTime &&
      nextDayLimit &&
      approveDateTime.getTime() <= nextDayLimit.getTime();

    const importKey = isApplicationLike
      ? [
          targetCode,
          formatDateForKey_(targetDate),
          formatDateTimeForKey_(applyDateTime),
          applyType || ''
        ].join('|')
      : '';

    const helper = [
      rowMonth,
      week.weekKey || '',
      week.weekStart || '',
      week.weekEnd || '',
      isTarget,
      closingDateTime || '',
      applyDateTime || '',
      approveDateTime || '',
      beforeApply ? 1 : 0,
      beforeApprove ? 1 : 0,
      nextDayApprove ? 1 : 0,
      category,
      importKey
    ];

    dataRows.push(base.concat(helper));
  });

  return { headers, dataRows };
}


/**
原本・取込データシートの補助列書式を整える。
A:P = TeamSpirit原本、Q:AC = 補助列。
*/
function formatRawAccumSheet_(sheet, columnCount) {
  const dataRows = Math.max(sheet.getMaxRows() - 1, 1);

  // A:P = TeamSpirit原本列。CSV由来文字列を数式として解釈させない。
  sheet.getRange(2, 1, dataRows, Math.min(RAW_TEXT_COLUMN_COUNT, columnCount)).setNumberFormat('@');

  // S:T = 週開始日・週終了日
  sheet.getRange(2, 19, dataRows, 2).setNumberFormat('yyyy/mm/dd');

  // V:X = 定時基準・申請日時_DT・承認日時_DT
  sheet.getRange(2, 22, dataRows, 3).setNumberFormat('yyyy/mm/dd hh:mm');

  // Y:AA = 定時前申請・定時前承認・翌日承認
  sheet.getRange(2, 25, dataRows, 3).setNumberFormat('0');

  sheet.setColumnWidths(1, columnCount, 120);
}

/**
原本シートへ書き込み
*/
function writeRawSheet_(sheet, processed) {
  sheet.clearContents();

  const output = [processed.headers].concat(processed.dataRows);
  ensureSheetSize_(sheet, output.length, output[0].length);
  prepareRawAccumSheetForWrite_(sheet, 1, output.length, output[0].length);

  sheet.getRange(1, 1, output.length, output[0].length)
    .setValues(sanitizeRawTextColumnsForSheet_(output, RAW_TEXT_COLUMN_COUNT));
  formatHeaderRow_(sheet, 1, output[0].length);
  sheet.setFrozenRows(1);
  formatRawAccumSheet_(sheet, output[0].length);
}


/**
取込データシートのヘッダー
*/
function ensureAccumHeader_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TS_CONFIG.SHEETS.ACCUM);
  const headers = TS_CONFIG.REQUIRED_HEADERS.concat(TS_CONFIG.HELPER_HEADERS);

  if (sheet.getLastRow() === 0 || isBlank_(sheet.getRange(1, 1).getValue())) {
    sheet.clearContents();
    ensureSheetSize_(sheet, 1, headers.length);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    formatHeaderRow_(sheet, 1, headers.length);
    sheet.setFrozenRows(1);
    formatRawAccumSheet_(sheet, headers.length);
    return;
  }

  const values = sheet.getDataRange().getValues();
  const currentHeaders = values[0].map(v => String(v || '').trim());

  if (JSON.stringify(currentHeaders) === JSON.stringify(headers)) {
    formatHeaderRow_(sheet, 1, headers.length);
    sheet.setFrozenRows(1);
    formatRawAccumSheet_(sheet, headers.length);
    return;
  }

  const migratedRows = values.slice(1)
    .filter(row => row.some(v => !isBlank_(v)))
    .map(row => migrateAccumRowToCurrentHeaders_(row, currentHeaders, headers));

  const output = [headers].concat(migratedRows);
  sheet.clearContents();
  ensureSheetSize_(sheet, output.length, headers.length);
  prepareRawAccumSheetForWrite_(sheet, 1, output.length, headers.length);
  sheet.getRange(1, 1, output.length, headers.length).setValues(output);
  formatHeaderRow_(sheet, 1, headers.length);
  sheet.setFrozenRows(1);
  formatRawAccumSheet_(sheet, headers.length);
}

/**
取込データシートへ取込キー単位で追加・更新する。
既存データと新規CSVをどちらもメモリ上のMapへ読み込み、V8内でUpsertを完結させる。
最後にヘッダー付き巨大配列を1回だけsetValues()して完全上書きすることで、
分割書き込み・セル走査・Sheets API batchUpdate・Dateシリアル手動変換を全廃する。
*/
function upsertAccumulatedRows_(headers, newRows) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TS_CONFIG.SHEETS.ACCUM);
  ensureAccumHeader_();

  const expectedHeaders = TS_CONFIG.REQUIRED_HEADERS.concat(TS_CONFIG.HELPER_HEADERS);
  const keyCol = expectedHeaders.indexOf('取込キー');
  if (keyCol < 0) {
    throw new Error('内部エラー：取込キー列が見つかりません。');
  }

  const values = sheet.getDataRange().getValues();
  const currentHeaders = values.length > 0 && values[0].some(v => !isBlank_(v))
    ? values[0].map(v => String(v || '').trim())
    : expectedHeaders;

  const rowMap = new Map();
  if (values.length > 1) {
    values.slice(1).forEach(row => {
      if (!row.some(v => !isBlank_(v))) return;
      const adjusted = migrateAccumRowToCurrentHeaders_(row, currentHeaders, expectedHeaders);
      const key = String(adjusted[keyCol] || '').trim();
      if (!key) return;
      // 既存重複キーは最後の行を採用して正規化する。
      rowMap.set(key, adjusted);
    });
  }

  let added = 0;
  let updated = 0;
  let skipped = 0;
  const incomingKeySet = new Set();

  (newRows || []).forEach(row => {
    const adjusted = migrateAccumRowToCurrentHeaders_(row, headers, expectedHeaders);
    const key = String(adjusted[keyCol] || '').trim();
    if (!key) {
      skipped++;
      return;
    }

    const existedBeforeImport = rowMap.has(key) && !incomingKeySet.has(key);
    if (existedBeforeImport) {
      updated++;
    } else if (!incomingKeySet.has(key)) {
      added++;
    }
    incomingKeySet.add(key);
    // 同一CSV内の重複キーも最後の行を採用する。
    rowMap.set(key, adjusted);
  });

  const output = [expectedHeaders].concat(Array.from(rowMap.values()));
  sheet.clearContents();
  ensureSheetSize_(sheet, output.length, expectedHeaders.length);
  prepareRawAccumSheetForWrite_(sheet, 1, output.length, expectedHeaders.length);
  sheet.getRange(1, 1, output.length, expectedHeaders.length)
    .setValues(sanitizeRawTextColumnsForSheet_(output, RAW_TEXT_COLUMN_COUNT));
  formatHeaderRow_(sheet, 1, expectedHeaders.length);
  sheet.setFrozenRows(1);
  formatRawAccumSheet_(sheet, expectedHeaders.length);
  SpreadsheetApp.flush();

  return { added, updated, skipped };
}

/**
列番号をアルファベットへ変換する。例：1 -> A、27 -> AA
*/
function columnToLetter_(column) {
  let letter = '';
  let temp = 0;

  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }

  return letter;
}


/**
旧ヘッダーの取込データを現行ヘッダーへ移行する。
対象週・週開始日・週終了日がない旧データでも、日付列から再計算して補完する。
*/
function migrateAccumRowToCurrentHeaders_(row, fromHeaders, toHeaders) {
  const migrated = adjustRowToHeaders_(row, fromHeaders, toHeaders);
  const headerIndex = buildHeaderIndex_(toHeaders);
  const dateIndex = findHeaderIndex_(headerIndex, '日付');
  const monthIndex = findHeaderIndex_(headerIndex, '対象年月');
  const weekIndex = findHeaderIndex_(headerIndex, '対象週');
  const weekStartIndex = findHeaderIndex_(headerIndex, '週開始日');
  const weekEndIndex = findHeaderIndex_(headerIndex, '週終了日');
  const targetDate = parseDate_(migrated[dateIndex]);

  if (targetDate) {
    if (isBlank_(migrated[monthIndex])) {
      migrated[monthIndex] = Utilities.formatDate(targetDate, TS_CONFIG.TIMEZONE, 'yyyy-MM');
    }

    const week = getWeekInfo_(targetDate);
    if (isBlank_(migrated[weekIndex])) migrated[weekIndex] = week.weekKey;
    if (isBlank_(migrated[weekStartIndex])) migrated[weekStartIndex] = week.weekStart;
    if (isBlank_(migrated[weekEndIndex])) migrated[weekEndIndex] = week.weekEnd;
  }

  return migrated;
}

/**
異なるヘッダー順の行を指定ヘッダー順に変換
*/
function adjustRowToHeaders_(row, fromHeaders, toHeaders) {
  const fromIndex = {};
  fromHeaders.forEach((header, index) => {
    fromIndex[String(header || '').trim()] = index;
    fromIndex[normalizeHeader_(header)] = index;
  });
  return toHeaders.map(header => {
    const exact = fromIndex[header];
    const normalized = fromIndex[normalizeHeader_(header)];
    const index = exact !== undefined ? exact : normalized;
    return index === undefined ? '' : row[index];
  });
}

/**
取込データシートの情報取得
getDataRange()依存を避け、実データ範囲だけを読む。
*/
function getAccumulatedDataInfo_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TS_CONFIG.SHEETS.ACCUM);
  ensureAccumHeader_();

  const expectedHeaders = TS_CONFIG.REQUIRED_HEADERS.concat(TS_CONFIG.HELPER_HEADERS);
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), expectedHeaders.length);

  if (lastRow < 2) {
    return {
      headers: expectedHeaders,
      rows: []
    };
  }

  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values[0].map(v => String(v || '').trim());
  return {
    headers: headers,
    rows: values.slice(1).filter(row => row.some(v => !isBlank_(v)))
  };
}

/**
CSVから対象月・対象週を判定
*/
function detectTargetPeriods_(rows, headerMap) {
  const idxDate = headerMap[normalizeHeader_('日付')];
  const idxTargetCode = headerMap[normalizeHeader_('残業申請:申請対象社員コード')];
  const idxApplyDateTime = headerMap[normalizeHeader_('残業申請:申請日時')];
  const idxStatus = headerMap[normalizeHeader_('残業申請:ステータス')];
  let latestDate = null;
  rows.forEach(row => {
    if (isExcludedStatus_(row[idxStatus]) || isBlank_(row[idxTargetCode]) || isBlank_(row[idxApplyDateTime])) return;
    const date = parseDate_(row[idxDate]);
    if (!date) return;
    if (!latestDate || date.getTime() > latestDate.getTime()) latestDate = date;
  });
  if (!latestDate) return { targetMonth: '', targetWeek: '', weekStart: null, weekEnd: null };
  const week = getWeekInfo_(latestDate);
  return {
    targetMonth: Utilities.formatDate(latestDate, TS_CONFIG.TIMEZONE, 'yyyy-MM'),
    targetWeek: week.weekKey,
    weekStart: week.weekStart,
    weekEnd: week.weekEnd
  };
}

/**
月次・週次サマリーをまとめて作成
1回の走査で、当月・前月・当週・前週を同時に集計する。
*/
function buildMonthlyWeeklySummaryBundle_(rows, headers, settings) {
  const currentMonth = String(settings.targetMonth || '').trim();
  const previousMonth = getPreviousMonthKey_(currentMonth);
  const currentWeek = String(settings.targetWeek || '').trim();
  const previousWeek = getPreviousWeekKey_(currentWeek);

  const headerIndex = buildHeaderIndex_(headers);
  const deptMaster = getDeptMaster_();
  const indexes = {
    targetFlag: findHeaderIndex_(headerIndex, '集計対象'),
    targetMonth: findHeaderIndex_(headerIndex, '対象年月'),
    targetWeek: findHeaderIndex_(headerIndex, '対象週'),
    deptName: findHeaderIndex_(headerIndex, '部署名'),
    category: findHeaderIndex_(headerIndex, '部門区分'),
    employeeCode: findHeaderIndex_(headerIndex, '残業申請:申請対象社員コード'),
    employeeName: findHeaderIndex_(headerIndex, '残業申請:申請対象社員名'),
    approver: findHeaderIndex_(headerIndex, '残業申請:承認者'),
    beforeApply: findHeaderIndex_(headerIndex, '定時前申請'),
    beforeApprove: findHeaderIndex_(headerIndex, '定時前承認'),
    nextDayApprove: findHeaderIndex_(headerIndex, '翌日承認'),
    approveDateTime: findHeaderIndex_(headerIndex, '承認日時_DT')
  };

  const bundle = {
    monthly: {
      current: createPeriodBucket_(),
      previous: createPeriodBucket_(),
      currentLabel: currentMonth,
      previousLabel: previousMonth
    },
    weekly: {
      current: createPeriodBucket_(),
      previous: createPeriodBucket_(),
      currentLabel: currentWeek,
      previousLabel: previousWeek
    }
  };

  rows.forEach(row => {
    const targetFlag = row[indexes.targetFlag];
    const isTarget = targetFlag === true || String(targetFlag).toUpperCase() === 'TRUE';

    if (!isTarget) {
      return;
    }

    const rowMonth = String(row[indexes.targetMonth] || '').trim();
    const rowWeek = String(row[indexes.targetWeek] || '').trim();

    if (
      rowMonth !== currentMonth &&
      rowMonth !== previousMonth &&
      rowWeek !== currentWeek &&
      rowWeek !== previousWeek
    ) {
      return;
    }

    const packet = {
      deptName: getCanonicalDeptName_(String(row[indexes.deptName] || ''), deptMaster),
      category: String(row[indexes.category] || ''),
      employeeCode: String(row[indexes.employeeCode] || ''),
      employeeName: String(row[indexes.employeeName] || ''),
      approver: String(row[indexes.approver] || ''),
      beforeApply: safeParseNumber_(row[indexes.beforeApply]),
      beforeApprove: safeParseNumber_(row[indexes.beforeApprove]),
      nextDayApprove: safeParseNumber_(row[indexes.nextDayApprove]),
      hasApproveDate: !isBlank_(row[indexes.approveDateTime])
    };

    if (rowMonth === currentMonth) {
      addPacketToPeriodBucket_(bundle.monthly.current, packet);
    }
    if (rowMonth === previousMonth) {
      addPacketToPeriodBucket_(bundle.monthly.previous, packet);
    }
    if (rowWeek === currentWeek) {
      addPacketToPeriodBucket_(bundle.weekly.current, packet);
    }
    if (rowWeek === previousWeek) {
      addPacketToPeriodBucket_(bundle.weekly.previous, packet);
    }
  });

  const threshold = parseRate_(settings.alertThreshold, 0.8);

  return {
    monthly: {
      current: finalizePeriodBucket_(bundle.monthly.current, threshold),
      previous: finalizePeriodBucket_(bundle.monthly.previous, threshold),
      currentLabel: bundle.monthly.currentLabel,
      previousLabel: bundle.monthly.previousLabel
    },
    weekly: {
      current: finalizePeriodBucket_(bundle.weekly.current, threshold),
      previous: finalizePeriodBucket_(bundle.weekly.previous, threshold),
      currentLabel: bundle.weekly.currentLabel,
      previousLabel: bundle.weekly.previousLabel
    }
  };
}

/**
ヘッダーインデックスを取得する。
*/
function findHeaderIndex_(headerIndex, name) {
  const exact = headerIndex[name];
  const normalized = headerIndex[normalizeHeader_(name)];
  const index = exact !== undefined ? exact : normalized;

  if (index === undefined) {
    throw new Error(`内部エラー：必要な列が見つかりません。列名：${name}`);
  }

  return index;
}

/**
期間別集計バケットを作成する。
*/
function createPeriodBucket_() {
  return {
    deptMap: {},
    personMap: {},
    approverMap: {},
    totalImportCount: 0
  };
}

/**
1行分の集計パケットを期間別バケットへ追加する。
*/
function addPacketToPeriodBucket_(bucket, packet) {
  bucket.totalImportCount++;

  const deptKey = [packet.deptName, packet.category].join('|');

  if (!bucket.deptMap[deptKey]) {
    bucket.deptMap[deptKey] = createSummaryBase_({
      deptName: packet.deptName,
      category: packet.category
    });
  }

  addSummaryCount_(
    bucket.deptMap[deptKey],
    packet.beforeApply,
    packet.beforeApprove,
    packet.nextDayApprove,
    packet.hasApproveDate
  );

  const personKey = [
    packet.category,
    packet.deptName,
    packet.employeeCode,
    packet.employeeName
  ].join('|');

  if (!bucket.personMap[personKey]) {
    bucket.personMap[personKey] = createSummaryBase_({
      category: packet.category,
      deptName: packet.deptName,
      employeeCode: packet.employeeCode,
      employeeName: packet.employeeName
    });
  }

  addSummaryCount_(
    bucket.personMap[personKey],
    packet.beforeApply,
    packet.beforeApprove,
    packet.nextDayApprove,
    packet.hasApproveDate
  );

  if (!isBlank_(packet.approver)) {
    const approverKey = [packet.category, packet.approver].join('|');

    if (!bucket.approverMap[approverKey]) {
      bucket.approverMap[approverKey] = createSummaryBase_({
        category: packet.category,
        approver: packet.approver
      });
    }

    addSummaryCount_(
      bucket.approverMap[approverKey],
      0,
      packet.beforeApprove,
      packet.nextDayApprove,
      packet.hasApproveDate
    );
  }
}

/**
期間別バケットを既存の出力形式へ変換する。
*/
function finalizePeriodBucket_(bucket, threshold) {
  const deptRows = Object.values(bucket.deptMap).map(s => toDeptSummaryRow_(s, threshold));
  const personRows = Object.values(bucket.personMap).map(toPersonSummaryRow_);
  const approverRows = Object.values(bucket.approverMap).map(toApproverSummaryRow_);

  deptRows.sort((a, b) => {
    if (a[1] !== b[1]) return String(a[1]).localeCompare(String(b[1]), 'ja');
    return safeParseNumber_(a[8]) - safeParseNumber_(b[8]);
  });

  personRows.sort((a, b) => {
    if (a[0] !== b[0]) return String(a[0]).localeCompare(String(b[0]), 'ja');
    if (a[1] !== b[1]) return String(a[1]).localeCompare(String(b[1]), 'ja');
    return safeParseNumber_(a[10]) - safeParseNumber_(b[10]);
  });

  approverRows.sort((a, b) => {
    if (a[0] !== b[0]) return String(a[0]).localeCompare(String(b[0]), 'ja');
    return safeParseNumber_(a[6]) - safeParseNumber_(b[6]);
  });

  return {
    deptRows,
    personRows,
    approverRows,
    totalImportCount: bucket.totalImportCount
  };
}

/**
週次集計作成
*/
function buildWeeklySummaries_(rawRows, targetWeek, settings, customHeaders) {
  const headers = customHeaders || TS_CONFIG.REQUIRED_HEADERS.concat(TS_CONFIG.HELPER_HEADERS);
  const headerIndex = buildHeaderIndex_(headers);
  const getByName = (row, name) => getValueByHeader_(row, headerIndex, name);
  const filteredRows = rawRows.filter(row => {
    const targetFlag = getByName(row, '集計対象');
    const isTarget = targetFlag === true || String(targetFlag).toUpperCase() === 'TRUE';
    if (!isTarget) return false;
    const rowWeek = String(getByName(row, '対象週') || '').trim();
    return rowWeek === String(targetWeek || '').trim();
  });
  return buildSummariesNoPeriodFilter_(filteredRows, settings, headers);
}

/**
期間フィルター済みデータの集計作成
*/
function buildSummariesNoPeriodFilter_(rawRows, settings, customHeaders) {
  const deptMap = {};
  const personMap = {};
  const approverMap = {};
  const headers = customHeaders || TS_CONFIG.REQUIRED_HEADERS.concat(TS_CONFIG.HELPER_HEADERS);
  const headerIndex = buildHeaderIndex_(headers);
  const getByName = (row, name) => getValueByHeader_(row, headerIndex, name);
  let totalImportCount = 0;

  rawRows.forEach(row => {
    totalImportCount++;
    const deptName = String(getByName(row, '部署名') || '');
    const category = String(getByName(row, '部門区分') || '');
    const employeeCode = String(getByName(row, '残業申請:申請対象社員コード') || '');
    const employeeName = String(getByName(row, '残業申請:申請対象社員名') || '');
    const approver = String(getByName(row, '残業申請:承認者') || '');
    const beforeApply = Number(getByName(row, '定時前申請') || 0);
    const beforeApprove = Number(getByName(row, '定時前承認') || 0);
    const nextDayApprove = Number(getByName(row, '翌日承認') || 0);
    const hasApproveDate = !isBlank_(getByName(row, '承認日時_DT'));
    const deptKey = [deptName, category].join('|');
    if (!deptMap[deptKey]) deptMap[deptKey] = createSummaryBase_({ deptName, category });
    addSummaryCount_(deptMap[deptKey], beforeApply, beforeApprove, nextDayApprove, hasApproveDate);
    const personKey = [category, deptName, employeeCode, employeeName].join('|');
    if (!personMap[personKey]) personMap[personKey] = createSummaryBase_({ category, deptName, employeeCode, employeeName });
    addSummaryCount_(personMap[personKey], beforeApply, beforeApprove, nextDayApprove, hasApproveDate);
    if (!isBlank_(approver)) {
      const approverKey = [category, approver].join('|');
      if (!approverMap[approverKey]) approverMap[approverKey] = createSummaryBase_({ category, approver });
      addSummaryCount_(approverMap[approverKey], 0, beforeApprove, nextDayApprove, hasApproveDate);
    }
  });

  const threshold = parseRate_(settings.alertThreshold, 0.8);
  const deptRows = Object.values(deptMap).map(s => toDeptSummaryRow_(s, threshold));
  const personRows = Object.values(personMap).map(toPersonSummaryRow_);
  const approverRows = Object.values(approverMap).map(toApproverSummaryRow_);
  deptRows.sort((a, b) => (a[1] !== b[1]) ? String(a[1]).localeCompare(String(b[1]), 'ja') : a[8] - b[8]);
  personRows.sort((a, b) => (a[0] !== b[0]) ? String(a[0]).localeCompare(String(b[0]), 'ja') : ((a[1] !== b[1]) ? String(a[1]).localeCompare(String(b[1]), 'ja') : a[10] - b[10]));
  approverRows.sort((a, b) => (a[0] !== b[0]) ? String(a[0]).localeCompare(String(b[0]), 'ja') : a[6] - b[6]);
  return { deptRows, personRows, approverRows, totalImportCount };
}

function buildHeaderIndex_(headers) {
  const index = {};
  headers.forEach((header, i) => {
    const raw = String(header || '').trim();
    const normalized = normalizeHeader_(raw);
    if (raw) index[raw] = i;
    if (normalized) index[normalized] = i;
  });
  return index;
}

function getValueByHeader_(row, headerIndex, name) {
  const exact = headerIndex[name];
  const normalized = headerIndex[normalizeHeader_(name)];
  const index = exact !== undefined ? exact : normalized;
  return index === undefined ? '' : row[index];
}

/**
集計作成
*/
function buildSummaries_(rawRows, targetMonth, settings, customHeaders) {
  const deptMap = {};
  const personMap = {};
  const approverMap = {};

  const allHeaders = customHeaders || TS_CONFIG.REQUIRED_HEADERS.concat(TS_CONFIG.HELPER_HEADERS);
  const headerIndex = {};

  allHeaders.forEach((header, index) => {
    const rawHeader = String(header || '').trim();
    const normalizedHeader = normalizeHeader_(rawHeader);

    if (rawHeader) {
      headerIndex[rawHeader] = index;
    }
    if (normalizedHeader) {
      headerIndex[normalizedHeader] = index;
    }
  });

  const getByName = (row, name) => {
    const exactIndex = headerIndex[name];
    const normalizedIndex = headerIndex[normalizeHeader_(name)];
    const index = exactIndex !== undefined ? exactIndex : normalizedIndex;

    if (index === undefined) {
      return '';
    }

    return row[index];
  };

  let totalImportCount = 0;

  rawRows.forEach(row => {
    const rowMonth = String(getByName(row, '対象年月') || '').trim();
    const targetMonthText = String(targetMonth || '').trim();
    const targetFlag = getByName(row, '集計対象');
    const isTarget = targetFlag === true || String(targetFlag).toUpperCase() === 'TRUE';

    if (!isTarget || rowMonth !== targetMonthText) {
      return;
    }

    totalImportCount++;

    const deptName = String(getByName(row, '部署名') || '');
    const category = String(getByName(row, '部門区分') || '');
    const employeeCode = String(getByName(row, '残業申請:申請対象社員コード') || '');
    const employeeName = String(getByName(row, '残業申請:申請対象社員名') || '');
    const approver = String(getByName(row, '残業申請:承認者') || '');

    const beforeApply = Number(getByName(row, '定時前申請') || 0);
    const beforeApprove = Number(getByName(row, '定時前承認') || 0);
    const nextDayApprove = Number(getByName(row, '翌日承認') || 0);
    const hasApproveDate = !isBlank_(getByName(row, '承認日時_DT'));

    const deptKey = [deptName, category].join('|');

    if (!deptMap[deptKey]) {
      deptMap[deptKey] = createSummaryBase_({ deptName, category });
    }

    addSummaryCount_(deptMap[deptKey], beforeApply, beforeApprove, nextDayApprove, hasApproveDate);

    const personKey = [category, deptName, employeeCode, employeeName].join('|');

    if (!personMap[personKey]) {
      personMap[personKey] = createSummaryBase_({ category, deptName, employeeCode, employeeName });
    }

    addSummaryCount_(personMap[personKey], beforeApply, beforeApprove, nextDayApprove, hasApproveDate);

    if (!isBlank_(approver)) {
      const approverKey = [category, approver].join('|');

      if (!approverMap[approverKey]) {
        approverMap[approverKey] = createSummaryBase_({ category, approver });
      }

      addSummaryCount_(approverMap[approverKey], 0, beforeApprove, nextDayApprove, hasApproveDate);
    }
  });

  const threshold = parseRate_(settings.alertThreshold, 0.8);
  const deptRows = Object.values(deptMap).map(s => toDeptSummaryRow_(s, threshold));
  const personRows = Object.values(personMap).map(toPersonSummaryRow_);
  const approverRows = Object.values(approverMap).map(toApproverSummaryRow_);

  deptRows.sort((a, b) => {
    if (a[1] !== b[1]) return String(a[1]).localeCompare(String(b[1]), 'ja');
    return a[8] - b[8];
  });

  personRows.sort((a, b) => {
    if (a[0] !== b[0]) return String(a[0]).localeCompare(String(b[0]), 'ja');
    if (a[1] !== b[1]) return String(a[1]).localeCompare(String(b[1]), 'ja');
    return a[10] - b[10];
  });

  approverRows.sort((a, b) => {
    if (a[0] !== b[0]) return String(a[0]).localeCompare(String(b[0]), 'ja');
    return a[6] - b[6];
  });

  return { deptRows, personRows, approverRows, totalImportCount };
}

/**
集計ベース
*/
function createSummaryBase_(props) {
  return Object.assign({
    count: 0,
    beforeApplyCount: 0,
    beforeApproveCount: 0,
    nextDayApproveCount: 0,
    approveDateCount: 0
  }, props);
}

/**
集計カウント加算
*/
function addSummaryCount_(summary, beforeApply, beforeApprove, nextDayApprove, hasApproveDate) {
  summary.count++;
  summary.beforeApplyCount += Number(beforeApply || 0);
  summary.beforeApproveCount += Number(beforeApprove || 0);
  summary.nextDayApproveCount += Number(nextDayApprove || 0);
  summary.approveDateCount += hasApproveDate ? 1 : 0;
}

/**
部署別集計行
*/
function toDeptSummaryRow_(s, threshold) {
  const beforeApplyRate = safeRate_(s.beforeApplyCount, s.count);
  const beforeApproveRate = safeRate_(s.beforeApproveCount, s.count);
  const nextDayApproveRate = safeRate_(s.nextDayApproveCount, s.count);
  const notApprovedCount = s.count - s.approveDateCount;
  const alert = beforeApproveRate < threshold ? '要確認' : '';

  return [
    s.deptName,
    s.category,
    s.count,
    s.beforeApplyCount,
    s.beforeApproveCount,
    s.nextDayApproveCount,
    s.approveDateCount,
    beforeApplyRate,
    beforeApproveRate,
    nextDayApproveRate,
    notApprovedCount,
    alert
  ];
}

/**
個人別集計行
*/
function toPersonSummaryRow_(s) {
  return [
    s.category,
    s.deptName,
    s.employeeCode,
    s.employeeName,
    s.count,
    s.beforeApplyCount,
    s.beforeApproveCount,
    s.nextDayApproveCount,
    s.approveDateCount,
    safeRate_(s.beforeApplyCount, s.count),
    safeRate_(s.beforeApproveCount, s.count),
    safeRate_(s.nextDayApproveCount, s.count)
  ];
}

/**
承認者別集計行
*/
function toApproverSummaryRow_(s) {
  return [
    s.category,
    s.approver,
    s.count,
    s.beforeApproveCount,
    s.nextDayApproveCount,
    s.approveDateCount,
    safeRate_(s.beforeApproveCount, s.count),
    safeRate_(s.nextDayApproveCount, s.count)
  ];
}

/**
集計シート書き込み
*/
function writeSummarySheets_(summaries) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  writeTable_(
    ss.getSheetByName(TS_CONFIG.SHEETS.SUMMARY_DEPT),
    [[
      '部署名',
      '部門区分',
      '申請件数',
      '定時前申請件数',
      '定時前承認件数',
      '翌日承認件数',
      '承認日時入力件数',
      '定時前申請率',
      '定時前承認率',
      '翌日承認率',
      '未承認件数',
      '要確認'
    ]].concat(summaries.deptRows),
    { percentColumns: [8, 9, 10] }
  );

  writeTable_(
    ss.getSheetByName(TS_CONFIG.SHEETS.SUMMARY_PERSON),
    [[
      '部門区分',
      '部署名',
      '社員コード',
      '社員名',
      '申請件数',
      '定時前申請件数',
      '定時前承認件数',
      '翌日承認件数',
      '承認日時入力件数',
      '定時前申請率',
      '定時前承認率',
      '翌日承認率'
    ]].concat(summaries.personRows),
    { percentColumns: [10, 11, 12] }
  );

  writeTable_(
    ss.getSheetByName(TS_CONFIG.SHEETS.SUMMARY_APPROVER),
    [[
      '部門区分',
      '承認者',
      '申請件数',
      '定時前承認件数',
      '翌日承認件数',
      '承認日時入力件数',
      '定時前承認率',
      '翌日承認率'
    ]].concat(summaries.approverRows),
    { percentColumns: [7, 8] }
  );
}


/**
月次・週次の集計シート書き込み
*/
function writeAllSummarySheets_(bundle) {
  writeSummarySheets_(bundle.monthly.current);
  writeSummarySet_(
    bundle.weekly.current,
    TS_CONFIG.SHEETS.SUMMARY_WEEK_DEPT,
    TS_CONFIG.SHEETS.SUMMARY_WEEK_PERSON,
    TS_CONFIG.SHEETS.SUMMARY_WEEK_APPROVER
  );
}

function writeSummarySet_(summaries, deptSheetName, personSheetName, approverSheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  writeTable_(
    ss.getSheetByName(deptSheetName),
    [[
      '部署名','部門区分','申請件数','定時前申請件数','定時前承認件数','翌日承認件数','承認日時入力件数','定時前申請率','定時前承認率','翌日承認率','未承認件数','要確認'
    ]].concat(summaries.deptRows),
    { percentColumns: [8, 9, 10] }
  );
  writeTable_(
    ss.getSheetByName(personSheetName),
    [[
      '部門区分','部署名','社員コード','社員名','申請件数','定時前申請件数','定時前承認件数','翌日承認件数','承認日時入力件数','定時前申請率','定時前承認率','翌日承認率'
    ]].concat(summaries.personRows),
    { percentColumns: [10, 11, 12] }
  );
  writeTable_(
    ss.getSheetByName(approverSheetName),
    [[
      '部門区分','承認者','申請件数','定時前承認件数','翌日承認件数','承認日時入力件数','定時前承認率','翌日承認率'
    ]].concat(summaries.approverRows),
    { percentColumns: [7, 8] }
  );
}

/**
ダッシュボード書き込み
*/
function writeDashboards_(summaries, targetMonth, settings) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  writeCategoryDashboard_(
    ss.getSheetByName(TS_CONFIG.SHEETS.DASHBOARD_STAFF),
    '【メイン】スタッフ部門：残業事前申請・事前承認 運用状況',
    TS_CONFIG.STAFF_CATEGORY,
    summaries.deptRows,
    targetMonth,
    settings
  );

  writeCategoryDashboard_(
    ss.getSheetByName(TS_CONFIG.SHEETS.DASHBOARD_SALES),
    '【参考】営業部門：残業事前申請・事前承認 運用状況',
    TS_CONFIG.SALES_CATEGORY,
    summaries.deptRows,
    targetMonth,
    settings
  );
}

/**
カテゴリ別ダッシュボード
*/
function writeCategoryDashboard_(sheet, title, category, deptRows, targetMonth, settings) {
  const rows = deptRows.filter(row => row[1] === category);

  const totals = rows.reduce((acc, row) => {
    acc.count += Number(row[2] || 0);
    acc.beforeApply += Number(row[3] || 0);
    acc.beforeApprove += Number(row[4] || 0);
    acc.nextDayApprove += Number(row[5] || 0);
    acc.approveDate += Number(row[6] || 0);
    return acc;
  }, { count: 0, beforeApply: 0, beforeApprove: 0, nextDayApprove: 0, approveDate: 0 });

  const note = settings.dashboardNote ||
    'TeamSpirit上の残業申請・承認データをもとに、事前申請・事前承認の運用状況を可視化するものです。';

  const summary = [
    [title],
    [note],
    [],
    ['対象年月', targetMonth, '取込区分', settings.importType, '最終取込日時', settings.lastImportTime],
    [],
    ['指標', '値'],
    ['申請件数', totals.count],
    ['定時前申請率', safeRate_(totals.beforeApply, totals.count)],
    ['定時前承認率', safeRate_(totals.beforeApprove, totals.count)],
    ['翌日承認率', safeRate_(totals.nextDayApprove, totals.count)],
    ['未承認件数', totals.count - totals.approveDate],
    [],
    [`${category} 部署別明細`],
    ['部署名', '申請件数', '定時前申請率', '定時前承認率', '翌日承認率', '未承認件数', '要確認']
  ];

  const detail = rows
    .sort((a, b) => a[8] - b[8])
    .map(row => [
      row[0],
      row[2],
      row[7],
      row[8],
      row[9],
      row[10],
      row[11]
    ]);

  const output = summary.concat(detail);
  const padded = padRows_(output);

  sheet.clearContents();
  ensureSheetSize_(sheet, padded.length, padded[0].length);

  sheet.getRange(1, 1, padded.length, padded[0].length).setValues(sanitizeValuesForSheet_(padded));
  sheet.getRange(1, 1).setFontWeight('bold').setFontSize(14);
  sheet.getRange(6, 1, 1, 2).setFontWeight('bold').setBackground('#17466f').setFontColor('#ffffff');
  sheet.getRange(14, 1, 1, 7).setFontWeight('bold').setBackground('#17466f').setFontColor('#ffffff');
  sheet.getRange(4, 6).setNumberFormat('yyyy/mm/dd hh:mm');
  sheet.getRange(8, 2, 3, 1).setNumberFormat('0.0%');

  if (detail.length > 0) {
    sheet.getRange(15, 3, detail.length, 3).setNumberFormat('0.0%');
  }

  sheet.setFrozenRows(14);
  sheet.setColumnWidths(1, 7, 120);
}


/**
月次・週次ダッシュボード書き込み
*/
function writeAllDashboards_(bundle, settings) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  writeCategoryDashboardWithCompare_(
    ss.getSheetByName(TS_CONFIG.SHEETS.DASHBOARD_STAFF),
    '【月次】スタッフ部門：残業事前申請・事前承認 運用状況',
    TS_CONFIG.STAFF_CATEGORY,
    bundle.monthly.current.deptRows,
    bundle.monthly.previous.deptRows,
    bundle.monthly.currentLabel,
    bundle.monthly.previousLabel,
    '前月',
    settings
  );
  writeCategoryDashboardWithCompare_(
    ss.getSheetByName(TS_CONFIG.SHEETS.DASHBOARD_SALES),
    '【月次】営業部門（参考）：残業事前申請・事前承認 運用状況',
    TS_CONFIG.SALES_CATEGORY,
    bundle.monthly.current.deptRows,
    bundle.monthly.previous.deptRows,
    bundle.monthly.currentLabel,
    bundle.monthly.previousLabel,
    '前月',
    settings
  );
  writeCategoryDashboardWithCompare_(
    ss.getSheetByName(TS_CONFIG.SHEETS.DASHBOARD_WEEK_STAFF),
    '【週次】スタッフ部門：残業事前申請・事前承認 運用状況',
    TS_CONFIG.STAFF_CATEGORY,
    bundle.weekly.current.deptRows,
    bundle.weekly.previous.deptRows,
    bundle.weekly.currentLabel,
    bundle.weekly.previousLabel,
    '前週',
    settings
  );
  writeCategoryDashboardWithCompare_(
    ss.getSheetByName(TS_CONFIG.SHEETS.DASHBOARD_WEEK_SALES),
    '【週次】営業部門（参考）：残業事前申請・事前承認 運用状況',
    TS_CONFIG.SALES_CATEGORY,
    bundle.weekly.current.deptRows,
    bundle.weekly.previous.deptRows,
    bundle.weekly.currentLabel,
    bundle.weekly.previousLabel,
    '前週',
    settings
  );
}

function writeCategoryDashboardWithCompare_(sheet, title, category, currentDeptRows, previousDeptRows, currentLabel, previousLabel, compareLabel, settings) {
  const currentRows = currentDeptRows.filter(row => row[1] === category);
  const previousRows = previousDeptRows.filter(row => row[1] === category);
  const previousMap = buildDeptRowMap_(previousRows);
  const currentTotals = calcTotals_(currentRows);
  const previousTotals = calcTotals_(previousRows);
  const note = settings.dashboardNote || 'TeamSpirit上の残業申請・承認データをもとに、事前申請・事前承認の運用状況を可視化するものです。';
  const currentApplyRate = safeRate_(currentTotals.beforeApply, currentTotals.count);
  const previousApplyRate = safeRate_(previousTotals.beforeApply, previousTotals.count);
  const currentApproveRate = safeRate_(currentTotals.beforeApprove, currentTotals.count);
  const previousApproveRate = safeRate_(previousTotals.beforeApprove, previousTotals.count);
  const currentNextDayRate = safeRate_(currentTotals.nextDayApprove, currentTotals.count);
  const previousNextDayRate = safeRate_(previousTotals.nextDayApprove, previousTotals.count);
  const summary = [
    [title],
    [note],
    [],
    ['対象期間', currentLabel, `${compareLabel}期間`, previousLabel, '取込区分', settings.importType, '最終取込日時', settings.lastImportTime],
    [],
    ['指標', '当期', compareLabel, '増減', '増減率／差分'],
    ['申請件数', currentTotals.count, previousTotals.count, currentTotals.count - previousTotals.count, safeGrowthRate_(currentTotals.count, previousTotals.count)],
    ['定時前申請率', currentApplyRate, previousApplyRate, currentApplyRate - previousApplyRate, currentApplyRate - previousApplyRate],
    ['定時前承認率', currentApproveRate, previousApproveRate, currentApproveRate - previousApproveRate, currentApproveRate - previousApproveRate],
    ['翌日承認率', currentNextDayRate, previousNextDayRate, currentNextDayRate - previousNextDayRate, currentNextDayRate - previousNextDayRate],
    ['未承認件数', currentTotals.count - currentTotals.approveDate, previousTotals.count - previousTotals.approveDate, (currentTotals.count - currentTotals.approveDate) - (previousTotals.count - previousTotals.approveDate), ''],
    [],
    [`${category} 部署別明細`],
    ['部署名', '申請件数', `${compareLabel}件数`, '増減', '増減率', '定時前申請率', `${compareLabel}差`, '定時前承認率', `${compareLabel}差`, '翌日承認率', `${compareLabel}差`, '未承認件数', '要確認']
  ];
  const detail = currentRows.sort((a, b) => a[8] - b[8]).map(row => {
    const prev = previousMap[row[0]] || emptyDeptSummaryRow_(row[0], category);
    return [
      row[0],
      row[2],
      prev[2],
      Number(row[2] || 0) - Number(prev[2] || 0),
      safeGrowthRate_(row[2], prev[2]),
      row[7],
      Number(row[7] || 0) - Number(prev[7] || 0),
      row[8],
      Number(row[8] || 0) - Number(prev[8] || 0),
      row[9],
      Number(row[9] || 0) - Number(prev[9] || 0),
      row[10],
      row[11]
    ];
  });
  const output = summary.concat(detail);
  const padded = padRows_(output);
  sheet.clearContents();
  ensureSheetSize_(sheet, padded.length, padded[0].length);
  sheet.getRange(1, 1, padded.length, padded[0].length).setValues(sanitizeValuesForSheet_(padded));
  sheet.getRange(1, 1).setFontWeight('bold').setFontSize(14);
  sheet.getRange(6, 1, 1, 5).setFontWeight('bold').setBackground('#17466f').setFontColor('#ffffff');
  sheet.getRange(14, 1, 1, 13).setFontWeight('bold').setBackground('#17466f').setFontColor('#ffffff');
  sheet.getRange(4, 8).setNumberFormat('yyyy/mm/dd hh:mm');
  sheet.getRange(7, 5, 1, 1).setNumberFormat('0.0%');
  sheet.getRange(8, 2, 3, 4).setNumberFormat('0.0%');
  if (detail.length > 0) {
    sheet.getRange(15, 5, detail.length, 1).setNumberFormat('0.0%');
    sheet.getRange(15, 6, detail.length, 1).setNumberFormat('0.0%');
    sheet.getRange(15, 7, detail.length, 1).setNumberFormat('+0.0%;-0.0%;0.0%');
    sheet.getRange(15, 8, detail.length, 1).setNumberFormat('0.0%');
    sheet.getRange(15, 9, detail.length, 1).setNumberFormat('+0.0%;-0.0%;0.0%');
    sheet.getRange(15, 10, detail.length, 1).setNumberFormat('0.0%');
    sheet.getRange(15, 11, detail.length, 1).setNumberFormat('+0.0%;-0.0%;0.0%');
  }
  sheet.setFrozenRows(14);
  sheet.setColumnWidths(1, 13, 120);
}

function buildDeptRowMap_(rows) {
  const map = {};
  rows.forEach(row => map[row[0]] = row);
  return map;
}

function emptyDeptSummaryRow_(deptName, category) {
  return [deptName, category, 0, 0, 0, 0, 0, 0, 0, 0, 0, ''];
}

function calcTotals_(rows) {
  return rows.reduce((acc, row) => {
    acc.count += Number(row[2] || 0);
    acc.beforeApply += Number(row[3] || 0);
    acc.beforeApprove += Number(row[4] || 0);
    acc.nextDayApprove += Number(row[5] || 0);
    acc.approveDate += Number(row[6] || 0);
    return acc;
  }, { count: 0, beforeApply: 0, beforeApprove: 0, nextDayApprove: 0, approveDate: 0 });
}

/**
エラー行作成
*/
function buildErrorRows_(rows, headerMap, deptMaster, targetMonth) {
  const errors = [];
  const seenKeys = {};
  const idx = name => {
    const normalized = normalizeHeader_(name);
    return headerMap[normalized] !== undefined ? headerMap[normalized] : headerMap[name];
  };

  rows.forEach((row, i) => {
    const displayRowNumber = i + 2;
    const dept = row[idx('部署名')];
    const dateValue = row[idx('日付')];
    const targetCode = row[idx('残業申請:申請対象社員コード')];
    const targetName = row[idx('残業申請:申請対象社員名')];
    const applyType = row[idx('残業申請:申請種類')];
    const status = String(row[idx('残業申請:ステータス')] || '');
    const applyDateTime = row[idx('残業申請:申請日時')];
    const approveDateTime = row[idx('残業申請:承認日時')];

    const isApplicationRow = !isBlank_(targetCode) || !isBlank_(applyDateTime) || !isBlank_(applyType);

    if (!isApplicationRow) {
      return;
    }

    const excludedStatus = isExcludedStatus_(status);

    if (excludedStatus) {
      return;
    }

    const parsedDate = parseDate_(dateValue);
    const parsedApply = parseDate_(applyDateTime);
    const parsedApprove = parseDate_(approveDateTime);

    const pushError = (type, message, statusText, memo) => {
      errors.push([
        new Date(),
        targetMonth,
        displayRowNumber,
        type,
        message,
        statusText,
        memo || ''
      ]);
    };

    if (isBlank_(dateValue) || !parsedDate) {
      pushError('エラー', '日付が空欄、または日付として読み取れません。', '未対応', `部署：${dept || ''}／対象者：${targetName || ''}`);
    }

    if (isBlank_(targetCode)) {
      pushError('エラー', '申請対象社員コードが空欄です。', '未対応', `部署：${dept || ''}／対象者：${targetName || ''}`);
    }

    if (isBlank_(applyDateTime) || !parsedApply) {
      pushError('エラー', '申請日時が空欄、または日時として読み取れません。', '未対応', `部署：${dept || ''}／対象者：${targetName || ''}`);
    }

    if (isBlank_(status)) {
      pushError('確認', 'ステータスが空欄です。集計対象として扱うか確認してください。', '確認中', `部署：${dept || ''}／対象者：${targetName || ''}`);
    }

    if (!isBlank_(approveDateTime) && !parsedApprove) {
      pushError('エラー', '承認日時が日時として読み取れません。', '未対応', `部署：${dept || ''}／対象者：${targetName || ''}`);
    }

    if (parsedApply && parsedApprove && parsedApprove.getTime() < parsedApply.getTime()) {
      pushError('確認', '承認日時が申請日時より前になっています。', '確認中', `部署：${dept || ''}／対象者：${targetName || ''}`);
    }

    if (isBlank_(approveDateTime)) {
      pushError('確認', '承認日時が空欄です。承認待ちまたは未承認の可能性があります。', '確認中', `ステータス：${status || ''}／対象者：${targetName || ''}`);
    }

    const category = getDeptCategory_(String(dept || ''), deptMaster);

    if (category === TS_CONFIG.UNCLASSIFIED_CATEGORY) {
      pushError('確認', '部署マスタ未登録のため、未分類として扱われます。', '未対応', `部署：${dept || ''}`);
    }

    if (!isBlank_(targetCode) && parsedDate && parsedApply) {
      const key = [
        targetCode,
        formatDateForKey_(parsedDate),
        formatDateTimeForKey_(parsedApply),
        applyType || ''
      ].join('|');

      if (seenKeys[key]) {
        pushError('確認', '取込キーが重複しています。', '確認中', `初出行：${seenKeys[key]}／対象者：${targetName || ''}`);
      } else {
        seenKeys[key] = displayRowNumber;
      }
    }
  });

  return errors;
}

/**
エラー一覧書き込み
*/
function writeErrorSheet_(errors, targetMonth) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TS_CONFIG.SHEETS.ERRORS);
  const header = ['検出日時', '対象年月', '行番号', '区分', '内容', '対応状況', 'メモ'];
  const output = [header];

  if (errors.length === 0) {
    output.push([new Date(), targetMonth, '-', '正常', 'エラー・確認事項はありません。', '対応不要', '']);
  } else {
    errors.forEach(row => output.push(row));
  }

  writeTable_(sheet, output, {});
  sheet.getRange('A:A').setNumberFormat('yyyy/mm/dd hh:mm');
}


/**
CSV由来の文字列が数式として解釈されないよう、TeamSpirit原本列をプレーンテキスト形式にする。
原本列以外の補助列は日付・数値として扱うため、ここでは書式を変えない。
*/
function prepareRawAccumSheetForWrite_(sheet, startRow, rowCount, columnCount) {
  if (!sheet || rowCount <= 0 || columnCount <= 0) {
    return;
  }

  const textColumnCount = Math.min(RAW_TEXT_COLUMN_COUNT, columnCount);
  if (textColumnCount > 0) {
    sheet.getRange(startRow, 1, rowCount, textColumnCount).setNumberFormat('@');
  }
}

/**
集計・ダッシュボード等の表示用シートで、CSV由来文字列の数式インジェクションを防ぐ。
*/

/**
原本列（A:P）に限定して数式インジェクション対策を行う。
補助列のDate/Boolean/Numberは型を維持し、DateオブジェクトをそのままsetValues()へ渡す。
*/
function sanitizeRawTextColumnsForSheet_(values, rawColumnCount) {
  return (values || []).map(row => (row || []).map((value, index) => {
    if (index < rawColumnCount) {
      return sanitizeValueForSheet_(value);
    }
    return value;
  }));
}

function sanitizeValuesForSheet_(values) {
  return (values || []).map(row => (row || []).map(sanitizeValueForSheet_));
}

function sanitizeValueForSheet_(value) {
  if (typeof value === 'string' && DANGEROUS_SHEET_TEXT_PREFIX.test(value)) {
    return "'" + value;
  }
  return value;
}

/**
表書き込み共通
*/
function writeTable_(sheet, values, options) {
  sheet.clearContents();

  if (!values || values.length === 0) {
    return;
  }

  const padded = padRows_(values);

  ensureSheetSize_(sheet, padded.length, padded[0].length);
  sheet.getRange(1, 1, padded.length, padded[0].length).setValues(sanitizeValuesForSheet_(padded));
  formatHeaderRow_(sheet, 1, padded[0].length);
  sheet.setFrozenRows(1);

  if (options && options.percentColumns) {
    options.percentColumns.forEach(col => {
      if (padded.length > 1) {
        sheet.getRange(2, col, padded.length - 1, 1).setNumberFormat('0.0%');
      }
    });
  }

  sheet.setColumnWidths(1, padded[0].length, 120);
}

/**
行の列数をそろえる
*/
function padRows_(values) {
  const maxCols = Math.max(...values.map(row => row.length));

  return values.map(row => {
    const copy = row.slice();

    while (copy.length < maxCols) {
      copy.push('');
    }

    return copy;
  });
}

/**
シートサイズ確保
*/
function ensureSheetSize_(sheet, requiredRows, requiredCols) {
  const currentRows = sheet.getMaxRows();
  const currentCols = sheet.getMaxColumns();

  if (currentRows < requiredRows) {
    sheet.insertRowsAfter(currentRows, requiredRows - currentRows);
  }

  if (currentCols < requiredCols) {
    sheet.insertColumnsAfter(currentCols, requiredCols - currentCols);
  }
}

/**
部署マスタ取得
*/
function getDeptMaster_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TS_CONFIG.SHEETS.DEPT_MASTER);
  const values = sheet.getDataRange().getValues();
  const map = {};

  for (let r = 1; r < values.length; r++) {
    const deptName = String(values[r][0] || '').trim();

    if (!deptName) {
      continue;
    }

    const deptKey = normalizeDeptName_(deptName);

    map[deptKey] = {
      originalName: deptName,
      category: String(values[r][1] || '').trim(),
      closingTime: String(values[r][3] || '').trim()
    };
  }

  return map;
}

/**
部署名正規化
*/
function normalizeDeptName_(deptName) {
  return String(deptName || '')
    .normalize('NFKC')
    .replace(/\s/g, '')
    .trim();
}

/**
部署名の表示名を部署マスタ側の正式表記に寄せる。
全角・半角カナなどの表記揺れがあっても、集計上は同一部署として扱う。
*/
function getCanonicalDeptName_(deptName, deptMaster) {
  const normalizedDeptName = normalizeDeptName_(deptName);

  if (normalizedDeptName && deptMaster[normalizedDeptName] && deptMaster[normalizedDeptName].originalName) {
    return deptMaster[normalizedDeptName].originalName;
  }

  return String(deptName || '').trim();
}

/**
部門区分判定
*/
function getDeptCategory_(deptName, deptMaster) {
  const normalizedDeptName = normalizeDeptName_(deptName);

  if (!normalizedDeptName) {
    return TS_CONFIG.UNCLASSIFIED_CATEGORY;
  }

  if (deptMaster[normalizedDeptName] && deptMaster[normalizedDeptName].category) {
    return deptMaster[normalizedDeptName].category;
  }

  if (REGEX_SALES_DEPT.test(normalizedDeptName)) {
    return TS_CONFIG.SALES_CATEGORY;
  }

  return TS_CONFIG.UNCLASSIFIED_CATEGORY;
}

/**
定時取得
*/
function getClosingTimeForDept_(deptName, deptMaster, defaultClosingTime) {
  const normalizedDeptName = normalizeDeptName_(deptName);

  if (deptMaster[normalizedDeptName] && deptMaster[normalizedDeptName].closingTime) {
    return deptMaster[normalizedDeptName].closingTime;
  }

  return String(defaultClosingTime || TS_CONFIG.DEFAULT_CLOSING_TIME);
}

/**
設定値まとめ取得
*/
function getSettings_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TS_CONFIG.SHEETS.SETTINGS);
  const settings = {
    targetMonth: '',
    targetWeek: '',
    weekStart: '',
    weekEnd: '',
    importType: '速報',
    closingTime: TS_CONFIG.DEFAULT_CLOSING_TIME,
    alertThreshold: 0.8,
    lastImportTime: '',
    dashboardNote: ''
  };

  if (!sheet) {
    return settings;
  }

  const values = sheet.getDataRange().getValues();

  for (let r = 0; r < values.length; r++) {
    const key = String(values[r][0] || '').trim();
    const val = values[r][1];

    if (key === '対象年月') settings.targetMonth = String(val || '').trim();
    if (key === '対象週') settings.targetWeek = String(val || '').trim();
    if (key === '週開始日') settings.weekStart = val;
    if (key === '週終了日') settings.weekEnd = val;
    if (key === '取込区分') settings.importType = val;
    if (key === '定時時刻') settings.closingTime = val || TS_CONFIG.DEFAULT_CLOSING_TIME;
    if (key === '承認率注意基準') settings.alertThreshold = parseRate_(val, 0.8);
    if (key === '最終取込日時') settings.lastImportTime = val;
    if (key === 'ダッシュボード注記') settings.dashboardNote = val;
  }

  return settings;
}

/**
設定値取得
*/
function getSettingValue_(settingName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TS_CONFIG.SHEETS.SETTINGS);

  if (!sheet) {
    return '';
  }

  const values = sheet.getDataRange().getValues();

  for (let r = 0; r < values.length; r++) {
    if (String(values[r][0]).trim() === settingName) {
      return values[r][1];
    }
  }

  return '';
}

/**
設定値更新
*/
function setSettingValue_(settingName, value) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TS_CONFIG.SHEETS.SETTINGS);

  if (!sheet) {
    throw new Error(`シートが見つかりません：${TS_CONFIG.SHEETS.SETTINGS}`);
  }

  const values = sheet.getDataRange().getValues();

  for (let r = 0; r < values.length; r++) {
    if (String(values[r][0]).trim() === settingName) {
      sheet.getRange(r + 1, 2).setValue(value);
      return;
    }
  }

  sheet.appendRow([settingName, value, '']);
}

/**
取込ログ追記
*/
function appendImportLog_(log) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TS_CONFIG.SHEETS.IMPORT_LOG);

  if (!sheet) {
    return;
  }

  const headerNeedsSetup =
    sheet.getLastRow() === 0 ||
    isBlank_(sheet.getRange(1, 1).getValue()) ||
    sheet.getLastColumn() < 11 ||
    String(sheet.getRange(1, 4).getValue() || '').trim() !== '対象週';

  if (headerNeedsSetup) {
    ensureSheetSize_(sheet, Math.max(sheet.getMaxRows(), 1), 11);
    sheet.getRange(1, 1, 1, 11).setValues([[
      '取込日時',
      '取込区分',
      '対象年月',
      '対象週',
      'ファイル名',
      '取込方法',
      '取込件数',
      '結果',
      'メモ',
      'メールID',
      'ファイルID'
    ]]);
    formatHeaderRow_(sheet, 1, 11);
  }

  sheet.appendRow([
    new Date(),
    log.importType || '',
    log.targetMonth || '',
    log.targetWeek || '',
    log.fileName || '',
    log.importMethod || '',
    log.importCount || '',
    log.result || '',
    log.memo || '',
    log.mailId || '',
    log.fileId || ''
  ]);

  sheet.getRange('A:A').setNumberFormat('yyyy/mm/dd hh:mm');
}

/**
スプレッドシートのタイムゾーン補正
*/
function ensureSpreadsheetTimeZone_(ss) {
  const current = ss.getSpreadsheetTimeZone();

  if (current !== TS_CONFIG.TIMEZONE) {
    ss.setSpreadsheetTimeZone(TS_CONFIG.TIMEZONE);
  }
}

/**
日付パース
*/
function parseDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return value;
  }

  if (isBlank_(value)) {
    return null;
  }

  if (typeof value === 'number') {
    const date = new Date(Math.round((value - 25569) * 86400 * 1000));
    return new Date(
      date.getUTCFullYear(),
      date.getUTCMonth(),
      date.getUTCDate(),
      date.getUTCHours(),
      date.getUTCMinutes(),
      date.getUTCSeconds()
    );
  }

  const text = String(value)
    .trim()
    .replace(/年/g, '/')
    .replace(/月/g, '/')
    .replace(/日/g, '')
    .replace(/-/g, '/');

  const match = text.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/);

  if (match) {
    return new Date(
      Number(match[1]),
      Number(match[2]) - 1,
      Number(match[3]),
      Number(match[4] || 0),
      Number(match[5] || 0),
      Number(match[6] || 0)
    );
  }

  const fallback = new Date(text);

  if (!isNaN(fallback.getTime())) {
    return fallback;
  }

  return null;
}

/**
時刻パース
*/
function parseTime_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return { hour: value.getHours(), minute: value.getMinutes() };
  }

  const text = String(value || TS_CONFIG.DEFAULT_CLOSING_TIME).trim();
  const match = text.match(/(\d{1,2}):(\d{1,2})/);

  if (!match) {
    return { hour: 17, minute: 30 };
  }

  return {
    hour: Number(match[1]),
    minute: Number(match[2])
  };
}

/**
日付＋時刻
*/
function combineDateAndTime_(date, timeValue) {
  const time = parseTime_(timeValue);

  return new Date(
    date.getFullYear(),
    date.getMonth(),
    date.getDate(),
    time.hour,
    time.minute,
    0
  );
}


/**
週情報取得（月曜始まり）
*/
function getWeekInfo_(date) {
  const d = parseDate_(date);
  if (!d) return { weekKey: '', weekStart: null, weekEnd: null };
  const base = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  const day = base.getDay();
  const diffToMonday = day === 0 ? -6 : 1 - day;
  const weekStart = new Date(base.getFullYear(), base.getMonth(), base.getDate() + diffToMonday);
  const weekEnd = new Date(weekStart.getFullYear(), weekStart.getMonth(), weekStart.getDate() + 6);
  return { weekKey: formatWeekKey_(weekStart), weekStart, weekEnd };
}

function formatWeekKey_(weekStart) {
  return Utilities.formatDate(weekStart, TS_CONFIG.TIMEZONE, 'yyyy-MM-dd') + '週';
}

function getPreviousWeekKey_(weekKey) {
  if (!weekKey) return '';
  const text = String(weekKey).replace('週', '');
  const start = parseDate_(text);
  if (!start) return '';
  const prevStart = new Date(start.getFullYear(), start.getMonth(), start.getDate() - 7);
  return formatWeekKey_(prevStart);
}

function getPreviousMonthKey_(monthKey) {
  if (!monthKey || !/^\d{4}-\d{2}$/.test(String(monthKey))) return '';
  const year = Number(String(monthKey).slice(0, 4));
  const month = Number(String(monthKey).slice(5, 7));
  const d = new Date(year, month - 2, 1);
  return Utilities.formatDate(d, TS_CONFIG.TIMEZONE, 'yyyy-MM');
}

function formatDateForDisplay_(date) {
  if (!date) return '';
  return Utilities.formatDate(date, TS_CONFIG.TIMEZONE, 'yyyy/MM/dd');
}

/**
日付キー
*/
function formatDateForKey_(date) {
  if (!date) {
    return '';
  }

  return Utilities.formatDate(date, TS_CONFIG.TIMEZONE, 'yyyy/MM/dd');
}

/**
日時キー
*/
function formatDateTimeForKey_(date) {
  if (!date) {
    return '';
  }

  return Utilities.formatDate(date, TS_CONFIG.TIMEZONE, 'yyyy/MM/dd HH:mm:ss');
}

/**
除外ステータス判定
*/
function isExcludedStatus_(status) {
  const text = String(status || '');
  return /取消|取り消し|却下/.test(text);
}

/**
安全な率計算
*/
function safeRate_(numerator, denominator) {
  const n = Number(numerator || 0);
  const d = Number(denominator || 0);

  if (d === 0) {
    return 0;
  }

  return n / d;
}


/**
増減率
*/
function safeGrowthRate_(currentValue, previousValue) {
  const current = Number(currentValue || 0);
  const previous = Number(previousValue || 0);
  if (previous === 0) {
    return current === 0 ? 0 : 1;
  }
  return (current - previous) / previous;
}

/**
設定値の率パース
*/
function parseRate_(value, defaultValue) {
  if (typeof value === 'number' && !isNaN(value)) {
    return value;
  }

  const text = String(value || '').trim();

  if (!text) {
    return defaultValue;
  }

  if (text.endsWith('%')) {
    const n = Number(text.replace('%', ''));

    if (!isNaN(n)) {
      return n / 100;
    }
  }

  const n = Number(text);

  if (!isNaN(n)) {
    return n;
  }

  return defaultValue;
}

/**
空欄判定
*/
function isBlank_(value) {
  return value === null || value === undefined || String(value).trim() === '';
}

/**
ヘッダー行の共通書式
*/
function formatHeaderRow_(sheet, rowNumber, columnCount) {
  if (!sheet || !rowNumber || !columnCount) {
    return;
  }

  sheet
    .getRange(rowNumber, 1, 1, columnCount)
    .setFontWeight('bold')
    .setBackground('#17466f')
    .setFontColor('#ffffff');
}

/**
カテゴリ合計件数
*/
function getCategoryTotalCount_(deptRows, category) {
  return deptRows
    .filter(row => row[1] === category)
    .reduce((sum, row) => sum + Number(row[2] || 0), 0);
}


/**
HTMLダッシュボードを表示
*/
function showHtmlDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSpreadsheetTimeZone_(ss);
  ensureBaseSheets_(ss);

  const html = HtmlService
    .createHtmlOutput(HTML_DASHBOARD)
    .setWidth(1280)
    .setHeight(820);

  SpreadsheetApp.getUi().showModalDialog(html, 'TeamSpirit 残業申請・承認ダッシュボード');
}

/**
HTMLダッシュボード用データ取得
*/

/**
HTMLダッシュボード用データ取得
画面が「読み込んでいます」のままにならないよう、サーバー側エラーも画面へ返す。
*/
function getHtmlDashboardData() {
  const startedAt = new Date();
  let payload;
  try {
    payload = getHtmlDashboardDataCore_();
    payload.serverElapsedMs = new Date().getTime() - startedAt.getTime();
  } catch (error) {
    let settings = {};
    let errorCount = 0;
    try {
      settings = buildDashboardSettingsPayload_(getSettings_());
    } catch (e) {
      settings = {};
    }
    try {
      errorCount = getCurrentErrorCount_();
    } catch (e) {
      errorCount = 0;
    }
    payload = {
      ok: false,
      message: 'ダッシュボードデータ作成時にエラーが発生しました。\n\n' +
        (error && error.message ? error.message : String(error)),
      settings: settings,
      latestLog: {},
      errorCount: errorCount,
      serverElapsedMs: new Date().getTime() - startedAt.getTime()
    };
  }

  // google.script.run can silently return an empty value when a complex server-side
  // object contains a value that cannot be marshalled to the iframe.  Serialize the
  // payload explicitly so the dashboard always receives a deterministic response.
  return JSON.stringify(payload);
}

function getHtmlDashboardDataCore_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSpreadsheetTimeZone_(ss);
  ensureBaseSheets_(ss);

  const settings = getSettings_();
  const accumInfo = getAccumulatedDataInfo_();

  const latestLog = getLatestImportLog_();
  const errorCount = getCurrentErrorCount_();

  if (!accumInfo.rows || accumInfo.rows.length === 0) {
    return {
      ok: false,
      message: '取込データ_残業申請にデータがありません。先にCSVファイルを取り込んでください。',
      settings: buildDashboardSettingsPayload_(settings),
      latestLog: latestLog,
      errorCount: errorCount
    };
  }

  const allSummaries = buildMonthlyWeeklySummaryBundle_(accumInfo.rows, accumInfo.headers, settings);

  return {
    ok: true,
    generatedAt: Utilities.formatDate(new Date(), TS_CONFIG.TIMEZONE, 'yyyy/MM/dd HH:mm:ss'),
    settings: buildDashboardSettingsPayload_(settings),
    latestLog: latestLog,
    errorCount: errorCount,
    monthly: {
      label: allSummaries.monthly.currentLabel,
      previousLabel: allSummaries.monthly.previousLabel,
      staff: buildDashboardCategoryPayload_(
        allSummaries.monthly.current.deptRows,
        allSummaries.monthly.previous.deptRows,
        TS_CONFIG.STAFF_CATEGORY
      ),
      sales: buildDashboardCategoryPayload_(
        allSummaries.monthly.current.deptRows,
        allSummaries.monthly.previous.deptRows,
        TS_CONFIG.SALES_CATEGORY
      )
    },
    weekly: {
      label: allSummaries.weekly.currentLabel,
      previousLabel: allSummaries.weekly.previousLabel,
      staff: buildDashboardCategoryPayload_(
        allSummaries.weekly.current.deptRows,
        allSummaries.weekly.previous.deptRows,
        TS_CONFIG.STAFF_CATEGORY
      ),
      sales: buildDashboardCategoryPayload_(
        allSummaries.weekly.current.deptRows,
        allSummaries.weekly.previous.deptRows,
        TS_CONFIG.SALES_CATEGORY
      )
    }
  };
}

/**
HTMLダッシュボード設定情報
*/
function buildDashboardSettingsPayload_(settings) {
  return {
    targetMonth: settings.targetMonth || '',
    targetWeek: settings.targetWeek || '',
    weekStart: settings.weekStart ? formatDateForDisplay_(settings.weekStart) : '',
    weekEnd: settings.weekEnd ? formatDateForDisplay_(settings.weekEnd) : '',
    importType: settings.importType || '',
    lastImportTime: settings.lastImportTime ? formatDateTimeForDisplay_(settings.lastImportTime) : '',
    alertThreshold: parseRate_(settings.alertThreshold, 0.8),
    dashboardNote: settings.dashboardNote || ''
  };
}

/**
HTMLダッシュボードカテゴリ別ペイロード
*/
function buildDashboardCategoryPayload_(currentDeptRows, previousDeptRows, category) {
  const currentRows = currentDeptRows.filter(row => row[1] === category);
  const previousRows = previousDeptRows.filter(row => row[1] === category);
  const previousMap = buildDeptRowMap_(previousRows);

  const currentTotals = calcTotals_(currentRows);
  const previousTotals = calcTotals_(previousRows);

  const currentSummary = buildDashboardTotalPayload_(currentTotals);
  const previousSummary = buildDashboardTotalPayload_(previousTotals);

  const details = currentRows
    .slice()
    .sort((a, b) => a[8] - b[8])
    .map(row => {
      const prev = previousMap[row[0]] || emptyDeptSummaryRow_(row[0], category);

      return {
        deptName: row[0],
        category: row[1],
        count: Number(row[2] || 0),
        previousCount: Number(prev[2] || 0),
        countDiff: Number(row[2] || 0) - Number(prev[2] || 0),
        countGrowth: safeGrowthRate_(row[2], prev[2]),
        beforeApplyRate: Number(row[7] || 0),
        beforeApplyRatePrev: Number(prev[7] || 0),
        beforeApplyRateDiff: Number(row[7] || 0) - Number(prev[7] || 0),
        beforeApproveRate: Number(row[8] || 0),
        beforeApproveRatePrev: Number(prev[8] || 0),
        beforeApproveRateDiff: Number(row[8] || 0) - Number(prev[8] || 0),
        nextDayApproveRate: Number(row[9] || 0),
        nextDayApproveRatePrev: Number(prev[9] || 0),
        nextDayApproveRateDiff: Number(row[9] || 0) - Number(prev[9] || 0),
        notApprovedCount: Number(row[10] || 0),
        notApprovedCountPrev: Number(prev[10] || 0),
        notApprovedCountDiff: Number(row[10] || 0) - Number(prev[10] || 0),
        alert: row[11] || ''
      };
    });

  return {
    category: category,
    current: currentSummary,
    previous: previousSummary,
    diff: {
      count: currentSummary.count - previousSummary.count,
      countGrowth: safeGrowthRate_(currentSummary.count, previousSummary.count),
      beforeApplyRate: currentSummary.beforeApplyRate - previousSummary.beforeApplyRate,
      beforeApproveRate: currentSummary.beforeApproveRate - previousSummary.beforeApproveRate,
      nextDayApproveRate: currentSummary.nextDayApproveRate - previousSummary.nextDayApproveRate,
      notApprovedCount: currentSummary.notApprovedCount - previousSummary.notApprovedCount
    },
    details: details
  };
}

/**
HTMLダッシュボード合計値
*/
function buildDashboardTotalPayload_(totals) {
  const count = Number(totals.count || 0);
  const approveDate = Number(totals.approveDate || 0);

  return {
    count: count,
    beforeApplyRate: safeRate_(totals.beforeApply, count),
    beforeApproveRate: safeRate_(totals.beforeApprove, count),
    nextDayApproveRate: safeRate_(totals.nextDayApprove, count),
    notApprovedCount: count - approveDate
  };
}

/**
最新取込ログ取得
*/
function getLatestImportLog_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TS_CONFIG.SHEETS.IMPORT_LOG);

  if (!sheet || sheet.getLastRow() < 2) {
    return {};
  }

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(v => String(v || '').trim());
  const last = values[values.length - 1];
  const index = {};
  headers.forEach((h, i) => index[h] = i);

  return {
    importTime: formatDateTimeForDisplay_(last[index['取込日時']]),
    importType: last[index['取込区分']] || '',
    targetMonth: last[index['対象年月']] || '',
    targetWeek: last[index['対象週']] || '',
    fileName: last[index['ファイル名']] || '',
    importMethod: last[index['取込方法']] || '',
    importCount: last[index['取込件数']] || '',
    result: last[index['結果']] || '',
    memo: last[index['メモ']] || ''
  };
}

/**
現在のエラー・確認件数
必要な列だけを読む。
*/
function getCurrentErrorCount_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TS_CONFIG.SHEETS.ERRORS);

  if (!sheet || sheet.getLastRow() < 2) {
    return 0;
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastCol < 4) {
    return 0;
  }

  const types = sheet.getRange(2, 4, lastRow - 1, 1).getValues();
  let count = 0;

  types.forEach(row => {
    const type = String(row[0] || '');
    if (type && type !== '正常') {
      count++;
    }
  });

  return count;
}

/**
日時表示
*/
function formatDateTimeForDisplay_(value) {
  const date = parseDate_(value);

  if (!date) {
    return value ? String(value) : '';
  }

  return Utilities.formatDate(date, TS_CONFIG.TIMEZONE, 'yyyy/MM/dd HH:mm:ss');
}
