# 1ファイル登録時のビュワーモード対応 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: このプロジェクトのCLAUDE.md指示により、実装はサブエージェント "implementer-codex" に委譲し、実装完了後にサブエージェント "plan-reviewer-codex" にレビュー済みの本計画を渡して進めること。通常の superpowers:subagent-driven-development / superpowers:executing-plans は使わない。Steps はチェックボックス（`- [ ]`）で進捗管理する。

**Goal:** `excel-diff.html` / `access-diff.html` / `word-diff.html` の3ツールで、ファイルAまたはBのどちらか一方だけが登録された状態でも、そのファイルの中身をそのまま閲覧できる「ビュワーモード」として機能するようにする。2ファイル目を追加すれば既存の差分比較モードに自動的に切り替わる。

**Architecture:** 「読み込まれていない側のファイルを、読み込み済みのファイルと同一視する」というアプローチを取る。具体的には各ツールに `effA()` / `effB()`（存在する方のファイルを返すヘルパー）を導入し、既存の比較計算（`compute()`）・描画（`render()`）関数の中で `state.A` / `state.B` を直接参照している箇所を `effA()` / `effB()` に置き換える。1ファイルのみの場合は `effA() === effB()`（同一データ）になるため、既存の比較ロジックが自然に「全セル一致」の結果を返し、追加の描画コンポーネントなしで通常表示になる。比較専用のUI部品（サマリー件数・警告・キー選択・表示モード切替など）は `isSingle()` の判定に応じて表示・非表示を切り替える。

**Tech Stack:** 素のHTML/CSS/JavaScript（ビルドツールなし、単一HTMLファイルに完結）。自動テストのインフラは無いため、検証は実ファイルを用いたブラウザ手動確認で行う。

## Global Constraints

- 各ファイルは単一HTMLファイルのまま完結させる。外部依存・別ファイル分割は増やさない。
- 既存のライブラリ埋め込み部分（SheetJS / mdb-reader / mammoth.js / jsdiff、各ファイル冒頭の圧縮コード行）は変更しない。
- Cookie/localStorageの設定キー名（`exceldiff` / `accessdiff` / `worddiff`）とその保存構造は変更しない。
- UI文言はすべて日本語のまま。
- 2ファイル揃った状態での既存の差分比較の挙動・見た目は変更しない（デグレ禁止）。
- 新規の専用ビューコンポーネントは作らず、既存の比較テーブル／本文描画を流用する（設計書の方針）。

---

## Task 1: excel-diff.html — 1ファイル登録時のビュワー対応

**Files:**
- Modify: `excel-diff.html`（HTML: 167行目付近, 179行目付近／CSS: 47行目付近／JS: 335-613行目のアプリロジック）

**Interfaces:**
- Produces: `effA()`, `effB()`, `isSingle()`, `updateModeUI()` — Task 2, 3では使わないが、同名パターンをAccess/Wordでも独自に実装する（ファイルが分かれているため共有はしない）。

- [ ] **Step 1: HTMLに比較専用UI要素のID付与とプレースホルダ文言修正**

`excel-diff.html:167` を次のように変更する（`class="modebar"` に `id="modebar"` を追加）:

変更前:
```html
    <div class="modebar">
      <label class="check"><input type="checkbox" id="manualPair" disabled> 比較するシートを選ぶ</label>
      <span class="modehint">オフ＝同名シートを比較</span>
    </div>
    <div class="tabs" id="tabs"><span class="tabhint">2ファイルを読み込むとシートが表示されます</span></div>
```

変更後:
```html
    <div class="modebar" id="modebar">
      <label class="check"><input type="checkbox" id="manualPair" disabled> 比較するシートを選ぶ</label>
      <span class="modehint">オフ＝同名シートを比較</span>
    </div>
    <div class="tabs" id="tabs"><span class="tabhint">ファイルを読み込むとシートが表示されます</span></div>
```

`excel-diff.html:179` を次のように変更する（`diffOnly` の `label` に `id="diffOnlyRow"` を追加）:

変更前:
```html
      <label class="check"><input type="checkbox" id="diffOnly"> 差分のある行のみ表示</label>
```

変更後:
```html
      <label class="check" id="diffOnlyRow"><input type="checkbox" id="diffOnly"> 差分のある行のみ表示</label>
```

`excel-diff.html:210` を次のように変更する（初期プレースホルダ文言）:

変更前:
```html
  <div class="tablehost" id="tableHost"><div class="empty">2つのファイルを読み込み、シートを選ぶと差分が表示されます。</div></div>
```

変更後:
```html
  <div class="tablehost" id="tableHost"><div class="empty">ファイルを読み込み、シートを選ぶと内容が表示されます。</div></div>
```

- [ ] **Step 2: CSSに汎用hiddenルールを追加**

`excel-diff.html:47` の既存行の直後に1行追加する:

変更前:
```css
.tabs.hidden,.pairsel.hidden{display:none}
```

変更後:
```css
.tabs.hidden,.pairsel.hidden{display:none}
.hidden{display:none}
```

- [ ] **Step 3: JSにヘルパー関数を追加**

`excel-diff.html:343`（`var $ = function (id) { return document.getElementById(id); };` の直後）に追加する:

変更前:
```javascript
  var $ = function (id) { return document.getElementById(id); };

  function saveSettings() {
```

変更後:
```javascript
  var $ = function (id) { return document.getElementById(id); };
  function effA() { return state.A || state.B; }
  function effB() { return state.B || state.A; }
  function isSingle() { return !!(state.A || state.B) && !(state.A && state.B); }
  function updateModeUI() {
    var single = isSingle();
    $("modebar").classList.toggle("hidden", single);
    $("diffOnlyRow").classList.toggle("hidden", single);
  }

  function saveSettings() {
```

- [ ] **Step 4: refreshAll() を単一ファイル対応にする**

`excel-diff.html:396-411` を次のように変更する:

変更前:
```javascript
  function refreshAll() {
    var tabsHost = $("tabs");
    if (!state.A || !state.B) {
      tabsHost.innerHTML = '<span class="tabhint">2ファイルを読み込むとシートが表示されます</span>';
      $("pairSel").classList.add("hidden"); $("manualPair").disabled = true;
      state.sheet = null; clearResult(); return;
    }
    $("manualPair").disabled = false;
    var common = intersect(state.A.sheets, state.B.sheets);
    state.sheetA = fillSelect($("sheetASel"), state.A.sheets, firstIn([pending && pending.sheetA, state.sheetA, common[0], state.A.sheets[0]], state.A.sheets));
    state.sheetB = fillSelect($("sheetBSel"), state.B.sheets, firstIn([pending && pending.sheetB, state.sheetB, common[0], state.B.sheets[0]], state.B.sheets));
    buildTabs(common, firstIn([pending && pending.sheet, state.sheet, common[0]], common));
    pending = null;
    toggleMode();
    computeActive();
  }
```

変更後:
```javascript
  function refreshAll() {
    updateModeUI();
    var tabsHost = $("tabs");
    if (!state.A && !state.B) {
      tabsHost.innerHTML = '<span class="tabhint">ファイルを読み込むとシートが表示されます</span>';
      $("pairSel").classList.add("hidden"); $("manualPair").disabled = true;
      state.sheet = null; clearResult(); return;
    }
    var single = isSingle();
    $("manualPair").disabled = single;
    if (single) state.manual = false;
    var common = intersect(effA().sheets, effB().sheets);
    state.sheetA = fillSelect($("sheetASel"), effA().sheets, firstIn([pending && pending.sheetA, state.sheetA, common[0], effA().sheets[0]], effA().sheets));
    state.sheetB = fillSelect($("sheetBSel"), effB().sheets, firstIn([pending && pending.sheetB, state.sheetB, common[0], effB().sheets[0]], effB().sheets));
    buildTabs(common, firstIn([pending && pending.sheet, state.sheet, common[0]], common));
    pending = null;
    toggleMode();
    computeActive();
  }
```

- [ ] **Step 5: buildTabs() を effA/effB ベースにする**

`excel-diff.html:412-428` を次のように変更する:

変更前:
```javascript
  function buildTabs(common, want) {
    var host = $("tabs"); host.innerHTML = "";
    var union = [], seen = {};
    state.A.sheets.concat(state.B.sheets).forEach(function (n) { if (!seen[n]) { seen[n] = 1; union.push(n); } });
    var inA = mapOf(state.A.sheets), inB = mapOf(state.B.sheets);
```

変更後:
```javascript
  function buildTabs(common, want) {
    var host = $("tabs"); host.innerHTML = "";
    var union = [], seen = {};
    effA().sheets.concat(effB().sheets).forEach(function (n) { if (!seen[n]) { seen[n] = 1; union.push(n); } });
    var inA = mapOf(effA().sheets), inB = mapOf(effB().sheets);
```

（以降の `buildTabs` 内のロジックは変更しない）

- [ ] **Step 6: currentPair() を単一ファイル対応にする**

`excel-diff.html:444-448` を次のように変更する:

変更前:
```javascript
  function currentPair() {
    if (!state.A || !state.B) return null;
    if (state.manual) { if (!state.sheetA || !state.sheetB) return null; return { a: state.sheetA, b: state.sheetB, manual: true }; }
    if (!state.sheet) return null; return { a: state.sheet, b: state.sheet, manual: false };
  }
```

変更後:
```javascript
  function currentPair() {
    if (!state.A && !state.B) return null;
    if (isSingle()) { if (!state.sheet) return null; return { a: state.sheet, b: state.sheet, manual: false }; }
    if (state.manual) { if (!state.sheetA || !state.sheetB) return null; return { a: state.sheetA, b: state.sheetB, manual: true }; }
    if (!state.sheet) return null; return { a: state.sheet, b: state.sheet, manual: false };
  }
```

- [ ] **Step 7: compute() を effA/effB ベースにし、single フラグを結果に含める**

`excel-diff.html:466-505` の該当行を変更する。まず469行目:

変更前:
```javascript
    var A = sheetModel(state.A.wb, p.a), B = sheetModel(state.B.wb, p.b);
```

変更後:
```javascript
    var A = sheetModel(effA().wb, p.a), B = sheetModel(effB().wb, p.b);
```

次に502-503行目（`state.result` の代入）:

変更前:
```javascript
    state.result = { rows: rows, cols: cols, cells: cells, rowHasDiff: rowHasDiff, counts: counts, mismatch: mismatch, pair: p,
      hasMerge: A.mergeCount > 0 || B.mergeCount > 0, mergeA: A.mergeCount, mergeB: B.mergeCount, dimA: { r: A.rows, c: A.cols }, dimB: { r: B.rows, c: B.cols } };
```

変更後:
```javascript
    state.result = { rows: rows, cols: cols, cells: cells, rowHasDiff: rowHasDiff, counts: counts, mismatch: mismatch, pair: p,
      hasMerge: A.mergeCount > 0 || B.mergeCount > 0, mergeA: A.mergeCount, mergeB: B.mergeCount, dimA: { r: A.rows, c: A.cols }, dimB: { r: B.rows, c: B.cols },
      single: isSingle() };
```

- [ ] **Step 8: render() でサマリー・警告・diffOnlyフィルタを single 対応にする**

`excel-diff.html:507-529` を次のように変更する:

変更前:
```javascript
  function render() {
    if (!state.result) return;
    var R = state.result, c = R.counts;
    $("sumSame").textContent = c.same; $("sumAdd").textContent = c.added; $("sumRem").textContent = c.removed; $("sumChg").textContent = c.changed;
    $("summary").classList.remove("hidden");

    $("phA").textContent = state.A ? state.A.name : "";
    $("phB").textContent = state.B ? state.B.name : "";
    $("phSheet").textContent = R.pair.manual ? (R.pair.a + " ↔ " + R.pair.b) : R.pair.a;
    $("phCounts").textContent = "一致 " + c.same + " / 追加 " + c.added + " / 削除 " + c.removed + " / 変更 " + c.changed;
    $("phTime").textContent = new Date().toLocaleString("ja-JP");

    var warn = $("warn"), msgs = [];
    if (R.pair.manual && R.pair.a !== R.pair.b) msgs.push("異なる名前のシートを比較中: A「" + R.pair.a + "」 ↔ B「" + R.pair.b + "」");
    if (R.dimA.r !== R.dimB.r || R.dimA.c !== R.dimB.c) msgs.push("範囲が異なります（A: " + R.dimA.r + "行×" + R.dimA.c + "列 / B: " + R.dimB.r + "行×" + R.dimB.c + "列）。位置基準で比較しています。");
    if (R.hasMerge) msgs.push("結合セルを検出（A: " + R.mergeA + " / B: " + R.mergeB + "）。結合範囲は値を展開して比較しています。");
    if (R.mismatch > 0) msgs.push("⚠ 結合構造がA/Bで異なる箇所が " + R.mismatch + " セルあります（該当セルは印付きでセル単位表示）。");
    if (msgs.length) { warn.innerHTML = msgs.map(escapeHtml).join("<br>"); warn.classList.remove("hidden"); } else warn.classList.add("hidden");

    var mergeView = R.hasMerge && !state.diffOnly;
    var rowsToShow = [];
    for (var r = 0; r < R.rows; r++) { if (!state.diffOnly || R.rowHasDiff[r]) rowsToShow.push(r); }
```

変更後:
```javascript
  function render() {
    if (!state.result) return;
    var R = state.result, c = R.counts;
    if (R.single) {
      $("summary").classList.add("hidden");
    } else {
      $("sumSame").textContent = c.same; $("sumAdd").textContent = c.added; $("sumRem").textContent = c.removed; $("sumChg").textContent = c.changed;
      $("summary").classList.remove("hidden");
    }

    $("phA").textContent = state.A ? state.A.name : "";
    $("phB").textContent = state.B ? state.B.name : "";
    $("phSheet").textContent = R.pair.manual ? (R.pair.a + " ↔ " + R.pair.b) : R.pair.a;
    $("phCounts").textContent = R.single ? "" : ("一致 " + c.same + " / 追加 " + c.added + " / 削除 " + c.removed + " / 変更 " + c.changed);
    $("phTime").textContent = new Date().toLocaleString("ja-JP");

    var warn = $("warn"), msgs = [];
    if (!R.single) {
      if (R.pair.manual && R.pair.a !== R.pair.b) msgs.push("異なる名前のシートを比較中: A「" + R.pair.a + "」 ↔ B「" + R.pair.b + "」");
      if (R.dimA.r !== R.dimB.r || R.dimA.c !== R.dimB.c) msgs.push("範囲が異なります（A: " + R.dimA.r + "行×" + R.dimA.c + "列 / B: " + R.dimB.r + "行×" + R.dimB.c + "列）。位置基準で比較しています。");
      if (R.hasMerge) msgs.push("結合セルを検出（A: " + R.mergeA + " / B: " + R.mergeB + "）。結合範囲は値を展開して比較しています。");
      if (R.mismatch > 0) msgs.push("⚠ 結合構造がA/Bで異なる箇所が " + R.mismatch + " セルあります（該当セルは印付きでセル単位表示）。");
    }
    if (msgs.length) { warn.innerHTML = msgs.map(escapeHtml).join("<br>"); warn.classList.remove("hidden"); } else warn.classList.add("hidden");

    var mergeView = R.hasMerge && !state.diffOnly;
    var rowsToShow = [];
    for (var r = 0; r < R.rows; r++) { if (R.single || !state.diffOnly || R.rowHasDiff[r]) rowsToShow.push(r); }
```

（以降の行番号は変わるが、`render()` の残りのロジック — CAP処理・thead/tbody生成 — は変更しない）

- [ ] **Step 9: clearResult() のプレースホルダ文言を修正**

`excel-diff.html:570-575` 付近の該当行を変更する:

変更前:
```javascript
    $("tableHost").innerHTML = '<div class="empty">シートを選ぶと差分が表示されます。</div>';
```

変更後:
```javascript
    $("tableHost").innerHTML = '<div class="empty">シートを選ぶと内容が表示されます。</div>';
```

- [ ] **Step 10: 見出し・サブタイトルの文言更新**

`excel-diff.html:145` を変更する:

変更前:
```html
    <span class="sub">2つの .xlsx をシートごとにセル比較</span>
```

変更後:
```html
    <span class="sub">1つならビュワー、2つなら .xlsx をシートごとにセル比較</span>
```

- [ ] **Step 11: ブラウザで手動確認する**

1. `excel-diff.html` をブラウザで直接開く（`file://` で可）。
2. 複数シートを持つ `.xlsx` ファイルをファイルAのドロップ枠にのみ読み込む。
3. 確認項目:
   - タブにそのファイルの全シート名が表示され、クリックで切り替えられる。
   - 選択中シートの内容が通常の見た目（差分色なし）でグリッド表示される。
   - サマリー行（一致/追加/削除/変更）が非表示。
   - 「比較するシートを選ぶ」チェックボックス行が非表示。
   - 「差分のある行のみ表示」チェックボックス行が非表示。
   - PDF出力ボタンが有効で、クリックして印刷プレビューが出る。
   - 結合セルを含むシートで、結合がそのまま表示される。
4. ファイルBに2つ目の `.xlsx` を読み込み、既存通りの同名シート差分比較（サマリー・警告・チェックボックス行が再表示）に切り替わることを確認する。
5. ファイルAを外し、ファイルBのみの状態でも同様に単体ビュワー表示になることを確認する。
6. コンソールにJavaScriptエラーが出ていないことを確認する。

- [ ] **Step 12: コミット**

```bash
git add excel-diff.html
git commit -m "$(cat <<'EOF'
Add single-file viewer mode to Excel diff tool

Loading only one workbook now renders its sheets directly instead of
showing an empty placeholder, by treating the missing side as identical
to the loaded one so the existing diff rendering shows no highlights.
EOF
)"
```

---

## Task 2: access-diff.html — 1ファイル登録時のビュワー対応

**Files:**
- Modify: `access-diff.html`（HTML: 267-276行目付近, 291行目付近／CSS: 93行目付近／JS: 343-800行目のアプリロジック）

**Interfaces:**
- Produces: `effA()`, `effB()`, `isSingle()`, `updateModeUI()`（Task 1と同じ役割だが、このファイル内で独立して定義する）。

- [ ] **Step 1: HTMLに比較専用UI要素のID付与とプレースホルダ文言修正**

`access-diff.html:266-269` を次のように変更する:

変更前:
```html
      <div class="field">
        <label for="tableSel">テーブル（同名）</label>
        <select id="tableSel" disabled><option>— 2ファイルを読み込んでください —</option></select>
      </div>
```

変更後:
```html
      <div class="field">
        <label for="tableSel">テーブル（同名）</label>
        <select id="tableSel" disabled><option>— ファイルを読み込んでください —</option></select>
      </div>
```

`access-diff.html:270-276` を次のように変更する（キー列フィールドに `id="keyField"` を追加）:

変更前:
```html
      <div class="field">
        <label>キー列（複合可・順序＝選択順）</label>
        <div class="multi" id="keyMulti">
          <button class="multibtn" id="keyBtn" type="button" disabled>キー列を選択… <span class="caret">&#9662;</span></button>
          <div class="pop hidden" id="keyPop"></div>
        </div>
      </div>
```

変更後:
```html
      <div class="field" id="keyField">
        <label>キー列（複合可・順序＝選択順）</label>
        <div class="multi" id="keyMulti">
          <button class="multibtn" id="keyBtn" type="button" disabled>キー列を選択… <span class="caret">&#9662;</span></button>
          <div class="pop hidden" id="keyPop"></div>
        </div>
      </div>
```

`access-diff.html:291` を次のように変更する（`diffOnly` の `label` に `id="diffOnlyRow"` を追加）:

変更前:
```html
      <label class="check"><input type="checkbox" id="diffOnly"> 差分のみ表示</label>
```

変更後:
```html
      <label class="check" id="diffOnlyRow"><input type="checkbox" id="diffOnly"> 差分のみ表示</label>
```

`access-diff.html:318` を次のように変更する（初期プレースホルダ文言）:

変更前:
```html
    <div class="empty">2つのファイルを読み込み、テーブルを選ぶと差分が表示されます。</div>
```

変更後:
```html
    <div class="empty">ファイルを読み込み、テーブルを選ぶと内容が表示されます。</div>
```

- [ ] **Step 2: CSSに汎用hiddenルールを追加**

`access-diff.html:93` の既存行の直後に1行追加する:

変更前:
```css
.pop.hidden{display:none}
```

変更後:
```css
.pop.hidden{display:none}
.hidden{display:none}
```

- [ ] **Step 3: JSにヘルパー関数を追加**

`access-diff.html:365`（`var $ = function (id) { return document.getElementById(id); };` の直後）に追加する:

変更前:
```javascript
  var $ = function (id) { return document.getElementById(id); };

  /* ---------- 設定の保存 / 復元（Cookie + localStorage 保険） ---------- */
```

変更後:
```javascript
  var $ = function (id) { return document.getElementById(id); };
  function effA() { return state.A || state.B; }
  function effB() { return state.B || state.A; }
  function isSingle() { return !!(state.A || state.B) && !(state.A && state.B); }
  function updateModeUI() {
    var single = isSingle();
    $("keyField").classList.toggle("hidden", single);
    $("diffOnlyRow").classList.toggle("hidden", single);
  }

  /* ---------- 設定の保存 / 復元（Cookie + localStorage 保険） ---------- */
```

- [ ] **Step 4: refreshTableChoices() を単一ファイル対応にする**

`access-diff.html:451-473` を次のように変更する:

変更前:
```javascript
  function refreshTableChoices() {
    var sel = $("tableSel"), prev = state.table;
    sel.innerHTML = "";
    if (!state.A || !state.B) {
      sel.disabled = true;
      addOpt(sel, "", "— 2ファイルを読み込んでください —");
      state.table = null; resetColControls(); clearResult(); return;
    }
    var common = intersect(state.A.tables, state.B.tables);
    if (common.length === 0) {
      sel.disabled = true;
      addOpt(sel, "", "— 同名テーブルがありません —");
      state.table = null; resetColControls(); clearResult(); return;
    }
    sel.disabled = false;
    addOpt(sel, "", "テーブルを選択…");
    common.forEach(function (t) { addOpt(sel, t, t); });

    var restoreTable = (pending && pending.table && common.indexOf(pending.table) >= 0) ? pending.table : null;
    var target = restoreTable || (prev && common.indexOf(prev) >= 0 ? prev : null);
    if (target) { sel.value = target; onTableChange(); }
    else { state.table = null; }
  }
```

変更後:
```javascript
  function refreshTableChoices() {
    updateModeUI();
    var sel = $("tableSel"), prev = state.table;
    sel.innerHTML = "";
    if (!state.A && !state.B) {
      sel.disabled = true;
      addOpt(sel, "", "— ファイルを読み込んでください —");
      state.table = null; resetColControls(); clearResult(); return;
    }
    var common = intersect(effA().tables, effB().tables);
    if (common.length === 0) {
      sel.disabled = true;
      addOpt(sel, "", "— 同名テーブルがありません —");
      state.table = null; resetColControls(); clearResult(); return;
    }
    sel.disabled = false;
    addOpt(sel, "", "テーブルを選択…");
    common.forEach(function (t) { addOpt(sel, t, t); });

    var restoreTable = (pending && pending.table && common.indexOf(pending.table) >= 0) ? pending.table : null;
    var target = restoreTable || (prev && common.indexOf(prev) >= 0 ? prev : null);
    if (target) { sel.value = target; onTableChange(); }
    else { state.table = null; }
  }
```

- [ ] **Step 5: commonColumns() を effA/effB ベースにする**

`access-diff.html:475-486` を次のように変更する:

変更前:
```javascript
  function commonColumns() {
    if (!state.table) return { common: [], onlyA: [], onlyB: [], all: [] };
    var ca = state.A.reader.getTable(state.table).getColumnNames();
    var cb = state.B.reader.getTable(state.table).getColumnNames();
```

変更後:
```javascript
  function commonColumns() {
    if (!state.table) return { common: [], onlyA: [], onlyB: [], all: [] };
    var ca = effA().reader.getTable(state.table).getColumnNames();
    var cb = effB().reader.getTable(state.table).getColumnNames();
```

（以降の `commonColumns` 内のロジックは変更しない）

- [ ] **Step 6: compute() を effA/effB ベースにし、single フラグを結果に含める**

`access-diff.html:579-580` を次のように変更する:

変更前:
```javascript
    var dataA = state.A.reader.getTable(state.table).getData();
    var dataB = state.B.reader.getTable(state.table).getData();
```

変更後:
```javascript
    var dataA = effA().reader.getTable(state.table).getData();
    var dataB = effB().reader.getTable(state.table).getData();
```

`access-diff.html:604` を次のように変更する:

変更前:
```javascript
    state.result = { rows: rows, counts: counts, cols: cols, dupWarn: hasDup(mapA) || hasDup(mapB) };
```

変更後:
```javascript
    state.result = { rows: rows, counts: counts, cols: cols, dupWarn: hasDup(mapA) || hasDup(mapB), single: isSingle() };
```

- [ ] **Step 7: render() でサマリー・警告・diffOnlyフィルタを single 対応にする**

`access-diff.html:641-665` を次のように変更する:

変更前:
```javascript
  function render() {
    if (!state.result) return;
    var c = state.result.counts;
    $("sumSame").textContent = c.same; $("sumAdd").textContent = c.added;
    $("sumRem").textContent = c.removed; $("sumChg").textContent = c.changed;
    $("summary").classList.remove("hidden");

    $("phA").textContent = state.A ? state.A.name : "";
    $("phB").textContent = state.B ? state.B.name : "";
    $("phTable").textContent = state.table || "";
    $("phKey").textContent = state.keyCols.join(" + ");
    $("phCounts").textContent = "一致 " + c.same + " / 追加 " + c.added + " / 削除 " + c.removed + " / 変更 " + c.changed;
    $("phTime").textContent = new Date().toLocaleString("ja-JP");

    var cols = state.result.cols;
    var warn = $("warn"), msgs = [];
    if (cols.onlyA.length) msgs.push("A のみの列: " + cols.onlyA.join(", "));
    if (cols.onlyB.length) msgs.push("B のみの列: " + cols.onlyB.join(", "));
    if (state.result.dupWarn) msgs.push("キーに重複があります（同キーは出現順で対応付け）。");
    if (msgs.length) { warn.innerHTML = msgs.map(escapeHtml).join(" &nbsp;/&nbsp; "); warn.classList.remove("hidden"); }
    else warn.classList.add("hidden");

    var rows = sortedRows();
    if (state.diffOnly) rows = rows.filter(function (r) { return r.status !== "same"; });
```

変更後:
```javascript
  function render() {
    if (!state.result) return;
    var c = state.result.counts;
    var single = state.result.single;
    if (single) {
      $("summary").classList.add("hidden");
    } else {
      $("sumSame").textContent = c.same; $("sumAdd").textContent = c.added;
      $("sumRem").textContent = c.removed; $("sumChg").textContent = c.changed;
      $("summary").classList.remove("hidden");
    }

    $("phA").textContent = state.A ? state.A.name : "";
    $("phB").textContent = state.B ? state.B.name : "";
    $("phTable").textContent = state.table || "";
    $("phKey").textContent = single ? "" : state.keyCols.join(" + ");
    $("phCounts").textContent = single ? "" : ("一致 " + c.same + " / 追加 " + c.added + " / 削除 " + c.removed + " / 変更 " + c.changed);
    $("phTime").textContent = new Date().toLocaleString("ja-JP");

    var cols = state.result.cols;
    var warn = $("warn"), msgs = [];
    if (!single) {
      if (cols.onlyA.length) msgs.push("A のみの列: " + cols.onlyA.join(", "));
      if (cols.onlyB.length) msgs.push("B のみの列: " + cols.onlyB.join(", "));
      if (state.result.dupWarn) msgs.push("キーに重複があります（同キーは出現順で対応付け）。");
    }
    if (msgs.length) { warn.innerHTML = msgs.map(escapeHtml).join(" &nbsp;/&nbsp; "); warn.classList.remove("hidden"); }
    else warn.classList.add("hidden");

    var rows = sortedRows();
    if (!single && state.diffOnly) rows = rows.filter(function (r) { return r.status !== "same"; });
```

（以降の `render()` の残りのロジック — CAP処理・thead/tbody生成 — は変更しない）

- [ ] **Step 8: clearResult() のプレースホルダ文言を修正**

`access-diff.html:727-733` の該当行を変更する:

変更前:
```javascript
    $("tableHost").innerHTML = "<div class=\"empty\">テーブルを選ぶと差分が表示されます。</div>";
```

変更後:
```javascript
    $("tableHost").innerHTML = "<div class=\"empty\">テーブルを選ぶと内容が表示されます。</div>";
```

- [ ] **Step 9: 見出し・サブタイトルの文言更新**

`access-diff.html:241` を変更する:

変更前:
```html
    <span class="sub">2つの .mdb / .accdb を突き合わせて同名テーブルを比較</span>
```

変更後:
```html
    <span class="sub">1つならビュワー、2つなら .mdb / .accdb を突き合わせて同名テーブルを比較</span>
```

- [ ] **Step 10: ブラウザで手動確認する**

1. `access-diff.html` をブラウザで直接開く。
2. 複数テーブルを持つ `.mdb` または `.accdb` ファイルをファイルAのドロップ枠にのみ読み込む。
3. 確認項目:
   - テーブル選択プルダウンにそのファイルの全テーブル名が表示される。
   - テーブルを選ぶと全レコードが通常の見た目（差分色なし）でグリッド表示される。
   - サマリー行・「キー列（複合可）」フィールド・「差分のみ表示」チェックボックス行が非表示。
   - 列ヘッダクリックによるソート（昇順/降順ボタン含む）が機能する。
   - PDF出力ボタンが有効で機能する。
4. ファイルBに2つ目のDBファイルを読み込み、既存通りの複合キー差分比較（サマリー・キー選択・警告・チェックボックス行が再表示）に切り替わることを確認する。
5. ファイルAを外し、ファイルBのみの状態でも同様に単体ビュワー表示になることを確認する。
6. コンソールにJavaScriptエラーが出ていないことを確認する。

- [ ] **Step 11: コミット**

```bash
git add access-diff.html
git commit -m "$(cat <<'EOF'
Add single-file viewer mode to Access diff tool

Loading only one database now renders its tables directly instead of
showing an empty placeholder, by treating the missing side as identical
to the loaded one so the existing diff rendering shows no highlights.
EOF
)"
```

---

## Task 3: word-diff.html — 1ファイル登録時のビュワー対応

**Files:**
- Modify: `word-diff.html`（HTML: 157-161行目付近／CSS: 60行目付近／JS: 238-433行目のアプリロジック）

**Interfaces:**
- Produces: `effA()`, `effB()`, `isSingle()`, `updateModeUI()`（Task 1/2と同じ役割だが、このファイル内で独立して定義する）。

- [ ] **Step 1: HTMLに比較専用UI要素のID付与とプレースホルダ文言修正**

`word-diff.html:157-161` を次のように変更する:

変更前:
```html
      <div class="field">
        <label for="modeSel">表示</label>
        <select id="modeSel"><option value="inline">インライン</option><option value="side">左右並列</option></select>
      </div>
      <label class="check"><input type="checkbox" id="collapse"> 一致部分を省略</label>
```

変更後:
```html
      <div class="field" id="modeField">
        <label for="modeSel">表示</label>
        <select id="modeSel"><option value="inline">インライン</option><option value="side">左右並列</option></select>
      </div>
      <label class="check" id="collapseRow"><input type="checkbox" id="collapse"> 一致部分を省略</label>
```

`word-diff.html:191` を次のように変更する（初期プレースホルダ文言）:

変更前:
```html
  <div class="dochost" id="docHost"><div class="empty">2つの .docx を読み込むと差分が表示されます。</div></div>
```

変更後:
```html
  <div class="dochost" id="docHost"><div class="empty">.docx を読み込むと内容が表示されます。</div></div>
```

- [ ] **Step 2: CSSに汎用hiddenルールを追加**

`word-diff.html:60` の既存行の直後に1行追加する:

変更前:
```css
.summary.hidden{display:none}
```

変更後:
```css
.summary.hidden{display:none}
.hidden{display:none}
```

- [ ] **Step 3: JSにヘルパー関数を追加**

`word-diff.html:246`（`var $ = function (id) { return document.getElementById(id); };` の直後）に追加する:

変更前:
```javascript
  var $ = function (id) { return document.getElementById(id); };

  function saveSettings() {
```

変更後:
```javascript
  var $ = function (id) { return document.getElementById(id); };
  function effA() { return state.A || state.B; }
  function effB() { return state.B || state.A; }
  function isSingle() { return !!(state.A || state.B) && !(state.A && state.B); }
  function updateModeUI() {
    var single = isSingle();
    $("modeField").classList.toggle("hidden", single);
    $("collapseRow").classList.toggle("hidden", single);
  }

  function saveSettings() {
```

- [ ] **Step 4: compute() を effA/effB ベースにし、single フラグを結果に含める**

`word-diff.html:293-314` を次のように変更する:

変更前:
```javascript
  function compute() {
    clearResult();
    if (!state.A || !state.B) return;
    var raw = JsDiff.diffLines(state.A.text || "", state.B.text || "");
    var segs = [];
    for (var i = 0; i < raw.length; i++) {
      var p = raw[i];
      if (p.removed && raw[i + 1] && raw[i + 1].added) { segs.push({ type: "rep", a: p.value, b: raw[i + 1].value }); i++; }
      else if (p.added) segs.push({ type: "ins", b: p.value });
      else if (p.removed) segs.push({ type: "del", a: p.value });
      else segs.push({ type: "equal", a: p.value, b: p.value });
    }
    var counts = { equal: 0, added: 0, removed: 0, changed: 0 };
    segs.forEach(function (s) {
      if (s.type === "equal") counts.equal += lineCount(s.a);
      else if (s.type === "ins") counts.added += lineCount(s.b);
      else if (s.type === "del") counts.removed += lineCount(s.a);
      else { counts.changed++; }
    });
    state.result = { segs: segs, counts: counts };
    render();
  }
```

変更後:
```javascript
  function compute() {
    clearResult();
    updateModeUI();
    if (!state.A && !state.B) return;
    var raw = JsDiff.diffLines(effA().text || "", effB().text || "");
    var segs = [];
    for (var i = 0; i < raw.length; i++) {
      var p = raw[i];
      if (p.removed && raw[i + 1] && raw[i + 1].added) { segs.push({ type: "rep", a: p.value, b: raw[i + 1].value }); i++; }
      else if (p.added) segs.push({ type: "ins", b: p.value });
      else if (p.removed) segs.push({ type: "del", a: p.value });
      else segs.push({ type: "equal", a: p.value, b: p.value });
    }
    var counts = { equal: 0, added: 0, removed: 0, changed: 0 };
    segs.forEach(function (s) {
      if (s.type === "equal") counts.equal += lineCount(s.a);
      else if (s.type === "ins") counts.added += lineCount(s.b);
      else if (s.type === "del") counts.removed += lineCount(s.a);
      else { counts.changed++; }
    });
    state.result = { segs: segs, counts: counts, single: isSingle() };
    render();
  }
```

- [ ] **Step 5: render() を single 対応にし、全文表示させる**

`word-diff.html:319-336` を次のように変更する:

変更前:
```javascript
  function render() {
    if (!state.result) return;
    var c = state.result.counts;
    $("sumEqual").textContent = c.equal; $("sumAdd").textContent = c.added;
    $("sumRem").textContent = c.removed; $("sumChg").textContent = c.changed;
    $("summary").classList.remove("hidden");

    $("phA").textContent = state.A ? state.A.name : "";
    $("phB").textContent = state.B ? state.B.name : "";
    $("phCounts").textContent = "一致 " + c.equal + "行 / 追加 " + c.added + "行 / 削除 " + c.removed + "行 / 変更 " + c.changed + "箇所";
    $("phTime").textContent = new Date().toLocaleString("ja-JP");

    var noDiff = c.added === 0 && c.removed === 0 && c.changed === 0;
    var host = $("docHost");
    if (noDiff) { host.innerHTML = '<div class="empty">差分はありません。2つの文書は一致しています。</div>'; $("pdfBtn").disabled = true; return; }
    host.innerHTML = state.mode === "side" ? renderSide() : renderInline();
    $("pdfBtn").disabled = false;
  }
```

変更後:
```javascript
  function render() {
    if (!state.result) return;
    var c = state.result.counts;
    var single = state.result.single;
    if (single) {
      $("summary").classList.add("hidden");
    } else {
      $("sumEqual").textContent = c.equal; $("sumAdd").textContent = c.added;
      $("sumRem").textContent = c.removed; $("sumChg").textContent = c.changed;
      $("summary").classList.remove("hidden");
    }

    $("phA").textContent = state.A ? state.A.name : "";
    $("phB").textContent = state.B ? state.B.name : "";
    $("phCounts").textContent = single ? "" : ("一致 " + c.equal + "行 / 追加 " + c.added + "行 / 削除 " + c.removed + "行 / 変更 " + c.changed + "箇所");
    $("phTime").textContent = new Date().toLocaleString("ja-JP");

    var host = $("docHost");
    if (!single) {
      var noDiff = c.added === 0 && c.removed === 0 && c.changed === 0;
      if (noDiff) { host.innerHTML = '<div class="empty">差分はありません。2つの文書は一致しています。</div>'; $("pdfBtn").disabled = true; return; }
    }
    host.innerHTML = single ? renderInline(true) : (state.mode === "side" ? renderSide() : renderInline());
    $("pdfBtn").disabled = false;
  }
```

- [ ] **Step 6: renderInline() / equalBlock() に forceFull 引数を追加する**

`word-diff.html:338-357` を次のように変更する:

変更前:
```javascript
  function renderInline() {
    var out = ['<div class="docdiff">'];
    state.result.segs.forEach(function (s) {
      if (s.type === "equal") { out.push(equalBlock(s.a)); }
      else if (s.type === "ins") { lines(s.b).forEach(function (l) { out.push('<div class="ln ins-l"><ins>' + escapeHtml(l) + "</ins></div>"); }); }
      else if (s.type === "del") { lines(s.a).forEach(function (l) { out.push('<div class="ln del-l"><del>' + escapeHtml(l) + "</del></div>"); }); }
      else { out.push('<div class="ln rep-l">' + wordDiffHtml(s.a, s.b) + "</div>"); }
    });
    out.push("</div>");
    return out.join("");
  }
  function equalBlock(text) {
    var ls = lines(text);
    if (state.collapse && ls.length > 6) {
      var head = ls.slice(0, 3), tail = ls.slice(-3), hidden = ls.length - 6;
      return head.map(eqLine).join("") + '<div class="collapsed">⋯ ' + hidden + " 行一致 ⋯</div>" + tail.map(eqLine).join("");
    }
    return ls.map(eqLine).join("");
  }
```

変更後:
```javascript
  function renderInline(forceFull) {
    var out = ['<div class="docdiff">'];
    state.result.segs.forEach(function (s) {
      if (s.type === "equal") { out.push(equalBlock(s.a, forceFull)); }
      else if (s.type === "ins") { lines(s.b).forEach(function (l) { out.push('<div class="ln ins-l"><ins>' + escapeHtml(l) + "</ins></div>"); }); }
      else if (s.type === "del") { lines(s.a).forEach(function (l) { out.push('<div class="ln del-l"><del>' + escapeHtml(l) + "</del></div>"); }); }
      else { out.push('<div class="ln rep-l">' + wordDiffHtml(s.a, s.b) + "</div>"); }
    });
    out.push("</div>");
    return out.join("");
  }
  function equalBlock(text, forceFull) {
    var ls = lines(text);
    if (!forceFull && state.collapse && ls.length > 6) {
      var head = ls.slice(0, 3), tail = ls.slice(-3), hidden = ls.length - 6;
      return head.map(eqLine).join("") + '<div class="collapsed">⋯ ' + hidden + " 行一致 ⋯</div>" + tail.map(eqLine).join("");
    }
    return ls.map(eqLine).join("");
  }
```

- [ ] **Step 7: clearResult() のプレースホルダ文言を修正**

`word-diff.html:397-402` の該当行を変更する:

変更前:
```javascript
    $("docHost").innerHTML = '<div class="empty">2つの .docx を読み込むと差分が表示されます。</div>';
```

変更後:
```javascript
    $("docHost").innerHTML = '<div class="empty">.docx を読み込むと内容が表示されます。</div>';
```

- [ ] **Step 8: 見出し・サブタイトルの文言更新**

`word-diff.html:134` を変更する:

変更前:
```html
    <span class="sub">2つの .docx をそのまま比較</span>
```

変更後:
```html
    <span class="sub">1つならビュワー、2つなら .docx を比較</span>
```

- [ ] **Step 9: ブラウザで手動確認する**

1. `word-diff.html` をブラウザで直接開く。
2. 複数段落を含む `.docx` ファイルをファイルAのドロップ枠にのみ読み込む。
3. 確認項目:
   - 本文全体が省略なく（折りたたみなしで）プレーン表示される。差分ハイライト（挿入・削除・変更の色）が付かない。
   - サマリー行・「表示」モード切替フィールド・「一致部分を省略」チェックボックス行が非表示。
   - PDF出力ボタンが有効で機能する。
4. ファイルBに2つ目の `.docx` を読み込み、既存通りの段落・語単位の差分比較（サマリー・モード切替・省略チェックボックスが再表示）に切り替わることを確認する。
5. ファイルAを外し、ファイルBのみの状態でも同様に単体ビュワー表示になることを確認する。
6. コンソールにJavaScriptエラーが出ていないことを確認する。

- [ ] **Step 10: コミット**

```bash
git add word-diff.html
git commit -m "$(cat <<'EOF'
Add single-file viewer mode to Word diff tool

Loading only one document now renders its full text directly instead
of showing an empty placeholder, by treating the missing side as
identical to the loaded one so no diff highlighting or collapsing applies.
EOF
)"
```

---

## Task 4: index.html — 文言更新

**Files:**
- Modify: `index.html`（74行目, 81行目, 89行目, 97行目）

**Interfaces:**
- Consumes: なし（静的文言の変更のみ）。

- [ ] **Step 1: サブタイトルの更新**

`index.html:74` を次のように変更する:

変更前:
```html
    <p class="sub">Access・Excel・Word の2ファイルをブラウザ内で比較（外部送信なし）</p>
```

変更後:
```html
    <p class="sub">Access・Excel・Word をブラウザ内で閲覧・比較（1ファイルならビュワー、2ファイルなら差分比較・外部送信なし）</p>
```

- [ ] **Step 2: 各カードの説明文を更新**

`index.html:81` を次のように変更する:

変更前:
```html
      <p class="desc">2つのDBの同名テーブルを、複合キーで行対応付けしてセル単位に比較。追加・削除・変更を色分け表示。</p>
```

変更後:
```html
      <p class="desc">1つならテーブルをそのまま閲覧、2つなら複合キーで行対応付けしてセル単位に比較。追加・削除・変更を色分け表示。</p>
```

`index.html:89` を次のように変更する:

変更前:
```html
      <p class="desc">2つのブックを同名シートごとに、セル位置基準（A1, B2 …）で比較。シートはタブで切り替え。</p>
```

変更後:
```html
      <p class="desc">1つならシートをそのまま閲覧、2つなら同名シートごとにセル位置基準（A1, B2 …）で比較。シートはタブで切り替え。</p>
```

`index.html:97` を次のように変更する:

変更前:
```html
      <p class="desc">2つの文書の本文を段落・語単位で比較。インライン／左右並列で表示し、変更箇所を強調。</p>
```

変更後:
```html
      <p class="desc">1つなら本文をそのまま閲覧、2つなら段落・語単位で比較。インライン／左右並列で表示し、変更箇所を強調。</p>
```

- [ ] **Step 3: ブラウザで手動確認する**

1. `index.html` をブラウザで直接開く。
2. サブタイトルと各カードの説明文が更新後の日本語表示になっていることを目視確認する。
3. 各カードのリンク（`access-diff.html` / `excel-diff.html` / `word-diff.html`）が引き続き正しく遷移することを確認する。

- [ ] **Step 4: コミット**

```bash
git add index.html
git commit -m "$(cat <<'EOF'
Update index.html copy to mention single-file viewer mode

Each tool card now notes it works as a standalone viewer with one
file and as a diff tool with two.
EOF
)"
```

---

## Plan Self-Review Notes

- **Spec coverage:** Excel/Access/Word 各ツールの仕様（全シート/全テーブル表示、差分色なし表示、比較専用UI非表示、2ファイル目追加時の自動切り替え、ソート機能の維持）をTask 1-3で実装。index.html・各ツールサブタイトルの文言更新をTask 1(Step10)/2(Step9)/3(Step8)/4で実装。テスト方針（手動確認）は各Taskの最終Stepに反映。
- **既存動作への影響:** `effA()`/`effB()` は2ファイル揃った状態では常に `state.A`/`state.B` をそのまま返すため、2ファイル時の既存ロジック・見た目は変更されない。
- **Wordの折りたたみ:** single時は `forceFull=true` により常に全文表示にする（`state.collapse` の値に関わらず省略しない）。これは「ビュワーとして中身を確認できる」という目的に沿うための実装判断。
