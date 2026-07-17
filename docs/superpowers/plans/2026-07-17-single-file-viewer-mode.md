# 1ファイル登録時のビュワーモード対応 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: ユーザーのグローバル設定（`~/.claude/CLAUDE.md`）の指示により、実装はサブエージェント "implementer-codex" に委譲し、実装完了後にメインセッションが結果を確認すること。通常の superpowers:subagent-driven-development / superpowers:executing-plans は使わない。Steps はチェックボックス（`- [ ]`）で進捗管理する。各Taskの「ブラウザで手動確認する」Stepは、実行環境にブラウザ操作の手段がない場合はメインセッションに差し戻し、人間のレビュアーに確認を依頼すること。

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

- Modify: `excel-diff.html`
  - HTML: 145行目（サブタイトル）, 167行目（modebar）, 171行目（tabsプレースホルダ）, 179行目（diffOnly行）, 210行目（初期プレースホルダ）
  - CSS: 47行目付近（`.hidden` ルール追加）
  - JS: 343行目付近（ヘルパー追加）, 396-411行目（`refreshAll`）, 412-428行目（`buildTabs`）, 444-448行目（`currentPair`）, 466-505行目（`compute`）, 507-529行目（`render`前半）, 570-575行目付近（`clearResult`）

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
    if (single) { state.manual = false; $("manualPair").checked = false; }
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

（`buildTabs` 内、タブ生成ループ自体のロジックは変更しない。ただし末尾の「同名シートがありません」の分岐は次のStepで変更する）

- [ ] **Step 5b: buildTabs() の0件メッセージを single 対応にする**

`excel-diff.html:426-427` を次のように変更する:

変更前:
```javascript
    if (common.length === 0) { host.innerHTML += '<span class="tabhint">同名シートがありません。右上「比較するシートを選ぶ」で手動指定できます。</span>'; state.sheet = null; return; }
    state.sheet = want || null;
```

変更後:
```javascript
    if (common.length === 0) {
      host.innerHTML += isSingle()
        ? '<span class="tabhint">このファイルにはシートがありません。</span>'
        : '<span class="tabhint">同名シートがありません。右上「比較するシートを選ぶ」で手動指定できます。</span>';
      state.sheet = null; return;
    }
    state.sheet = want || null;
```

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

    var mergeView = R.hasMerge && (R.single || !state.diffOnly);
    var rowsToShow = [];
    for (var r = 0; r < R.rows; r++) { if (R.single || !state.diffOnly || R.rowHasDiff[r]) rowsToShow.push(r); }
```

（以降の行番号は変わるが、`render()` の残りのロジック — CAP処理・thead/tbody生成 — は変更しない）

> **注記（Critical修正）:** `state.diffOnly` はCookie/localStorageから復元される値であり、単一ファイル表示時に前回セッションの値が `true` のまま残っている場合がある。`mergeView` の判定に `R.single` を含めないと、結合セルを含むシートを単一ファイル表示した際に「差分のある行のみ」相当の展開表示（結合解除）になってしまい、設計書の「結合セルはそのまま結合表示を再現」という要件に反する。単一ファイル時は常に結合表示を優先する。

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
   - 「比較するシートを選ぶ」チェックボックス行が非表示、チェック自体も外れている。
   - 「差分のある行のみ表示」チェックボックス行が非表示。
   - PDF出力ボタンが有効で、クリックして印刷プレビューが出る。印刷ヘッダのB欄が空欄のまま崩れずに表示される。
   - 結合セルを含むシートで、結合がそのまま表示される。
   - シートが0件のブック（あれば）で「このファイルにはシートがありません。」と表示される。
4. 事前に「差分のある行のみ表示」にチェックを入れた状態（Cookie/localStorageに `diffOnly=true` を保存させた状態）でページを再読み込みし、ファイルAのみを読み込む。結合セルを含むシートで結合表示が解除されず、そのまま結合表示されることを確認する（`mergeView` の `R.single` 考慮の検証）。
5. 4001行を超える大きいシートを持つファイルで、「先頭 4000 行のみ表示。」の注記が単一ファイル表示でも出ることを確認する。
6. ファイルBに2つ目の `.xlsx` を読み込み、既存通りの同名シート差分比較（サマリー・警告・チェックボックス行が再表示、チェック状態も正しく反映）に切り替わることを確認する。
7. ページを再読み込みし、今度はファイルBのみを読み込んで、同様に単体ビュワー表示になることを確認する。
8. 「比較するシートを選ぶ」を手動でオンにした状態（2ファイル時）でシートA/Bを選び、その状態を保ったままページを再読み込みし、2ファイルを再度読み込んで手動ペア設定が復元されることを確認する（既存動作の回帰確認）。
9. コンソールにJavaScriptエラーが出ていないことを確認する。

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

- Modify: `access-diff.html`
  - HTML: 241行目（サブタイトル）, 267行目（テーブル選択ラベル）, 270-276行目（キー列フィールド）, 291行目（diffOnly行）, 318行目（初期プレースホルダ）
  - CSS: 93行目付近（`.hidden` ルール追加）
  - JS: 365行目付近（ヘルパー追加）, 451-473行目（`refreshTableChoices`）, 475-486行目（`commonColumns`）, 572-606行目（`compute`）, 641-665行目（`render`前半）, 727-733行目（`clearResult`）

**Interfaces:**

- Produces: `effA()`, `effB()`, `isSingle()`, `updateModeUI()`（Task 1と同じ役割だが、このファイル内で独立して定義する）。

- [ ] **Step 1: HTMLに比較専用UI要素のID付与とプレースホルダ文言修正**

`access-diff.html:266-269` を次のように変更する（ラベルに `id="tableSelLabel"` を追加し、単一ファイル時に「テーブル（同名）」→「テーブル」に文言を切り替えられるようにする）:

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
        <label for="tableSel" id="tableSelLabel">テーブル（同名）</label>
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
    $("tableSelLabel").textContent = single ? "テーブル" : "テーブル（同名）";
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
      addOpt(sel, "", isSingle() ? "— テーブルがありません —" : "— 同名テーブルがありません —");
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

> **注記:** `common` は単一ファイル時 `effA() === effB()` により当該ファイルの全テーブル名になるため、`common.length === 0` は「読み込んだDBにテーブルが1つもない」という異常系のみを指す。この場合のメッセージも比較前提の文言にならないよう `isSingle()` で出し分ける。

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
   - テーブル選択の見出しが「テーブル（同名）」ではなく「テーブル」になっている。
   - テーブル選択プルダウンにそのファイルの全テーブル名が表示される。
   - テーブルを選ぶと全レコードが通常の見た目（差分色なし）でグリッド表示される。
   - サマリー行・「キー列（複合可）」フィールド・「差分のみ表示」チェックボックス行が非表示。
   - 列ヘッダクリックによるソート（昇順/降順ボタン含む）が機能する。
   - PDF出力ボタンが有効で機能する。
   - 6001件を超える大きいテーブルで「先頭 6000 行のみ表示」の注記が単一ファイル表示でも出ることを確認する。
4. ファイルBに2つ目のDBファイルを読み込み、既存通りの複合キー差分比較（サマリー・キー選択・警告・チェックボックス行が再表示、テーブル選択見出しが「テーブル（同名）」に戻る）に切り替わることを確認する。
5. ページを再読み込みし、今度はファイルBのみを読み込んで、同様に単体ビュワー表示になることを確認する。
6. 2ファイル時に列ヘッダクリックで並べ替え・「差分のみ表示」・キー列を変更した状態でページを再読み込みし、2ファイルを再度読み込んで、既存通りにこれらの設定が復元されることを確認する（既存動作の回帰確認）。
7. コンソールにJavaScriptエラーが出ていないことを確認する。

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

- Modify: `word-diff.html`
  - HTML: 134行目（サブタイトル）, 157-161行目（モード切替・省略チェックボックス）, 191行目（初期プレースホルダ）
  - CSS: 60行目付近（`.hidden` ルール追加）
  - JS: 246行目付近（ヘルパー追加）, 293-314行目（`compute`）, 319-336行目（`render`）, 338-357行目（`renderInline`/`equalBlock`）, 397-402行目付近（`clearResult`）

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
2. 事前に「一致部分を省略」にチェックを入れて7段落以上ある文書で省略表示を確認した後（Cookie/localStorageに `collapse=true` を保存させる）、ページを再読み込みする。
3. 7段落以上ある `.docx` ファイルをファイルAのドロップ枠にのみ読み込む。
4. 確認項目:
   - `collapse=true` が保存されていても、本文全体が省略なく（折りたたみなしで）プレーン表示される。差分ハイライト（挿入・削除・変更の色）が付かない。
   - サマリー行・「表示」モード切替フィールド・「一致部分を省略」チェックボックス行が非表示。
   - PDF出力ボタンが有効で機能する。
5. 本文が空（0段落）の `.docx` を読み込んだ場合に、コンソールエラーなく空の表示領域になることを確認する（既存に「空です」等の専用メッセージはないため、真っ白な表示で問題ない）。
6. ファイルBに2つ目の `.docx` を読み込み、既存通りの段落・語単位の差分比較（サマリー・モード切替・省略チェックボックスが再表示、`collapse` 設定も反映）に切り替わることを確認する。
7. ページを再読み込みし、今度はファイルBのみを読み込んで、同様に単体ビュワー表示になることを確認する。
8. コンソールにJavaScriptエラーが出ていないことを確認する。

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

## Task 5: 統合スモークテスト（3ツール横断・回帰確認）

**Files:**

- なし（コード変更なし。Task 1-4完了後の最終確認のみ）

**Interfaces:**

- Consumes: Task 1-4で変更した4ファイルすべて。

- [ ] **Step 1: 3ツールを一通り単体ビュワーとして確認する**

1. `index.html` から各ツールに遷移し、更新後の説明文からリンクで正しく各ツールに移動できることを確認する。
2. Excel・Access・Wordそれぞれで、ファイルA単体・ファイルB単体の両方のパターンで単体ビュワー表示になることを一通り再確認する（Task 1/2/3の手動確認Stepで個別に確認済みだが、同一ブラウザセッション内で3ツールを連続して開いた場合に設定（Cookie/localStorage）が混線しないことを確認する目的）。

- [ ] **Step 2: 2ファイル揃った状態での既存動作に回帰がないことを確認する**

1. Excel: 同名シートがある2ファイル、同名シートが無い2ファイル（手動ペア選択が必要になるケース）、結合セルを含む2ファイルで、既存の差分比較結果が変更前と同じ見た目になることを確認する。
2. Access: 同名テーブルがある2ファイル、キー重複があるテーブルを含む2ファイルで、既存の差分比較結果（重複警告・A/Bのみ列バッジ含む）が変更前と同じ見た目になることを確認する。
3. Word: 内容が完全に一致する2ファイル（「差分はありません。2つの文書は一致しています。」の表示が引き続き出ること）、差分がある2ファイルで、既存の差分比較結果が変更前と同じ見た目になることを確認する。
4. 3ツールとも、片方のファイルの読み込みに失敗した場合（対応形式外のファイルを読ませる等）に、既存通りのエラーメッセージが `#err` に表示されることを確認する。

- [ ] **Step 3: コンソールエラーの最終確認**

1. 上記すべての操作を通して、ブラウザの開発者コンソールにJavaScriptエラーや警告が出ていないことを確認する。
2. 問題が見つかった場合は該当Taskに戻って修正し、該当ファイルのみ再コミットする。

## Plan Self-Review Notes

- **Spec coverage:** Excel/Access/Word 各ツールの仕様（全シート/全テーブル表示、差分色なし表示、比較専用UI非表示、2ファイル目追加時の自動切り替え、ソート機能の維持）をTask 1-3で実装。index.html・各ツールサブタイトルの文言更新をTask 1(Step10)/2(Step9)/3(Step8)/4で実装。テスト方針（手動確認）は各Taskの最終Stepおよび設計書の「実装上の既知の制約」節に反映。
- **既存動作への影響:** `effA()`/`effB()` は2ファイル揃った状態では常に `state.A`/`state.B` をそのまま返すため、2ファイル時の既存ロジック・見た目は変更されない。
- **Wordの折りたたみ:** single時は `forceFull=true` により常に全文表示にする（`state.collapse` の値に関わらず省略しない）。これは「ビュワーとして中身を確認できる」という目的に沿うための実装判断。

### plan-reviewer-codex によるレビュー反映（2026-07-17）

1回目のレビューで指摘されたCritical 2件・主要Warningを本計画に反映済み:

- **[Critical→対応済み] Excel: `diffOnly` 残留による結合セル表示破綻** — Task 1 Step 8で `mergeView` の判定に `R.single` を追加し、単一ファイル時は `state.diffOnly` の値に関わらず結合表示を維持するよう修正。
- **[Critical→対応済み] Excel: `manual` 復元状態とチェックボックスUIの不整合** — Task 1 Step 4で、単一ファイル時に `state.manual = false` とあわせて `$("manualPair").checked = false` も設定し、内部状態とDOMを同期させるよう修正。
- **[Warning→対応済み] 0シート/0テーブルの空データケースの文言** — Excel `buildTabs()`（Task 1 Step 5b 新設）、Access `refreshTableChoices()`（Task 2 Step 4）で、単一ファイル時は比較を前提にした文言を出さないよう `isSingle()` で出し分け。
- **[Warning→対応済み] Accessの「テーブル（同名）」ラベルの残留** — `id="tableSelLabel"` を追加し、`updateModeUI()` で単一ファイル時は「テーブル」に切り替えるよう対応（Task 2 Step 1, Step 3）。
- **[Warning→対応済み] 設定復元・表示上限・PDF出力の手動確認漏れ** — 各Taskの手動確認Stepに、Cookie/localStorage復元確認（`diffOnly`/`manual`/`collapse`等の残留値が単一ファイル表示を壊さないこと）、表示上限（Excel 4000行/Access 6000行）、PDF出力の見た目確認、「ファイルを外す」の代替手段（ページ再読み込み＋別スロットのみ読み込み）を追加。
- **[Warning→対応済み] Files欄の変更箇所一覧が不完全** — 各Taskの「Files」を実際に変更する行番号の完全なリストに更新。
- **[Warning→対応済み] 統合スモークテストの欠如** — Task 5として新設し、3ツール横断の確認と2ファイル時の既存動作の回帰確認をまとめた。
- **[Warning→設計書に反映] Wordの「diff処理なし」という表現、Wordの空段落消失、Accessのキーバッジ残留** — 実装を変更する代わりに、設計書に「実装上の既知の制約」節を新設し、既存ロジック流用に伴う許容事項として明記した。
- **[Question→回答] 「ファイルを外す」操作** — 現行UIに解除ボタンは無いため、手動確認手順は「ページを再読み込みして別スロットのみ読み込む」という操作に置き換えた。
- **[Question→対応不要と判断] CLAUDE.md参照・bash heredocの前提** — 計画冒頭の記述をユーザーのグローバル設定である旨に明確化。bash heredocはこの環境のBashツール（Git Bash）で問題なく動作するため対応不要と判断。
