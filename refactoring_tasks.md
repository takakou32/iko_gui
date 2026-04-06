# Functions.ps1 リファクタリング（UI生成との分離）進捗管理

各タスク（枝番）完了直後に必ずGUIが起動し、システムの全機能が正常動作すること（アトミック性、リリース可能性）を保証する手順で再構成しました。

## コーディングルール
- **エンコーディングの厳守:** PowerShellスクリプト（`.ps1`）を新規作成・編集した際は、文字化けや意図しない動作を防ぐため、**必ず「UTF-8（BOM付き）」で保存**してください。

## タスクリスト

### フェーズ1：UI描画ロジックの別ファイルへの隔離
- [x] **1.1:** 空の `UIPageRenderer.ps1` を作成し、`Main.ps1` からドットソースで読み込むよう追記する。（起動確認）
- [x] **1.2:** `Create-AfterImportButton` や `Create-MaintButton` などの、外部依存が少ない「コントロール生成補助関数群」だけを `Functions.ps1` から `UIPageRenderer.ps1` へ移動する。（起動・描画確認）
- [x] **1.3:** `Update-ProcessControls` 本体を `Functions.ps1` から `UIPageRenderer.ps1` へ移動する。（起動・描画確認、UI描画ファイルの隔離完了）

### フェーズ2：巨大な Update-ProcessControls の細分化
- [x] **2.1:** `Update-ProcessControls` 内の「1・2ページ目のレイアウト作成」処理を抽出し、`UIPageRenderer.ps1` 内に新関数 `Render-Page1And2Row` を作成する。同時に元の巨大ループからこの関数を呼び出すように書き換える。（1・2ページの動作確認）
- [ ] **2.2:** 「3ページ目の特有レイアウト」処理を抽出し、新関数 `Render-Page3Row` を作成。元のループから呼び出すように書き換える。（3ページの動作確認）
- [ ] **2.3:** 「4ページ目の特有レイアウト」処理を抽出し、新関数 `Render-Page4Row` を作成。元のループから呼び出すように書き換える。（4ページの動作確認、細分化完了）

### フェーズ3：イベント（Add_Click）とロジックの分離
- [ ] **3.1:** 「ファイル選択ダイアログ処理（Save-BatchFilePath含む）」全体を独立させ、`Functions.ps1` に新関数 `Invoke-ActionFileSelect` を作成。UI側（ボタン）の `Add_Click` をこの関数呼び出しに置き換える。（ファイル選択の動作確認）
- [ ] **3.2:** 「バッチ実行処理（Invoke-BatchFile含む）」全体を独立させ、`Functions.ps1` に新関数 `Invoke-ActionBatchExecute` を作成。UI側（ボタン）の `Add_Click` をこの関数呼び出しに置き換える。（バッチ実行の動作確認、ロジック分離完了）

### フェーズ4：状態更新（State Update）の分離
- [ ] **4.1:** 現在 `Update-ProcessControls` の末尾で行っている「グレーアウト処理」と「編集モードによるテキスト切替」ロジックを抽出し、新関数 `Update-ProcessControlsState` を作成して末尾から呼び出すようにする。（各ページでの状態更新の動作確認）
- [ ] **4.2:** 編集モード切替イベント（`EditModeButton` の `Add_Click` 時）において、現状「UI全体を全て破棄して再作成している部分」を、要素を再利用して `Update-ProcessControlsState` のみを呼び出す（プロパティ更新のみを行う）高速なロジックに変更する。（編集モード切替時のチラつき解消・テキスト反映の動作確認）
