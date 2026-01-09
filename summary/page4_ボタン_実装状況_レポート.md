# ページ4のボタン実装状況レポート

## ボタン一覧表

| No. | ボタン名 | 表示行 | 表示条件 | 機能実装状況 | クリック時の動作 | 備考 |
|-----|---------|--------|---------|-------------|----------------|------|
| 1 | KDL変換CSV格納元の移動設定 | 1行目・2行目 | 常に表示 | ✅ 実装済み | 編集モードON: `Show-FileMoveSettingsDialog`<br>編集モードOFF: `Invoke-FileMoveOperation` | 編集モードOFF時は「移動」と表示、紺色 |
| 2 | KDL変換CSV格納先の移動設定 | 1行目・2行目 | 常に表示 | ✅ 実装済み | 編集モードON: `Show-FileMoveSettingsDialog`<br>編集モードOFF: `Invoke-FileMoveOperation` | 編集モードOFF時は「移動」と表示、紺色 |
| 3 | V1抽出CSV格納先の移動設定 | 全行 | 常に表示 | ✅ 実装済み | 編集モードON: `Show-FileMoveSettingsDialog`<br>編集モードOFF: `Invoke-FileMoveOperation` | 編集モードOFF時は「移動」と表示、紺色 |
| 4 | KDL取込 | 1行目・2行目 | 常に表示 | ✅ 実装済み | 編集モードOFF: `BatchFiles[0]`を実行<br>編集モードON: ファイル選択ダイアログで`BatchFiles[0]`を保存 | 編集モードON時は「参照」と表示 |
| 5 | 直接取込 | 全行 | 常に表示 | ✅ 実装済み | 編集モードOFF: `BatchFiles[1]`を実行<br>編集モードON: ファイル選択ダイアログで`BatchFiles[1]`を保存 | 編集モードON時は「参照」と表示 |
| 6 | 取込後 | 全行 | 常に表示 | ✅ 実装済み | 編集モードOFF: `BatchFiles[2]`を実行<br>編集モードON: ファイル選択ダイアログで`BatchFiles[2]`を保存 | 編集モードON時は「参照」と表示 |
| 7 | ログ確認 | 全行 | 常に表示 | ✅ 実装済み | `Show-ProcessLog`を呼び出し | 編集モード時は「参照」と表示 |

## 実装状況サマリー

### ✅ 実装済み（全ボタン実装完了）
- **移動設定ボタン（3種類）**: すべて実装済み
  - KDL変換CSV格納元の移動設定（1行目・2行目）
  - KDL変換CSV格納先の移動設定（1行目・2行目）
  - V1抽出CSV格納先の移動設定（全行）
  - 編集モードON時: `Show-FileMoveSettingsDialog`関数を呼び出し、movefilesフォルダにファイルリストを保存
  - 編集モードOFF時: `Invoke-FileMoveOperation`関数を呼び出し、ファイル移動を実行
  - 編集モードOFF時は「移動」と表示、紺色（`#1e3a8a`）に変更
- **KDL取込ボタン**: 実装済み
  - 1行目・2行目のみに表示
  - 編集モードOFF時: `page4.json.Processes[i].BatchFiles[0]`のバッチファイルを実行
  - 編集モードON時: 「参照」と表示、ファイル選択ダイアログで`BatchFiles[0]`のパスを保存
  - `Save-BatchFilePath`関数を使用してJSONに保存
- **直接取込ボタン**: 実装済み
  - 全行に表示
  - 編集モードOFF時: `page4.json.Processes[i].BatchFiles[1]`のバッチファイルを実行
  - 編集モードON時: 「参照」と表示、ファイル選択ダイアログで`BatchFiles[1]`のパスを保存
  - `Save-BatchFilePath`関数を使用してJSONに保存
- **取込後ボタン**: 実装済み
  - 全行に表示
  - 編集モードOFF時: `page4.json.Processes[i].BatchFiles[2]`のバッチファイルを実行
  - 編集モードON時: 「参照」と表示、ファイル選択ダイアログで`BatchFiles[2]`のパスを保存
  - `Save-BatchFilePath`関数を使用してJSONに保存
- **ログ確認ボタン**: 実装済み
  - `Show-ProcessLog`関数を呼び出し、プロセスログを表示
  - 編集モード時は「参照」と表示され、ログ出力フォルダを選択可能

## 詳細説明

### 1. 移動設定ボタン（3種類）
- **実装関数**: 
  - 編集モードON時: `Show-FileMoveSettingsDialog`
  - 編集モードOFF時: `Invoke-FileMoveOperation`
- **動作**:
  - **編集モードON時**:
    - ボタンテキスト: 「移動設定」
    - ボタン色: 水色（`#dae8fc`）
    - クリックするとファイル移動設定ダイアログを表示
    - ダイアログでファイルリスト（移動元パス|移動先パス形式）を入力
    - `movefiles/***.txt`ファイルに保存（***はプロセス名から生成）
  - **編集モードOFF時**:
    - ボタンテキスト: 「移動」
    - ボタン色: 紺色（`#1e3a8a`）
    - クリックすると`movefiles/***.txt`を読み込み、ファイル移動を実行
    - `Invoke-FileMoveOperation`関数を使用
- **JSON要素との関連**: なし（movefilesフォルダにテキストファイルとして保存）

### 2. KDL取込ボタン
- **実装状況**: ✅ 実装済み
- **表示条件**: 1行目・2行目のみ
- **実装関数**: 
  - 編集モードOFF時: `Invoke-BatchFile`（`BatchFiles[0]`を実行）
  - 編集モードON時: `Save-BatchFilePath`（`BatchFiles[0]`を保存）
- **動作**:
  - **編集モードOFF時**:
    - ボタンテキスト: 「KDL取込」
    - `page4.json.Processes[i].BatchFiles[0]`のバッチファイルを実行
    - 実行中はボタンを無効化（`$this.Enabled = $false`）
  - **編集モードON時**:
    - ボタンテキスト: 「参照」
    - ファイル選択ダイアログを表示
    - 選択したバッチファイルのパスを`BatchFiles[0]`に保存
- **JSON要素との関連**: `page4.json.Processes[i].BatchFiles[0].Path`

### 3. 直接取込ボタン
- **実装状況**: ✅ 実装済み
- **表示条件**: 全行
- **実装関数**: 
  - 編集モードOFF時: `Invoke-BatchFile`（`BatchFiles[1]`を実行）
  - 編集モードON時: `Save-BatchFilePath`（`BatchFiles[1]`を保存）
- **動作**:
  - **編集モードOFF時**:
    - ボタンテキスト: 「直接取込」
    - `page4.json.Processes[i].BatchFiles[1]`のバッチファイルを実行
    - 実行中はボタンを無効化（`$this.Enabled = $false`）
  - **編集モードON時**:
    - ボタンテキスト: 「参照」
    - ファイル選択ダイアログを表示
    - 選択したバッチファイルのパスを`BatchFiles[1]`に保存
- **JSON要素との関連**: `page4.json.Processes[i].BatchFiles[1].Path`

### 4. 取込後ボタン
- **実装状況**: ✅ 実装済み
- **表示条件**: 全行
- **実装関数**: 
  - 編集モードOFF時: `Invoke-BatchFile`（`BatchFiles[2]`を実行）
  - 編集モードON時: `Save-BatchFilePath`（`BatchFiles[2]`を保存）
- **動作**:
  - **編集モードOFF時**:
    - ボタンテキスト: 「取込後」
    - `page4.json.Processes[i].BatchFiles[2]`のバッチファイルを実行
    - 実行中はボタンを無効化（`$this.Enabled = $false`）
  - **編集モードON時**:
    - ボタンテキスト: 「参照」
    - ファイル選択ダイアログを表示
    - 選択したバッチファイルのパスを`BatchFiles[2]`に保存
- **JSON要素との関連**: `page4.json.Processes[i].BatchFiles[2].Path`

### 5. ログ確認ボタン
- **実装関数**: `Show-ProcessLog`
- **動作**:
  - 編集モードOFF時: 「ログ確認」と表示、プロセスログを表示
  - 編集モードON時: 「参照」と表示、ログ出力フォルダを選択可能
- **JSON要素との関連**: `page4.json.Processes[i].LogButtonText`（表示テキストのカスタマイズ）

## ボタンの表示条件まとめ

### 常に表示（編集モードに関わらず）
- **移動設定ボタン（3種類）**: 常に表示
  - 編集モードON時: 「移動設定」、水色
  - 編集モードOFF時: 「移動」、紺色
- **KDL取込**: 1行目・2行目のみ、常に表示
  - 編集モードON時: 「参照」
  - 編集モードOFF時: 「KDL取込」
- **直接取込**: 全行、常に表示
  - 編集モードON時: 「参照」
  - 編集モードOFF時: 「直接取込」
- **取込後**: 全行、常に表示
  - 編集モードON時: 「参照」
  - 編集モードOFF時: 「取込後」
- **ログ確認**: 全行、常に表示
  - 編集モードON時: 「参照」
  - 編集モードOFF時: 「ログ確認」

## 技術的な実装詳細

### バッチファイル実行処理
- **関数**: `Invoke-BatchFile`
- **実行方法**: `Start-Process`を使用してバッチファイルを実行
- **ログ出力**: `process_${currentPage}_${ProcessIndex}_stdout.log`と`process_${currentPage}_${ProcessIndex}_stderr.log`に出力
- **エラーハンドリング**: バッチファイルが存在しない場合、エラーメッセージを表示

### バッチファイルパス保存処理
- **関数**: `Save-BatchFilePath`
- **保存先**: `page4.json.Processes[i].BatchFiles[BatchIndex].Path`
- **パス変換**: スクリプトルート内の場合は相対パスに変換、それ以外は絶対パスを保存
- **エンコーディング**: UTF-8 BOM付きで保存

### 編集モード切り替え時の更新処理
- **関数**: `Update-ProcessControls`
- **処理内容**: 編集モードに応じてボタンのテキストと色を更新
- **対象ボタン**: KDL取込、直接取込、取込後、移動設定ボタン（3種類）

---
**レポート作成日**: 2026-01-09
**最終更新日**: 2026-01-09
**対象ファイル**: `Functions.ps1`, `page4.json`
**実装状況**: 全ボタン実装完了 ✅
