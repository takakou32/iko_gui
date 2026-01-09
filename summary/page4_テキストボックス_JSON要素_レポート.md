# ページ4の編集モードON時の編集可能テキストボックス - JSON要素参照レポート

## テキストボックス一覧表

| No. | テキストボックス名 | 表示行 | 参照元JSON要素 | 初期値読み込み | 保存機能 | 備考 |
|-----|-------------------|--------|---------------|--------------|---------|------|
| 1 | V1抽出CSV格納元 | 1行目のみ | `page4.json.SourcePath` | ✅ 実装済み | ✅ 実装済み | Load-PagePathsで読み込み、Save-PagePathsで保存 |
| 2 | KDL変換CSV格納元 | 1行目・2行目 | `page4.json.Processes[i].KdlSourcePath` | ✅ 実装済み | ✅ 実装済み | Update-ProcessControlsで読み込み、Save-ProcessKdlSourcePathで保存 |
| 3 | KDL変換CSV格納先 | 1行目・2行目 | `page4.json.Processes[i].KdlDestPath` | ✅ 実装済み | ✅ 実装済み | Update-ProcessControlsで読み込み、Save-ProcessKdlDestPathで保存 |
| 4 | V1抽出CSV格納先 | 全行 | `page4.json.Processes[i].V1CsvDestPath` | ✅ 実装済み | ✅ 実装済み | Update-ProcessControlsで読み込み、Save-ProcessV1CsvDestPathで保存 |
| 5 | タスク名 | 全行 | `page4.json.Processes[i].Name` | ✅ 実装済み | ❌ 未実装 | ReadOnly = trueのため編集不可 |
| 6 | ログ格納先 | 共通セクション | `page4.json.LogStoragePath` | ✅ 実装済み | ✅ 実装済み | UILayout.ps1で実装済み |

## 実装状況サマリー

### ✅ 実装済み
- **V1抽出CSV格納元**: 読み込み・保存ともに実装済み
- **KDL変換CSV格納元**: 読み込み・保存ともに実装済み
- **KDL変換CSV格納先**: 読み込み・保存ともに実装済み
- **V1抽出CSV格納先**: 読み込み・保存ともに実装済み
- **タスク名**: 読み込みのみ（編集不可）
- **ログ格納先**: 読み込み・保存ともに実装済み

### ❌ 未実装
なし（すべて実装済み）

## 詳細説明

### 1. V1抽出CSV格納元テキストボックス
- **参照元**: `page4.json.SourcePath`
- **実装状況**: 
  - 初期値読み込みは実装済み（`Load-PagePaths`関数でページ4の場合も対応）
  - 保存機能は実装済み（`Save-PagePaths -SourcePath`を呼び出し）

### 2. KDL変換CSV格納元テキストボックス
- **参照元**: `page4.json.Processes[i].KdlSourcePath`
- **実装状況**: 
  - 初期値読み込みは実装済み（`Update-ProcessControls`関数で読み込み）
  - 保存機能は実装済み（`Save-ProcessKdlSourcePath`関数を呼び出し）

### 3. KDL変換CSV格納先テキストボックス
- **参照元**: `page4.json.Processes[i].KdlDestPath`
- **実装状況**: 
  - 初期値読み込みは実装済み（`Update-ProcessControls`関数で読み込み）
  - 保存機能は実装済み（`Save-ProcessKdlDestPath`関数を呼び出し）

### 4. V1抽出CSV格納先テキストボックス
- **参照元**: `page4.json.Processes[i].V1CsvDestPath`
- **実装状況**: 
  - 初期値読み込みは実装済み（`Update-ProcessControls`関数で読み込み）
  - 保存機能は実装済み（`Save-ProcessV1CsvDestPath`関数を呼び出し）

### 5. タスク名テキストボックス
- **参照元**: `page4.json.Processes[i].Name`
- **実装状況**: 
  - 初期値読み込みは実装済み（1577行目）
  - `ReadOnly = $true`のため編集不可

### 6. ログ格納先テキストボックス
- **参照元**: `page4.json.LogStoragePath`
- **実装状況**: 
  - 初期値読み込みは実装済み（`Load-PagePaths`関数）
  - 保存機能は実装済み（`UILayout.ps1`で`Save-PagePaths -LogStoragePath`を呼び出し）

---
**レポート作成日**: 2026-01-09
**対象ファイル**: `Functions.ps1`, `page4.json`, `UILayout.ps1`
