# CHECKLIST

## ビルド/配布

- `dotnet build .\CoffeeAutoButton.sln` が警告 0 / エラー 0 で完了する。
- `.\tools\publish.ps1` が `publish\CoffeeAutoButton` に出力する。
- publish 出力の `CoffeeAutoButton.exe` が起動する。
- `.\tools\build-msi.ps1` が UI 付きの `installer\out\CoffeeAutoButtonSetup.msi` を出力する。
- MSI から `%LOCALAPPDATA%\CoffeeAutoButton` にインストールできる。
- スタートメニューに `Coffee AutoButton` と `CoffeeAutoButton` のショートカットが作成される。
- 必要に応じて `.\tools\build-installer.ps1` が簡易 EXE インストーラーを出力する。

## 基本動作

- キー連打で単一キーを送信できる。
- 修飾キー付きキー入力を送信できる。
- キー長押しを送信できる。
- キーシーケンスを順番に送信できる。
- 間隔、継続時間、開始待機が指定通り動作する。
- 一時停止/再開で残り時間が進まない。
- 停止/一時停止ホットキーがアプリ非フォーカスでも動作する。
- 停止/一時停止ホットキーを変更して保存/復元できる。

## クリック動作

- 位置取得で対象情報が表示される。
- 非干渉クリックで物理カーソル位置が変わらない。
- 非干渉テスト送信が動作する。
- 左クリック、右クリック、左ダブルクリック、左長押しが動作する。
- 非干渉クリックが失敗した場合、設定に応じて物理クリックへ切り替わる。
- 物理クリック後に元のカーソル位置へ戻る。

## 設定/プリセット

- プリセット保存、読込、削除が動作する。
- 前回設定が `%APPDATA%\CoffeeAutoButton\last-settings.json` から復元される。
- プリセットが `%APPDATA%\CoffeeAutoButton\presets.json` に保存される。
- 旧カレントディレクトリの JSON が AppData 側へ移行される。
- 保存 JSON に `SchemaVersion` が含まれる。
