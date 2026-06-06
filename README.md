# Coffee AutoButton

Coffee AutoButton は Windows 向けの自動入力ツールです。キー入力や固定位置クリックを一定間隔で繰り返しつつ、作業中のマウス位置やアクティブウィンドウをできるだけ奪わないことを重視しています。

現在の実装メモ: v1.0.8

## コンセプト

- 作業を邪魔しない自動入力を第一にする。
- ブラウザ向けクリックは通常の物理クリックではなく、専用ブラウザへの CDP 送信を使う。
- 互換性が必要な場面だけ、明示的に物理クリックへ切り替える。
- 設定とプリセットを保存し、同じ作業を再開しやすくする。
- 実装は WPF と小さなサービスクラスに分け、保守しやすさを優先する。

## 主な機能

- キー連打
- 修飾キー付き入力
- キー長押し
- キーシーケンス入力
- 固定位置クリック連打
- 左クリック、右クリック、左ダブルクリック、左長押し
- 専用ブラウザへの非干渉クリック
- 非ブラウザアプリ向けの Win32 メッセージクリック
- 物理クリック互換モード
- プリセット保存、削除、前回設定の復元
- 停止、一時停止、再開ホットキー
- MSI インストーラー

## ブラウザクリック

Chrome 系ブラウザへの非干渉クリックは、通常の Chrome ウィンドウではなく Coffee AutoButton から起動した専用ブラウザを対象にします。

1. `専用ブラウザURL` に対象ページを入力する。
2. `専用ブラウザ` ボタンで Chrome を起動する。
3. 専用ブラウザ内で対象ページを開く。
4. `位置を取得する` でクリック位置を取得する。
5. 対象情報に `専用ブラウザ認識: CDP:9223` が表示されていることを確認する。
6. `非干渉テスト` または `スタート` を実行する。

Chrome 系ブラウザが専用ブラウザとして認識されていない場合、アプリは Chrome へクリックを送信しません。これはフォーカスを奪う `PostMessage` や物理クリックへ勝手に戻らないための仕様です。

## クリック方式

- 専用ブラウザ CDP 方式: Chrome 系ブラウザ用。カーソル移動やフォーカス取得を行わず、`Input.dispatchMouseEvent` でクリックを送信します。
- Win32 メッセージ方式: 非ブラウザアプリ用。対象ウィンドウへ `PostMessage` でクリックメッセージを送信します。
- 物理クリック方式: 互換性優先の明示的な方式です。カーソル移動を伴うため、通常は非干渉目的では使いません。

## バージョン確認

アプリのタイトルバーと画面下部にバージョンが表示されます。v1.0.8 以降は、起動中の画面から現在版を確認できます。

## 設定保存

設定は `%APPDATA%\CoffeeAutoButton` に保存されます。

- `last-settings.json`: 前回設定
- `presets.json`: プリセット
- `dedicated-browser-profile`: 専用ブラウザのプロファイル

## ビルド

必要な SDK は .NET 10 です。`global.json` で `10.0.300` を基準にしています。

```powershell
dotnet build .\CoffeeAutoButton.sln
```

## Publish

```powershell
.\tools\publish.ps1
```

出力先がロックされている場合は、バージョン付きの別フォルダを使います。

```powershell
.\tools\publish.ps1 -OutputPath "publish\CoffeeAutoButton-1.0.8"
```

## Installer

MSI は次のコマンドで生成します。

```powershell
.\tools\build-msi.ps1
```

特定の publish フォルダから MSI を作る場合:

```powershell
.\tools\build-msi.ps1 -PublishPath "publish\CoffeeAutoButton-1.0.8"
```

出力先:

```text
installer\out\CoffeeAutoButtonSetup.msi
```

MSI はユーザー単位で `%LOCALAPPDATA%\CoffeeAutoButton` にインストールし、スタートメニューに `Coffee AutoButton` のショートカットを作成します。

## 改良方針

- Chrome 系ブラウザは専用ブラウザ認識を前提にし、誤ってフォーカスを奪う経路を増やさない。
- UI は状態がすぐ分かることを優先し、対象認識、クリック方式、バージョンを明示する。
- 物理クリックは互換性のために残すが、自動フォールバックは慎重に扱う。
- 配布時はアプリ版、MSI 版、表示バージョンをそろえる。
- 仕様変更後は `README.md`、`DESIGN.md`、必要に応じて `.codex/skills/coffee-auto-button/SKILL.md` を更新する。

詳細な設計は [DESIGN.md](DESIGN.md)、手動確認観点は [CHECKLIST.md](CHECKLIST.md) を参照してください。
