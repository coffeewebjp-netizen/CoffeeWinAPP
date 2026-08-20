# DESIGN

この文書は Coffee AutoButton の現在仕様と改良方針を固定するための設計メモです。

現在の実装メモ: v1.0.10

## 目的

Coffee AutoButton は、指定したキー入力またはクリック操作を一定間隔で繰り返す Windows WPF アプリです。特にクリック連打では、作業中のユーザーのマウス位置やアクティブウィンドウを奪わないことを重視します。

## 非目標

- 通常の Chrome ウィンドウへ無理に非干渉クリックを送ること。
- Chrome 系ブラウザで失敗した時に、自動的に `PostMessage` や物理クリックへ戻すこと。
- ゲーム、管理者権限アプリ、独自描画 UI など、外部入力を拒否するアプリを完全対応と見なすこと。

## 全体構成

- `MainWindow`: WPF UI ファサード。設定/ホットキーは `MainWindow.Presets.cs`、実行制御は `MainWindow.RunLoop.cs`、対象位置取得は `MainWindow.TargetWindow.cs`。
- `KeyboardInputService`: キー入力、修飾キー、長押し、シーケンス送信。
- `MouseClickService`: 非ブラウザ向け Win32 メッセージクリックと物理クリック。
- `BrowserDirectClickService`: 専用ブラウザの CDP 検出とクリック送信。
- `NativeMethods`: Win32 API の P/Invoke 宣言。
- `AppSettingsManager`: 前回設定の保存、復元、正規化。
- `PresetManager`: プリセット保存、削除、正規化。

## クリック戦略

### 1. Chrome 系ブラウザ

Chrome、Edge、Chromium、Brave、Vivaldi、Opera、Electron 系は Chromium 系として扱います。

Chromium 系では専用ブラウザ CDP 方式だけを使います。専用ブラウザとして認識できない場合はクリック送信を止めます。

実装条件:

- 専用ブラウザは CDP ポート `9223` で起動する。
- CDP 探索は `127.0.0.1:9223` と `[::1]:9223` の両方を確認する。CoffeeBook など同じポートを使う別ツールとの同時利用時に、片方が IPv4、もう片方が IPv6 ループバックへ分かれる場合があるため。
- プロファイルは `%APPDATA%\CoffeeAutoButton\dedicated-browser-profile` を使う。
- 位置取得時に `BrowserClickTarget` を解決し、CDP の `targetId`、URL、タイトルを保持する。
- クリック送信時は保持した `BrowserClickTarget` を再解決してから `Input.dispatchMouseEvent` を送る。
- `window.focus`、フォーカスエミュレーション、Topmost 化は行わない。
- Chrome 系の失敗時に `PostMessage` や物理クリックへ自動フォールバックしない。

この方針は、非干渉クリック中にアクティブウィンドウを奪われる問題を避けるための中核仕様です。

### 2. 非ブラウザアプリ

非ブラウザアプリでは、対象ウィンドウに対して `PostMessage` でマウスメッセージを送信します。カーソル位置を変えないため、物理クリックよりも非干渉性が高い方式です。

ただし、アプリ側が外部メッセージを受け付けない場合があります。その場合は明示的な互換モードとして物理クリックを使います。

### 3. 物理クリック

物理クリックはカーソル移動を伴います。指定位置へ移動してクリックし、元の位置へ戻す方式です。互換性は高い一方、作業中の操作を妨げる可能性があるため、非干渉の第一選択にはしません。

## 対象認識

位置取得では、スクリーン座標、クライアント座標、対象ウィンドウ、クラス名、プロセス名、PID、権限状態を保存します。

Chrome 系では、子ウィンドウのタイトルが `Chrome Legacy Window` などになることがあります。そのため Chromium 系または Chrome 系クラス名の場合は、ルートウィンドウのタイトルも使って CDP 対象を解決します。

UI では次の状態を表示します。

- `専用ブラウザ認識: CDP:9223 / ...`
- `専用ブラウザ未認識`
- 対象プロセス、ウィンドウクラス、権限状態

## UI 方針

- タイトルバーと画面下部にバージョンを表示する。
- クリック方式と対象認識状態を、テスト前に確認できるようにする。
- 専用ブラウザ URL は UI から指定できるようにする。
- 失敗時は、Chrome 系と非ブラウザで案内を分ける。
- Chrome 系の失敗時は、物理クリックへの切り替え確認を出さない。

## 設定とプリセット

保存先は `%APPDATA%\CoffeeAutoButton` です。

- 前回設定: `last-settings.json`
- プリセット: `presets.json`
- 専用ブラウザプロファイル: `dedicated-browser-profile`

設定とプリセットには `SchemaVersion` を持たせ、読み込み時に現在仕様へ正規化します。専用ブラウザ URL も設定とプリセットに含めます。

## バージョンと配布

配布時は次を同じバージョンにそろえます。

- `CoffeeAutoButton.csproj` の `Version`
- `AssemblyVersion`
- `FileVersion`
- `tools/build-msi.ps1` の既定 MSI バージョン
- UI に表示されるバージョン

MSI は `installer\out\CoffeeAutoButtonSetup.msi` に生成します。ユーザー単位で `%LOCALAPPDATA%\CoffeeAutoButton` にインストールし、スタートメニューに `Coffee AutoButton` のショートカットを作成します。

## 改良方針

今後の優先順位:

1. 専用ブラウザ認識の分かりやすさを上げる。
2. CDP 対象の再取得や URL/タイトル変化への耐性を上げる。
3. 失敗時のメッセージを、ユーザーが次に取る操作へ直結させる。
4. UI をさらにモダンにしつつ、状態確認の密度を落とさない。
5. インストーラー、ショートカット、表示バージョンの検証を release 手順に固定する。
6. 手動確認項目を `CHECKLIST.md` に集約する。

## 検証観点

- `dotnet build .\CoffeeAutoButton.sln` が成功する。
- アプリ画面に現在バージョンが表示される。
- 専用ブラウザ起動後、位置取得で `専用ブラウザ認識: CDP:9223` が表示される。
- 専用ブラウザへのテストクリックで、現在のアクティブウィンドウが奪われない。
- Chrome 系で専用ブラウザ未認識の場合、物理クリックへの自動切り替えが起きない。
- 非ブラウザアプリでは Win32 メッセージクリックが動作する。
- MSI インストール後、スタートメニューから `Coffee AutoButton` を起動できる。
- インストール済みアプリの表示バージョンと exe のバージョンが一致する。
