# Design Index

Coffee AutoButton の設計情報を読むための入口です。
毎回すべての docs を読まず、触る面だけに進みます。

## 読む順番

1. `docs/DESIGN_INDEX.md`
2. `DESIGN.md`
3. `CHECKLIST.md`
4. 関連する C# だけ

Do not search `bin/`, `obj/`, `publish/`, `installer/out/`, `installer/payload/`, or `.vs/`.

| 作業 | 先に読む文書 | 主な実装入口 |
| --- | --- | --- |
| いまの設計メモ | `DESIGN.md` | `CoffeeAutoButton/` |
| 検証手順 | `CHECKLIST.md` | ビルド、ホットキー、クリック、プリセット |
| 起動とビルド | `README.md` | `CoffeeAutoButton.sln` |
| UI / 設定 / ホットキー | `DESIGN.md` | `MainWindow.xaml.cs` と `MainWindow.Presets.cs` |
| 実行ループ | `DESIGN.md` | `MainWindow.RunLoop.cs` |
| 対象ウィンドウ / CDP | `DESIGN.md` | `MainWindow.TargetWindow.cs`, `BrowserDirectClickService.cs` |

## 文書の役割

- `README.md` は起動・ビルド案内。
- `DESIGN.md` は現在仕様とクリック戦略の置き場。
- `CHECKLIST.md` はビルドと手動確認の手順。
- 性能メモは `docs/PERFORMANCE_*.md`。通常作業の入口にはしない。
