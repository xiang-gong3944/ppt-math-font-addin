# PowerPoint Math Font Add-in

PowerPoint の Office 数式 (Math Zone) のフォントだけをまとめて変更する Windows 向け VBA アドインです。

既定のフォントは **Latin Modern Math** です。Computer Modern 系の見た目を PowerPoint 数式で使いたい用途を想定しています。

## Features

PowerPoint のリボンに **Math Font** タブを追加します。

- **Selection** — 選択範囲に含まれる数式だけ変更
- **Current slide** — 現在のスライドの数式を一括変更
- **All slides** — プレゼン全体の数式を一括変更
- **Auto format** — 選択が変わるたび、現在のスライドの数式を自動で整える
- **Set font...** — 使用する数式フォント名を変更
- **Restore Cambria** — プレゼン全体を Cambria Math に戻す

テキストボックス全体のフォントを置換するのではなく、PowerPoint の `TextRange2.MathZones` を使って数式ゾーンだけを対象にします。グループ化された図形と表セルも再帰的に処理します。

## Requirements

- Windows 版 Microsoft PowerPoint
- VBA が利用可能なデスクトップ版 Office
- Latin Modern Math を使う場合、そのフォントが Windows にインストール済みであること

## Install

1. このリポジトリを ZIP でダウンロードするか clone します。
2. PowerPoint をすべて閉じます。
3. 初回ビルドのため、PowerPoint で一時的に以下を ON にします。

   `ファイル > オプション > トラスト センター > トラスト センターの設定 > マクロの設定 > VBA プロジェクト オブジェクト モデルへのアクセスを信頼する`

4. PowerShell で実行します。

   ```powershell
   powershell -ExecutionPolicy Bypass -File .\install.ps1
   ```

5. PowerPoint を開き、リボンに **Math Font** タブが出ることを確認します。
6. インストール後は「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」を OFF に戻して構いません。

生成されたアドインは以下に保存されます。

```text
%APPDATA%\Microsoft\AddIns\PptMathFontAddin\PptMathFontAddin.ppam
```

## Usage

既存スライドなら **Math Font > All slides** を押すだけです。

新しく数式を追加し続ける場合は **Auto format** を ON にすると、選択変更時に現在スライドの Math Zone を既定フォントへ戻します。

別の OpenType Math フォントを使いたい場合は **Set font...** から Windows 上の正確なフォント名を入力してください。

## Caveat

PowerPoint の数式エンジンは Cambria Math 以外の OpenType Math フォントを完全には扱えない場合があります。Latin Modern Math でも、伸縮括弧など一部の構造が保存・再オープン後に崩れる可能性があります。

このアドインは PowerPoint の数式レンダラを置換するものではなく、既存の Math Zone にフォント名を設定する作業を自動化します。

最終版では、保存後に PowerPoint を一度閉じて再オープンし、数式表示を確認してから PDF 化する運用を推奨します。

## Uninstall

PowerPoint をすべて閉じてから以下を実行してください。

```powershell
powershell -ExecutionPolicy Bypass -File .\uninstall.ps1
```
