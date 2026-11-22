# PowerPoint Math Font Add-in

デスクトップ版 PowerPoint の数式（Math Zone）のフォントをまとめて変更する VBA アドインです。Windows と Mac で動作確認済みです。Mac で使う場合は [Mac の手動インストール](#mac-の手動インストール)を参照してください。

既定のフォントは **Latin Modern Math** です。Computer Modern 系の見た目を PowerPoint 数式で使いたい用途を想定しています。

## Features

PowerPoint のリボンに **Math Font** タブを追加します。

- **Selection** — 選択範囲に含まれる数式だけ変更
- **Current slide** — 現在のスライドの数式を一括変更
- **All slides** — プレゼン全体の数式を一括変更
- **Font** — 候補から数式フォントを選択。別のフォント名の直接入力も可能
- **Auto format** — 選択が変わるたび、現在のスライドの数式を自動で整える
- **Restore Cambria** — プレゼン全体の数式を Cambria Math に戻し、Auto format を OFF にする

テキストボックス全体のフォントを置換するのではなく、PowerPoint の `TextRange2.MathZones` を使って数式ゾーンだけを対象にします。グループ化された図形と表セルも再帰的に処理します。

通常の適用・復元・フォント選択では完了ポップアップを表示しません。操作できる数式がない場合も何も変更せず終了します。エラーの通知だけを表示します。

## Requirements

- VBA を実行できるデスクトップ版 Microsoft PowerPoint
- 使用する OpenType Math フォントが OS にインストール済みであること
- `.ps1` でのインストールと `.ppam` のビルドには Windows 版 PowerPoint が必要

## Fonts

**使用する数式フォントは、アドインとは別にインストールしてください。** リボンの **Font** には、インストール状況にかかわらず候補が表示されます。

未インストールのフォント名を指定しても、そのフォントの見た目にはなりません。フォント名だけが設定され、表示には代替フォントが使われるため、見た目が変わらない場合もあります。別の PC で編集・発表する場合も、使用フォントの導入状況を確認してください。[Microsoft のフォント導入・共有の説明](https://support.microsoft.com/en-us/office/fonts/download-and-install-custom-fonts-to-use-with-office)

### 数式フォントの候補

| フォント | 特徴 | 数式用ファイルの例 | 配布元 |
| --- | --- | --- | --- |
| Latin Modern Math | Latin Modern の本文書体に合う、Computer Modern 系の数式フォント。 | `latinmodern-math.otf` | [CTAN](https://ctan.org/pkg/lm-math) |
| New Computer Modern Math | Computer Modern 系の拡張。Regular と、やや太い Book などがある。 | `NewCMMath-Regular.otf`、`NewCMMath-Book.otf` | [CTAN](https://ctan.org/pkg/newcomputermodern) |
| STIX Two Math | STIX Two Text に合う数式フォント。科学・技術出版向けの幅広い数式記号を収録。 | `STIXTwoMath-Regular.otf` | [CTAN](https://ctan.org/pkg/stix2-otf) |
| TeX Gyre Math | TeX Gyre の本文書体に合わせて選べる数式フォント群。Termes、Pagella、Bonum、Schola がある。 | `texgyretermes-math.otf`、`texgyrepagella-math.otf` | [CTAN](https://ctan.org/pkg/tex-gyre-math) |

表の配布元から数式用の `.otf` を取得し、Windows では右クリックから［インストール］、Mac では **Font Book** でインストールします。PowerPoint を開き直し、**Math Font > Font** から選びます。[Windows の手順](https://support.microsoft.com/en-us/windows/experience/personalization/manage-fonts-in-windows) · [Mac の手順](https://support.apple.com/en-ca/guide/font-book/fntbk1000/mac)

ドロップダウンには `NewComputerModernMath`、`TeX Gyre Termes Math`、`TeX Gyre Pagella Math` などが入っています。版・ウェイトによって名前が異なる場合や一覧にない書体は、OS のフォント一覧で正確な名前を確認し、**Font** 欄に入力してください。各フォントの互換性は、使う数式で確認してください。

## Install

### Windows 方法 1 — `.ps1` でビルド・登録

1. このリポジトリを ZIP でダウンロードして展開するか、clone します。
2. PowerPoint で、ビルドに必要な以下の設定を一時的に ON にします。

   `ファイル > オプション > トラスト センター > トラスト センターの設定 > マクロの設定 > VBA プロジェクト オブジェクト モデルへのアクセスを信頼する`

3. PowerPoint を終了し、展開先またはリポジトリのディレクトリで PowerShell を開いて実行します。

   ```powershell
   powershell -ExecutionPolicy Bypass -File .\install.ps1
   ```

4. PowerPoint を開き、リボンに **Math Font** タブが出ることを確認します。上の VBA プロジェクトへのアクセス設定は元に戻せます。

### Windows 方法 2 — `.ppam` を手動で登録

1. このリポジトリの [`PptMathFontAddin.ppam`](./PptMathFontAddin.ppam) を取得し、次の場所に置きます。GitHub のファイル画面では **Download raw file** を使います。フォルダーは必要に応じて作成します。

   ```text
   %APPDATA%\Microsoft\AddIns\PptMathFontAddin\PptMathFontAddin.ppam
   ```

2. PowerPoint の **ファイル > オプション > アドイン** を開きます。**管理** で **PowerPoint アドイン** を選んで **設定** を押します。
3. **新規追加** で `.ppam` を選び、有効にします。リボンに **Math Font** タブが出れば登録完了です。

既に `PptMathFontAddin` が登録されている場合は、PowerPoint を終了して同じ場所の `.ppam` を差し替えます。手動登録では、PowerShell の実行や VBA プロジェクトへのアクセス許可は不要です。

ダウンロードしたファイルが Windows でブロックされた場合は、配布元を確認し、信頼できる場合に限って、その `.ppam` のプロパティからブロックを解除して再度登録します。[Microsoft の手順](https://learn.microsoft.com/en-us/microsoft-365-apps/security/internet-macros-blocked#macro-enabled-add-in-files-for-powerpoint-and-excel)

### Mac の手動インストール

Mac では同梱の `.ppam` を手動で登録します。Mac 版 PowerPoint 16.115.1 で、アドインの読み込み、数式フォントの STIX Two Math への変更、Cambria Math への復元を確認しました。

1. [`PptMathFontAddin.ppam`](./PptMathFontAddin.ppam) を、登録後も使い続けるフォルダーに保存します。
2. デスクトップ版 PowerPoint のメニューバーで **ツール > PowerPoint アドイン**（**Tools > PowerPoint Add-ins**）を開き、**＋** で `.ppam` を追加します。
3. リボンに **Math Font** タブが表示されれば登録完了です。

動作を確認するには、テスト用スライドに PowerPoint 標準の数式を挿入し、**Font** で導入済みのフォントを選んで **Current slide** を押します。**Auto format** を試す場合は、オンにして新しい数式を追加し、選択を変えたときに指定フォントが適用されることを確認します。**Restore Cambria** を押すと数式を Cambria Math に戻せます。

## Usage

**Math Font > Font** でインストール済みの数式フォントを選び、既存スライドなら **All slides** を押します。フォントを選ぶ操作は使用するフォントの設定だけを変更します。

新しく数式を追加し続ける場合は **Auto format** を ON にすると、選択変更時に現在スライドの Math Zone を既定フォントへ戻します。

直立体・斜体は、数式内の対象文字を選択して PowerPoint の **ホーム > フォント > 斜体（I）** で切り替えます。本アドインに専用ボタンは設けていません。Selection / Current slide / All slides / Auto format は、指定済みの直立体・斜体を保ったままフォント名だけを変更します。

**Restore Cambria** は数式フォントを Cambria Math に変更する機能です。操作前の書式への Undo ではありません。復元直後に再変換されないよう Auto format も OFF にし、チェック表示と保存設定を更新します。再び自動変換したいときは Auto format を ON にしてください。

一覧にない OpenType Math フォントを使う場合は、**Font** 欄に正確なフォント名を入力し、Enter で確定してください。空欄で確定した場合は直前の設定を維持します。

## Caveat

PowerPoint の数式エンジンは Cambria Math 以外の OpenType Math フォントを完全には扱えない場合があります。Latin Modern Math でも、伸縮括弧など一部の構造が保存・再オープン後に崩れる可能性があります。

このアドインは PowerPoint の数式レンダラを置換するものではなく、既存の Math Zone にフォント名を設定する作業を自動化します。

最終版では、保存後に PowerPoint を一度閉じて再オープンし、数式表示を確認してから PDF 化する運用を推奨します。

> [!NOTE]
> Mac 版 PowerPoint の **電子配布とアクセシビリティに最適** は、Microsoft のオンラインサービスで PDF に変換します。変換サービスで指定の数式フォントを使えない場合、Cambria Math に置き換わることがあります。[Microsoft のオンライン変換の説明](https://support.microsoft.com/ja-jp/accessibility/office-accessibility/create-accessible-pdfs) · [フォント置換の説明](https://support.microsoft.com/en-gb/office/collab-files/save-or-convert-to-pdf-or-xps-in-office-desktop-apps)
>
> **印刷に最適** は Mac 上で変換し、インストール済みのフォントを使えます。数式の見た目を維持したいときは、この設定で書き出した PDF を確認してください。Windows デスクトップ版の **ファイル > エクスポート > PDF/XPS の作成** も PC 上で変換します。[Windows 版の PDF 書き出し手順](https://support.microsoft.com/ja-jp/powerpoint/training/save-powerpoint-presentations-as-pdf-files)

## Uninstall

Windows のスクリプトでインストールした場合は、PowerPoint を終了して以下を実行します。

```powershell
powershell -ExecutionPolicy Bypass -File .\uninstall.ps1
```

手動でインストールした場合は、PowerPoint のアドイン画面から `PptMathFontAddin` の登録を解除し、PowerPoint を終了して `.ppam` を削除します。

## Build the PPAM

配布用 `.ppam` を更新するには、Windows で VBA プロジェクトへのアクセスを許可し、PowerPoint を終了して実行します。リポジトリ直下の `PptMathFontAddin.ppam` が更新されます。PowerPoint への登録は行いません。

```powershell
powershell -ExecutionPolicy Bypass -File .\install.ps1 -BuildOnly
```

## License

アドインと関連コードは MIT ライセンスで配布します。[ライセンス全文](./LICENSE)

数式フォントには各配布元のライセンスが適用されます。
