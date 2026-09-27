<a name="readme-top"></a>

# powerpoint-converter

PowerPoint の資料（.pptx）を，**ローカルのブラウザで PowerPoint と同じ並びの画面に出し，見ながら直す**ツールです．
Claude Code などのツールがファイルを直した結果も，数秒で画面に出ます．

Docker も sudo も使いません．Python と LibreOffice だけで動きます．

<details>
  <summary>目次</summary>
  <ol>
    <li><a href="#どんなものか">どんなものか</a></li>
    <li><a href="#画面">画面</a></li>
    <li><a href="#動作環境">動作環境</a></li>
    <li><a href="#使いはじめる">使いはじめる</a></li>
    <li><a href="#使い方">使い方</a></li>
    <li><a href="#キー操作">キー操作</a></li>
    <li><a href="#claude-code-に直してもらう">Claude Code に直してもらう</a></li>
    <li><a href="#しくみ">しくみ</a></li>
    <li><a href="#コマンド一覧">コマンド一覧</a></li>
    <li><a href="#ファイル構成">ファイル構成</a></li>
    <li><a href="#git管理上の注意">Git管理上の注意</a></li>
    <li><a href="#うまくいかないとき">うまくいかないとき</a></li>
    <li><a href="#関連資料">関連資料</a></li>
    <li><a href="#作成者">作成者</a></li>
    <li><a href="#ライセンス">ライセンス</a></li>
  </ol>
</details>

## どんなものか

VS Code などの PPTX の拡張機能の多くは，PPTX を JavaScript で描き直すため，
テンプレートの背景やマスターの図形が抜けたり（透けたり），ファイルを直しても表示が変わらなかったりします．
このツールは **LibreOffice 本体で描いた絵をブラウザに出し，直した内容は PPTX の中身（XML）に直接書き込みます**．

- **見た目が崩れない**：背景・テンプレート・画像の透過まで，LibreOffice で開いたときと同じに見えます
- **ほかで直したものもすぐ出る**：Claude Code・LibreOffice・PowerPoint でファイルを直すと，約 1 秒で気づき，変わったスライドだけを描き直します（1 枚 0.5〜1 秒）
- **ブラウザで直せる**：図形の移動・大きさ・回転，文字（一部の文字だけの書式も），表，画像，スライドの追加・並べ替え，ノート，アニメーション，画面切り替え
- **アニメーションも動く**：スライドショーと，編集画面の上でのプレビュー
- **PowerPoint で開いても崩れにくい**：LibreOffice で保存し直さず，触った部分の XML だけを書き換えます．変えていない部品はバイト単位で同じままです
- **.ppt・.odp なども見られます**（直せるのは .pptx）

<p align="right">(<a href="#readme-top">上に戻る</a>)</p>

## 画面

```text
┌ ↶↷ │ 📋✂⎘ │ ＋新しいスライド │ 𝐓 ◇図形 ▦表 🖼画像 │ ⊞配置 ⧉ 🗑 │ 🔍検索 │ ▶最初から ▷このスライドから │ ⤓ ┐
├──────┬──────────────────────────────────────────┬────────────────────────────────┤
│  1 ▭ │ （検索と置換の帯：Ctrl+F）                │ 書式│アニメーション│画面切替│選択 │
│  2 ▭ │                                          │                                │
│★ 3 ▭ │                スライド                   │  位置とサイズ・文字・段落        │
│⇢ 4 ▭ │                                          │  塗り・線・配置・重なりの順番     │
│  …   ├──────────────────────────────────────────┤                                │
│      │ ノート                                    │                                │
└──────┴──────────────────────────────────────────┴────────────────────────────────┘
 ◧ スライド 3 / 20   ✓ 描画済み                       ⚠ フォント   − 合わせる ＋   ◨
```

| 場所 | できること |
| --- | --- |
| **左：スライド一覧** | クリックで開く．ドラッグで並べ替え．右クリックで新規・複製・削除・非表示・画像で保存．★ はアニメーション，⇢ は画面切り替えあり．外で直されたスライドは黄色く光ります |
| **中央：スライド** | クリックで選ぶ（Shift で追加），空いた所をドラッグで範囲選択，**右クリックでメニュー**．ドラッグで移動（吸着のガイドが出ます），四隅と辺で大きさ，上の ⟳ で回転．**ダブルクリックで文字を編集**．グループはダブルクリックで中に入り，表はダブルクリックで編集します |
| **下：ノート** | 入力が止まって約 1 秒で保存します |
| **右：書式** | 位置とサイズ・回転・反転，フォント・大きさ・太字・斜体・下線・取り消し線・上付き・下付き・色・蛍光ペン，揃え・上下の位置・箇条書き・段落番号・インデント・行間・自動調整，塗り・線（色・太さ・種類・矢印），図の変更，配置．何も選ばないと，スライドの背景・崩れの警告・フォントの置き換え |
| **右：アニメーション** | クリックの順番ごとの一覧．▲▼ で順番，開始のしかた（クリック時／同時／後），✕ で削除．効果を選ぶと種類・長さ・遅延を変えられます．効果の追加と **▶ プレビュー**（スライドの上で効果を順に再生します．追加したときも自動で再生します） |
| **右：画面切替** | スライドに入るときの効果（18 種類），速さ，クリックで次へ・自動で次へ．すべてのスライドにも当てられます |
| **右：選択** | PowerPoint の選択ウィンドウと同じです．重なりの順番の一覧，👁 で表示・非表示，ダブルクリックで名前を変えます |

<p align="right">(<a href="#readme-top">上に戻る</a>)</p>

## 動作環境

| 項目 | 内容 |
| --- | --- |
| OS | Linux（Ubuntu 24.04 で確認） |
| Python | 3.10 以上（標準ライブラリだけで動きます） |
| LibreOffice | Impress と python3-uno（Ubuntu なら `sudo apt install libreoffice-impress python3-uno`） |
| ブラウザ | Chrome・Firefox など |

画面で使う部品はすべてリポジトリに入っているので，インターネットにつながっていなくても動きます．
LibreOffice は裏で動かすだけで，普段使っている LibreOffice の設定や見た目は変えません．

<p align="right">(<a href="#readme-top">上に戻る</a>)</p>

## 使いはじめる

```bash
git clone https://github.com/kurokara-YK/powerpoint-converter.git
cd powerpoint-converter
bash install.sh
```

`install.sh` は次のことをします（何度実行しても大丈夫です）．

1. LibreOffice と python3-uno があるか確かめる（無ければ入れ方を表示します．sudo は実行しません）
2. `powerpoint-converter` コマンドを `~/.local/bin` に作る
3. 裏で使う LibreOffice の専用の設定を作っておく（初回の表示を速くするため）

起動します．

```bash
powerpoint-converter
```

ブラウザで `data/` の一覧が開きます．止めるときは端末で **Ctrl+C** を押します．

| 起動のしかた | 開くもの |
| --- | --- |
| `powerpoint-converter` | `data/` の一覧 |
| `powerpoint-converter data/フォルダ/` | そのフォルダの一覧 |
| `powerpoint-converter data/フォルダ/資料.pptx` | その資料 |
| `powerpoint-converter ~/どこか/資料.pptx` | `data/` の外の資料（そのフォルダを置き場として開きます） |
| `powerpoint-converter --no-browser` | ブラウザを開かない（表示された URL を自分で開きます） |
| `powerpoint-converter --port 9000` | ポート番号を変える（既定は 8766．使用中なら次の番号を使います） |

<p align="right">(<a href="#readme-top">上に戻る</a>)</p>

## 使い方

### 1. 資料を置く

`data/` の下に，好きな単位でフォルダを作って資料を置きます．
ブラウザの一覧の点線の枠へ**ドラッグ＆ドロップ**しても取り込めます（同じ名前があれば `_2` を付けます）．

### 2. 開いて直す

一覧で資料を押すと，PowerPoint と同じ並びの画面が開きます．
**保存は操作のたびに自動です．** 描き直すまでの約 1 秒は，前の絵を新しい位置に置いて見せます．

- **文字**：図形をダブルクリックして直します．Shift+Enter で段落の中の改行（↵），Esc で確定します．
  文字を選んで Ctrl+B などを押すと，その文字だけに書式が付きます
- **コピーと貼り付け**：別のスライド・別の資料にも貼れます．ほかのアプリでコピーした文字はテキスト ボックスに，画像は画像になります
- **画像**：ボタンのほか，貼り付けやスライドへのドロップでも入ります
- **表**：「表 ▾」で大きさを選んで入れます．ダブルクリックで，セルの文字と行・列を直します
- **検索と置換**：Ctrl+F．すべてのスライド・表・ノートを探します．置き換えなかった部分の書式は保ちます

### 3. アニメーションとスライドショー

- 右の「アニメーション」で効果を付けると，スライドの上で動きを再生します（**▶ プレビュー** でも再生できます）．
  下の帯で一時停止・次へ・最初から．スライドをクリックすると 1 つずつ進み，最後の状態のまま止まります（Esc か ✕ で閉じます）
- **▶ 最初から**（F5）・**▷ このスライドから**（Shift+F5）でスライドショーです．アニメーションと画面切り替えも動きます
- クリックか → で進み，**Esc ですぐ戻ります**．最後のスライドの次は「スライド ショーの最後です」と出ます

### 4. ほかで直したものも同じ画面に出る

| どこで直すか | 何が起きるか |
| --- | --- |
| ブラウザの画面 | すぐにファイルに書き込み，そのスライドを描き直します |
| Claude Code・LibreOffice・PowerPoint | 約 1 秒で気づき，変わったスライドだけを描き直します |

> **注意**
> ファイルは 1 つなので，**最後に保存したほうが残ります**．
> LibreOffice の画面で開いたまま，ブラウザで直してから LibreOffice で保存すると，ブラウザでの変更は上書きされます．
> 同じファイルを直すのは，どちらか一方にしてください．

### 5. ダウンロード

右上の「⤓ ダウンロード」から，PPTX・PDF（非表示のスライドは除きます）・このスライドの画像（PNG）を落とせます．

**初めて開く大きな資料**は，全部のスライドを描き終えるまで十数秒かかります（35 枚で約 12 秒．今見ているスライドは 1 秒ほど）．
一度描いたスライドは覚えておくので，次からはすぐに出ます．

<p align="right">(<a href="#readme-top">上に戻る</a>)</p>

## キー操作

| キー | すること |
| --- | --- |
| Ctrl+Z / Ctrl+Y | 元に戻す / やり直し |
| Ctrl+C / Ctrl+X / Ctrl+V | コピー / 切り取り / 貼り付け |
| Ctrl+D | 複製 |
| Delete | 選んだ図形を消す |
| Ctrl+G / Ctrl+Shift+G | グループ化 / グループ解除 |
| Ctrl+B / I / U | 太字 / 斜体 / 下線（文字を編集中なら，選んだ文字だけ） |
| Ctrl+] / Ctrl+[ | 文字を大きく / 小さく |
| Ctrl+L / E / R / J | 左 / 中央 / 右 / 両端揃え |
| Ctrl+Shift+] / Ctrl+Shift+[ | 最前面へ / 最背面へ |
| Ctrl+A / Tab | すべて選ぶ / 次の図形を選ぶ |
| 矢印（Shift で大きく・Ctrl で細かく） | 選んだ図形を少し動かす．何も選んでいなければスライドを移る |
| Enter / F2 | 選んだ図形の文字（表なら表）を編集 |
| Ctrl+F / Ctrl+H | 検索 / 置換 |
| Ctrl+M | 新しいスライド |
| F5 / Shift+F5 | 最初から / このスライドから再生 |
| Ctrl+ホイール | 拡大・縮小 |

<p align="right">(<a href="#readme-top">上に戻る</a>)</p>

## Claude Code に直してもらう

Claude Code に直したい内容を頼みます．Claude Code は [AGENTS.md](AGENTS.md) の規則に従い，
`outline` で中身を見て，`op` で書式を保ったまま直し，`check` で確かめます．
ブラウザを開いておけば，直された箇所がそのまま見えます．

```bash
powerpoint-converter outline data/フォルダ/資料.pptx --slide 3   # 図形の id・位置・文字・アニメーション
powerpoint-converter op      data/フォルダ/資料.pptx replace '{"find": "旧名称", "replace": "新名称"}'
powerpoint-converter check   data/フォルダ/資料.pptx             # 全スライドの PNG と，崩れ・壊れの一覧
```

`check` は全スライドを PNG にし（非表示のスライドも），文字のあふれ・スライドの外へのはみ出し・
PowerPoint が修復を求める原因・フォントの置き換えを出します．壊れがあれば終了コード 1 を返します．

<p align="right">(<a href="#readme-top">上に戻る</a>)</p>

## しくみ

```text
data/資料.pptx         … 正本（ブラウザも Claude Code も，このファイルを直接直す）
    ↓ 1 枚ずつ
LibreOffice（裏で常駐） … 背景（マスター込み）と図形ごとの透明な絵を描く
    ↓
ブラウザ               … 絵を重ねて出す．図形を動かすと，その図形の絵だけが動く
    ↓ 操作
PPTX の XML            … 触った部分だけを書き換えて保存（LibreOffice には保存させない）
```

**速く描くため，次のようにしています．**

- 直したスライドは，そのスライドだけを含む PPTX を作って描きます（1 枚 0.5〜1 秒）
- 初めて開いたときなど描く枚数が多いときは，資料をまるごと 1 回読み込み，LibreOffice を 4 つまで並べて分担します
- 描いた結果は，スライドの中身から作った印ごとに `~/.cache/powerpoint-converter/` に覚えておきます．同じ中身なら描き直しません

**LibreOffice の図形と PPTX の図形は，描く用の写しで図形の説明欄に id を書き込んで対応づけています．**
このため，並び順に頼らずに，ドラッグした図形が XML のどの図形かを正しく決められます．

詳しい設計と理由は [AGENTS.md](AGENTS.md) の「設計で決めたこと」にあります．

<p align="right">(<a href="#readme-top">上に戻る</a>)</p>

## コマンド一覧

| コマンド | すること |
| --- | --- |
| `powerpoint-converter [資料 か フォルダ]` | ブラウザで開く（`serve` を省いた形） |
| `powerpoint-converter check <資料> [--json] [--strict]` | 全スライドを PNG にし，崩れ・壊れ・フォントの置き換えを出す |
| `powerpoint-converter outline <資料> [--slide N]` | スライドと図形（id・位置・文字）・アニメーション・ノートの一覧 |
| `powerpoint-converter op <資料> <操作> '<JSON>'` | 画面と同じ操作を当てる（下の表） |
| `powerpoint-converter render <資料> [-o 出力先]` | 全スライドを PNG にする（非表示も含む） |
| `powerpoint-converter pdf <資料> [-o 出力.pdf]` | PDF にする |
| `powerpoint-converter import <資料> [フォルダ]` | `data/` に写す |

`op` で使える操作：

| 種類 | 操作 |
| --- | --- |
| 文字 | `text` `text_style` `style` `replace` |
| 図形 | `frame` `rotate` `flip` `visibility` `rename` `delete` `duplicate` `zorder` `add` `group` `ungroup` |
| 表 | `table_text` `table_row` `table_col` |
| スライド | `slide_new` `slide_dup` `slide_delete` `slide_move` `slide_hide` `notes` `transition` `background` |
| アニメーション | `anim_add` `anim_delete` `anim_move` `anim_trigger` `anim_timing`（長さ・遅延） `anim_change`（種類） |

引数の例は [AGENTS.md](AGENTS.md) の「Claude Code が資料を直すときの規則」にあります．

<p align="right">(<a href="#readme-top">上に戻る</a>)</p>

## ファイル構成

```text
powerpoint-converter/
├── README.md                  このファイル
├── AGENTS.md                  要件と方針（Claude Code などのエージェントが読む）
├── CLAUDE.md                  Claude Code の入口（@AGENTS.md）
├── install.sh                 初回のセットアップ
├── powerpoint_converter.sh    起動スクリプト
├── pyproject.toml
├── powerpoint_converter/      プログラム本体
│   ├── cli.py                 コマンド
│   ├── server.py              ローカルサーバ（127.0.0.1 のみ）
│   ├── deck.py                開いている資料（ファイルの監視・描画の順番・元に戻す）
│   ├── office.py              LibreOffice の常駐と描画
│   ├── render.py              描画用の PPTX とキャッシュ
│   ├── pptx.py                PPTX の読み書き
│   ├── edit.py ほか            操作（clip・slides・tables・search・anim）
│   └── static/                画面（index.html・app.css・js/）
└── data/                      資料の置き場（Git では管理しない）
```

<p align="right">(<a href="#readme-top">上に戻る</a>)</p>

## Git管理上の注意

`data/` の中身は `.gitignore` で除外しています．
**資料には個人名・所属・未発表の内容が含まれることが多いため，このリポジトリには含めないでください．**

資料の履歴を残したい場合は，`data/` の中のフォルダごとに `git init` し，**非公開のリポジトリ**で管理してください．

描画の結果（スライドの画像）は `~/.cache/powerpoint-converter/` に置くので，リポジトリには入りません．

<p align="right">(<a href="#readme-top">上に戻る</a>)</p>

## うまくいかないとき

| 症状 | 確認すること |
| --- | --- |
| 「python3-uno が無い」 | `sudo apt install python3-uno` |
| 「LibreOffice が見つからない」 | `sudo apt install libreoffice-impress` |
| `powerpoint-converter` が見つからない | 端末を開き直す．または `bash powerpoint_converter.sh` で起動する |
| 改行の位置が PowerPoint と違う | 右の「書式」（何も選ばない状態）の「フォントの置き換え」を見る．元のフォントを `~/.fonts` に入れて `fc-cache -f` |
| ブラウザで直した内容が消えた | LibreOffice の画面でも同じファイルを開いて保存していないか確認する（最後に保存したほうが残ります） |
| 「ほかで保存された内容がある」と出る | 外でファイルが直されました．画面は最新を読み込んだので，もう一度操作します |
| LibreOffice ではゆっくり動くのに，ブラウザでは速い | ブラウザは資料に書かれた長さどおりに動きます（PowerPoint と同じ）．LibreOffice は描画が重いと効果を引き延ばして見せることがあります．ゆっくりにしたいときは，右の「アニメーション」で効果を選び，長さを変えます |
| 描画がおかしいまま | 端末で Ctrl+C して起動し直す．`~/.cache/powerpoint-converter/renders` を消すと全部描き直します |
| ポートが使用中 | 自動で次の番号を使います．表示された URL を開きます |

<p align="right">(<a href="#readme-top">上に戻る</a>)</p>

## 関連資料

- [LibreOffice](https://www.libreoffice.org/) ／ [LibreOffice API（UNO）](https://api.libreoffice.org/)
- [Office Open XML（ECMA-376）](https://ecma-international.org/publications-and-standards/standards/ecma-376/)
- [overleaf-compiler](https://github.com/kurokara-YK/overleaf_compiler)（同じ考え方で作った LaTeX 用のツール）

<p align="right">(<a href="#readme-top">上に戻る</a>)</p>

## 作成者

| 項目 | 内容 |
| --- | --- |
| 作成者 | kurokara-YK |
| 連絡先 | kurokara1226@gmail.com |
| リンク集 | https://lit.link/kurokara |

不具合の報告や改善の提案は，Issues または上記のメールアドレスまでお願いします．

<p align="right">(<a href="#readme-top">上に戻る</a>)</p>

## ライセンス

このリポジトリのコードは自由に利用・改変してください．

<p align="right">(<a href="#readme-top">上に戻る</a>)</p>
