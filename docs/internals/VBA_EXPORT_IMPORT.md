# VBA エクスポート/インポート — 内部技術メモ

## 概要

`vba-extractor` は `.xlsm`（および類似形式）ファイル内の `xl/vbaProject.bin` を読み書きする。
本ドキュメントはバイナリ形式の仕様と実装上の要点をまとめたものである。

---

## vbaProject.bin の CFB 構造

```
Root Entry/
  VBA/
    dir           — OVBA 圧縮された dir ストリーム（モジュール一覧）
    _VBA_PROJECT  — プロジェクトレベルのコンパイル済み P-Code キャッシュ
    __SRP_0       — プロジェクトレベルのパフォーマンスキャッシュ（マジック: 93 4B）
    __SRP_1 … N   — モジュールごとのパフォーマンスキャッシュ（マジック: 72 55）
    Module1       — OVBA 圧縮された VBA ソース（+ 省略可能なプリアンブル）
    Class1
    Sheet1
    ThisWorkbook
    UserForm1     — UserForm のコード部分ストリーム
  PROJECT         — テキストストリーム: モジュール宣言 + [Workspace]
  PROJECTwm       — バイナリ: ASCII→UTF-16LE モジュール名マッピング
  UserForm1/      — ストレージ: フォームのバイナリレイアウト（UserForm のみ）
    f             — フォームバイナリデータ（コントロール・配置・プロパティ）
    o             — OLE ストレージ
    CompObj       — COM クラス識別子
    VBFrame       — テキスト: フォームのバージョン/キャプションヘッダー
```

### FAT と miniSAT のセクター割り当て

- ストリームサイズ **≥ 4096 バイト** → FAT（通常）セクター
- ストリームサイズ **< 4096 バイト** → miniSAT セクター
- `cfb.js`（`CFB.write`）はサイズ変更時の FAT↔miniSAT の切り替えを正しく処理する

---

## dir ストリームの形式（[MS-OVBA] 2.3.4.2）

dir ストリームは OVBA 圧縮されている。展開後はバイナリレコードが並ぶ：

```
[id: u16][size: u32][data: size バイト]
```

### プロジェクトレベルのレコード（抜粋）

| ID     | 名前                  | 備考                                             |
|--------|-----------------------|--------------------------------------------------|
| 0x0001 | PROJECTSYSKIND        | 0=Win16, 1=Win32, 2=Mac, 3=Win64               |
| 0x0003 | PROJECTCODEPAGE       | 例: 日本語環境では 932                           |
| 0x0009 | PROJECTVERSION        | 特殊: size バイト + 追加の MinorVersion 2 バイト |
| 0x000F | PROJECTMODULES        | モジュール数（u16）                              |
| 0x0013 | PROJECTCOOKIE         | 書き込み時は 0xFFFF にすること                  |
| 0x0010 | PROJECTTERMINATOR     | 終端マーカー                                     |

### モジュールごとのレコード（順序が重要）

| ID     | 名前                        | データサイズ | 備考                        |
|--------|-----------------------------|--------------|-----------------------------|
| 0x0019 | MODULENAME                  | 可変         | コードページ名              |
| 0x0047 | MODULENAMEUNICODE           | 可変         | UTF-16LE 名                 |
| 0x001A | MODULESTREAMNAME            | 可変         | コードページ版ストリーム名  |
| 0x0032 | MODULESTREAMNAME Unicode    | 可変         | UTF-16LE 版ストリーム名     |
| 0x001C | MODULEDOCSTRING             | 可変         | 通常は空                    |
| 0x0048 | MODULEDOCSTRING Unicode     | 可変         | 予約済み、通常は空          |
| 0x0031 | MODULEOFFSET                | 4            | 圧縮ソースの開始バイトオフセット |
| 0x001E | MODULEHELPCONTEXT           | 4            | 4 バイト u32               |
| 0x002C | MODULECOOKIE                | 2            | 書き込み時は 0xFFFF にすること |
| 0x0021 | MODULETYPE（手続き型）      | 0            | 0x0022 と排他               |
| 0x0022 | MODULETYPE（クラス/ドキュメント）| 0       | 0x0021 と排他               |
| 0x002B | MODULETERMINATOR            | 0            | モジュールブロックの終端    |

### MODULEOFFSET

- モジュールストリーム内で OVBA 圧縮ソースが始まるバイト位置を示す
- バイト 0..(MODULEOFFSET-1) = **プリアンブル**: コンパイル済み P-Code のリンクデータ
- MODULEOFFSET=0 はプリアンブルなしを意味し、圧縮ソースがバイト 0 から始まる
- **インポート時の方針**: 常に MODULEOFFSET=0 に設定する（プリアンブルなし）
- Excel は VBE 経由で保存するとき、常にプリアンブルを保持・再生成する（MODULEOFFSET=0 にはしない）

---

## P-Code キャッシュの無効化（ソースのみのプロジェクトにする）

モジュールのソースを差し替えると、コンパイル済みキャッシュが陳腐化する：

1. **各モジュールストリームのプリアンブル**（バイト 0..MODULEOFFSET-1）= そのモジュールの P-Code
2. `_VBA_PROJECT`: プロジェクトレベルのコンパイル済み P-Code（PerformanceCache）
3. `__SRP_0..N`: プロジェクト/モジュールごとのパフォーマンスキャッシュ

これらは互いに整合した「コンパイル済みのひとそろい」であり、**一部だけ残すと不整合**になる
（マクロ有効化でクラッシュ／シートコピー不可など）。インポートでは
**プロジェクト全体を「ソースのみ（デコンパイル状態）」にして、Excel に再コンパイルさせる**のが正解。

これは EPPlus（ソースから動作する .xlsm を生成する実績あるライブラリ）と同じ方式である。

### 方法（3 点セット）

1. **各モジュールストリーム** = 圧縮ソースのみ（プリアンブルを剥がす）。dir の **MODULEOFFSET=0**。
2. **`_VBA_PROJECT`** = 7 バイトの正規ヘッダーに置き換える。Version を `0xFFFF` にすると、
   どの VBA バージョンとも一致せず Excel が P-Code を破棄してソースから再コンパイルする。
3. **`__SRP_*`** = 全ストリームを削除する（type=0）。

```typescript
// [MS-OVBA 2.3.4.1] _VBA_PROJECT:
//   CC 61  Reserved1 (0x61CC)
//   FF FF  Version (0xFFFF → 再コンパイル強制)
//   00     Reserved3
//   00 00  Reserved4
// 長さ 7 = PerformanceCache 空（len = size - 7 = 0）
const vp = CFB.find(cfb, '/VBA/_VBA_PROJECT') as any;
if (vp) { vp.content = Buffer.from([0xCC, 0x61, 0xFF, 0xFF, 0x00, 0x00, 0x00]); vp.size = 7; }

// __SRP_* は全削除（残すと Excel のプロジェクト操作を壊す）
for (let i = 0; i < (cfb as any).FullPaths.length; i++) {
    if (((cfb as any).FullPaths[i] as string).includes('/__SRP_')) {
        (cfb as any).FileIndex[i].type = 0; // STGTY_INVALID
    }
}
```

> ⚠️ **やってはいけない無効化**: `_VBA_PROJECT` をサイズ維持のままボディだけゼロ埋めする方法は
> 不可。巨大なゼロ埋め PerformanceCache や Version=0x0000 は Excel の
> 「シートをコピーできません」エラーの原因になる。7 バイトに切り詰め、Version=0xFFFF にすること。

---

## CFB 署名ストリーム `Sh33tJ5` の扱い（重要な落とし穴）

`cfb.js`（SheetJS）は `CFB.write` のたびに `Sh33tJ5` という署名ストリームを**必ず**注入する
（`seed_cfb` が `rebuild_cfb` 経由で毎回再追加するため、API では抑止できない）。

**このストリームはそのまま有効ストリーム（type=2）として残すこと。Office は未知の CFB ストリームを無視する。**

> ⚠️ **やってはいけない**: 書き込み後のバイト列で署名エントリの type を 0（STGTY_INVALID）に
> 書き換えて「消す」方法は不可。隣接ノードの赤黒木ポインター（L/R/C）が無効ノードを指したまま残り、
> Excel がディレクトリを辿る際に破損（「読み取れないコンテンツ」「シートをコピーできません」）を起こす。
> 特にモジュール増減でストリーム構成が変わると、署名が辿り順の途中に来て顕在化する。

---

## CFB ストリームの追加（新規モジュール）

`CFB.utils.cfb_add` を使う。cfb.js は CFB ディレクトリの管理に赤黒木を使っており、
`cfb_add` は内部で `cfb_gc` を呼んで木を再構築する。

```typescript
CFB.utils.cfb_add(cfb, `/VBA/${streamName}`, compressedContent);
```

### 新規クラスモジュールの内部属性

VBE の「ファイルのエクスポート」で得られる通常の `.cls` には、ブック内部で
Excel が保存するクラスモジュールに含まれる属性の一部がない場合がある。新規クラスを
source-only 形式で追加した後に VBE で保存できるよう、`import` は不足している場合だけ
次の属性を補う。

```vb
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
```

既に source に同名の属性がある場合は、その値を変更しない。改行形式も source のまま保持する。

#### `VB_Base` の扱いと根拠の範囲

`0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}` は UUID を都度生成する値ではない。また、
MS-OVBA がこの値を固定値として規定しているわけでもない。通常クラス互換のための
**実測済み既定値**としてのみ扱う。根拠は次のとおり。

1. 手作業で作成し Excel で保存できた `book.xlsm` の通常クラス2件がこの値を持つ。
2. 同じ再現条件で、属性を欠く新規クラスは VBE 保存に失敗し、上記3属性を補うと成功した。
3. 独立した VBA 抽出例にも、通常クラスで同じ値が現れる。

この根拠は Excel の全バージョン・すべての Office ホストアプリに一般化できるものではない。
今後は別バージョンおよび別ホストで検証を追加する。

`VB_Base` 単独では UserForm の判定に使えない。Excel が生成した通常クラスにも含まれるため、
新規追加を拒否するデザイナー／ドキュメントモジュールは `VB_Base` **かつ**
`VB_PredeclaredId = True` を持つ場合として判定する。

## CFB ストリームの削除（モジュール削除時）

type を 0（STGTY_INVALID）に設定する。`cfb.js` の `rebuild_cfb` が書き込み時にスキップする。

```typescript
const entry = CFB.find(cfb, `/VBA/${mod.streamName}`);
if (entry) (entry as any).type = 0;
```

UserForm 削除時はフォームレイアウトストレージも削除する：

```typescript
for (let i = cfbAny.FullPaths.length - 1; i >= 0; i--) {
    const p: string = cfbAny.FullPaths[i];
    if (p.startsWith(`${formName}/`) || p === formName) {
        cfbAny.FileIndex[i].type = 0;
    }
}
```

---

## PROJECT ストリーム（`/PROJECT`）

全モジュールを列挙するテキストファイル（CRLF 改行、コードページエンコード）：

```
ID="{GUID}"
Document=ThisWorkbook/&H00000000
Document=Sheet1/&H00000000
Module=Module1
Class=Class1
Package={GUID}              ← UserForm の Package レコード
BaseClass=UserForm1         ← UserForm の BaseClass レコード（Package= の直後）
Name="VBAProject"
...

[Host Extender Info]
...

[Workspace]
Module1=22, 22, 1172, 464, C
```

**挿入位置**: 新しい `Module=`/`Class=` 行は `Package=+BaseClass=` ブロック（UserForm）より前に挿入する。

**ワークスペースエントリ**: モジュール追加時は `[Workspace]` ヘッダー行の直後に
`モジュール名=0, 0, 2000, 1000, C` を追加する。

**削除時**: `Module=`/`Class=` 行と対応する `[Workspace]` エントリを削除する。
UserForm の場合は `Package=+BaseClass=` のペアも削除する。

---

## PROJECTwm ストリーム（`/PROJECTwm`）

ロケール対応の名前解決に使われるモジュール名マッピングのバイナリストリーム。

形式: `(ASCII名 NUL UTF-16LE名 NUL NUL)* NUL NUL`

- モジュール削除時: 削除したモジュール名のエントリを取り除く
- モジュール追加時: 新しいエントリを末尾に追加する

---

## UserForm のバイナリレイアウト — 現在の制限

UserForm のバイナリレイアウトは `Root Entry/<formName>/` 配下に格納されている：
- `f` — フォームバイナリデータ（コントロール・配置・プロパティ）
- `o` — OLE ストレージ
- `CompObj` — COM クラス識別子
- `VBFrame` — テキスト: フォームのバージョン/キャプションヘッダー

このバイナリデータは `vba-extractor export` で `.cls` ファイルにエクスポートされない。
したがって、`import` による新規 UserForm の追加は現在サポートしていない。

**UserForm / ドキュメントモジュールの新規追加検出**: `Attribute VB_Base = "..."` **かつ**
`Attribute VB_PredeclaredId = True` を含む場合は、ホストまたはデザイナーに紐付く
モジュールとして扱う。`VB_Base` は Excel が生成した通常のクラスモジュールにも存在するため、
それだけで UserForm と判定してはいけない。

`import` は UserForm の追加を明確なエラーで拒否する（Excel で先にフォームを作成してから
コード部分のみ import で更新すること）。

**既存 UserForm のコード更新**: バイナリレイアウトストレージはそのまま保持し、VBA ソースストリームのみ差し替える。

**UserForm の削除**: VBA ソースストリームとフォームレイアウトストレージ（`Root Entry/<名前>/` 以下）を両方削除する。

---

## 検証記録: 新規クラス追加後の VBE 保存エラー（2026-07-19）

入力は `empty_with_macro.xlsm` と `rare-features-check` source directory。完全同期により
`ThisWorkbook`、`Sheet1`、`Module1` を削除する条件で検証した。

| 出力 | source モジュール | Alt+F11 を開き未編集で保存 |
|---|---|---|
| `bas-only.xlsm` | 標準モジュール2件 | 成功 |
| `one-class.xlsm` | 標準モジュール2件 + `EventSrcCheck.cls` | 失敗 |
| `two-class.xlsm` | 標準モジュール2件 + `.cls` 2件 | 失敗 |

したがって再現境界はクラス数ではなく、**新規クラスを1件以上追加すること**だった。
手作業で保存できる `book.xlsm` と DIR / PROJECT / PROJECTwm / クラスソースを比較した結果は以下。

| 項目 | 手作業クラス | import した新規クラス（修正前） |
|---|---|---|
| `MODULETYPE` | クラス（`0x0022`） | 同じ |
| `MODULEPRIVATE` | `0x0028` あり | なし |
| クラスソース属性 | `VB_Base`、`VB_TemplateDerived`、`VB_Customizable` あり | なし |
| `MODULECOOKIE` | Excel 生成値 | `0xFFFF`（仕様どおり） |
| document 関連の `PROJECT` / `PROJECTwm` | 完全同期後に差異あり | `bas-only` でも成功するため直接原因とは扱わない |

`MODULEPRIVATE` は「当該 VBA プロジェクト内でのみ使用可能」を表すレコードであり、通常クラスの
参照ファイルに合わせて新規クラスにも出力する。一方、保存エラーの解消を確認できた決定的な差分は
上記のクラスソース属性である。
