# 見積での拡張設定について
フレームワークの仕様上、明細（Detail）を取り扱う場合は専用の `DetailCalculator` や `Reader` が準備されています。処理を他PGに転用する際のコードイメージは以下のようになります。

### 1. DA層：DBから定義を読み込む（`Initializer.vb` 等）
データアクセス層で専用のReaderをインスタンス化し、対象機能（例：見積処理の明細）のレイアウト設定をDBから取得します。

```vb
' Reader のインスタンスを生成
Dim target As IHANExpandItemEntryInputReader = _
    CommonFactory.CreateInstance(Of IHANExpandItemEntryInputReader)(Me, GetType(HANExpandItemEntryInputReader))

' 対象の入力情報・レイアウト設定情報を取得してテーブルに格納
'Initializer.vb-SetExpandItemInfo-拡張項目入力処理画面表示=Trueの時
target.SetExpandItemEntryInputInfo( _
    IHANExpandItemEntryInputReader.HAN拡張入力処理型.見積処理, _
    IHANExpandItemEntryInputReader.明細区分型.明細, _
    inputCalculateTable)

' 上記テーブル（inputCalculateTable）をUIへ渡す初期化情報等に格納する
' initInfo("プログラム専用.入力設定計算式情報テーブル") = inputCalculateTable
```

### 2. UI層：画面側での振分計算機バインド（`EntryForm.vb` 等）
画面の明細行ごとに定義を解釈するための「Calculator（計算機）」をインスタンス化し、退避情報として保持させます。ポップアップかグリッド内入力かはこのCalculatorが内部判断します。

```vb
' クラス変数等で明細用の Calculator を宣言
Private _inputCalculater伝票 As IHANExpandItemEntryInputDetailCalculator

' 初期化メソッド内等で Calculator のインスタンスを生成
_inputCalculater伝票 = CommonFactory.CreateInstance(Of IHANExpandItemEntryInputDetailCalculator)( _
    Me, GetType(HANExpandItemEntryInputDetailCalculator))

' DA層から引き継いだ定義データをもとに、計算機を初期化
Dim targetReader As IHANExpandItemEntryInputDataReader = CreateDataAccessHANExpandItemEntryInputDataReader()
_inputCalculater伝票.InitializeEntryInput(targetReader, IHANExpandItemEntryInputDataReader.HAN拡張入力処理型.見積処理, ...)

' ※明細行が追加されるタイミング等で、対象行のコントロール情報と共に退避情報オブジェクトへ格納する
' _退避情報.明細部項目(rowIndex, 退避情報型.明細項目型.入力計算情報) = _inputCalculater明細
```

### 3. Controller層：ポップアップ画面の呼び出し（`SubController.vb` 等）
画面上で「拡張」ボタン（ポップアップ呼出）が押された時に、共通コンポーネントである `ExpandItemEntryForm` を起動して画面制御を移します。

```vb
' クラス変数で共通画面のインターフェースを宣言
Private _expandItemEntry自明細 As IExpandItemEntryForm

' コントローラー初期化時にポップアップ画面のインスタンスを生成
_expandItemEntry自明細 = CommonFactory.CreateInstance(Of IExpandItemEntryForm)(Me, GetType(ExpandItemEntryForm))

' ----------------------------------------------------
' ボタン押下時の処理（コマンド受取メソッドなど）
' ----------------------------------------------------
Public Function ExpandItemEntryDisplay(ByRef 拡張項目入力情報 As I拡張項目入力情報型) As Boolean
    
    Dim expandItemEntryForm As IExpandItemEntryForm = _expandItemEntry自明細
    
    ' StartExpandItemEntry を呼ぶことで共通ポップアップ画面が ShowDialog される
    If expandItemEntryForm.StartExpandItemEntry(拡張項目入力情報) = True Then
        
        ' ユーザーがポップアップ上で「OK（確定）」した場合の設定反映処理を記述
        ' 例：入力された値を画面のデータテーブルや実計算処理へ同期させる
        Return True
    End If
    
    Return False
End Function
```

**ポイント**
実際のポップアップ画面ファイル（`ExpandItemEntryForm`）の中身自体は共通ライブラリ側（`OSK.X1.HAN.LIB.ExpandItemEntry.UserInterface.dll`）に隠蔽されているため、このように振分器（Calculator）とポップアップ起動（StartExpandItemEntry）を繋ぐ記述を書くだけで完了します。


- 部品でテーブル判断して画面セットしたりしてるってことはEOに載せる場合、同じように画面設定処理と入力項目処理の設定をEO化して専用のテーブルも設計して、適用させるっていう必要があるってことだね。まあ調査として続ける前提であまり工数使わないようにしておかないと
- 明日はもう少し深めに調査してみる。


## 2026/03/05
- 作業無し。

# 拡張項目（Expand Items）の表示制御ロジック調査

本ドキュメントは、「見積処理」プログラムにおける拡張項目の仕組み（グリッド内に表示するか、ボタン経由でポップアップ表示させるか）の振り分けロジックおよび関連するソースコードの実装箇所をまとめたものです。

---

## 1. 概要と振り分けルールの結論

拡張項目はデータベース側の設定で「この画面でどう見せるか」が登録されており、Data Access（DA）層がそれを読み込んで UI 層（EntryForm）に伝達します。
UI 層では以下の基準で「グリッドへの配置」と「ポップアップでの入力」を振り分けています。

- **グリッドへの配置（インライン入力）**
  `expandItemHeader` などのグリッドにおいて、対象列の `EditType` が `Nothing` や `Empty` **以外**であれば、グリッド上で入力可能な項目として扱われます。
- **ボタン経由のポップアップ表示**
  ヘッダー部や明細部の「拡張」ボタンが押下されると、共通のポップアップ画面（`ExpandItemEntryForm`）が呼び出されます。ポップアップの起動には、常に `I拡張項目入力情報型`（退避情報に格納されているオブジェクト）が引数として使用されます。

---

## 2. 実装の全体フロー

1. **データアクセス（DA）**
   `Initializer.vb` でデータベースから入力設定や計算式情報を読み込み、`initInfo`（初期化情報辞書）に格納します。
2. **UI 初期化・表示判定**
   `EntryForm.vb` 内で、列の属性（`_headerColumnInfoList` など）に基づき、各カラムの表示・非表示（`Hide` の切り替えや `EditType = Nothing` の設定）が行われます。
3. **入力・確定処理**
   - **グリッド入力時**: `Is配置拡張項目入力判定` を経由し、有効なセルの場合は `DecideCalculate拡張` にて入力値を検証・計算します。
   - **ポップアップ入力時（ボタン押下）**: `expandHeader.Click` などのイベントから `拡張項目入力画面起動(...)` が呼ばれ、`SubController` 経由で共有画面を呼び出します。

---

## 3. 具体的なソースコードの該当箇所

### 3.1. 配置判定（EditType を用いた判定）

グリッドに配置された項目に入力できるか（有効か）を判別するロジックです。

- **ファイル**: [A00080/UI/EntryForm.vb](A00080/UI/EntryForm.vb#L31170-L31177)
- **関連メソッド**: `Is配置拡張項目入力判定`

```vb
    ''' <summary>
    ''' 配置拡張項目入力判定
    ''' </summary>
    ''' <param name="gridEdit"></param>
    ''' <param name="index"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function Is配置拡張項目入力判定(ByVal gridEdit As OSK.X1.Foundation.Controls.GridEdit, ByVal index As Integer) As Boolean    'X1変換Tool-20160524

        Select Case gridEdit.Columns(index).EditType
            Case GridEditType.Empty,
                 GridEditType.Nothing
                Return False
        End Select
        Return True

    End Function
```

### 3.2. グリッドの表示・非表示制御（EditType の設定）

ヘッダー拡張等において、特定の条件（例：仕入拡張項目など）でセルの編集を不可（`GridEditType.Nothing`）にする処理です。

- **ファイル**: [A00080/UI/EntryForm.vb](A00080/UI/EntryForm.vb#L11580-L11680)
- **関連メソッド**: `Display_ヘッダー拡張`

```vb
    Private Sub Display_ヘッダー拡張()
        ' (中略)
        For colIndex As Integer = 0 To expandItemHeader.ColumnCount - 1
            Dim colName As String = expandItemHeader.Columns(colIndex).Name
            Dim 拡張種類 As ColumnInfoList.拡張種類型 = _headerColumnInfoList.拡張種類(colName)
            If 拡張種類 <> ColumnInfoList.拡張種類型.仕入伝票 Then
                Continue For
            End If
            ' (中略)
            expandItemHeader.Columns(colIndex).EditType = GridEditType.Nothing ' ←非表示・入力不可に設定している箇所

        Next
        expandItemHeader.Redraw = True

    End Sub
```

### 3.3. ボタン経由のポップアップ呼び出し

「拡張ボタン」を押下した際に呼び出される処理口です。ここでボタンの種類（ヘッダー/明細/消費税）を判別し、退避情報にある拡張項目情報を引き渡しています。

- **ファイル**: [A00080/UI/EntryForm.vb](A00080/UI/EntryForm.vb#L42220-L42340)
- **関連メソッド**: `拡張項目入力画面起動`

```vb
    ''' <summary>
    ''' 拡張入力画面の表示
    ''' </summary>
    ''' <param name="ボタン種類"></param>
    ''' <param name="付箋色"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function 拡張項目入力画面起動(ByVal ボタン種類 As 退避情報型.ボタン種類型, ByRef 付箋色 As System.Drawing.Color, ByRef 付箋イメージ As System.Drawing.Image, ByVal targetControl As Control) As Boolean    'A-20160613-X1-付箋アイコン表示

        Dim 拡張項目定義 As I拡張項目定義型
        Dim 拡張項目入力情報 As I拡張項目入力情報型
        Dim 同時区分 As ISubController.同時型
        Dim fusenIcon As IFusenIcon     'A-20160613-X1-付箋アイコン表示

        Select Case ボタン種類
            Case 退避情報型.ボタン種類型.ヘッダー
                拡張項目入力情報 = CType(_退避情報.ヘッダー部項目(退避情報型.ヘッダー項目型.拡張項目), I拡張項目入力情報型)
                拡張項目定義 = _拡張項目定義_見積伝票
                同時区分 = ISubController.同時型.自
                ' (中略)
        End Select

        If _subControler Is Nothing Then
            _subControler = New SubController
            _subControler.Initialize(Me)
        End If
        If _subControler.ExpandItemEntryDisplay(同時区分, 拡張項目入力情報) = True Then
            If Is拡張ボタン色表示可否(ボタン種類) = True Then
                ' 確定時の付箋色更新など
            End If
            Return True
        End If
        Return False

    End Function
```

### 3.4. ポップアップフォームの実体（SubController）

UI から渡された情報を元に、見積ヘッダー／明細などの状況に応じた `ExpandItemEntryForm`（共通の拡張項目入力画面）のインスタンスを生成して起動します。

- **ファイル**: [A00080/UI/SubController.vb](A00080/UI/SubController.vb#L1-L120)
- **関連メソッド**: `ExpandItemEntryDisplay`

```vb
    Friend Function ExpandItemEntryDisplay( _
                                        ByVal 同時区分 As ISubController.同時型, _
                                        ByRef 拡張項目入力情報 As I拡張項目入力情報型) As Boolean Implements ISubController.ExpandItemEntryDisplay

        Dim expandItemEntryForm As IExpandItemEntryForm
        If 同時区分 = ISubController.同時型.自 Then
            If 拡張項目入力情報.行番号 = 0 AndAlso 拡張項目入力情報.明細区分 = I拡張項目入力情報型.明細区分型.通常 Then
                '自ヘッダー
                If _expandItemEntry自ヘッダー Is Nothing Then
                    _expandItemEntry自ヘッダー = CommonFactory.CreateInstance(Of IExpandItemEntryForm)(Me, GetType(ExpandItemEntryForm))
                    _expandItemEntry自ヘッダー.Initialize(Me)
                End If
                expandItemEntryForm = _expandItemEntry自ヘッダー
            ' (中略)
            End If
        End If

        If expandItemEntryForm.StartExpandItemEntry(拡張項目入力情報) = True Then
            Return True
        End If

        Return False

    End Function
```

### 3.5. DA側からの設定読み込み（起点）

設定情報や計算式情報はプログラム実行時に読み込まれ、辞書形式（`initInfo`）でやり取りされます。これが拡張項目の動作パラメータ・レイアウトの源泉です。

- **ファイル**: [A00080/DA/Initializer.vb](A00080/DA/Initializer.vb#L616-L656)
- **関連メソッド**: `SetExpandItemEntryInputInfo` などを含む `Initialize` 関連処理

```vb
            '入力設定／計算式情報取得（ヘッダー）
            Dim target As IHANExpandItemEntryInputReader = CommonFactory.CreateInstance(Of IHANExpandItemEntryInputReader)(Me, GetType(HANExpandItemEntryInputReader))      'A-2007/10/30_7
            target.Initialize(Me)
            Dim inputCalculateTable As IDictionary = New Hashtable
            target.SetExpandItemEntryInputInfo(connection, _
                                               inputCalculateTable, _
                                               IHANExpandItemEntryInputReader.HAN拡張入力処理型.見積処理, _
                                               IHANExpandItemEntryInputReader.明細区分型.ヘッダー, _
                                               "0000")
            initInfo.Add("プログラム専用.入力設定計算式情報テーブル", inputCalculateTable)
```

---

## 4. 他プログラムへの転用時のポイント

この仕組みを別のプログラムに転用・展開する際は、以下の点に留意してください。

1. **データベースの拡張定義**
   DA 層で `SetExpandItemEntryInputInfo` および `GetExpandItemInfo` を呼び出して設定情報を `initInfo` に格納します。
2. **コントロールの宣言と関連付け**
   UI クラスの Designer で対象のグリッドやボタン（例：`expandItemHeader`, `expandHeader`）を定義し、設定情報を元に `ColumnInfoList` などで列情報をマッピングします。
3. **イベントとポップアップ連携**
   ボタンのクリックイベントから `I拡張項目入力情報型` を抽出し、`SubController` 等を経由して `ExpandItemEntryForm.StartExpandItemEntry` を呼び出すフローを構成します。
4. **グリッド入出力の検証**
   `Is配置拡張項目入力判定` のように、グリッドに入力可能かどうかの制御（`EditType` ベース等）を利用して計算処理（Calculator）へ渡すタイミングを制御します。

- 拡張入力画面
 ・拡張部品で作成<br>
 ・設定処理が別で必要<Br>
 ・その情報を保持しておくテーブル（画面設定分と入力分）が必要<Br>
・それ用のフレームワークが必要？<br>


### EO
・①Field Builder部品の内設「4.1 起動時パラメータ」に
「初期フォーカス項目」がありますが、見積処理の内設を確認したところ、
　引渡パラメータに該当項目が存在しないように見受けられました。
　本項目は必須で渡す必要があるのか、省略可能で、見積処理サンプルでは
　渡していないのかご教示いただけますでしょうか。<br>
→入力チェックを呼出元で実施した時に画面起動時の初期フォーカスを設定する用<br>
②Field Builder部品の内設「4.1 起動時パラメータ」に
　「処理名」が追加されていますが、見積処理の内設では、
　引数の項目名が「起動処理」となっております。
　どちらが正しい記載となりますでしょうか。<br.>
→実装を正とし、処理名が正しい<br.>
③Field Builder部品への引渡パラメータとして
「起動処理」が追加されておりますが、本項目は必須となりますでしょうか。
原価管理マスターなど、複数処理で参照しない種別の場合は
　省略可能かどうかご教示いただけますでしょうか。<br>
→「処理名」となりますが、必須入力ではございません。XML側で処理別設定がある場合に使用するものになるのでその際に設定していれば使えるものとなる。


## 2026/03/05
- EOの質問もらっていたのでその回答と設計書の修正作業を行う。
- どんなに時間がかかっても午前中で終わる作業かな。

- 拡張の作業続きまとめ
全体の流れ（上から下へ）

データ層（起動処理）

Initializer が DB から拡張項目定義（例：HAN10Z030 系）と配置情報（HAN10Z033GAMENSET）を読み、Me.InitInfo にキー（例："HAN拡張項目.見積伝票", "HAN拡張配置情報"）として格納する。
UI 初期化（EntryForm 側）

InitializeExpandItemPostInfo で空の 配置情報型 インスタンスを生成し、HANExpandItemEntryPostInfoGetter.GetExpandItemEntryPostInfo(Me.InitInfo, 配置情報) を呼ぶ（ByRef）。これにより Me.InitInfo の "HAN拡張配置情報" を基に 配置情報型 がそのまま埋められる。<br>
Set配置順情報() が配列/ハッシュ化して、ヘッダー用・明細用の _配置順情報_ヘッダー／_配置順情報_明細（内部名 → 入力順番号）を作る。
SetHeaderGrid の内部処理（主要3段階）

SetUpHANControlInfo(): 基本コントロール情報の初期化（非拡張項目の準備）。
SetUpExpandItemInfo()（準備段階）:
HANExpandItemInfoGetter 系で「4つの拡張項目定義」（見積伝票・見積明細・仕入伝票・仕入明細）を読み出す（桁数、属性、参照定義など）。
HANExpandItemEntryPostInfoGetter で HAN10Z033GAMENSET を読み、配置（どの内部名をどの順で表示するか）を階層化して 配置情報型 に格納する。
上記の結果を ColumnInfoList に格納する（主に3つの Hashtable：拡張項目属性、拡張項目桁、参照定義コード と 配置情報）。
→ この段階で「表示に必要なメタ情報」は揃っているが、グリッドの列リスト（_itemInfoList）はまだ未構築。
SetUpHeaderItemInfo(): _配置順情報_ヘッダー を参照してヘッダー側の列順序や基本列を ColumnInfoList／_itemInfoList に反映し始める。
明細側の組立（SetUpDetailItemInfoLine → SetUpItemInfoExpandItemLine）

この段階で SetAddExpandItemInfo* 系メソッドが順次呼ばれる（例：SetAddExpandItemInfo商品コード_X1、SetAddExpandItemInfo商品名_X1、SetAddExpandItemInfo商品名２_X1 等）。
各 SetAddExpandItemInfo* は ColumnInfoList にある「属性・幅・参照コード・配置順」を参照し、AddItemInfo() を使って具体的な列定義（セル型、表示名、タブ順、列幅 など）を _itemInfoList に追加する。
つまり SetUpExpandItemInfo = データ準備（メタ情報集約）、SetAddExpandItemInfo* = 具体的列の組立（メタ → 実体）。
GridHelper 組み立てと表示

_itemInfoList が完成すると GridHelper のコンストラクタへ渡され、SetupGridHelper() で EditManager 等と結びつけられる。
InitializeGrid_Header()/InitializeGridDetail で ColumnInfoList／_配置順情報 を使って各列の DisplayName・Alignment・RelativeColumn（表示幅比）・TabIndex 等を設定する。
最後に SetDataTable() 等でデータをバインドし、UI に反映される。
重要な実装上のポイント（短く）

配置情報は HAN10Z033GAMENSET に完全に依存する（内部名の生成規則：先頭2文字で種別 DU/MU/DS/MS、末尾2桁で番号）。
GetExpandItemEntryPostInfo は ByRef で渡された 配置情報型 を直接更新するため、呼び出し元では戻り値ではなく渡したオブジェクトが埋まる。
準備フェーズが正しく動作していないと、SetAddExpandItemInfo* が正しい列を生成できない（順序・幅・型が欠ける）。
配置情報が存在しない場合はデフォルト処理にフォールバックする実装が各所にある（コード上で判定している）。


```text
明細区分：HANZ033002
配置行No：HANZ033003
明細行内配置：HANZ033004
配置項目売上：HANZ033005
配置項目仕入：HANZ033006 
内部名：HANZ033005とHANZ033006をベースに作成（DU/MU)
表示桁：HANZ033007
参照名称表示桁：HNZ033008
入力順区分：HANZ033009
入力順N0:HANZ033010

SetExpandItemEntryPostInfo：個別設定情報に項目配置情報を追加する

拡張項目入力設定テーブル：Z034テーブルの取得結果
拡張項目計算情報テーブル：Z035テーブルの取得結果


個別設定情報に、入力設定計算式情報テーブルのキーを入れ、そこにZ034 ,Z035テーブルの情報を追加する
_拡張項目定義_見積伝票のデータの入り方。
→GetExpandItemInfoで個別設定情報にDAのInitializeで得た拡張項目の情報を返すようにしている
ついでに、使用区分（拡張項目使ってるかどうか）も返している

InitializeExpandItemPostInfoで配置情報の初期化をして情報をセットする
それと同時に、_配置情報使用区分のTrueFalseの情報もセットされる
_配置使用可否_ヘッダー、_配置使用回避_明細でも同様にTrueFalseが設定される。

```
