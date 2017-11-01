Attribute VB_Name = "EntryPoints"
'必須　宣言していない変数をコンパイルエラー扱いとする
Option Explicit

'当モジュールには、マクロとして呼び出し可能なPublic Subのみ記載する


Public Sub 複数選択セルにある中身をパースする()

    'javaっぽく　 クラス名と変数名を同じにしないほうがいい　（ 変数名側が優先されて、なぜか型が小文字になってしまう)
    Dim log As LogUtil
    Set log = New LogUtil
    log.DebugMode = True
    
    log.info message:="----------------------------------"
    log.info message:="選択セルの辞書等確認用パース処理開始"
    
    Dim cu As CollectionUtil
    Set cu = New CollectionUtil
    cu.LogObj = log

    '実処理
    Dim result As collection
    Set result = cu.SelectionToCollection(Application.selection)
    Set result = cu.Sanitize(result, "＝", "=")
    Set result = cu.Sanitize(result, "．", ".")
    Set result = cu.Sanitize(result, "，", ",")
    Set result = cu.SplitBy(result, Space$(1))
    Set result = cu.SplitBy(result, ",")
    Set result = cu.SplitBy(result, "=")
    Call log.info("★最終結果★")
    Call log.info(cu.toString(result))

    GoTo LogicEnd
ErrorHandler:
    Call MsgBox(Err.Description, vbCritical + vbOKOnly)

LogicEnd:
    Set result = Nothing
    Set cu = Nothing
    Set log = Nothing

End Sub


' vb6エラーコーディング本にあるパターン
' functionの戻り値をエラー判定で固定するパターン
' 関数の戻りをすべて、引数のbyref 参照渡しにする必要があり、作りが窮屈になる
' 実質的にはC言語の初期に近い
Public Sub エラーの泡パターン1()
    
    Dim ret As String
    ret = ""
    
    Dim errMessage As String
    
    '名前空間を作成することができる（というか、モジュールたくさん作った時に、どこにあるかを探しやすくする
    '(Shift + F2で定義に飛べばいい）
    '同じ名前のpublic関数を作成しても大丈夫なようにする
    errMessage = Fromvb6ErrCoding.C(ret)
    
    Debug.Print errMessage
    Debug.Print ret


End Sub

Public Sub エラーの泡パターン2()
On Error GoTo ErrorHandler
    
    Dim result As Integer
    result = FromQuiita.C

    Debug.Print "終了しました"
    GoTo LogicEnd
ErrorHandler:
    Debug.Print Err.Source & vbCrLf & Err.Description
LogicEnd:
    
End Sub
