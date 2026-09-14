Attribute VB_Name = "Module7"
Option Explicit

Sub ユニットBL2とBL3をマージ()
    On Error GoTo ErrorHandler ' エラーハンドリングを設定
    Debug.Print "\n\n\n_______Start  @Sub ユニットBL2とBL3をマージ()\n\n\n"

    Dim i As Integer
    Dim MaxRow As Integer
    Dim MaxRow_of_TargetUnit As Integer
    Dim TargetUnit As String
    Dim TargetSheet As String
    Dim Sonzai_flg_BL2 As Boolean: Sonzai_flg_BL2 = False
    Dim Sonzai_flg_BL3 As Boolean: Sonzai_flg_BL3 = False
    Dim Sonzai_flg_Merged As Boolean: Sonzai_flg_Merged = False
    Dim wb As Workbook
    Dim BNAME_SHUKEI As String
    
    If Not CheckServerAccess_FSO(BNAME_MATOME) Then
        Exit Sub
    End If
    ' wb_MATOMEを開く
    Dim wb_MATOME As Workbook    ' ちゃんと宣言しないと、関数SheetExistsの引数が異なると怒られる
    Set wb_MATOME = OpenBook(BNAME_MATOME, False) ' フルパスを指定
    If wb_MATOME Is Nothing Then Call Fin("ブックが開けませんでした。パスの異なる同じ名前のブックが既に開かれてる可能性があります。", 3)
    wb_MATOME.Activate
    If ActiveWorkbook.Name <> wb_MATOME.Name Then
        Call Fin("現在アクティブなブック名が異常です。終了します。" & vbCrLf & "ActiveWorkbook.Name:  " & ActiveWorkbook.Name & vbCrLf & "BNAME_SHUKEI:  " & BNAME_SHUKEI, 3)
    End If
    
            
    wb_MATOME.Windows(1).WindowState = xlMaximized
    wb_MATOME.Worksheets("Fault集計").Activate 'これ大事
    wb_MATOME.Worksheets("Fault集計").Cells(1, 1).Select ' 選択範囲が残ってるの気持ち悪いのでとりあえず
    
    For i = 1 To sheetS.Count
        Debug.Print sheetS(i).Name
        If sheetS(i).Name = "まとめ " Then 'シート「まとめ 」の次のシートが対象となるユニット
            TargetSheet = sheetS(i + 1).Name
            Debug.Print "Hit-------" & TargetSheet
            Exit For
        End If
    Next
    Debug.Print "TargetSheet = " & TargetSheet
    
    

    
    '「ユニット(BL*)」というパターン表現の場合次ぎすすむ ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Dim Hakken As Boolean
    Dim regex As Object
    Dim testString As String
    Dim matches As Object
    Dim match As Object
    Dim pattern As String
    Hakken = False

    ' 検索したい文字列
    testString = TargetSheet

    ' 正規表現オブジェクトの作成
    Set regex = CreateObject("VBScript.RegExp")

    
'    pattern = "\d+-\d+\(BL\d\)"  ' 正規表現パターンの設定（部分一致を含む）
    pattern = "^\d+-\d+\(BL\d\)$"  ' 正規表現パターンの設定（完全一致）

    ' 正規表現のプロパティを設定
    With regex
        .Global = True         ' すべての一致を検索
        .IgnoreCase = True     ' 大文字と小文字を区別しない
        .pattern = pattern     ' 検索パターンを指定
    End With

    ' 文字列内の一致を検索
    Set matches = regex.Execute(testString)

    ' 一致した結果を表示
    For Each match In matches
        Debug.Print "見つかったパターン: " & match.Value
        Hakken = True
    Next match

    ' オブジェクトのクリーンアップ
    Set regex = Nothing
    Set matches = Nothing
    
    If Hakken = False Then
        Call Fin("まとめシートの次のシート名が、" & vbCrLf & "「" & TargetSheet & "」" & vbCrLf & "です。" & vbCrLf & "「ユニット(BL*)」というパターン表現ではありません。" & vbCrLf & "ユニット(BL2)とユニット(BL3)とう名前のシートを結合したいのでこれではできません。", 3)
    End If
    
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        
    
    
    
    Dim tmp As Variant
    tmp = Split(TargetSheet, "(")
    TargetUnit = tmp(0)
    Debug.Print "TargetUnit=" & TargetUnit
    

    'シートの存在確認
    Sonzai_flg_BL2 = SheetExists(wb_MATOME, TargetUnit & "(BL2)")
    Sonzai_flg_BL3 = SheetExists(wb_MATOME, TargetUnit & "(BL3)")
    Sonzai_flg_Merged = SheetExists(wb_MATOME, TargetUnit)
    If Sonzai_flg_Merged Then
        Call Fin("既に結合されたシートが存在します。", 3)
    End If
    If Not Sonzai_flg_BL2 Or Not Sonzai_flg_BL3 Then
        Call Fin("ユニット、または、ユニット(BL2) または ユニット(BL3) のシートが出来てません。", 3)
    End If

    If MsgBox("このマクロは「SACLA運転状況集計まとめ.xlsm」の" & vbCrLf & "シート「ユニット(BL2)と(BL3)」を結合します。" & vbCrLf & "結合しようとしているユニットは「" & TargetUnit & "」です。" & vbCrLf & "いいですか？？", vbYesNo + vbQuestion, "確認") = vbNo Then
        Call Fin("「No」が選択されました", 1)
    End If
        
    
    
    
    
    'BL2のシートをコピーしてベースにする
    wb_MATOME.Worksheets(TargetUnit & "(BL2)").Copy After:=wb_MATOME.Worksheets("まとめ ") ' なぜかSCSS+ログノート用PCで実行すると、「名前'~～'は既に存在します」とメッセージが出る場合がある。とりあえず「はい」で進めるしかないので進めると、出来たシートに条件付き書式でなぜか赤印がつく。OFFICEのプロフェッショナルだと発生する。
    ActiveSheet.Name = TargetUnit
    Cells.Select '　コピーしたことを視覚的に分かりやすくするため。なくてもいい
    MsgBox "シート「" & TargetUnit & "(BL2)" & "」をコピーしました。" & vbCrLf & "これを下地にします。", Buttons:=vbInformation
        




    
    '(a)運転時間　期間毎  の部分の処理
    wb_MATOME.Worksheets(TargetUnit & "(BL3)").Activate ' これ大事　これしないと .Selectできない
    wb_MATOME.Worksheets(TargetUnit & "(BL3)").Range("I9:L9").Copy
    wb_MATOME.Worksheets(TargetUnit & "(BL3)").Range("I9:L9").Select
    If MsgBox("選択されてる部分をコピーしました。シート「" & TargetUnit & "」" & vbCrLf & "に張り付けます。" & vbCrLf & "いいです？", vbYesNo + vbQuestion, "確認") = vbYes Then
        wb_MATOME.Worksheets(TargetUnit).Activate
        wb_MATOME.Worksheets(TargetUnit).Range("I9").PasteSpecial Paste:=xlPasteValues
        MsgBox "貼り付けました。" & vbCrLf & "次に進みます。", Buttons:=vbInformation
    End If
    
    
    '(b)運転時間　シフト毎  の部分の処理
    'MsgBox getLineNum("(b-2)BL3", 2, wb_MATOME.Worksheets(TargetUnit & "(BL3)"))
    'MsgBox getLineNum("(c)運転条件", 2, wb_MATOME.Worksheets(TargetUnit & "(BL3)"))
    wb_MATOME.Worksheets(TargetUnit & "(BL3)").Activate ' これ大事　これしないと .Selectできない
    wb_MATOME.Worksheets(TargetUnit & "(BL3)").Rows(getLineNum("(b-2)BL3", 2, wb_MATOME.Worksheets(TargetUnit & "(BL3)")) & ":" & getLineNum("(c)運転条件", 2, wb_MATOME.Worksheets(TargetUnit & "(BL3)")) - 1).Copy
    wb_MATOME.Worksheets(TargetUnit & "(BL3)").Rows(getLineNum("(b-2)BL3", 2, wb_MATOME.Worksheets(TargetUnit & "(BL3)")) & ":" & getLineNum("(c)運転条件", 2, wb_MATOME.Worksheets(TargetUnit & "(BL3)")) - 1).Select
    If MsgBox("選択されてる部分をコピーしました。シート「" & TargetUnit & "」" & vbCrLf & "に張り付けます。" & vbCrLf & "いいです？", vbYesNo + vbQuestion, "確認") = vbYes Then
        wb_MATOME.Worksheets(TargetUnit).Activate
        wb_MATOME.Worksheets(TargetUnit).Cells(getLineNum("(c)運転条件", 2, wb_MATOME.Worksheets(TargetUnit)) - 1, 1).Insert xlDown
        wb_MATOME.Worksheets(TargetUnit).Cells(getLineNum("(b-2)BL3", 2, wb_MATOME.Worksheets(TargetUnit)), 2).Select
        MsgBox "貼り付けました。" & vbCrLf & "次に進みます。", Buttons:=vbInformation
    End If

    '(c)運転条件　  の部分の処理
    wb_MATOME.Worksheets(TargetUnit & "(BL3)").Activate ' これ大事　これしないと .Selectできない
    MaxRow = wb_MATOME.Worksheets(TargetUnit & "(BL3)").UsedRange.Rows(wb_MATOME.Worksheets(TargetUnit & "(BL3)").UsedRange.Rows.Count).Row 'UsedRangeの注意点　罫線なども含んだ使用されている領域
    wb_MATOME.Worksheets(TargetUnit & "(BL3)").Rows(getLineNum("(c-2)BL3", 2, wb_MATOME.Worksheets(TargetUnit & "(BL3)")) & ":" & MaxRow - 1).Copy
    wb_MATOME.Worksheets(TargetUnit & "(BL3)").Rows(getLineNum("(c-2)BL3", 2, wb_MATOME.Worksheets(TargetUnit & "(BL3)")) & ":" & MaxRow - 1).Select
    If MsgBox("選択されてる部分をコピーしました。シート「" & TargetUnit & "」" & vbCrLf & "に張り付けます。" & vbCrLf & "いいです？", vbYesNo + vbQuestion, "確認") = vbYes Then
        wb_MATOME.Worksheets(TargetUnit).Activate
        MaxRow_of_TargetUnit = wb_MATOME.Worksheets(TargetUnit).UsedRange.Rows(wb_MATOME.Worksheets(TargetUnit).UsedRange.Rows.Count).Row
        wb_MATOME.Worksheets(TargetUnit).Cells(MaxRow_of_TargetUnit + 1, 1).Insert xlDown
        wb_MATOME.Worksheets(TargetUnit).Cells(getLineNum("(c-2)BL3", 2, wb_MATOME.Worksheets(TargetUnit)), 2).Select
        MsgBox "貼り付けました。" & vbCrLf & "。", Buttons:=vbInformation
    End If


    'wb_MATOME.Worksheets(TargetUnit).ResetAllPageBreaks ' 全ての改ページをクリア
    wb_MATOME.Worksheets(TargetUnit).PageSetup.printArea = False ' 全ての印刷範囲をクリア

    
    Call Fin("これで終了です。" & vbCrLf & "シート「ユニット(BL*)」は手動で削除して下さい", 1)

    Exit Sub  ' 通常の処理が完了したらエラーハンドラをスキップ
ErrorHandler:
    Call Fin("エラーです。内容は　 " & Err.Description, 3)
    Exit Sub
    
End Sub
