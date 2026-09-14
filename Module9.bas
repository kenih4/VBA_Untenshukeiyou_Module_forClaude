Attribute VB_Name = "Module9"
Option Explicit

Sub ユニットBLの結果をシートまとめに張り付ける()
    On Error GoTo ErrorHandler
            
    'MsgBox "マクロの内容" & vbCrLf & "「シート「ユニット」の項目aとシート「ユニット(BL2)と(BL3)」の項目b、cをシート「まとめ」に貼り付け」" & vbCrLf & "です。", Buttons:=vbInformation
    
    Dim i As Integer
    Dim TargetUnit As String
    Dim TargetSheet As String
    Dim Sonzai_flg_BL2 As Boolean: Sonzai_flg_BL2 = False
    Dim Sonzai_flg_BL3 As Boolean: Sonzai_flg_BL3 = False
    Dim Sonzai_flg_Merged As Boolean: Sonzai_flg_Merged = False
    Dim Category As String
    Dim BNAME_SHUKEI As String
    Dim result As Boolean
    
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
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
    ActiveWindow.Zoom = 60
    'Application.DisplayFullScreen = True
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
'    pattern = "^\d+-\d+\(BL\d\)$"  ' 正規表現パターンの設定（完全一致）
    pattern = "^\d+-\d+"  ' 正規表現パターンの設定（完全一致）

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
        Call Fin("まとめシートの次のシート名が、" & vbCrLf & "「" & TargetSheet & "」" & vbCrLf & "です。" & vbCrLf & "「ユニット」というパターン表現ではありません。", 3)
    End If
    
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    
    
    
    
    
    
    Dim tmp As Variant
    tmp = Split(TargetSheet, "(")
    Debug.Print "UBound(tmp) = " & UBound(tmp)

    TargetUnit = tmp(0)
    Debug.Print "TargetUnit=" & TargetUnit
    
    
    
    
    'シートの存在確認
    Sonzai_flg_BL2 = SheetExists(wb_MATOME, TargetUnit & "(BL2)")
    Sonzai_flg_BL3 = SheetExists(wb_MATOME, TargetUnit & "(BL3)")
    Sonzai_flg_Merged = SheetExists(wb_MATOME, TargetUnit)
    If Not Sonzai_flg_Merged Or Not Sonzai_flg_BL2 Or Not Sonzai_flg_BL3 Then
        Call Fin("ユニット、または、ユニット(BL2) または ユニット(BL3) のシートが出来てません。", 3)
    End If
    
    
    If MsgBox("このマクロは「シート「ユニット」の項目aとシート「ユニット(BL2)と(BL3)」の項目b、cをシート「まとめ」に貼り付けします。" & vbCrLf & "対処ユニットは「" & TargetUnit & "」です。" & vbCrLf & "いいですか？？", vbYesNo + vbQuestion, "確認") = vbNo Then
        Call Fin("「No」が選択されました", 1)
    End If

    
    
    
    
    
    '(a)運転時間　期間毎  の部分の処理
    Category = "(a)運転時間　期間毎"
    result = Check_Unit_and_Copy(Category, 2, wb_MATOME.Worksheets(TargetUnit), TargetUnit)
    If result = False Then Call Fin("関数[Check_Unit_and_Copy]が失敗しました", 3)
    If MsgBox("選択されてる部分をコピーしました。" & vbCrLf & "次はシート「まとめ」に張り付けです。" & vbCrLf & "進みますか？", vbYesNo + vbQuestion, "確認") = vbNo Then Call Fin("「No」が選択されました", 1)
    result = Find_targetcell_and_paste(Category, 2, wb_MATOME.Worksheets("まとめ "))
    
    
    
    
    Debug.Print "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
    '(b)運転時間　シフト毎  の部分の処理
    Category = "(b-1)BL2"
    result = Check_Unit_and_Copy(Category, 2, wb_MATOME.Worksheets(TargetUnit & "(BL2)"), TargetUnit)
    If result = False Then Call Fin("関数[Check_Unit_and_Copy]が失敗しました", 3)
    If MsgBox("選択されてる部分をコピーしました。" & vbCrLf & "次はシート「まとめ」に張り付けです。" & vbCrLf & "進みますか？", vbYesNo + vbQuestion, "確認") = vbNo Then Call Fin("「No」が選択されました", 1)
    result = Find_targetcell_and_paste(Category, 2, wb_MATOME.Worksheets("まとめ "))
    
    '(b)運転時間　シフト毎  の部分の処理
    Category = "(b-2)BL3"
    result = Check_Unit_and_Copy(Category, 2, wb_MATOME.Worksheets(TargetUnit & "(BL3)"), TargetUnit)
    If result = False Then Call Fin("関数[Check_Unit_and_Copy]が失敗しました", 3)
    If MsgBox("選択されてる部分をコピーしました。" & vbCrLf & "次はシート「まとめ」に張り付けです。" & vbCrLf & "進みますか？", vbYesNo + vbQuestion, "確認") = vbNo Then Call Fin("「No」が選択されました", 1)
    result = Find_targetcell_and_paste(Category, 2, wb_MATOME.Worksheets("まとめ "))



    Debug.Print "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
    '(c)運転条件　  の部分の処理
    Category = "(c-1)BL2"
    result = Check_Unit_and_Copy(Category, 2, wb_MATOME.Worksheets(TargetUnit & "(BL2)"), TargetUnit)
    If result = False Then Call Fin("関数[Check_Unit_and_Copy]が失敗しました", 3)
    If MsgBox("選択されてる部分をコピーしました。" & vbCrLf & "次はシート「まとめ」に張り付けです。" & vbCrLf & "進みますか？", vbYesNo + vbQuestion, "確認") = vbNo Then Call Fin("「No」が選択されました", 1)
    result = Find_targetcell_and_paste(Category, 2, wb_MATOME.Worksheets("まとめ "))
 
    '(c)運転条件　  の部分の処理
    Category = "(c-2)BL3"
    result = Check_Unit_and_Copy(Category, 2, wb_MATOME.Worksheets(TargetUnit & "(BL3)"), TargetUnit)
    If result = False Then Call Fin("関数[Check_Unit_and_Copy]が失敗しました", 3)
    If MsgBox("選択されてる部分をコピーしました。" & vbCrLf & "次はシート「まとめ」に張り付けです。" & vbCrLf & "進みますか？", vbYesNo + vbQuestion, "確認") = vbNo Then Call Fin("「No」が選択されました", 1)
    result = Find_targetcell_and_paste(Category, 2, wb_MATOME.Worksheets("まとめ "))
 
 
 
 


    'wb_MATOME.Worksheets(TargetUnit).ResetAllPageBreaks ' 全ての改ページをクリア
    wb_MATOME.Worksheets(TargetUnit).PageSetup.printArea = False ' 全ての印刷範囲をクリア

    
    Call Fin("これで終了です。", 1)
    Exit Sub  ' 通常の処理が完了したらエラーハンドラをスキップ
ErrorHandler:
    Call Fin("エラーです。内容は　 " & Err.Description, 3)
    Exit Sub
    
End Sub
















'シートsheetのCategoryのセルに正しいユニットが入ってるか確認して、コピー==============================================================================================================================
Function Check_Unit_and_Copy(ByVal Category As String, ByVal TARGET_COL As Integer, ByVal sheet As Worksheet, ByVal TargetUnit As String) As Boolean
    
    Check_Unit_and_Copy = False
    sheet.Activate
    ActiveWindow.Zoom = 60
    Dim r As Integer: r = 2 ' 「Category 」行から「ユニット名」行までの行数
    
    Debug.Print "OK0:  " & sheet.Cells(getLineNum(Category, TARGET_COL, sheet) + 0, TARGET_COL)  ' 「Category 」
    Debug.Print "OK1:  " & sheet.Cells(getLineNum(Category, TARGET_COL, sheet) + 1, TARGET_COL) ' 「ユニット」
    Debug.Print "OK2:  " & sheet.Cells(getLineNum(Category, TARGET_COL, sheet) + 2, TARGET_COL) ' 上の行が結合されてる場合は空、　そうでない場合は「ユニット名」の筈
    Debug.Print "OK3:  " & sheet.Cells(getLineNum(Category, TARGET_COL, sheet) + 3, TARGET_COL)
    
    If sheet.Cells(getLineNum(Category, TARGET_COL, sheet) + 1, TARGET_COL).MergeCells Then  '  B列　セルが結合されている場合
       r = r + sheet.Cells(getLineNum(Category, TARGET_COL, sheet) + 1, TARGET_COL).mergeArea.Rows.Count - 1
    End If

    If sheet.Cells(getLineNum(Category, TARGET_COL, sheet) + r, TARGET_COL).Value <> TargetUnit Then
        MsgBox "シート「" & sheet.Name & "」の" & Category & "のユニットが一致しません。　終了します。" & vbCrLf & " TargetUnit　= " & TargetUnit & vbCrLf & "セル：" & sheet.Cells(getLineNum(Category, TARGET_COL, sheet) + 3, TARGET_COL).Value, Buttons:=vbCritical
    Else
        Debug.Print "OK:  r = " & r & "     Category = "; Category & "    TargetUnit = " & TargetUnit & "    セル= " & sheet.Cells(getLineNum(Category, TARGET_COL, sheet) + r, TARGET_COL).Value
        sheet.Rows(getLineNum(Category, TARGET_COL, sheet) + r & ":" & getLineNum(Category, TARGET_COL, sheet) + r + sheet.Cells(getLineNum(Category, TARGET_COL, sheet) + r, TARGET_COL).mergeArea.Rows.Count - 1).Copy
        sheet.Rows(getLineNum(Category, TARGET_COL, sheet) + r & ":" & getLineNum(Category, TARGET_COL, sheet) + r + sheet.Cells(getLineNum(Category, TARGET_COL, sheet) + r, TARGET_COL).mergeArea.Rows.Count - 1).Select
        Check_Unit_and_Copy = True
    End If
        
End Function




'==============================================================================================================================
Function Find_targetcell_and_paste(ByVal Category As String, ByVal TARGET_COL As Integer, ByVal sheet As Worksheet) As Boolean
    Dim i As Integer
    
    Find_targetcell_and_paste = False
        
        sheet.Activate
        ActiveWindow.Zoom = 60
                
        Debug.Print "sheet.UsedRange.Rows.Count: " & sheet.UsedRange.Rows.Count & "     sheet.UsedRange.Rows(sheet.UsedRange.Rows.Count).Row: " & sheet.UsedRange.Rows(sheet.UsedRange.Rows.Count).Row
        
        'For i = getLineNum(Category, TARGET_COL, sheet) To sheet.Cells(Rows.Count, TARGET_COL).End(xlUp).Row
        For i = getLineNum(Category, TARGET_COL, sheet) To sheet.UsedRange.Rows(sheet.UsedRange.Rows.Count).Row
            'Debug.Print "行番号: " & i & "    Value: " & sheet.Cells(i, TARGET_COL).Value & "      Cells(i, TARGET_COL).MergeArea.Rows.Count = " & Cells(i, 2).MergeArea.Rows.Count
            If sheet.Cells(i, TARGET_COL).Value = "" Then '
                Debug.Print "空なので、ここに貼り付けます！！！！　行番号: " & i & "    Value: " & sheet.Cells(i, TARGET_COL).Value
                sheet.Cells(i, 1).Select
                If MsgBox("ここに値を貼り付けていいですか？", vbYesNo + vbQuestion, "確認") = vbYes Then
                    sheet.Cells(i, 1).Insert xlDown
                    If MsgBox("貼り付けましたがOKですか？？" & vbCrLf & "次に進むにはYes", vbYesNo + vbQuestion, "確認") = vbNo Then Exit Function
                Else
                    Exit Function
                End If
                Exit For
            End If
            If Cells(i, TARGET_COL).MergeCells Then  '  B列　セルが結合されている場合、iに結合されてる分だけ足して次のループへ
                i = i + Cells(i, TARGET_COL).mergeArea.Rows.Count - 1
            End If
        Next
        
        Find_targetcell_and_paste = True
        
End Function




'==============================================================================================================================
Sub ログノートをHTML出力と調整時間がログノートに記載されてるか確認_ユニット月(Nen As Integer, Tsuki As Integer)
    Dim Command As String
    Dim LogNOTE_from As String
    Dim LogNOTE_to As String
    Dim result As Boolean
    LogNOTE_from = Nen & "_" & Format(Tsuki, "00") & ".xlsm"
    LogNOTE_to = Nen & "_" & Format(Tsuki, "00") & "_SACLA.xlsm"
'    MsgBox TARGET_PATH & "\" & Nen & "\" & Tsuki & "\" & LogNOTE
    MsgBox LogNOTE_from
    MsgBox LogNOTE_to
    result = CopyFileSafely(TARGET_PATH & "\" & Nen & "\" & Format(Tsuki, "00") & "\" & LogNOTE_from, DIST_PATH & "\" & LogNOTE_to)
    If Not result Then
        MsgBox "コピー失敗…　終了します。", vbCritical
    Else
        Command = "cd /c/Users/kenic/Documents/operation_log_NEW" & ";" & _
                   "./excelgrep_by_XMLparse.sh SACLA/" & LogNOTE_to & " '$|引渡' '$|引き渡' '$|波長変更依頼' '$|ユニット' '$|利用終了' '$|運転終了'" & ";" & _
                   "read -p '処理が完了しました。Enterキーを押すと終了します...'"
        Debug.Print "コマンドは以下です。"
        Debug.Print Command
        ExecuteGitBashCommand Command
    End If

End Sub

