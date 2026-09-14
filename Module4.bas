Attribute VB_Name = "Module4"
Option Explicit

Sub 適切な箇所に改ページを入れるVer2(ByVal sheet As Worksheet)
    Dim RPL() As Variant ' RequiredPagebreakList  必須の改ページのリスト
    Const TARGET_COL As Integer = 2          '対象列B
    Dim Before_PB As Integer: Before_PB = 0
    Dim add_up As Integer: add_up = 0
    Dim i As Integer
    Dim h As Integer
    Dim msg As String
    
    Debug.Print "============================================================================================================"
    Debug.Print "============適切な箇所に改ページを入れるVer2================================================================"
    Debug.Print "============================================================================================================"
        
    sheet.Activate 'これ大事
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
    ActiveWindow.Zoom = 35
    'Application.DisplayFullScreen = True

    MsgBox "このマクロは SACLA運転状況集計まとめ.xlsm のシート「" & sheet.Name & "」" & vbCrLf & "の適切な所に貝いれます", Buttons:=vbInformation



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '自動貝が一体何行で入るのかを確認
    Dim AUTO_PB_LINE_CNT As Integer: AUTO_PB_LINE_CNT = 50 ' 初期値は適当に設定した。だいたい50行くらいで自動貝されるようなので。
    Dim HPageBreak As HPageBreak
    sheet.ResetAllPageBreaks ' 全ての貝をクリア
    ' 水平方向の自動貝を確認
    If sheet.HPageBreaks.Count > 0 Then
        Debug.Print "水平方向の自動貝の数: " & sheet.HPageBreaks.Count
        For Each HPageBreak In sheet.HPageBreaks
            If HPageBreak.Type = xlPageBreakAutomatic Then
                Debug.Print "自動貝（水平）位置: 行 " & HPageBreak.Location.Row
                AUTO_PB_LINE_CNT = HPageBreak.Location.Row - 3 ' ちょっと余裕見て-5してる
            End If
            Exit For
        Next HPageBreak
    Else
        Debug.Print "水平方向の自動貝はありません。"
    End If
    Debug.Print "自動貝はこれを越えると入ります。 AUTO_PB_LINE_CNT = " & AUTO_PB_LINE_CNT




    Select Case sheet.Name
        Case "まとめ "
                ReDim RPL(4, 1) ' 配列サイズ設定
                RPL(0, 0) = "(b)運転時間　シフト毎"
                RPL(0, 1) = getLineNum(RPL(0, 0), TARGET_COL, sheet)
                
                RPL(1, 0) = "(b-2)BL3"
                RPL(1, 1) = getLineNum(RPL(1, 0), TARGET_COL, sheet)
                
                
                RPL(2, 0) = "(c)運転条件"
                RPL(2, 1) = getLineNum(RPL(2, 0), TARGET_COL, sheet)
                
                RPL(3, 0) = "(c-2)BL3"
                RPL(3, 1) = getLineNum(RPL(3, 0), TARGET_COL, sheet)
                
                
        Case "Fault集計"
                ReDim RPL(1, 1) ' 配列サイズ設定
                RPL(0, 0) = "SACLA Fault間隔(BL3)"
                RPL(0, 1) = getLineNum(RPL(0, 0), TARGET_COL, sheet)
        Case Else
            Debug.Print "Zzz..."
            Call Fin("このシートでは特にやることないです。" & vbCrLf & "sheet.name:  " & sheet.Name, 3)
    End Select
    
    
    Debug.Print "Last: " & Cells(Rows.Count, TARGET_COL).End(xlUp).Row
    For i = 0 To UBound(RPL, 1) - 1 ' UBound(RPL, 1)の1は行の意味。　2だと列のようだ
        Debug.Print "RPL(" & i & ", 0) : " & RPL(i, 0) & "  この行が必須貝:" & RPL(i, 1)
    Next
    
    Debug.Print "============================================================================================================"
    
    sheet.ResetAllPageBreaks ' 全ての貝をクリア
    sheet.PageSetup.printArea = "" ' 印刷範囲のクリア
    'MsgBox "sheet.UsedRange.Rows(sheet.UsedRange.Rows.Count).Row: " & sheet.UsedRange.Rows(sheet.UsedRange.Rows.Count).Row & "  AUTO_PB_LINE_CNT=" & AUTO_PB_LINE_CNT
    If sheet.UsedRange.Rows(sheet.UsedRange.Rows.Count).Row < AUTO_PB_LINE_CNT Then
        MsgBox "SACLA運転状況集計まとめ.xlsm のシート「" & sheet.Name & "」" & vbCrLf & "行数が少ないので貝ページを入れる必要はありません。", Buttons:=vbInformation
        Exit Sub
    End If

    For i = 1 To sheet.UsedRange.Rows(sheet.UsedRange.Rows.Count).Row
        
        '必須の貝
        For h = 0 To UBound(RPL, 1) - 1 ' UBound(RPL, 1)の1は行の意味。　2だと列のようだ
'            Debug.Print "RPL(" & h & ", *) : 「" & RPL(h, 0) & "」   は      " & RPL(h, 1) & " 行目にあります"
            If i = RPL(h, 1) Then
                Debug.Print "必須の貝 " & i & "       Before_PB = " & Before_PB & "    Value: " & Cells(i, 2).Value
                If i - Before_PB < 10 Then
                    Debug.Print "前の貝とあまりに近すぎるのでパスします。" & i & "       Before_PB = " & Before_PB & "    Value: " & Cells(i, 2).Value
                Else
                    If Not PageBreakInsert(sheet, i, msg) Then
                        MsgBox "改ページを設定できませんでした: " & msg
                    End If
                    Cells(i, TARGET_COL).Activate
                    Before_PB = i
                    If MsgBox("Debug:このセルの上に　必須　貝いれました" & vbCrLf & "次に進むにはYes", vbYesNo + vbQuestion, "確認") = vbNo Then Exit Sub
                    add_up = 0
                End If
            End If
        Next
        
            

            
            
        If (Not IsEmpty(Cells(i, TARGET_COL).Value) And Cells(i, TARGET_COL).MergeCells) Then  '　セルに値が入ってて結合される場合で、
            'Debug.Print "A  行番号: " & i & "   add_up = " & add_up & "    Value: " & Cells(i, 2).Value & "    値が入って、結合されている"
            If (i + Cells(i, TARGET_COL).mergeArea.Rows.Count - Before_PB) > AUTO_PB_LINE_CNT Then ' i+結合行-前貝行が　AUTO_PB_LINE_CNT　を越えたら
                Debug.Print "この行は結合されてて、次はOverしちゃうのでここに貝 " & i & "       Before_PB = " & Before_PB & "    Value: " & Cells(i, 2).Value
                If Not PageBreakInsert(sheet, i, msg) Then
                    MsgBox "改ページを設定できませんでした: " & msg
                End If
                Cells(i, TARGET_COL).Activate
                If MsgBox("Debug:このセルの上に　A　貝いれました" & vbCrLf & "次に進むにはYes", vbYesNo + vbQuestion, "確認") = vbNo Then Exit Sub
                add_up = 0
                Before_PB = i
            End If
        End If

        If (Not Cells(i, TARGET_COL).MergeCells) And add_up > AUTO_PB_LINE_CNT Then  '　セルが結合されてなくて、かつ時期
             Debug.Print "この行は結合されてなくて、かつ時期    行番号: " & i & "   add_up = " & add_up & "    Value: " & Cells(i, 2).Value
            If Not PageBreakInsert(sheet, i, msg) Then
                MsgBox "改ページを設定できませんでした: " & msg
            End If
             Cells(i, TARGET_COL).Activate
             If MsgBox("Debug:このセルの上に　B　貝いれました" & vbCrLf & "次に進むにはYes", vbYesNo + vbQuestion, "確認") = vbNo Then Exit Sub
             add_up = 0
             Before_PB = i
        End If
        
        add_up = add_up + 1
    Next
    
    
    
    
    
    

    Select Case sheet.Name
        Case "まとめ " ' 下の方に、計算用なのかいらない部分があるので、とりあえず列Bの最終行までを
            Dim lastRow As Long
            lastRow = sheet.Cells(Rows.Count, "B").End(xlUp).Row     '列Bの最終行を取得
            lastRow = lastRow + sheet.Cells(lastRow, "B").mergeArea.Rows.Count '列Bの最終行が結合されている場合があるので、結合行を追加
            sheet.PageSetup.printArea = "A1:N" & lastRow ' 印刷範囲をA1:N最終行に設定
        Case "Fault集計"
        Case Else
            Debug.Print "Zzz..."
    End Select


    
    
    
    sheet.DisplayPageBreaks = True
   
'    If MsgBox("プレビュー表示しますか？", vbYesNo + vbQuestion, "確認") = vbYes Then
'        sheet.PrintPreview
'    End If
    Call Fin("終了しました。シート「" & sheet.Name & "」" & vbCrLf & "の適切な所に貝いれました", 1)
    Debug.Print "Fin================================================================================================================="
    
End Sub














'Not USE~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Sub 適切な箇所に改ページを入れる(ByVal sheet As Worksheet)
    
    Dim RPL() As Variant ' RequiredPagebreakList  必須の改ページのリスト
    Dim TARGET_COL As Integer           '対象列
    Dim LINE_CNT_PAGEBREAK As Integer   'デフォルト:50    50行数くらいで改行   それ以上にすると1ページに収まりきらない事があり、エクセルが自動で改ページを挿入してしまう
    Dim MARGIN_PAGEBREAK   As Integer   'デフォルト:20    次の改ページまで20行以上ある場合、改ページセット
    Dim i As Integer
    Dim p As Integer

    Debug.Print "============================================================================================================"
    Debug.Print "============改ページを自動で設定============================================================================"
    Debug.Print "============================================================================================================"
        
    sheet.Activate 'これ大事
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
    ActiveWindow.Zoom = 35
    'Application.DisplayFullScreen = True

    MsgBox "このマクロは SACLA運転状況集計まとめ.xlsm のシート「" & sheet.Name & "」" & vbCrLf & "の適切な所に改ページいれます", Buttons:=vbInformation


    Select Case sheet.Name
        Case "まとめ "
                TARGET_COL = 2  ' 対象列 B
                LINE_CNT_PAGEBREAK = 30   'デフォルト:50    50行数くらいで改行   それ以上にすると1ページに収まりきらない事があり、エクセルが自動で改ページを挿入してしまう
                MARGIN_PAGEBREAK = 20  '　デフォルト:20    次の改ページまで20行以上ある場合、改ページセット
                'Const sheet.name As String = "まとめ "        'なんだよこれ、まとめの後ろに半角スペースが入ってるよ
                ReDim RPL(3, 1) ' 配列サイズ設定
                RPL(0, 0) = "(a)運転時間　期間毎"
                RPL(0, 1) = getLineNum(RPL(0, 0), TARGET_COL, sheet)
                RPL(1, 0) = "(b)運転時間　シフト毎"
                RPL(1, 1) = getLineNum(RPL(1, 0), TARGET_COL, sheet)
                RPL(2, 0) = "(c)運転条件"
                RPL(2, 1) = getLineNum(RPL(2, 0), TARGET_COL, sheet)
        Case "Fault集計"
                TARGET_COL = 2  ' 対象列 B
                LINE_CNT_PAGEBREAK = 25  'デフォルト:50    50行数くらいで改行   それ以上にすると1ページに収まりきらない事があり、エクセルが自動で改ページを挿入してしまう
                MARGIN_PAGEBREAK = 10  '　デフォルト:20    次の改ページまで20行以上ある場合、改ページセット
                'Const sheet.name As String = "Fault集計"
                ReDim RPL(2, 1) ' 配列サイズ設定
                RPL(0, 0) = "SACLA Fault間隔(BL2)"
                RPL(0, 1) = getLineNum(RPL(0, 0), TARGET_COL, sheet)
                RPL(1, 0) = "SACLA Fault間隔(BL3)"
                RPL(1, 1) = getLineNum(RPL(1, 0), TARGET_COL, sheet)
        Case Else
            Debug.Print "Zzz..."
            Call Fin("このシートでは特にやることないです。" & vbCrLf & "sheet.name:  " & sheet.Name, 3)
    End Select
    
    Debug.Print "sheet.Name: " & sheet.Name
    Debug.Print "Last: " & Cells(Rows.Count, TARGET_COL).End(xlUp).Row
    For i = 0 To UBound(RPL, 1) - 1 ' UBound(RPL, 1)の1は行の意味。　2だと列のようだ
        Debug.Print "RPL(i, 0) : " & RPL(i, 0) & "  " & RPL(i, 1)
    Next
    
    Debug.Print "============================================================================================================"

    
    sheet.ResetAllPageBreaks ' 全ての改ページをクリア
    
'    Debug.Print "UBound(RPL, 1) : " & UBound(RPL, 1)
    For i = 0 To UBound(RPL, 1) - 1 ' UBound(RPL, 1)の1は行の意味。　2だと列のようだ
        Debug.Print "Debug: RPL(" & i & ", *) : 「" & RPL(i, 0) & "」   は      " & RPL(i, 1) & " 行目にあります"
        sheet.Rows(RPL(i, 1)).PageBreak = xlPageBreakManual ' 必須の改ページをセット
        If i = (UBound(RPL, 1) - 1) Then ' RPL配列の最後の場合
            Call SetPagebreak(RPL(i, 1), Cells(Rows.Count, TARGET_COL).End(xlUp).Row, TARGET_COL, LINE_CNT_PAGEBREAK, sheet)
        Else
            Call SetPagebreak(RPL(i, 1), RPL(i + 1, 1) - 1, TARGET_COL, LINE_CNT_PAGEBREAK, sheet)
        End If
    Next
    
    
    For p = 1 To sheet.HPageBreaks.Count
        Debug.Print "Page = " & p & "   sheet.HPageBreaks(p).Location.Row = " & sheet.HPageBreaks(p).Location.Row
    Next p
    
    Debug.Print "~~~~~~~~~~~~~~~~~~~~~    とりあえずのトータルの改ページ数：" & sheet.HPageBreaks.Count
    Debug.Print "~~~~~~~~~~~~~~~~~~~~~    必須の改ページ数： " & UBound(RPL, 1)
    
      
    
      
    Debug.Print "============================================================================================================"
    Debug.Print "===改ページの行数と必須かどうかを2次元配列PagebreakListに格納して、それを基に改めて改ページをセットし直す==="
    Debug.Print "============================================================================================================"
    
    Dim PagebreakList() As Integer
    ReDim PagebreakList(sheet.HPageBreaks.Count, 1) '配列サイズ設定
    
    

    
    '改ページの行数と必須かどうかを2次元配列PagebreakListに格納
    Dim totalP As Integer
    totalP = 0
    For p = 1 To sheet.HPageBreaks.Count
        'Debug.Print "Page = " & p & "   sheet.HPageBreaks(p).Location.Row = " & sheet.HPageBreaks(p).Location.Row
        
        If sheet.HPageBreaks(p).Location.Row > MARGIN_PAGEBREAK Then  'セットする改ページが MARGIN_PAGEBREAK 行より下の場合、改ページセット
            PagebreakList(totalP, 0) = sheet.HPageBreaks(p).Location.Row
            PagebreakList(totalP, 1) = 0  'とりあえず、必須でない改ページ:0を入れとく
            
            For i = 0 To UBound(RPL, 1) - 1 ' UBound(RPL, 1)の1は行の意味。　2だと列のようだ
                If sheet.HPageBreaks(p).Location.Row = RPL(i, 1) Then
                    'Debug.Print "必須RPL(i, 0) = " & RPL(i, 0) & "  RPL(i, 1)=" & RPL(i, 1) & " sheet.HPageBreaks(p).Location.Row= " & sheet.HPageBreaks(p).Location.Row
                    PagebreakList(totalP, 1) = 1 '必須の改ページのところにフラグ:1を立てる
                End If
            Next
        
            Debug.Print "* PagebreakList " & totalP & ":    " & PagebreakList(totalP, 0) & "   行目      必須フラグ: " & PagebreakList(totalP, 1)
            totalP = totalP + 1
        End If
            
    Next p
    
    
    
    
    'End 'for DEBUG
    
    
    '一旦改ページを削除
    sheet.ResetAllPageBreaks
    
    
    'PagebreakListを基に改めて改ページをセットし直す
    For i = 0 To UBound(PagebreakList, 1) - 1
        'Debug.Print ">PagebreakList(" & i; ",):   " & PagebreakList(i, 0) & "  " & PagebreakList(i, 1)
    
        If (PagebreakList(i + 1, 0) - PagebreakList(i, 0)) > MARGIN_PAGEBREAK Then   '次の改ページまで MARGIN_PAGEBREAK 行以上ある場合、改ページセット
            sheet.Rows(PagebreakList(i, 0)).PageBreak = xlPageBreakManual
            Debug.Print "Set PagebreakList(" & i & ", 0" & "):   " & PagebreakList(i, 0) & "  Next: " & PagebreakList(i + 1, 0)
            
            
            sheet.Rows(PagebreakList(i, 0)).Select
            MsgBox "Debug   このセルの上に改ページセット", Buttons:=vbInformation


        End If
    
        If (PagebreakList(i, 1)) = 1 Then  '必須の改ページの場合、改ページセット
            sheet.Rows(PagebreakList(i, 0)).PageBreak = xlPageBreakManual
            Debug.Print "Set Required PagebreakList(" & i & ", 0" & "):   " & PagebreakList(i, 0) & "  Next: " & PagebreakList(i + 1, 0)
       
            sheet.Rows(PagebreakList(i, 0)).Select
            MsgBox "Debug   このセルの上に　必須　改ページセット", Buttons:=vbInformation

        End If
        
    Next
    
    
    
    
    sheet.DisplayPageBreaks = True
   
    If MsgBox("終了。プレビュー表示しますか？", vbYesNo + vbQuestion, "確認") = vbYes Then
        sheet.PrintPreview
    End If

    MsgBox "終了しました。シート「" & sheet.Name & "」" & vbCrLf & "の適切な所に改ページいれました", Buttons:=vbInformation
    Call Fin("ブックが開けませんでした。パスの異なる同じ名前のブックが既に開かれてる可能性があります。", 3)
    Debug.Print "Fin================================================================================================================="
    
    

End Sub

















'Not USE    改行は***行ごとくらいでいれる==============================================================================================================================
Function SetPagebreak(ByVal startLine As Integer, ByVal endLine As Integer, ByVal TARGET_COL As Integer, ByVal LINE_CNT_PAGEBREAK As Integer, ByVal sheetName As Worksheet)
    'Debug.Print "SetPagebreak~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~startLine = "; startLine & "       endLine: " & endLine
    If (endLine - startLine) <= 0 Then
        MsgBox "引数異常"
        End
    End If

    Dim line_cnt As Integer
    line_cnt = 0

    Dim i As Integer
    For i = startLine To endLine
        'Debug.Print "行番号: " & i & "   line_cnt = " & line_cnt & "    Value: " & Cells(i, 2).Value
        
        If line_cnt > LINE_CNT_PAGEBREAK Then
        
            If IsEmpty(Cells(i, TARGET_COL).Value) And Cells(i, TARGET_COL).MergeCells Then
                'Debug.Print "セルが空っぽで、結合されている、、"
            Else
            
                If Cells(i, TARGET_COL).MergeCells Then  '  B列が結合されている
                    Debug.Print "行番号: " & i & "   line_cnt = " & line_cnt & "    Value: " & Cells(i, 2).Value & "    値が入って、結合されている     結合行：" & Cells(i, TARGET_COL).mergeArea.Rows.Count & " の下に貝"
                    
'                    Cells(i, TARGET_COL).Activate
'                    MsgBox "A   Debug:このセルの上に改ページいれます"
                    
                    '下のコードで失敗する場合、ページレイアウトから「印刷範囲のクリア」をして再度マクロを実行するとなぜかＯＫ
                    sheetName.Rows(i + Cells(i, 2).mergeArea.Rows.Count).PageBreak = xlPageBreakManual  '  2より大きくないとCells(i, 2).MergeArea.Rows.Countで、エラー　　　2だと2行目の上に引かれる
                Else
                    Debug.Print "行番号: " & i & "   line_cnt = " & line_cnt & "    Value: " & Cells(i, 2).Value & "    値が入ってなくて、結合されてない　貝: i = " & i
                    
'                    Cells(i, TARGET_COL).Activate
'                    MsgBox "B   Debug:このセルの上に改ページいれます"
                    
                            
                    sheetName.Rows(i).PageBreak = xlPageBreakManual
                    
                End If
                line_cnt = 0
            
            End If

        End If
        
        line_cnt = line_cnt + 1
    Next

End Function















'=======================================

Sub PrintPDF(ByVal sheet As Worksheet)
    sheet.DisplayPageBreaks = True
   
'    If MsgBox("終了。プレビュー表示しますか？", vbYesNo + vbQuestion, "確認") = vbYes Then
'        sheet.PrintPreview
'    End If
    Call Fin("終了しました。シート「" & sheet.Name & "」" & vbCrLf & "の適切な所に貝いれました", 1)
    Debug.Print "Fin================================================================================================================="
    
End Sub




'=======================================
'改ページをいれられない原因が複数あるとのことだけど、これらの原因を考慮して改ページを入れる関数を作ってほしい。具体的にはFunction PageBreakInsert(ByVal sheet As Worksheet, Line As Int)のようにシート名と挿入したい改ページ行を引数にとって、成否を返り値に欲しい
'=======================================
Function PageBreakInsert(ByVal sheet As Worksheet, _
                         ByVal Line As Long, _
                         ByRef Reason As String) As Boolean
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = sheet

    '--- 1. 行番号チェック
    If Line <= 1 Or Line > ws.Rows.Count Then
        Reason = "行番号が不正（1行目または範囲外）"
        PageBreakInsert = False
        Exit Function
    End If

    '--- 2. 行が非表示
    If ws.Rows(Line).Hidden = True Then
        Reason = "対象行が非表示のため改ページ不可"
        PageBreakInsert = False
        Exit Function
    End If

    '--- 3. 印刷タイトル（繰り返し行）が設定されている
    Dim titleRows As String
    titleRows = ws.PageSetup.PrintTitleRows
    If titleRows <> "" Then
        ws.PageSetup.PrintTitleRows = ""
    End If

    '--- 4. 改ページをリセット（自動改ページの干渉回避）
    'ws.ResetAllPageBreaks

    '--- 5. 列方向の改ページを解除
    'Dim col As Long
    'For col = 1 To ws.UsedRange.Columns.Count
    '    ws.Columns(col).PageBreak = xlPageBreakNone
    'Next col

    '--- 6. 改ページ設定
    ws.Rows(Line).PageBreak = xlPageBreakManual

    PageBreakInsert = True
    Reason = ""
    Exit Function

ErrHandler:
    Reason = "Excel が改ページを拒否（詳細: " & Err.Description & "）"
    PageBreakInsert = False
End Function
