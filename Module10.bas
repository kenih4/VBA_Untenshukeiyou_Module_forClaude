Attribute VB_Name = "Module10"
Option Explicit

Sub Fault集計m(BL As Integer)

    '/追加部分----------------------------
    On Error GoTo ErrorHandler
    Dim i As Integer
    Dim targetline As Integer
    Dim BNAME_SHUKEI As String
    Dim SNAME_FAULT As String
    Dim beginL As Integer
    Dim EndL As Integer
    MsgBox "マクロ「Fault集計m」を実行します。" & vbCrLf & "このマクロは、" & vbCrLf & "SACLA運転状況集計BL" & BL & ".xlsmにシート「Fault集計(BL" & BL & ")」を作るマクロです。", vbInformation, "BL" & BL

    Select Case BL
        Case 1
            Debug.Print "SCSS+"
            BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\運転状況集計\最新\SCSS\SCSS運転状況集計BL1.xlsm"
            SNAME_FAULT = "Fault集計(BL1)"
        Case 2
            Debug.Print "BL2"
            BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\運転状況集計\最新\SACLA\SACLA運転状況集計BL2.xlsm"
            SNAME_FAULT = "Fault集計(BL2)"
        Case 3
            Debug.Print ">>>BL3"
            BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\運転状況集計\最新\SACLA\SACLA運転状況集計BL3.xlsm"
            SNAME_FAULT = "Fault集計(BL3)"
        Case Else
            MsgBox "BLが不正です。終了します。", vbCritical
            End
    End Select
    
    ' wb_SHUKEIを開く
    Dim wb_SHUKEI As Workbook    ' ちゃんと宣言しないと、関数SheetExistsの引数が異なると怒られる
    Debug.Print "Debug<<<   Before  Function OpenBook(" & BNAME_SHUKEI & ")"
    Set wb_SHUKEI = OpenBook(BNAME_SHUKEI, False) ' フルパスを指定
    Debug.Print "Debug>>>   After  Function OpenBook(" & BNAME_SHUKEI & ")"
    If wb_SHUKEI Is Nothing Then Call Fin("ブックが開けませんでした。パスの異なる同じ名前のブックが既に開かれてる可能性があります。", 3)
    wb_SHUKEI.Activate
    If ActiveWorkbook.Name <> wb_SHUKEI.Name Then
        Call Fin("現在アクティブなブック名が異常です。終了します。" & vbCrLf & "ActiveWorkbook.Name:  " & ActiveWorkbook.Name & vbCrLf & "BNAME_SHUKEI:  " & BNAME_SHUKEI, 3)
    End If
    
    wb_SHUKEI.Windows(1).WindowState = xlMaximized
    wb_SHUKEI.Worksheets("Fault間隔(ユニット)").Activate
    wb_SHUKEI.Worksheets("Fault間隔(ユニット)").PageSetup.printArea = "" ' 20241113追加　印刷範囲が狭かった場合、範囲外が灰色なので選択しても見えないので印刷範囲をクリア
    If MsgBox("選択されてるユニット(シート「利用時間（期間）」のセルB2)は    " & vbCrLf & wb_SHUKEI.Worksheets("利用時間（期間）").Range("B2") & "   です。 " & vbCrLf & "間違いないですか？" & vbCrLf & "進むにはYESを押して下さい", vbYesNo + vbQuestion, "BL" & BL) = vbNo Then
        Call Fin("「No」が選択されました", 1)
    End If
    
    If CheckForErrors(wb_SHUKEI.Worksheets("Fault間隔(ユニット)")) Then
        MsgBox "シート「Fault間隔(ユニット)」 に数式エラーが発生してます。見直して下さい。終了します。", Buttons:=vbCritical
        End
    End If
    '追加部分----------------------------/

    Dim 最終行 As Integer
    
    Call 高速化処理開始

    Application.DisplayAlerts = False  '--- 確認メッセージを非表示
    If SheetDetect(SNAME_FAULT) Then
            wb_SHUKEI.Worksheets(SNAME_FAULT).Delete
    End If
    Application.DisplayAlerts = True   '--- 確認メッセージを表示
    
    ActiveSheet.Copy After:=ActiveSheet 'シートのコピー'
    ActiveSheet.Name = SNAME_FAULT 'シート名変更'
    
    MsgBox "デバック　1048579でいい。Rows.Count = " & Rows.Count
    最終行 = Cells(Rows.Count, 8).End(xlUp).Row
    MsgBox "デバック　最終行 = " & 最終行
    Range("A1:R" & 最終行).Value = Range("A1:R" & 最終行).Value '値の代入'
    
    Call Fault_セル結合
    
    Call 空白削除(8, 1000, 7)
    Call Fault_合計セル挿入
    Columns("J:R").Delete
    
    Call 高速化処理終了
        
    '/追加部分----------------------------
    Dim UnitLine As Integer
    UnitLine = getLineNum(wb_SHUKEI.Worksheets("利用時間（期間）").Range("B2"), 2, wb_SHUKEI.Worksheets(SNAME_FAULT))
    If UnitLine = -1 Then
        MsgBox "ユニット「" & wb_SHUKEI.Worksheets("利用時間（期間）").Range("B2") & "」が見つかりませんでした。ユーザー運転がなかったのかな？？" & vbCrLf & "本来なら、[" & SNAME_FAULT & "]シートのユニットの部分をコピーして、「SACLA運転状況集計まとめ.xlsm」に貼り付けるんですが、ないので、プログラムを終了します。" & vbCrLf & "後の手順を説明します。" & vbCrLf & "[SACLA運転状況集計まとめ.xlsm]のシート[Fault集計]を開いて「ユーザー運転無し」の行を手動で追加して下さい。", Buttons:=vbCritical
        Exit Sub
    Else
        wb_SHUKEI.Worksheets(SNAME_FAULT).Range("B" & UnitLine, "I" & UnitLine + wb_SHUKEI.Worksheets(SNAME_FAULT).Range("B" & UnitLine).mergeArea.Rows.Count - 1).Select
        wb_SHUKEI.Worksheets(SNAME_FAULT).Range("B" & UnitLine, "I" & UnitLine + wb_SHUKEI.Worksheets(SNAME_FAULT).Range("B" & UnitLine).mergeArea.Rows.Count - 1).Copy
    End If
    If MsgBox("選択されてる部分をコピーしました" & vbCrLf & "次は、「SACLA運転状況集計まとめ.xlsm」の「Fault集計」の張り付けです。" & vbCrLf & "ファイルを開きますか？", vbYesNo + vbQuestion, "BL" & BL) = vbYes Then
        
        ' wb_MATOMEを開く
        Dim wb_MATOME As Workbook    ' ちゃんと宣言しないと、関数SheetExistsの引数が異なると怒られる
        Set wb_MATOME = OpenBook(BNAME_MATOME, False) ' フルパスを指定
        If wb_MATOME Is Nothing Then Call Fin("ブックが開けませんでした。パスの異なる同じ名前のブックが既に開かれてる可能性があります。", 3)
        
        wb_MATOME.Worksheets("Fault集計").Activate 'これ大事
        Select Case BL
            Case 1
                Debug.Print "SCSS+"
            Case 2
                Debug.Print "BL2"
                beginL = getLineNum("SACLA Fault間隔(BL2)", 2, wb_MATOME.Worksheets("Fault集計"))
                EndL = getLineNum("SACLA Fault間隔(BL3)", 2, wb_MATOME.Worksheets("Fault集計"))
            Case 3
                Debug.Print ">>>BL3" 'BL3 の場合はB列の最終行からさかのぼる
                Dim xlLastRow As Long
                xlLastRow = wb_MATOME.Worksheets("Fault集計").UsedRange.Rows(wb_MATOME.Worksheets("Fault集計").UsedRange.Rows.Count).Row 'UsedRangeの注意点　罫線なども含んだ使用されている領域
                beginL = getLineNum("SACLA Fault間隔(BL3)", 2, wb_MATOME.Worksheets("Fault集計"))
                EndL = wb_MATOME.Worksheets("Fault集計").Cells(xlLastRow, 2).End(xlUp).Row   'B列の最終行を取得
            Case Else
                MsgBox "BLが不正です。終了します。", vbCritical
                End
        End Select
        
        For i = getLineNum_RS("ユニット", 2, beginL, EndL, wb_MATOME.Worksheets("Fault集計")) To wb_MATOME.Worksheets("Fault集計").UsedRange.Rows(wb_MATOME.Worksheets("Fault集計").UsedRange.Rows.Count).Row
            Debug.Print "i = " & i & "  " & Cells(i, 2).Value
            If IsEmpty(wb_MATOME.Worksheets("Fault集計").Cells(i, 2).Value) And Not wb_MATOME.Worksheets("Fault集計").Cells(i, 2).MergeCells Then
                targetline = i
                'MsgBox "セルが空っぽで、結合されてない、、" & vbCrLf & "", Buttons:=vbInformation
                Exit For
            End If
        Next
                
        wb_MATOME.Worksheets("Fault集計").Cells(targetline, 2).Select
        If MsgBox("ここに貼り付けします。" & vbCrLf & "いいですか？？", vbYesNo + vbQuestion, "BL" & BL) = vbYes Then
            wb_MATOME.Worksheets("Fault集計").Cells(targetline, 2).Insert xlDown
        End If
    
    End If
    
'    Call Fin("マクロ終了" & vbCrLf & "次はマクロ「運転集計_形式処理m」をしましょう！", 1) ' 親「Sub マクロいろいろxlsmからSACLA運転状況集計BLxlsmにマクロを流し込んで実行()」に戻りたいのでコメントアウトした
    Exit Sub  ' 通常の処理が完了したらエラーハンドラをスキップ
ErrorHandler:
    Call Fin("エラーです。内容は　 " & Err.Description, 3)
    Exit Sub
    '追加部分----------------------------/
    
End Sub









