Attribute VB_Name = "Module11"
Option Explicit

Sub 運転集計_形式処理m(BL As Integer)

    '/追加部分----------------------------
    On Error GoTo ErrorHandler
    Dim BNAME_SHUKEI As String
    Dim DOWNTIME_ROW As Integer
    MsgBox "マクロ「運転集計_形式処理m」を実行します。" & vbCrLf & "このマクロは、" & vbCrLf & "ひな形シート「運転状況（対象ユニット）」からシート「24-*(BL" & BL & ")」を作成します。", vbInformation, "BL" & BL
    
    Select Case BL
        Case 1
            Debug.Print "SCSS+"
            BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\運転状況集計\最新\SCSS\SCSS運転状況集計BL1.xlsm"
        Case 2
            Debug.Print "BL2"
            BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\運転状況集計\最新\SACLA\SACLA運転状況集計BL2.xlsm"
            DOWNTIME_ROW = 8
        Case 3
            Debug.Print ">>>BL3"
            BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\運転状況集計\最新\SACLA\SACLA運転状況集計BL3.xlsm"
            DOWNTIME_ROW = 9
        Case Else
            MsgBox "BLが不正です。終了します。", vbCritical
            End
    End Select

    ' wb_SHUKEIを開く
    Dim wb_SHUKEI As Workbook    ' ちゃんと宣言しないと、関数SheetExistsの引数が異なると怒られる
    Set wb_SHUKEI = OpenBook(BNAME_SHUKEI, False) ' フルパスを指定
    If wb_SHUKEI Is Nothing Then Call Fin("ブックが開けませんでした。パスの異なる同じ名前のブックが既に開かれてる可能性があります。", 3)
    wb_SHUKEI.Activate
    If ActiveWorkbook.Name <> wb_SHUKEI.Name Then
        Call Fin("現在アクティブなブック名が異常です。終了します。" & vbCrLf & "ActiveWorkbook.Name:  " & ActiveWorkbook.Name & vbCrLf & "BNAME_SHUKEI:  " & BNAME_SHUKEI, 3)
    End If
        
    wb_SHUKEI.Windows(1).WindowState = xlMaximized
    wb_SHUKEI.Worksheets("運転状況(対象ユニット)").Activate
    
    If MsgBox("選択されてるユニット(シート「利用時間（期間）」のセルB2)は    " & vbCrLf & wb_SHUKEI.Worksheets("利用時間（期間）").Range("B2") & "   です。 " & vbCrLf & "間違いないですか？" & vbCrLf & "進むにはYESを押して下さい", vbYesNo + vbQuestion, "BL" & BL) = vbNo Then
        Call Fin("「No」が選択されました", 1)
    End If
    '追加部分----------------------------/


    Dim 最終行 As Integer
    Dim シート名 As String

    最終行 = Cells(Rows.Count, 16).End(xlUp).Row
    シート名 = (Cells(8, 2).Value & "(BL" & BL & ")")
    Call 高速化処理開始
    
    'シートの重複処理'
    Application.DisplayAlerts = False  '--- 確認メッセージを非表示
    If SheetDetect(シート名) Then
            Worksheets(シート名).Delete
    End If
    Application.DisplayAlerts = True   '--- 確認メッセージを表示

    sheetS("運転状況(対象ユニット)").Copy After:=ActiveSheet 'シートのコピー'
    ActiveSheet.Name = シート名 'シート名変更'
    Range("A1:P" & 最終行).Value = Range("A1:P" & 最終行).Value '数式⇒値へ変換'

    If Cells(Range("P1:P500").Find("条件_開始行").Row + 1, 7) = "" Then 'ユーザーがいないとき'
       Rows(Range("P1:P500").Find("シフト毎_開始行").Row + 1 & ":" & Range("P1:P500").Find("シフト毎_終了行").Row).Delete
       Rows(Range("P1:P500").Find("条件_開始行").Row + 1 & ":" & Range("P1:P500").Find("シフトユーザー_終了行").Row).Delete
    Else
       Call 空白削除(Range("P1:P500").Find("シフト毎_開始行").Row + 1, Range("P1:P500").Find("シフト毎_終了行").Row - 1, 3)  'シフト毎_空白削除'
       Call 空白削除(Range("P1:P500").Find("条件_開始行").Row + 1, Range("P1:P500").Find("条件_終了行").Row - 1, 3) '条件_空白削除'
       Call シフトユーザー行挿入
       Call シフトユーザー行_削除
       Call 条件行_罫線
    End If

    Call 印刷設定

    Columns("O:P").Delete

    Call 高速化処理終了 ' なぜかここでエクセルが落ちる。手動でSACLA運転状況集計BL2.xlsmを立ち上げて、オリジナルの Sub 運転集計_形式処理 を走らせても落ちないのに。
                        ' 原因判明：　なんらかの理由で　「高速化処理終了」がなされなかった状態で、再度このプログラムを走らすと落ちる事が判明。その場合、手動で「高速化処理終了」を実行してやる。

    '/追加部分----------------------------
    wb_SHUKEI.Worksheets(シート名).Cells(1, 1).Select
    If wb_SHUKEI.Worksheets(シート名).Cells(DOWNTIME_ROW, 9).Value = 0 Then
        MsgBox "利用調整運転(BL調整orBL-study)はなかったんですね。　" & vbCrLf & "", vbExclamation, "BL" & BL
    End If
    
    If wb_SHUKEI.Worksheets(シート名).Cells(DOWNTIME_ROW, 11).Value = 0 Then
        MsgBox "利用運転(ユーザー)はなかったんですね。　" & vbCrLf & "" & vbCrLf & "「ユーザー運転無し」と手動で処理しないといけない部分があります。", vbExclamation, "BL" & BL
    Else
        If wb_SHUKEI.Worksheets(シート名).Cells(DOWNTIME_ROW, 12).Value = 0 Then
            MsgBox "ダウンタイムは　" & wb_SHUKEI.Worksheets(シート名).Cells(DOWNTIME_ROW, 12).Value & " です。一回もトリップしてないって事？確認した方がよいです。" & vbCrLf & "シート「集計記録」に数式が入っていない可能性があります", vbExclamation, "BL" & BL
        End If
    End If

    If MsgBox("今表示さているシート「" & シート名 & "」が作成されたものです。" & vbCrLf & "これを「SACLA運転状況集計まとめ.xlsm」にコピーしますか？", vbYesNo + vbQuestion, "BL" & BL) = vbYes Then
        ' wb_MATOMEを開く
        Dim wb_MATOME As Workbook    ' ちゃんと宣言しないと、関数SheetExistsの引数が異なると怒られる
        Set wb_MATOME = OpenBook(BNAME_MATOME, False) ' フルパスを指定
        If wb_MATOME Is Nothing Then Call Fin("ブックが開けませんでした。パスの異なる同じ名前のブックが既に開かれてる可能性があります。", 3)
        
        wb_SHUKEI.Worksheets(シート名).Copy After:=wb_MATOME.Worksheets("まとめ ")
        wb_MATOME.Worksheets(シート名).Activate
        MsgBox wb_MATOME.Name & "に" & vbCrLf & "シートのコピーが完了しまた。" & vbCrLf & "BL2/BL3両方終わったらマージしましょう！", Buttons:=vbInformation
    End If
    
'    Call Fin("マクロ終了" & vbCrLf & "！", 1) ' 親「Sub マクロいろいろxlsmからSACLA運転状況集計BLxlsmにマクロを流し込んで実行()」に戻りたいのでコメントアウトした
    Exit Sub  ' 通常の処理が完了したらエラーハンドラをスキップ
ErrorHandler:
    Call Fin("エラーです。内容は　 " & Err.Description, 3)
    Exit Sub
    '追加部分----------------------------/
    
    
End Sub

