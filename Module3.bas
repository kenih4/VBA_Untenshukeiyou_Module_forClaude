Attribute VB_Name = "Module3"
Option Explicit

Sub cp_paste_faulttxt_UNTENZYOKYOSYUKEI(BL As Integer, ROW_COUNT As Integer)
    On Error GoTo ErrorHandler

    Debug.Print "============================================================================================================"
    Debug.Print "============cp_paste_faulttxt_UNTENZYOKYOSYUKEI========== BL=" & BL & "    ROW_COUNT=" & ROW_COUNT & "======"
    Debug.Print "============================================================================================================"

    Dim targetline As Integer
    Dim tempText As String
    Dim BNAME_SHUKEI As String
    Dim SNAME_KEIKAKU_BL As String


    ' ウィンドウを標準サイズにする
    Application.WindowState = xlMaximized
    ' 最前面に持ってくる
    Application.ActiveWindow.Activate

    Dim WSH
    Set WSH = CreateObject("Wscript.Shell")
    '    Dim BL As Integer  ' 対象BL

    Dim CB As Variant, i As Long
    CB = Application.ClipboardFormats
    If CB(1) = True Then
        MsgBox "クリップボードは空です。python getBlFaultSummary_LOCALTEST.pyを走らせたら何かしらクリップボードに入るはずなのでなにかおかしいです。将又、一度もトリップがなかったか？？", vbCritical, "BL" & BL
        Exit Sub
    Else

        With New DataObject
            .GetFromClipboard
            tempText = .GetText
        End With

        If MsgBox("python getBlFaultSummary_LOCALTEST.pyの出力、" & vbCrLf & "「falut.txt」をSACLA運転状況集計BL " & BL & " .xlsmのシート「集計記録」に張り付けるマクロです。" & vbCrLf & vbCrLf & "クリップボードの中身は以下です。進みますか？" & vbCrLf & vbCrLf & "「" & vbCrLf & tempText & vbCrLf & "」", vbYesNo + vbQuestion, "BL" & BL) = vbNo Then
            Exit Sub
        End If

    End If





    '    Dim s
    '    s = Application.InputBox("BLを入力して下さい。", "確認")
    '    If s = False Then
    '        Exit Sub
    '    ElseIf s = "" Then
    '        MsgBox "何も入力されていません"
    '        Exit Sub
    '    Else
    '        BL = s
    '    End If

    Select Case BL
    Case 1
        Debug.Print "SCSS+"
    Case 2
        Debug.Print "BL2"
        BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\運転状況集計\最新\SACLA\SACLA運転状況集計BL2.xlsm"
        SNAME_KEIKAKU_BL = "bl2"
    Case 3
        Debug.Print ">>>BL3"
        BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\運転状況集計\最新\SACLA\SACLA運転状況集計BL3.xlsm"
        SNAME_KEIKAKU_BL = "bl3"
    Case Else
        MsgBox "BLが不正です。終了します。", vbCritical
        Exit Sub
    End Select






    ' wb_SHUKEIを開く
    Dim wb_SHUKEI As Workbook    ' ちゃんと宣言しないと、関数SheetExistsの引数が異なると怒られる
    Set wb_SHUKEI = OpenBook(BNAME_SHUKEI, False)    ' フルパスを指定
    If wb_SHUKEI Is Nothing Then Call Fin("ブックが開けませんでした。パスの異なる同じ名前のブックが既に開かれてる可能性があります。", 3)
    wb_SHUKEI.Activate
    If ActiveWorkbook.Name <> wb_SHUKEI.Name Then
        Call Fin("現在アクティブなブック名が異常です。終了します。" & vbCrLf & "ActiveWorkbook.Name:  " & ActiveWorkbook.Name & vbCrLf & "BNAME_SHUKEI:  " & BNAME_SHUKEI, 3)
    End If

    wb_SHUKEI.Windows(1).WindowState = xlMaximized
    wb_SHUKEI.Worksheets("集計記録").Activate

'   targetline = wb_SHUKEI.Worksheets("集計記録").Range("C8").End(xlDown).ROW + 1 ' セルC8「開始時間」から最終行へ（データが連続している場合は、空白セルの手前のデータを取得）
    targetline = wb_SHUKEI.Worksheets("集計記録").Cells(wb_SHUKEI.Worksheets("集計記録").Rows.Count, "C").End(xlUp).Row + 1 ' 列Cの最下行から上方向にデータを探すので、空白があっても無視できます。
    If Check(Array(7, 8, 9), 3, ROW_COUNT + 10, wb_SHUKEI.Worksheets("集計記録")) <> 0 Then Call Fin("貼付け先のシートに数式が入っていない箇所が見つかりました。終了します。" & vbCrLf & "数式を直してから再度行って下さい。", 3)
    wb_SHUKEI.Worksheets("集計記録").Cells(targetline, 1).Activate
    MsgBox "ここに　Ctrl+Vして、「fault.txt」を貼り付けて下さい。" & vbCrLf & "注意点：先に片方のBLを引き渡した場合など、この段階で調整時間（ユニット切替えなど）の時間を確認しておく！！", vbInformation, "BL" & BL
    'wb_SHUKEI.Worksheets("集計記録").Cells(targetline, 1) = tempText  これだと1つのセルの中にtempTextが入ってしまうぅ



'    If MsgBox("貼り付け終わったらシート「利用時間(User)」にエネルギーなどを手動入力しましょう！！" & vbCrLf & "シフトサマリーを開きますか？？", vbYesNo + vbQuestion, "BL" & BL) = vbYes Then
'
'        Dim sourceWorkbook As Workbook    ' ちゃんと宣言しないと、関数SheetExistsの引数が異なると怒られる
'        Set sourceWorkbook = OpenBook(BNAME_SOURCE, False) ' フルパスを指定
'        If sourceWorkbook Is Nothing Then Call Fin("ブックが開けませんでした。パスの異なる同じ名前のブックが既に開かれてる可能性があります。", 3)
'        sourceWorkbook.Worksheets("編集用_利用時間(User)").Activate
'
'        wb_SHUKEI.Worksheets("利用時間(User)").Activate
'
'        Windows.Arrange ArrangeStyle:=xlVertical
'
'        WSH.RUN "http://saclaopr19.spring8.or.jp/~summary/display_ui.html?sort=date%20desc%2Cstart%20desc&limit=0%2C100&search_root=BL" & BL & "#STATUS", 3
'        Set WSH = Nothing
'    End If



    '================================================
    '            案1 ダメ
    '            Dim DataObj As Object
    '            Dim ClipText As String
    '            ' クリップボードのデータを取得
    '            Set DataObj = CreateObject("MSForms.DataObject") 'ActiveX コンポーネントはオブジェクトを作成できません
    '            DataObj.GetFromClipboard
    '            ' クリップボードの内容をテキストとして取得
    '            ClipText = DataObj.GetText
    '
    '            ' テキストが空でない場合に貼り付け
    '            If Len(ClipText) > 0 Then
    '                ' アクティブなセルに貼り付け
    '                ActiveCell.Value = ClipText
    '            Else
    '                MsgBox "クリップボードの内容はテキストではありません。", vbExclamation
    '            End If
    '================================================

    '            Dim CB As Variant, i As Long
    '            CB = Application.ClipboardFormats
    '            If CB(1) = True Then
    '                MsgBox "クリップボードは空です。python getBlFaultSummary_LOCALTEST.pyを走らせたら何かしらクリップボードに入るはずなのでなにかおかしいです。"
    '                End
    '            Else
    '                With New DataObject
    '                    .GetFromClipboard
    '                    tempText = .GetText
    '                End With
    '                MsgBox "ここに　Ctrl+Vして、「fault.txt」を貼り付けて下さい" & vbCrLf & "それで終了です。" & vbCrLf & "クリップボードの中身は以下です。" & vbCrLf & vbCrLf & tempText, Buttons:=vbInformation
    '
    '                If MsgBox("貼り付け終わったらシート「利用時間(User)」にエネルギーなどを手動入力しましょう！！" & vbCrLf & "シフトサマリーを開きますか？？", vbYesNo + vbQuestion, "確認") = vbYes Then
    '                    WSH.Run "http://saclaopr19.spring8.or.jp/~summary/display_ui.html?sort=date%20desc%2Cstart%20desc&limit=0%2C100&search_situation=ユーザー運転&&search_root=BL" & BL & "#STATUS", 3 ' 第2引数: 3は最大化
    '                    Set WSH = Nothing
    '                End If
    '
    '            End If


    'If MsgBox("ここに値「fault.txt」を貼り付けていいですか？", vbYesNo + vbQuestion, "確認") = vbYes Then
    ' 危険    'Application.SendKeys "^v" ' Ctrl+Vで貼り付け
    'End If



    'なぜか貼り付け不能。
    '            If MsgBox("ここに値「fault.txt」を貼り付けていいですか？", vbYesNo + vbQuestion, "確認") = vbYes Then
    '                Dim CB As Variant, i As Long
    '                CB = Application.ClipboardFormats
    '                If CB(1) = True Then
    '                    MsgBox "クリップボードは空です。python getBlFaultSummary_LOCALTEST.pyを走らせたら何かしらクリップボードに入るはずなのでなにかおかしいです。"
    '                Else
    '                    With New DataObject
    '                        .GetFromClipboard
    '                        tempText = .GetText
    '                    End With
    '                    MsgBox tempText
    '                    wb_SHUKEI.Worksheets("集計記録").Cells(targetline, 1).Paste
    '上の行では、貼り付けできないのでダメなので一旦以下のように、クリップボードに再度いれてみたがだめ
    '                    Dim cbData As New DataObject
    '                    Dim cbFormat As Variant
    '                    'DataObjectにメッセージを格納
    '                    cbData.SetText tempText
    '                    'DataObjectのデータをクリップボードに格納
    '                    cbData.PutInClipboard
    '                    wb_SHUKEI.Worksheets("集計記録").Cells(targetline, 1).Paste
    'DAME                    wb_SHUKEI.Worksheets("集計記録").Cells(targetline, 1).PasteSpecial Paste:=xlPasteValues
    '                End If
    '            End If

    Call Fin("これで終了です。", 1)
    Exit Sub  ' 通常の処理が完了したらエラーハンドラをスキップ
ErrorHandler:
    Call Fin("エラーです。内容は　 " & Err.Description, 3)
    Exit Sub

End Sub












