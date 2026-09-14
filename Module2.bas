Attribute VB_Name = "Module2"
Option Explicit

Sub cp_paste_KEIKAKUZIKAN_UNTENZYOKYOSYUKEI(BL As Integer)
    On Error GoTo ErrorHandler

    Dim arr() As String
    Dim BNAME_SHUKEI As String
    Dim SNAME_KEIKAKU_BL As String
    Dim RANGE_GUN_HV_OFF As String
    Dim Col_GUN_HV_OFF As Integer
    Dim tr As Variant
    Dim result As Boolean
    Dim PasteSheet As Worksheet
    Dim PasteRow As Integer
    Dim pattern As String
    Dim newunit As String
    Debug.Print "============================================================================================================"


'    Dim buttonName As String
'    If TypeName(Application.Caller) = "String" Then
'        buttonName = Application.Caller
'    Else
'        MsgBox "このマクロはシート上のボタンからのみ実行してください。", Buttons:=vbCritical
'        End
'    End If
'
'    If buttonName = "ボタン 6" Then
'        BL = 2
'    ElseIf buttonName = "ボタン 7" Then
'        BL = 3
'    Else
'        MsgBox "異常です。終了します。" & vbCrLf & "buttonName = " & buttonName, Buttons:=vbCritical
'        End
'    End If
    MsgBox "「計画時間.xlsx」を「SACLA運転状況集計BL" & BL & ".xlsm」にコピーするマクロです。", vbInformation, "BL" & BL

    '    Dim s
    '    s = Application.InputBox("「計画時間.xlsx」を「SACLA運転状況集計BL" & BL & ".xlsm」にコピーするマクロです。 " & vbCrLf & vbCrLf & "BLを入力して下さい。", "BL" & BL)
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
        RANGE_GUN_HV_OFF = "A3:C"
        Col_GUN_HV_OFF = 1
    Case 3
        Debug.Print ">>>BL3"
        BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\運転状況集計\最新\SACLA\SACLA運転状況集計BL3.xlsm"
        SNAME_KEIKAKU_BL = "bl3"
        RANGE_GUN_HV_OFF = "G3:I"
        Col_GUN_HV_OFF = 7
    Case Else
        MsgBox "BLが不正です。終了します。" & vbCrLf & "！", Buttons:=vbInformation
        Exit Sub
    End Select


    ' wb_SHUKEIを開く
    Dim wb_SHUKEI As Workbook    ' ちゃんと宣言しないと、関数SheetExistsの引数が異なると怒られる
    Set wb_SHUKEI = OpenBook(BNAME_SHUKEI, False)    ' フルパスを指定
    If wb_SHUKEI Is Nothing Then Call Fin("ブックが開けませんでした。パスの異なる同じ名前のブックが既に開かれてる可能性があります。", 3)

    ' wb_KEIKAKUを開く
    Dim wb_KEIKAKU As Workbook    ' ちゃんと宣言しないと、関数SheetExistsの引数が異なると怒られる
    Set wb_KEIKAKU = OpenBook(BNAME_KEIKAKU, False)    ' フルパスを指定
    wb_KEIKAKU.Activate
    If wb_KEIKAKU Is Nothing Then Call Fin("ブックが開けませんでした。パスの異なる同じ名前のブックが既に開かれてる可能性があります。", 3)
    If ActiveWorkbook.Name <> wb_KEIKAKU.Name Then
        Call Fin("現在アクティブなブック名が異常です。終了します。" & vbCrLf & "ActiveWorkbook.Name:  " & ActiveWorkbook.Name & vbCrLf & "BNAME_SHUKEI:  " & BNAME_SHUKEI, 3)
    End If

    wb_KEIKAKU.Windows(1).WindowState = xlMaximized
    wb_KEIKAKU.Worksheets("GUN HV OFF").Select    '最前面に表示

    'コピーして貼り付け
    Set PasteSheet = wb_SHUKEI.Worksheets("GUN HV OFF時間記録")
    '
    If PasteSheet.Range("C6").Value = "" Then 'このセルが空欄だった場合、Range(" ").End(xlDown).Rowは最終行1048576を返すのでif分にした IsEmpty は“完全に空”だけを判定するため、空白チェックならこちらの方が安全
        PasteRow = 6
    Else
        PasteRow = PasteSheet.Range("C6").End(xlDown).Row + 1
    End If
    result = CpPaste(wb_KEIKAKU.Worksheets("GUN HV OFF"), RANGE_GUN_HV_OFF, Col_GUN_HV_OFF, PasteSheet, PasteSheet.Cells(PasteRow, 3), Array(2, 6, 7), 3)    '「シート GUN HV OFF」をコピーして貼り付け
    
    Set PasteSheet = wb_SHUKEI.Worksheets("運転予定時間")
    If PasteSheet.Range("B4").Value = "" Then 'このセルが空欄だった場合、Range(" ").End(xlDown).Rowは最終行1048576を返すのでif分にした IsEmpty は“完全に空”だけを判定するため、空白チェックならこちらの方が安全
        PasteRow = 4
    Else
        PasteRow = PasteSheet.Range("B4").End(xlDown).Row + 1
    End If
    result = CpPaste(wb_KEIKAKU.Worksheets(SNAME_KEIKAKU_BL), "A2:C", 1, PasteSheet, PasteSheet.Cells(PasteRow, 2), Array(1, 3, 5, 6, 8, 9, 10, 11, 12, 13), 2)    '「シート bl*」をコピーして貼り付け
    result = CpPaste(wb_KEIKAKU.Worksheets(SNAME_KEIKAKU_BL), "D2:D", 1, PasteSheet, PasteSheet.Cells(PasteRow, 7), -1, -1)    '「シート bl*の備考列」をコピーして貼り付け　' 前の行で、Check Array(1, 3, 5, 6, 8, 9, 10, 11, 12, 13), 2  してるから本来いらないので-1


    If PasteRow = 4 Then
        MsgBox "年度初めですね。ユニットは" & ThisWorkbook.sheetS("手順").Range(UNITNAME & UNITROW) & vbCrLf & "ですね。", Buttons:=vbInformation
        newunit = ThisWorkbook.sheetS("手順").Range(UNITNAME & UNITROW)
    Else
        '「新しいユニット名を計算」
        Dim before_unit As String
        Dim latest_unit As Integer
        PasteSheet.Cells(PasteRow - 1, 1).Select
        
        pattern = "^[1-9][0-9]*-[1-9][0-9]*$" ' パターン: 先頭(^)から、1-9で始まる数字の塊、ハイフン、1-9で始まる数字の塊、末尾($)まで
        If Not IsValidFormat(PasteSheet.Cells(PasteRow - 1, 1), pattern) Then
            Call CMsg("セル [" & PasteSheet.Cells(PasteRow - 1, 1).Value & "] の値が ユニットの形式（例: 2-11）ではありません。終了します。", vbCritical, PasteSheet.Cells(PasteRow - 1, 1))
            Exit Sub
        End If
        
        before_unit = PasteSheet.Cells(PasteRow - 1, 1)
        Debug.Print "before_unit: " & before_unit
        arr = Split(before_unit, "-")
        If Not IsNumeric(arr(1)) Then
            MsgBox "新しいユニット名を見繕うとしましたがユニット名がヘンです。 " & before_unit & vbCrLf & "終了します。", Buttons:=vbInformation
            Exit Sub
        End If
        latest_unit = Val(arr(1))
        latest_unit = latest_unit + 1
        newunit = arr(0) + "-" + CStr(latest_unit)
        Debug.Print "newunit: " & newunit
        If newunit <> ThisWorkbook.sheetS("手順").Range(UNITNAME & UNITROW) Then
            MsgBox "ユニット名が連続になりませんけど。今から出力しようとしているユニット名：" & ThisWorkbook.sheetS("手順").Range(UNITNAME & UNITROW) & vbCrLf & "  newunit: " & newunit, Buttons:=vbExclamation
        Else
            MsgBox "OK!" & vbCrLf & "新しいユニット名合致!!!", Buttons:=vbInformation
        End If
    End If


    PasteSheet.Activate
    PasteSheet.Cells(PasteSheet.Range("B3").End(xlDown).Row, 1).Activate    ' セルB3[運転種別]の最終行へ
    If MsgBox("ここに新しいユニット " & newunit & "を入れていいですか？？", vbYesNo + vbQuestion, "newunit") = vbYes Then
        PasteSheet.Cells(PasteSheet.Range("B3").End(xlDown).Row, 1) = newunit
    End If

    pattern = "^[1-9][0-9]*-[1-9][0-9]*$" ' パターン: 先頭(^)から、1-9で始まる数字の塊、ハイフン、1-9で始まる数字の塊、末尾($)まで
    If Not IsValidFormat(PasteSheet.Cells(PasteSheet.Range("B3").End(xlDown).Row, 1), pattern) Then
        Call CMsg("セル [" & PasteSheet.Cells(PasteSheet.Range("B3").End(xlDown).Row, 1).Value & "] の値が ユニットの形式（例: 2-11）ではありません。終了します。", vbCritical, PasteSheet.Cells(PasteSheet.Range("B3").End(xlDown).Row, 1))
        Exit Sub
    End If

    MsgBox "終了しました。" & vbCrLf & "保存してから、" & vbCrLf & "次、fault.txt出力(getBlFaultSummary.py)に進みましょう！", vbInformation, "BL" & BL

    If MsgBox("次の準備の為に、 シート「利用時間（期間）」の上の所に、ユニット[" & newunit & "]を入れていいですか？？", vbYesNo + vbQuestion, "newunit") = vbYes Then
        wb_SHUKEI.Worksheets("利用時間（期間）").Range("B2") = newunit
    End If

    Call Fin("これで終了です。", 1)
    Exit Sub    ' 通常の処理が完了したらエラーハンドラをスキップ
ErrorHandler:
    Call Fin("エラーです。内容は　 " & Err.Description, 3)
    Exit Sub

End Sub











Function CpPaste(sheetS As Worksheet, rangeS As String, colS As Integer, sheetT As Worksheet, pasteCELL As Variant, Arr_forCheck As Variant, Col_forCheck As Integer) As Boolean
'   rng1 As Range,
'    MsgBox sheetS.Columns(3).Address ' $C:$C
'    MsgBox sheetS.Range("$A:$A").Find(What:="*", LookIn:=xlValues, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row    ' 5
'    MsgBox sheetS.Range("$C:$C").Find(What:="*", LookIn:=xlValues, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row    ' 2
'    MsgBox sheetS.Range(Range(HeaderCELL).Columns.Address).Find(What:="*", LookIn:=xlValues, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row    ' エラー
    Dim tr As Variant
    Dim SetLastRow As Integer
    sheetS.Activate
    
    If CheckIsValidFormat(rangeS) Then
        SetLastRow = ExtractNumber(rangeS)
'        MsgBox "セル範囲のフォーマットは正しいです。" & SetLastRow, vbInformation
    Else
        MsgBox "セル範囲のフォーマットは正しくありません。", vbCritical
        Exit Function
    End If
    
    If SetLastRow > Cells(Rows.Count, colS).End(xlUp).Row Then
        MsgBox "コピーする範囲はありません。", vbInformation
        Exit Function
    End If
    
    Set tr = Range(rangeS & Cells(Rows.Count, colS).End(xlUp).Row)
    tr.Copy
    tr.Select
    If MsgBox("選択部分をコピーしました。" & "行数は " & tr.Rows.Count & vbCrLf & "次に進むにはYes", vbYesNo + vbQuestion) = vbNo Then Exit Function
    
    If Col_forCheck > 0 Then
        If Check(Arr_forCheck, Col_forCheck, tr.Rows.Count + 10, sheetT) <> 0 Then Call Fin("貼付け先のシートに数式が入っていない箇所が見つかりました。終了します。" & vbCrLf & "数式を直してから再度行って下さい。", 3)
    End If
    sheetT.Activate    ' これ必要。これないと、次の行で、セルをアクティブにできない
    pasteCELL.Activate

    If MsgBox("ここに貼り付けていいですか？", vbYesNo + vbQuestion) = vbYes Then
        pasteCELL.PasteSpecial Paste:=xlPasteValues
        If MsgBox("貼り付けましたがOKですか？？" & vbCrLf & "次に進むにはYes", vbYesNo + vbQuestion) = vbNo Then Exit Function
    End If

End Function


Function CheckIsValidFormat(str As String) As Boolean
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' 正規表現のパターンを設定
    regex.pattern = "^[A-Za-z0-9]+:[A-Za-z]+$"
    regex.IgnoreCase = True
    regex.Global = False
    
    ' フォーマットが一致するかどうかを判定
    CheckIsValidFormat = regex.Test(str)
End Function

Function ExtractNumber(str As String) As String
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    
    With reg
        .pattern = "\d+"    ' 数字が1回以上連続するパターン
        .Global = False     ' 最初の1致のみ
    End With
    
    Dim matches As Object
    Set matches = reg.Execute(str)
    
    If matches.Count > 0 Then
        ExtractNumber = matches(0).Value
    Else
        ExtractNumber = ""
    End If
End Function
