Attribute VB_Name = "Module6"
Option Explicit

Sub 利用時間Userに手動入力(BL As Integer)
    On Error GoTo ErrorHandler

    Debug.Print "============================================================================================================"
    Debug.Print "============利用時間Userに手動入力================================================================ BL=" & BL
    Debug.Print "============================================================================================================"

    Dim BNAME_SHUKEI As String
    Dim LineSta As Integer
    Dim LineSto As Integer
    Dim i As Integer
        
    Dim result As Double
    result = Application.WorksheetFunction.RoundUp(Now - ThisWorkbook.sheetS("手順").Cells(UNITROW, 5).Value, 0)
    
    MsgBox "python Pickup_from_shiftsummary.pyで、" & result & "日分のシフトサマリーを取得します。"
    If RunPythonScript("Pickup_from_shiftsummary.py BL" & BL & " " & result * 3, "C:\Users\kenic\Dropbox\gitdir\Pickup_from_shiftsummary") = False Then
        MsgBox "pythonでエラー発生です。シフトサマリーから取得できませんでした。手動で行ってください。", Buttons:=vbExclamation
    End If
     
    
    ' ウィンドウを標準サイズにする
    Application.WindowState = xlNormal
    ' 最前面に持ってくる
    Application.ActiveWindow.Activate
    
    Select Case BL
    Case 1
        Debug.Print "SCSS+"
    Case 2
        Debug.Print "BL2"
        BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\運転状況集計\最新\SACLA\SACLA運転状況集計BL2.xlsm"
    Case 3
        Debug.Print ">>>BL3"
        BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\運転状況集計\最新\SACLA\SACLA運転状況集計BL3.xlsm"
    Case Else
        MsgBox "BLが不正です。終了します。", vbCritical
        Exit Sub
    End Select

'    Dim WSH
'    Set WSH = CreateObject("Wscript.Shell")
'    WSH.RUN "http://saclaopr19.spring8.or.jp/~summary/display_ui.html?sort=date%20desc%2Cstart%20desc&limit=0%2C100&search_root=BL" & BL & "#STATUS", 3
'    Set WSH = Nothing

   ' マクロいろいろを開く　既に開かれてるが
    Dim sourceWorkbook As Workbook    ' ちゃんと宣言しないと、関数SheetExistsの引数が異なると怒られる
    Set sourceWorkbook = OpenBook(BNAME_SOURCE, False) ' フルパスを指定
    If sourceWorkbook Is Nothing Then Call Fin("ブックが開けませんでした。パスの異なる同じ名前のブックが既に開かれてる可能性があります。", 3)
    sourceWorkbook.Worksheets("編集用_利用時間(User)BL" & BL).Activate

    ' wb_SHUKEIを開く
    Dim wb_SHUKEI As Workbook    ' ちゃんと宣言しないと、関数SheetExistsの引数が異なると怒られる
    Set wb_SHUKEI = OpenBook(BNAME_SHUKEI, False)    ' フルパスを指定
    If wb_SHUKEI Is Nothing Then Call Fin("ブックが開けませんでした。パスの異なる同じ名前のブックが既に開かれてる可能性があります。", 3)
    wb_SHUKEI.Worksheets("利用時間(User)").Activate
    
    LineSta = getLineNum("ユニット", 2, wb_SHUKEI.Worksheets("利用時間(User)"))
    LineSto = wb_SHUKEI.Worksheets("利用時間(User)").Cells(Rows.Count, "B").End(xlUp).Row
    Debug.Print " LineSto :   " & LineSto
    Dim Kokokara As Long
    Dim Kokomade As Long
    For i = LineSta To LineSto
        Debug.Print "DEBUG 　    i = " & i & "  " & Cells(i, 2).Value
        If wb_SHUKEI.Worksheets("利用時間(User)").Cells(i, 2).Value = wb_SHUKEI.Worksheets("利用時間（期間）").Range("B2") Then
            Debug.Print "この行　i = " & i & " が、　　ユニット： " & Cells(i, 2).Value
            'Cells(i, 15).Select
            Kokokara = i
            Exit For
        End If
    Next
    Debug.Print "利用時間(User)の最終行 = " & wb_SHUKEI.Worksheets("利用時間(User)").Range(wb_SHUKEI.Worksheets("利用時間(User)").Columns(15).Address).Find(What:="*", LookIn:=xlValues, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row '罫線含まない最終行
    Kokomade = wb_SHUKEI.Worksheets("利用時間(User)").Range(wb_SHUKEI.Worksheets("利用時間(User)").Columns(15).Address).Find(What:="*", LookIn:=xlValues, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    Range("O" & Kokokara & ":" & "O" & Kokomade).Select

    Windows.Arrange ArrangeStyle:=xlVertical



    Call Fin("マクロはこれで終了です。" & vbCrLf & "あとはシフトサマリーからエネルギー、繰り返し、、波長、強度をピックアップして下さい", 1)
    Exit Sub  ' 通常の処理が完了したらエラーハンドラをスキップ
ErrorHandler:
    Call Fin("エラーです。内容は　 " & Err.Description, 3)
    Exit Sub

End Sub
