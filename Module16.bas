Attribute VB_Name = "Module16"
Option Explicit ' 未定義の変数は使用できないように


Sub TEST_Button_Click()
    Debug.Print "TEST"
    MsgBox "TEST_Button_Click" & vbCrLf & " " & vbCrLf & "test" & vbCrLf & " " & vbCrLf & " ", vbInformation
   
    Application.StatusBar = "TEST_Button_Clickしました。"

    Application.VBE.MainWindow.Visible = True
    
'    Range(Cells(1, 1), Cells(3, 3)).Interior.Color = RGB(0, 255, 0) ' 緑色
    
'    Dim pattern As String
'    pattern = "^[1-9][0-9]*-[1-9][0-9]*$" ' パターン: 先頭(^)から、1-9で始まる数字の塊、ハイフン、1-9で始まる数字の塊、末尾($)まで
'    If Not IsValidFormat(ThisWorkbook.sheetS("手順").Range(UNITNAME & UNITROW), pattern) Then
'        Call CMsg("Err セル [" & ThisWorkbook.sheetS("手順").Range(UNITNAME & UNITROW).Value & "] の値が ユニットの形式（例: 2-11）ではありません。終了します。", vbCritical, ThisWorkbook.sheetS("手順").Range(UNITNAME & UNITROW))
'    Else
'        Call CMsg("OK セル [" & ThisWorkbook.sheetS("手順").Range(UNITNAME & UNITROW).Value & "] の値が ユニットの形式（例: 2-11）です", vbInformation, ThisWorkbook.sheetS("手順").Range(UNITNAME & UNITROW))
'    End If


'    MsgBox vbInformation

'    MsgBox vbExclamation

'    Call RunGitBashCommands


'    MsgBox Month(ThisWorkbook.sheetS("手順").Range(BEGIN_COL & UNITROW))
'    Exit Sub
    
End Sub



