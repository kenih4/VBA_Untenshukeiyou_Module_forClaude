Attribute VB_Name = "Module12"
Option Explicit

Function RunPythonScript(scriptPath As String, workDir As String) As Boolean
    On Error GoTo ErrorHandler
    RunPythonScript = False
'    Dim scriptPath As String
    Dim Command As String
    Dim buttonName As String
'    Const pythonExe As String = "python"
    Dim pythonExe As String
    pythonExe = "python"
    Debug.Print "Debug   " & workDir
    
'    If TypeName(Application.Caller) = "String" Then
'        buttonName = Application.Caller
'    Else
'        MsgBox "このマクロはシート上のボタンからのみ実行してください。" & vbCrLf & "終了します。", Buttons:=vbCritical
'        End
'    End If
'
'    If buttonName = "ボタン 8" Then
'        scriptPath = "getGunHvOffTime_LOCALTEST.py"
'    ElseIf buttonName = "ボタン 9" Then
'        scriptPath = "getBlFaultSummary_LOCALTEST.py bl2"
'    ElseIf buttonName = "ボタン 10" Then
'        scriptPath = "getBlFaultSummary_LOCALTEST.py bl3"
'    Else
'        MsgBox "異常です。終了します。" & vbCrLf & "buttonName = " & buttonName, Buttons:=vbCritical
'        End
'    End If
    Debug.Print "Debug   scriptPath=" & scriptPath
    MsgBox "python " & scriptPath & "を" & vbCrLf & "実行します。", Buttons:=vbInformation
    
    ' コマンドを組み立て：まず指定フォルダに移動し、その後Pythonを実行
    Command = "cmd.exe /c cd /d " & Chr(34) & workDir & Chr(34) & " && " & pythonExe & " " & workDir & "\" & scriptPath
    Debug.Print "Debug  Command=" & Command
    'Shell command, vbNormalFocus ' Shell関数でPythonスクリプトを実行 終了を待たない
    
    Dim shell As Object
    Dim exitCode As Long
    Set shell = CreateObject("WScript.Shell")
    exitCode = shell.RUN(Command, vbMaximizedFocus, True)   ' WScript.ShellのRunメソッドでコマンドを実行し、終了を待つ
    If exitCode = 0 Then
        RunPythonScript = True
        Call Fin("Pythonスクリプトが正常に終了しました。 " & vbCrLf & "[" & scriptPath & "]", 1)
    Else
        Call Fin("Pythonスクリプトがエラーコード " & exitCode & " で終了しました。", 3)
    End If
    
    
    Exit Function ' 通常の処理が完了したらエラーハンドラをスキップ
ErrorHandler:
    Call Fin("エラーです。内容は　 " & Err.Description, 3)
    Exit Function
    
End Function

