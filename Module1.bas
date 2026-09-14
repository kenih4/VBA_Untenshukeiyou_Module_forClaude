Attribute VB_Name = "Module1"
Option Explicit

Sub マクロいろいろxlsmからSACLA運転状況集計BLxlsmにマクロを流し込んで実行(BL As Integer, macroName As String)
    On Error GoTo ErrorHandler
    Debug.Print "Debug   マクロいろいろxlsmからSACLA運転状況集計BLxlsmにマクロを流し込んで実行"
        
    Dim result As Boolean
    Dim sourceModule As Object
    Dim targetModule As Object
       
    'Dim BL As Integer
    Dim BNAME_SHUKEI As String
    'Dim macroName As String
    Dim vbComp As VBIDE.vbComponent
    
    Dim dict As Dictionary
    Set dict = New Dictionary
    dict.Add "Fault集計m", "Module10"
    dict.Add "運転集計_形式処理m", "Module11"

'    Dim buttonName As String
'    If TypeName(Application.Caller) = "String" Then
'        buttonName = Application.Caller
'    Else
'        Call Fin("このマクロはシート上のボタンからのみ実行してください。" & vbCrLf & "終了します。", 3)
'    End If
'
'    If buttonName = "ボタン 1" Then
'        BL = 2
'        macroName = "Fault集計m"
'    ElseIf buttonName = "ボタン 2" Then
'        BL = 2
'        macroName = "運転集計_形式処理m"
'    ElseIf buttonName = "ボタン 4" Then
'        BL = 3
'        macroName = "Fault集計m"
'    ElseIf buttonName = "ボタン 5" Then
'        BL = 3
'        macroName = "運転集計_形式処理m"
'    Else
'        MsgBox "異常です。終了します。" & vbCrLf & "buttonName = " & buttonName, Buttons:=vbCritical
'        End
'    End If
    
    BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\運転状況集計\最新\SACLA\SACLA運転状況集計BL" & BL & ".xlsm"

        
    
    ' sourceWorkbookを開く
    Dim sourceWorkbook As Workbook    ' ちゃんと宣言しないと、関数SheetExistsの引数が異なると怒られる
    Set sourceWorkbook = OpenBook(BNAME_SOURCE, False) ' フルパスを指定
    If sourceWorkbook Is Nothing Then Call Fin("ブックが開けませんでした。パスの異なる同じ名前のブックが既に開かれてる可能性があります。", 3)
    
'    For Each vbComp In sourceWorkbook.VBProject.VBComponents
'        Debug.Print "Debug   vbComp.name =  " & vbComp.Name & "     vbComp.Type: " & vbComp.Type
'    Next vbComp
'MOTO    Set sourceModule = sourceWorkbook.VBProject.VBComponents(dict(macroName)) ' モジュール名を確認       Module10 = Fault集計m()
        
    ' targetWorkbookを開く
    Dim targetWorkbook As Workbook
    Set targetWorkbook = OpenBook(BNAME_SHUKEI, False) ' フルパスを指定
    If targetWorkbook Is Nothing Then Call Fin("ブックが開けませんでした。パスの異なる同じ名前のブックが既に開かれてる可能性があります。", 3)

    
    result = sourceWorkbookからtargetWorkbookにmoduleNameを流し込む(sourceWorkbook, targetWorkbook, "Module8", "RunBatchFile", False) ' 共通関数
    If result Then
        Debug.Print "成功 「sourceWorkbookからtargetWorkbookにmoduleNameを流し込む」（共通関数）"
    Else
        Call Fin("失敗 「sourceWorkbookからtargetWorkbookにmoduleNameを流し込む」（共通関数）", 3)
    End If
    
    result = sourceWorkbookからtargetWorkbookにmoduleNameを流し込む(sourceWorkbook, targetWorkbook, dict(macroName), macroName, False)
    If result Then
        Debug.Print "成功 「sourceWorkbookからtargetWorkbookにmoduleNameを流し込む」"
    Else
        Call Fin("失敗 「sourceWorkbookからtargetWorkbookにmoduleNameを流し込む」", 3)
    End If
    
    
'
    If MsgBox("流し込んだマクロを実行します。" & vbCrLf & "いいですか？？", vbYesNo + vbQuestion, "BL" & BL) = vbYes Then
        Debug.Print "<<<<<<ブック「" & targetWorkbook.Name & "」　の　マクロ「" & macroName & "」 を実行します"
        Application.RUN "'" & targetWorkbook.Name & "'!" & macroName, BL
        MsgBox "マクロ「" & macroName & "」 が完了しました！", Buttons:=vbInformation
        Debug.Print "マクロ「" & macroName & "」 が完了しました>>>>>>>>>>"
    End If


    'マクロmacroNameを片づける
    result = sourceWorkbookからtargetWorkbookにmoduleNameを流し込む(sourceWorkbook, targetWorkbook, "Module8", "RunBatchFile", True) ' 共通関数
    result = sourceWorkbookからtargetWorkbookにmoduleNameを流し込む(sourceWorkbook, targetWorkbook, dict(macroName), macroName, True)
    MsgBox "流し込んだマクロの片付けが終了しました。", Buttons:=vbInformation
    
    ' ワークブックを閉じる
    'sourceWorkbook.Close SaveChanges:=False
    'targetWorkbook.Close SaveChanges:=True
    
   
    Call Fin("これで終了です。", 1)
    Exit Sub  ' 通常の処理が完了したらエラーハンドラをスキップ
ErrorHandler:
    Call Fin("エラーです。内容は　 " & Err.Description, 3)
    Exit Sub
    
    
End Sub





'Not USE    マクロmacroNameが、workbookNameに存在するか確認して「モジュール」を削除する  返り値が欲しいのでFunctionにした===========================================================================
Function CheckAndDeleteModuleContainingMacro(WorkBookName As String, macroName As String) As Boolean
    Dim targetWorkbook As Workbook
    Dim vbComponent As VBIDE.vbComponent
    Dim exists As Boolean

    ' 指定したブックを設定
    On Error Resume Next
    Set targetWorkbook = Workbooks.Open(WorkBookName) ' 指定したブック名で開いているか確認
    targetWorkbook.Windows(1).WindowState = xlMaximized
    On Error GoTo 0

    If targetWorkbook Is Nothing Then
        MsgBox "指定したブック '" & WorkBookName & "' が開いていません。"
        CheckAndDeleteModuleContainingMacro = False
        Exit Function
    End If

    ' モジュールをループ
    exists = False
    For Each vbComponent In targetWorkbook.VBProject.VBComponents
        If vbComponent.Type = vbext_ct_StdModule Or vbComponent.Type = vbext_ct_ClassModule Then
            ' モジュールが空でない場合のみ確認
            If vbComponent.CodeModule.CountOfLines > 0 Then
                ' モジュールのコードを確認
                If InStr(vbComponent.CodeModule.Lines(1, vbComponent.CodeModule.CountOfLines), "Sub " & macroName & "(") > 0 Then
                    exists = True
                    ' モジュールを削除
                    targetWorkbook.VBProject.VBComponents.Remove vbComponent
                    Exit For
                End If
            End If
        End If
    Next vbComponent

    CheckAndDeleteModuleContainingMacro = exists
End Function













Function sourceWorkbookからtargetWorkbookにmoduleNameを流し込む(ByVal sourceWorkbook As Workbook, ByVal targetWorkbook As Workbook, ByVal moduleName As String, ByVal macroName As String, ByVal ONLY_DELETE) As Boolean
    ' moduleNameには、追加するモジュールに含まれるマクロ名を指定。
    ' モジュールを追加するだけ（ONLY_ADD=TRUE）なら、targetWorkbookに既に含まれているかの確認に使うため、moduleNameに含まれるマクロ名ならなんでもいいが、Subの方で！！！！！！！！！！！！！
        
    On Error GoTo ErrorHandler
    Debug.Print "Debug   Start  sourceWorkbookからtargetWorkbookにmoduleNameを流し込む"
    sourceWorkbookからtargetWorkbookにmoduleNameを流し込む = True
            
    Dim sourceModule As Object
    Dim targetModule As Object
    Dim vbComp As VBIDE.vbComponent
        
    If sourceWorkbook Is Nothing Then Call Fin("ブックが開けませんでした。パスの異なる同じ名前のブックが既に開かれてる可能性があります。", 3)
    If targetWorkbook Is Nothing Then Call Fin("ブックが開けませんでした。パスの異なる同じ名前のブックが既に開かれてる可能性があります。", 3)
     
'    For Each vbComp In sourceWorkbook.VBProject.VBComponents
'        Debug.Print "Debug   vbComp.name =  " & vbComp.Name & "     vbComp.Type: " & vbComp.Type
'    Next vbComp
    Set sourceModule = sourceWorkbook.VBProject.VBComponents(moduleName) ' モジュール名
    
    targetWorkbook.Windows(1).WindowState = xlMaximized
    'マクロmacroNameが、BNAME_SHUKEIに存在したら、削除する
    ' モジュールをループ
    Dim exists As Boolean: exists = False
    For Each vbComp In targetWorkbook.VBProject.VBComponents
        If vbComp.Type = vbext_ct_StdModule Or vbComp.Type = vbext_ct_ClassModule Then
            ' モジュールが空でない場合のみ確認
            ' Debug.Print "Debug  モジュール名 vbComp.CodeModule.Lines(1, vbComp.CodeModule.CountOfLines) = " & vbComp.CodeModule.Lines(1, vbComp.CodeModule.CountOfLines)
            If vbComp.CodeModule.CountOfLines > 0 Then
                ' モジュールのコードを確認
                If InStr(vbComp.CodeModule.Lines(1, vbComp.CodeModule.CountOfLines), "Sub " & macroName & "(") > 0 Then
                    exists = True
                    ' モジュールを削除
                    targetWorkbook.VBProject.VBComponents.Remove vbComp
                    Debug.Print "Debug   モジュールが既に存在したので、削除しました！！！！ " & moduleName & "  " & macroName
                    Exit For
                End If
            End If
        End If
    Next vbComp
    
    If ONLY_DELETE = True Then Exit Function
        
    If exists Then
        MsgBox "マクロ 「" & macroName & "」 が含まれるモジュール[" & moduleName & "] は " & vbCrLf & targetWorkbook.Name & " に存在したので、一旦、モジュールを削除して、" & vbCrLf & "マクロを流し込みます。。", Buttons:=vbInformation
    Else
        MsgBox "マクロ 「" & macroName & "」 が含まれるモジュール[" & moduleName & "] を " & vbCrLf & targetWorkbook.Name & vbCrLf & "に流し込みます。", Buttons:=vbInformation
    End If
    
    Set targetModule = targetWorkbook.VBProject.VBComponents.Add(1) ' vbext_ct_StdModule = 1  標準モジュールを追加
    targetModule.CodeModule.AddFromString sourceModule.CodeModule.Lines(1, sourceModule.CodeModule.CountOfLines)
    Debug.Print "Debug   targetWorkbookに、[" & moduleName & "]を追加しました！"
    
    Debug.Print "Debug   Function sourceWorkbookからtargetWorkbookにmoduleNameを流し込む    が  終了"

    Exit Function ' 通常の処理が完了したらエラーハンドラをスキップ
ErrorHandler:
    MsgBox "エラーです。内容は　 " & Err.Description, Buttons:=vbCritical
    sourceWorkbookからtargetWorkbookにmoduleNameを流し込む = False
    
End Function



