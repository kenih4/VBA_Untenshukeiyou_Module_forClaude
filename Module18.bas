Attribute VB_Name = "Module18"
Option Explicit

'=== git管理用: VBAプロジェクトの各モジュールを、ImportFileで戻せる形式でエクスポート/インポートする ===

' vbext_ComponentType の値を直書き(VBIDEの参照設定を追加しなくて済むように)
Private Const CT_STD_MODULE As Long = 1    ' vbext_ct_StdModule   標準モジュール → .bas
Private Const CT_CLASS_MODULE As Long = 2  ' vbext_ct_ClassModule クラスモジュール → .cls
Private Const CT_MS_FORM As Long = 3       ' vbext_ct_MSForm      ユーザーフォーム → .frm(+.frx)
Private Const CT_DOCUMENT As Long = 100    ' vbext_ct_Document    ThisWorkbook/シート ※インポート不可

Sub ExportModulesForGit()
    Dim exportFolder As String
    exportFolder = "C:\Users\kenic\Dropbox\gitdir\VBA_Untenshukeiyou_Module_Export-exclusive\"

    If Dir(exportFolder, vbDirectory) = "" Then MkDir exportFolder

    ' 前回エクスポート分の残骸(リネーム/削除されたモジュールのファイル)を先に掃除しておく
    CleanExportFolder exportFolder

    Dim vbComp As Object
    Dim ext As String
    Dim skippedDocs As String

    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case CT_STD_MODULE
                ext = ".bas"
            Case CT_CLASS_MODULE
                ext = ".cls"
            Case CT_MS_FORM
                ext = ".frm"   ' .frx(バイナリ側)はExportFileが自動で一緒に書き出す
            Case CT_DOCUMENT
                ' ThisWorkbook/シートモジュールはインポートし直せないので、
                ' コードだけ参考用テキストとして残す(拡張子を変えて誤インポート防止)
                skippedDocs = skippedDocs & vbComp.Name & vbCrLf
                ExportDocumentCodeAsText vbComp, exportFolder
                GoTo NextComp
            Case Else
                GoTo NextComp   ' 想定外の種類は念のためスキップ
        End Select

        vbComp.Export exportFolder & vbComp.Name & ext

NextComp:
    Next vbComp

    Dim msg As String
    msg = "エクスポート完了: " & exportFolder
    If Len(skippedDocs) > 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              "以下はThisWorkbook/シートモジュールのためインポート対象外です。" & vbCrLf & _
              "コードは *.doccode.txt に参考保存したので、戻す時は手動で貼り付けてください:" & _
              vbCrLf & skippedDocs
    End If
    MsgBox msg, vbInformation, "エクスポート結果"
    
    Dim Command As String
    Dim vscodePath As String
    Dim folderPath As String
    If MsgBox("すべてのモジュールがエクスポートされました。" & vbCrLf & exportFolder & vbCrLf & "vscodeを開きますか？" & vbCrLf & "git add -A" & vbCrLf & "git commit -m comment", vbYesNo + vbQuestion, "確認") = vbNo Then
        MsgBox "No"
    Else
        vscodePath = "C:\Users\kenic\AppData\Local\Programs\Microsoft VS Code\Code.exe"
        Command = """" & vscodePath & """ """ & Left(exportFolder, Len(exportFolder) - 1) & """"
        shell Command, vbNormalFocus
    End If
    
End Sub

Private Sub ExportDocumentCodeAsText(vbComp As Object, exportFolder As String)
    Dim fileNum As Integer, lineCount As Long
    lineCount = vbComp.CodeModule.CountOfLines
    If lineCount = 0 Then Exit Sub

    fileNum = FreeFile
    Open exportFolder & vbComp.Name & ".doccode.txt" For Output As #fileNum
    Print #fileNum, vbComp.CodeModule.Lines(1, lineCount)
    Close #fileNum
End Sub

Private Sub CleanExportFolder(exportFolder As String)
    Dim f As String, e As Variant
    For Each e In Array("*.bas", "*.cls", "*.frm", "*.frx", "*.doccode.txt")
        f = Dir(exportFolder & e)
        Do While f <> ""
            Kill exportFolder & f
            f = Dir()
        Loop
    Next e
End Sub

'=== 上でエクスポートしたファイルを元のブックに戻す ===
Sub ImportModulesFromGit()
    Dim importFolder As String
    importFolder = "C:\Users\kenic\Dropbox\gitdir\VBA_Untenshukeiyou_Module_forClaude\"

    Dim vbProj As Object
    Set vbProj = ThisWorkbook.VBProject

    ImportByExtension vbProj, importFolder, "*.bas", CT_STD_MODULE
    ImportByExtension vbProj, importFolder, "*.cls", CT_CLASS_MODULE
    ImportByExtension vbProj, importFolder, "*.frm", CT_MS_FORM

    MsgBox "インポート完了。" & vbCrLf & _
           "ThisWorkbook/シートモジュールの *.doccode.txt は自動反映していないので、" & vbCrLf & _
           "中身を該当モジュールに手動でコピー&ペーストしてください。", vbInformation
End Sub

Private Sub ImportByExtension(vbProj As Object, folder As String, pattern As String, compType As Long)
    Dim f As String, moduleName As String
    Dim existing As Object

    f = Dir(folder & pattern)
    Do While f <> ""
        moduleName = Left(f, InStrRev(f, ".") - 1)

        ' 同名モジュールが既にあると "Module1_1" のような別名でインポートされてしまうため、
        ' 先に既存のものを削除してから入れ直す
        On Error Resume Next
        Set existing = Nothing
        Set existing = vbProj.VBComponents(moduleName)
        On Error GoTo 0
        If Not existing Is Nothing Then
            vbProj.VBComponents.Remove existing
        End If

        vbProj.VBComponents.Import folder & f
        f = Dir()
    Loop
End Sub

