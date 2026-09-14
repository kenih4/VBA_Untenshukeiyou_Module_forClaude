Attribute VB_Name = "Module8"
Option Explicit ' 未定義の変数は使用できないように
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const UNITROW As Integer = 21
Public Const UNITNAME As String = "D"
Public Const BEGIN_COL As String = "E"
Public Const END_COL As String = "G"
Public Const CPATH As String = "\\saclaopr18.spring8.or.jp\common\運転状況集計\最新\"
Public Const WHICH As String = "SACLA"
Public Const BNAME_UNTENSHUKEIKIROKU_SACLA As String = CPATH & WHICH & "\SACLA運転集計記録.xlsm"
Public Const BNAME_KEIKAKU As String = CPATH & "計画時間.xlsx"
Public Const BNAME_SOURCE As String = "C:\me\unten\マクロいろいろ.xlsm"
Public Const OperationSummaryDir As String = "C:\me\unten\OperationSummary"
Public Const IcalDir As String = "C:\Users\kenic\Dropbox\gitdir\ical_to_ganto"
Public Const BNAME_MATOME As String = CPATH & WHICH & "\SACLA運転状況集計まとめ.xlsm"

Public Const edgePath As String = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

Public Const TARGET_PATH  As String = "\\saclaoprfs01.spring8.or.jp\log_note\SACLA\operation_log"
Public Const DIST_PATH  As String = "C:\Users\kenic\Documents\operation_log_NEW\SACLA"


'ブックを開く ==============================================================================================================================
'リモートサーバー上のファイルを開く際、開くのに時間がかかっているためにタイミングの問題でエラーが発生している可能性があります。この場合、待機時間を設けて再試行することで、エラーを回避できることがあります。以下の方法で、指定された時間待機しながらエラーを再試行するコードを実装できます。
Function OpenBook(ByVal WorkBookName As String, ByVal RO As Boolean) As Workbook
    
    Debug.Print "Debug---   Start  Function OpenBook(" & WorkBookName & ")"
    Dim OWB As Workbook
    Dim wb As Workbook
    Dim retryCount As Integer
    retryCount = 1  ' 再試行の回数

    ' 開いているブックの中に指定されたパスのブックがあるかを確認
    For Each wb In Workbooks
        'Debug.Print "Debug   wb.Name =  " & wb.Name & " は開かれています"
        If wb.FullName = WorkBookName Then
            Set OWB = wb
            Debug.Print "Debug---   OpenBook.Name =  [" & OWB.Name & "] は既に開かれています"
            Exit For
       End If
    Next wb

    On Error Resume Next
    If OWB Is Nothing Then
        Do While retryCount > 0
            Set OWB = Workbooks.Open(WorkBookName, ReadOnly:=RO)
            If Err.Number = 0 Then Exit Do  ' 正常に開けたらループを抜ける
            Debug.Print "Debug--- エラーが発生しました。再試行します。残り再試行回数: " & retryCount - 1
            Err.Clear
            retryCount = retryCount - 1
            Application.Wait Now + TimeValue("0:00:05")  ' 5秒待機
        Loop
        
        ' 最後にエラーが残っている場合の対応
        If Err.Number <> 0 Then
            MsgBox "ブックが見つからないか、開けませんでした。エラー番号: " & Err.Number & vbCrLf & _
                   "エラーメッセージ: " & Err.Description & vbCrLf & _
                   "ファイル名やパスを確認してください: " & WorkBookName, vbExclamation
            Set OWB = Nothing
            Err.Clear
        Else
            Debug.Print "Debug---   OpenBook.Name =  [" & OWB.Name & "] を開きました"
        End If
    End If
    On Error GoTo 0  ' エラーハンドリング解除

    Set OpenBook = OWB
    
    Debug.Print "Debug---   Finish  Function OpenBook(" & WorkBookName & ")"
End Function









' 使ってない
Function OpenBookOLD(ByVal WorkBookName As String) As Workbook
    Debug.Print "Debug   ブックを開きます。-----------  " & WorkBookName
    Dim OWB As Workbook
    Dim wb As Workbook

    ' 開いているブックの中に指定されたパスのブックがあるかを確認
    For Each wb In Workbooks
        'Debug.Print "Debug   wb.Name =  " & wb.Name & " は開かれています"
        If wb.FullName = WorkBookName Then
            Set OWB = wb
            Debug.Print "Debug   OpenBook.Name =  [" & OWB.Name & "] は既に開かれています"
            Exit For
        End If
    Next wb

    ' エラーハンドリング開始
    On Error Resume Next
    If OWB Is Nothing Then
        ' 指定したブックが開かれていない場合、新たに開こうとする
        Set OWB = Workbooks.Open(WorkBookName, ReadOnly:=False)    ' SACLA運転状況集計BL*.xlsm　を開こうとすると、なぜかエラーが発生するので以下コメントアウトした
        If Err.Number <> 0 Then
            ' エラーが発生した場合、エラーメッセージを表示
            MsgBox "ブックが見つからないか、開けませんでした。エラー番号: " & Err.Number & vbCrLf & _
                   "エラーメッセージ: " & Err.Description & vbCrLf & _
                   "ファイル名やパスを確認してください: " & WorkBookName, vbExclamation
            Set OWB = Nothing  ' エラー発生時は Nothing を返す
            Err.Clear
        Else
            Debug.Print "Debug   OpenBook.Name =  [" & OWB.Name & "] を開きました"
        End If
        Debug.Print "Debug   OpenBook.Name =  [" & OWB.Name & "] を開きました   開けていない可能性あり　エラー処理をパスしてるので"
    End If
    On Error GoTo 0  ' エラーハンドリング解除

    ' 関数の戻り値として設定
    Set OpenBookOLD = OWB

    Debug.Print "Debug   OpenBook Finish"
End Function





'========================================================================================================
Sub CMsg(ByVal msg As String, ByVal Level As Integer, tc As Variant)

    Debug.Print "_____Msg(" & msg & ")_____"

    tc.Select
    tc.Font.Bold = True
    Select Case Level
    Case vbInformation
        tc.Font.Color = RGB(0, 200, 0)
        tc.Interior.Color = RGB(0, 255, 255)
        MsgBox msg, vbInformation, "お知らせ"
    Case vbExclamation
        tc.Interior.Color = RGB(255, 255, 0)
        MsgBox msg, vbExclamation, "注意"
    Case vbCritical
        tc.Interior.Color = RGB(255, 0, 0)
        MsgBox msg, vbCritical, "警告"
    Case Else
        Debug.Print "Zzz..."
    End Select
    
End Sub


'========================================================================================================
Sub Fin(ByVal msg As String, ByVal Level As Integer)

    Debug.Print "_____Fin(" & msg & ")_____"
    Select Case Level
        Case 1
            MsgBox msg, vbInformation, "終了処理"
        Case 2
            MsgBox msg, vbExclamation, "終了処理"
        Case 3
            MsgBox msg, vbCritical, "終了処理"
        Case Else
            Debug.Print "Zzz..."
    End Select
    
'    ActiveWindow.Zoom = 100
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
    Application.DisplayFullScreen = False
    ' 開いているすべてのブックをループ
    Dim wb As Workbook
    For Each wb In Workbooks
        wb.Windows(1).Zoom = 100 ' 各ブックのウィンドウに対してズームを設定
    Next wb
'    End   これいる？？？
End Sub





'----------------------------------------------------------------------------------------------------------------------
'シート内のエラーセルを検出し、メッセージを表示する
Function CheckForErrors(ByVal sheet As Worksheet) As Boolean
  Dim cell As Range
  Dim errorRange As Range
  CheckForErrors = True
  
  If sheet Is Nothing Then
    MsgBox "' のシート '" & sheet & "' は存在しません。", vbOKOnly + vbCritical
    Exit Function
  End If
  sheet.Activate
  
  For Each cell In sheet.UsedRange
    'Debug.Print "Debug  Value =  " & cell.Value & "  Row = " & cell.Row & " Columuns = " & cell.Column
    If IsError(cell.Value) Then
      ' 最初のエラーセルであれば、errorRangeに設定
      If errorRange Is Nothing Then
        Set errorRange = cell
      Else
        ' 2つ目以降のエラーセルであれば、errorRangeに追加
        Set errorRange = Union(errorRange, cell)
        cell.Select
      End If
    End If
  Next cell

  ' エラーセルが見つかった場合、メッセージを表示
  If Not errorRange Is Nothing Then
        MsgBox "シート '" & sheet.Name & "' にエラーセルがあります。＜注意＞年度初め、まだ真っ新の場合、沢山エラーセルがあるのでエラーセルがあると警告されますが無視していいです。" & vbCrLf & "エラーセル: " & errorRange.Address, vbOKOnly + vbCritical
  Else
'    MsgBox "安心です。シート '" & sheet.Name & "' にエラーセルはありませんでした。", vbOKOnly + vbInformation
        Debug.Print "安心です。シート '" & sheet.Name & "' にエラーセルはありませんでした。"
        CheckForErrors = False
  End If

  Set errorRange = Nothing
End Function




'指定さてた文字列が存在する行を取得  シート内全て==============================================================================================================================
Function getLineNum(ByVal str As String, ByVal TARGET_COL As Integer, ByVal sheet As Worksheet) As Integer
    getLineNum = getLineNum_RS(str, TARGET_COL, 1, sheet.Cells(Rows.Count, TARGET_COL).End(xlUp).Row, sheet)
End Function


'指定さてた文字列が存在する行を取得 Range Specification版==============================================================================================================================
Function getLineNum_RS(ByVal str As String, ByVal TARGET_COL As Integer, ByVal beginLine As Integer, ByVal endLine As Integer, ByVal sheet As Worksheet) As Integer
    Dim i As Integer: i = -1
    getLineNum_RS = i
    For i = beginLine To endLine
        'Debug.Print "getLineNum_RS　行番号: " & i & "    Value: " & Cells(i, 2).Value
        If sheet.Cells(i, TARGET_COL).Value = str Then ' #DIV/0!などのエラーセルがあると、正しく途中で止まります。
            getLineNum_RS = i
            Debug.Print "Hit!!!!!!!!!!!!!!!!!!!!!!!!!   getLineNum_RS　行番号: " & i & "    Value: " & Cells(i, 2).Value
            Exit Function
        End If
    Next
'    Call Fin("@getLineNum_RS    文字列「" & str & "」と一致するセルは見つかりませんでした。", 3)
    MsgBox "@getLineNum_RS    文字列「" & str & "」と一致するセルは見つかりませんでした。", vbExclamation, "警告"
End Function




'シート存在を確認==============================================================================================================================
Function SheetExists(wb As Workbook, sname As String) As Boolean
    On Error Resume Next ' エラーが発生しても処理を継続
    Dim ws As Worksheet
    Set ws = wb.sheetS(sname) ' 指定したシートをセット
    SheetExists = Not ws Is Nothing ' シートが存在すればTrue
    Debug.Print "@SheetExists   Sheetname: [" & sname & "]  " & SheetExists
    On Error GoTo 0 ' エラーハンドリングをリセット
End Function






'ActiveWorkbookシート存在を確認 Not Use ==============================================================================================================================
Function SheetExist_ActiveWorkbook(ByVal WorkSheetName As String) As Boolean
  Dim sht As Worksheet
  For Each sht In ActiveWorkbook.Worksheets
    If sht.Name = WorkSheetName Then
        SheetExist_ActiveWorkbook = True
        Exit Function
    End If
  Next sht
  SheetExist_ActiveWorkbook = False
End Function




'==============================================================================================================================
Sub RunBatchFile(batchFilePath As String)

    ' バッチファイルのパスが指定されているか確認
    If batchFilePath = "" Then
        MsgBox "バッチファイルのパスを指定してください", vbExclamation
        Exit Sub
    End If
    
    ' Shell関数でバッチファイルを実行
    shell batchFilePath, vbNormalFocus
End Sub




'エクセルブックが開かれたフォルダを取得==============================================================================================================================
Function GetWorkbookFolder() As String
    Dim folderPath As String
    
    ' ブックが保存されていない場合、Path は空文字列になる
    folderPath = ThisWorkbook.path
    
    ' 保存されていない場合、空文字列を返す
    If folderPath = "" Then
        GetWorkbookFolder = "" ' 空文字列を返す
    Else
        GetWorkbookFolder = folderPath ' フォルダパスを返す
    End If
End Function


Sub GetWorkbookFolderToCell()
' ThisWorkbook.Path でカレントブックの保存されているパスを取得
    Dim folderPath As String
    folderPath = ThisWorkbook.path
'    MsgBox folderPath
    
    If folderPath <> "" Then
        ThisWorkbook.sheetS("手順").Range("B1").Value = folderPath
        
'        MsgBox folderPath
        If folderPath = "C:\me\unten" Then
            MsgBox "OK: " & vbCrLf & "ワーキングフォルダ = " & folderPath, Buttons:=vbInformation
        Else
            MsgBox "チェック: " & vbCrLf & "ワーキングフォルダ = " & folderPath & vbCrLf & "ワーキングフォルダが「C:\me\unten」でありません！！", Buttons:=vbInformation
        End If
        
    Else
        ThisWorkbook.sheetS(1).Range("A1").Value = "ワーキングフォルダが取得できませんでした"
        MsgBox "異常: " & vbCrLf & "ワーキングフォルダが取得できませんでした", Buttons:=vbCritical
    End If
End Sub






' 循環参照を検出
Sub CheckCircularReference()
    Application.Calculate ' 先に計算を実行  Application.CircularReference は、計算後に値を返すため、計算がまだ実行されていない場合には使えない為

    On Error Resume Next ' エラーを無視
    Dim circRef As Range
    Set circRef = Application.CircularReference
    On Error GoTo 0 ' エラー処理を戻す

    If circRef Is Nothing Then
'        MsgBox "循環参照は見つかりませんでした。", vbInformation
    Else
        MsgBox "循環参照が見つかりました: " & circRef.Address, vbExclamation
    End If
End Sub







Sub ToggleButton()    '---------------------------------------------------------------------------------
' ボタンの外観を変更する
    If ActiveSheet.Shapes("Button 18").Fill.ForeColor.RGB = RGB(255, 255, 255) Then
        ActiveSheet.Shapes("Button 18").Fill.ForeColor.RGB = RGB(0, 0, 0)  ' 黒に変更
        ActiveSheet.Shapes("Button 18").TextFrame.Characters.Text = "押し込み中"
    Else
        ActiveSheet.Shapes("Button 18").Fill.ForeColor.RGB = RGB(255, 255, 255)  ' 白に戻す
        ActiveSheet.Shapes("Button 18").TextFrame.Characters.Text = "押してください"
    End If
End Sub






' シートに文字列が存在するか確認する
Function CheckStringInSheet(ByVal ws As Worksheet, ByVal searchString As String) As Boolean
    Dim foundCell As Range

    Set foundCell = ws.Cells.Find(What:=searchString, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    
    If foundCell Is Nothing Then
        CheckStringInSheet = False
    Else
        CheckStringInSheet = True
    End If
    
End Function







Function Check_checkbox_status(obj_name) As Boolean
    Dim chk As Shape
    Check_checkbox_status = False
    For Each chk In ActiveSheet.Shapes
'        Debug.Print chk.Name
        If chk.Type = msoFormControl Then
            If chk.FormControlType = xlCheckBox Then
                Debug.Print "Checked checkbox:  " & chk.Name
                If chk.Name = obj_name Then
                    Debug.Print "Checked!  True @Check_checkbox_status"
                    'chk.OLEFormat.Object.Value = xlOff
                    Check_checkbox_status = True
                End If
            Else
                'Debug.Print "Checked checkbox:  " & chk.Name
            End If
        End If
    Next chk
End Function





'=== git管理したい為、モジュールごとに別々のテキストファイルにエクスポートする ===
Sub ExportModulesToSeparateTextFiles()
    Dim vbComp As Object
    Dim filePath As String
    Dim fileNum As Integer
    Dim moduleContent As String
    
    ' 出力ファイルのパスを指定（同じフォルダに保存）
'    filePath = ThisWorkbook.path & "\"
    filePath = "C:\Users\kenic\Dropbox\gitdir\VBA_Untenshukeiyou_ModuleText\"
    
    ' 各モジュールをループして内容を取得
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        ' モジュールの行数を取得
        Dim lineCount As Long
        lineCount = vbComp.CodeModule.CountOfLines
        
        ' モジュールが空でない場合のみ内容を取得
        If lineCount > 0 Then
            ' モジュールの内容を取得
            moduleContent = vbComp.CodeModule.Lines(1, lineCount)
            
            ' モジュール名に基づいてファイル名を作成
            Dim moduleFileName As String
            moduleFileName = filePath & vbComp.Name & ".vba"
            
            ' 新しいテキストファイルを作成
            fileNum = FreeFile
            Open moduleFileName For Output As #fileNum
            
            ' モジュール名とその内容をファイルに書き込む
            Print #fileNum, "Module: " & vbComp.Name
            Print #fileNum, moduleContent
            
            ' ファイルを閉じる
            Close #fileNum
        Else
            ' 空のモジュールの場合はスキップ
            Debug.Print "モジュール " & vbComp.Name & " は空です。"
        End If
    Next vbComp
     
    
    ThisWorkbook.SaveCopyAs "C:\Users\kenic\Dropbox\gitdir\VBA_Untenshukeiyou_ModuleText" & "\マクロいろいろ_" & Format(Date, "yyyymmdd") & ".xlsm"
    
    If MsgBox("すべてのモジュールがテキストにエクスポートされました。" & vbCrLf & filePath & vbCrLf & "vscodeを開きますか？" & vbCrLf & "git add -A" & vbCrLf & "git commit -m comment", vbYesNo + vbQuestion, "確認") = vbNo Then
        MsgBox "No"
    Else
        Dim vscodePath As String
        Dim folderPath As String
        Dim Command As String
        vscodePath = "C:\Users\kenic\AppData\Local\Programs\Microsoft VS Code\Code.exe"
        folderPath = "C:\Users\kenic\Dropbox\gitdir\VBA_Untenshukeiyou_ModuleText"
        Command = """" & vscodePath & """ """ & folderPath & """"
        shell Command, vbNormalFocus
    End If
        
End Sub





'=== フルパスで書かれたファイル名の拡張子だけを取り除きたい ================================
Function RemoveFileExtension(fullPath As String) As String
    Dim fileName As String
    Dim dotPosition As Long
    
    ' フルパスからファイル名を取得
    'fileName = Dir(fullPath)
    fileName = fullPath
    
    ' 最後のドットの位置を取得
    dotPosition = InStrRev(fileName, ".")
    
    ' ドットが見つかった場合、拡張子を取り除く
    If dotPosition > 0 Then
        RemoveFileExtension = Left(fileName, dotPosition - 1)
    Else
        ' ドットが見つからない場合は、元のファイル名を返す
        RemoveFileExtension = fileName
    End If
End Function





'=== 値が入ってる最終行を取得したい ================================
'     最終行番号（例えば1048576行目）を返す可能性がある関数には、Long でないと オーバーフローの危険があります。
Function GetLastDataRow(ws As Worksheet, colName As String) As Long
    Dim i As Long
    Dim lastRow As Long
    Dim colNum As Long
    colNum = ws.Range(colName & "1").Column
    lastRow = ws.Cells(ws.Rows.Count, colNum).End(xlUp).Row

    For i = lastRow To 1 Step -1
'        If ws.Cells(i, colNum).HasFormula Then
            If ws.Cells(i, colNum) = "" Then
'                Debug.Print i & "   KARA  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ " & ws.Cells(i, colNum).Value
            Else
'                Debug.Print i & "   Not KARA  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ " & ws.Cells(i, colNum).Value
                GetLastDataRow = i ' 値が入ってる場合、ココだ
                Exit Function
            End If
'        Else
            'Debug.Print i & "   Not HasFormula  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ " & ws.Cells(i, colNum).Value
'        End If
    Next i

    GetLastDataRow = 0 ' 該当セルが無い場合
End Function





'=== 2つのセルが一致するのかしないのか。日時の比較を行う際には、非常に小さい差を許容する方法が有効 ================================
' 日付の値、セルA1には「45769.66667」、セルA2にも同じく「45769.66667」が入っているのですが、一致しません。疑問に思って、試しに、セルA2 の セルA1の差分をとったら、「-7.27695761418343E-12」となりました。
Function CheckCellsMatch(cell1 As Range, cell2 As Range) As Boolean
    Dim tolerance As Double

    ' 許容誤差を設定
    tolerance = 0.0000000001 ' 10の-10乗より小さい差は無視

    If Abs(cell1.Value - cell2.Value) < tolerance Then
        CheckCellsMatch = True ' 一致している場合
    Else
        CheckCellsMatch = False ' 一致していない場合
    End If
End Function

'=== 2つの値が一致するのかしないのか。日時の比較を行う際には、非常に小さい差を許容する方法が有効 ================================
Function CheckValMatch(v1 As Variant, v2 As Variant) As Boolean
    Dim tolerance As Double

    ' 許容誤差を設定
    tolerance = 0.0000000001 ' 10の-10乗より小さい差は無視

    If Abs(v1 - v2) < tolerance Then
        CheckValMatch = True ' 一致している場合
    Else
        CheckValMatch = False ' 一致していない場合
    End If
End Function



'=== 範囲内のすべての重複値を検出して、まとめて警告メッセージで表示 ======================================
Sub CheckAllDuplicatesByRange(targetRange As Range)
    Dim cell As Range
    Dim dict As Object
    Dim duplicates As Collection
    Set dict = CreateObject("Scripting.Dictionary")
    Set duplicates = New Collection
    targetRange.Select
    'ActiveWindow.Zoom = True ' 選択したセルの範囲が画面に最適化
    Call ZoomToSelectionWithMaxLimit
    
    For Each cell In targetRange
        If Not IsEmpty(cell.Value) Then
            If dict.exists(cell.Value) Then
                duplicates.Add cell
            Else
                dict.Add cell.Value, cell
            End If
        End If
    Next cell
    
    If duplicates.Count > 0 Then
        Dim msg As String
        msg = "以下の重複があります:" & vbCrLf
        
        Dim dupCell As Range
        For Each dupCell In duplicates
            msg = msg & dupCell.Address(False, False) & ": " & dupCell.Value & vbCrLf
            dupCell.Interior.Color = RGB(255, 200, 200)
        Next dupCell
        
        MsgBox msg, vbCritical, "重複検出結果"
    Else
        MsgBox "重複はありませんでした", vbInformation, "確認完了"
    End If
    ActiveWindow.Zoom = 100
End Sub




'=== 新規フォルダ作成 ======================================
Function CreateFolderWithAutoRename(parentPath As String, newFolderName As String) As String
    Dim fullPath As String
    Dim counter As Long
    Dim candidateName As String

    ' 最初の候補
    candidateName = newFolderName
    fullPath = parentPath & "\" & candidateName

    ' フォルダが存在する限り、連番を付けて試す
    Do While Dir(fullPath, vbDirectory) <> ""
        counter = counter + 1
        candidateName = newFolderName & "_(" & counter & ")"
        fullPath = parentPath & "\" & candidateName
    Loop

    On Error GoTo ErrorHandler
    MkDir fullPath
    CreateFolderWithAutoRename = fullPath
    Exit Function

ErrorHandler:
    CreateFolderWithAutoRename = vbNullString  ' 作成失敗時は空文字を返す
End Function



'=== ファイルコピー ======================================
Function CopyFileSafely(sourcePath As String, destinationPath As String) As Boolean
    On Error GoTo ErrorHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' コピー実行（Trueで上書き許可）
    fso.CopyFile sourcePath, destinationPath, True
    CopyFileSafely = True
    Exit Function

ErrorHandler:
    CopyFileSafely = False
End Function


'=== 読み取り専用属性を設定 ======================================
Function SetFileReadOnly(filePath As String) As Boolean
    On Error GoTo ErrorHandler

    If Dir(filePath, vbNormal) = "" Then
        SetFileReadOnly = False  ' ファイルが存在しない
        Exit Function
    End If

    SetAttr filePath, vbReadOnly
    SetFileReadOnly = True
    Exit Function

ErrorHandler:
    SetFileReadOnly = False
End Function


'=== セルが空欄かどうか判定 ======================================
Function IsBlankCell(rng As Range) As Boolean
    IsBlankCell = Trim(rng.Value) = ""
End Function




'=== 実行したいBashコマンドを引数 CommandsToRun で受け取るプロシージャ ======================================
Sub ExecuteGitBashCommand(ByVal CommandsToRun As String)
   
    Dim GitBashPath As String
    Dim FullCommand As String
    
    GitBashPath = "C:\Program Files\Git\bin\bash.exe"
    
    ' Git Bashに渡す完全なコマンド文字列を作成します
    ' --login -c の後に、引数で渡されたコマンド文字列を二重引用符で囲んで渡します
    FullCommand = GitBashPath & " --login -c """ & CommandsToRun & """"
    
    On Error GoTo ErrorHandler
    
    ' シェルを実行します (ウィンドウを通常サイズで表示: 1)
    ' 0: 非表示, 1: 通常表示, 2: 最小化, 3: 最大化
    shell FullCommand, 1
    
    Exit Sub

ErrorHandler:
    MsgBox "Git Bashの実行に失敗しました。パスを確認してください。", vbCritical
End Sub


'=== ズーム値が100を超えないように制御 ======================================
Sub ZoomToSelectionWithMaxLimit()
    Dim rng As Range
    Dim tempZoom As Variant

    ' 一時的に選択範囲をズーム
    ActiveWindow.Zoom = True

    ' ズーム値を取得
    tempZoom = ActiveWindow.Zoom

    ' 最大ズーム値を100に制限
    If tempZoom > 100 Then
        ActiveWindow.Zoom = 100
    End If
End Sub

