Attribute VB_Name = "Module13"
Option Explicit

'==============================================================================================================================
Sub Button14()
'    Call RunBatchFile("C:\Users\kenichi\Documents\operation_log_NEW\vscode_operation_log.bat")
    Call RunBatchFile("C:\Users\kenic\Dropbox\gitdir\vscode_open\vscode_open.bat C:\Users\kenic\Documents\operation_log_NEW")
End Sub



Sub ボタン15_Click()
    Dim wb_MATOME As Workbook    ' ちゃんと宣言しないと、関数SheetExistsの引数が異なると怒られる
    Set wb_MATOME = OpenBook(BNAME_MATOME, False) ' フルパスを指定
    If wb_MATOME Is Nothing Then Call Fin("ブックが開けませんでした。パスの異なる同じ名前のブックが既に開かれてる可能性があります。", 3)
        
    Dim Sonzai_flg As Boolean: Sonzai_flg = False
    Sonzai_flg = SheetExists(wb_MATOME, "まとめ ")
    If Not Sonzai_flg Then
        MsgBox "シートが存在しません。" & vbCrLf & " 終了します。", Buttons:=vbExclamation
    Else
        Call 適切な箇所に改ページを入れるVer2(wb_MATOME.Worksheets("まとめ "))
    End If
End Sub



Sub ボタン16_Click()
    Dim wb_MATOME As Workbook    ' ちゃんと宣言しないと、関数SheetExistsの引数が異なると怒られる
    Set wb_MATOME = OpenBook(BNAME_MATOME, False) ' フルパスを指定
    If wb_MATOME Is Nothing Then Call Fin("ブックが開けませんでした。パスの異なる同じ名前のブックが既に開かれてる可能性があります。", 3)
    
    Dim Sonzai_flg As Boolean: Sonzai_flg = False
    Sonzai_flg = SheetExists(wb_MATOME, "Fault集計")
    If Not Sonzai_flg Then
        MsgBox "シートが存在しません。" & vbCrLf & " 終了します。", Buttons:=vbExclamation
    Else
        Call 適切な箇所に改ページを入れるVer2(wb_MATOME.Worksheets("Fault集計"))
    End If
End Sub





Sub ボタン18_Click() ' TEST!!!!!!!!!!!!
'    Dim wb_MATOME As Workbook    '　ちゃんと宣言しないと、関数SheetExistsの引数が異なると怒られる
'    Set wb_MATOME = OpenBook("\\saclaopr18.spring8.or.jp\common\運転状況集計\最新\SACLA\SACLA運転状況集計まとめ.xlsm") ' フルパスを指定
'    If wb_MATOME Is Nothing Then Call Fin("ブックが開けませんでした。パスの異なる同じ名前のブックが既に開かれてる可能性があります。",3)
    
    
'     Dim BL As Integer: BL = 2
'     Select Case BL
'        Case 1
'            Debug.Print "SCSS+"
'        Case 2
'            Debug.Print "BL2"
'            BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\運転状況集計\最新\SACLA\SACLA運転状況集計BL2.xlsm"
'            SNAME_KEIKAKU_BL = "bl2"
'        Case 3
'            Debug.Print ">>>BL3"
'            BNAME_SHUKEI = "\\saclaopr18.spring8.or.jp\common\運転状況集計\最新\SACLA\SACLA運転状況集計BL3.xlsm"
'            SNAME_KEIKAKU_BL = "bl3"
'        Case Else
'            MsgBox "BLが不正です。終了します。", vbCritical
'            Exit Sub
'    End Select
'
'    ' sourceWorkbookを開く
'    Dim sourceWorkbook As Workbook    ' ちゃんと宣言しないと、関数SheetExistsの引数が異なると怒られる
'    Set sourceWorkbook = OpenBook(BNAME_SOURCE) ' フルパスを指定
'    If sourceWorkbook Is Nothing Then Call Fin("ブックが開けませんでした。パスの異なる同じ名前のブックが既に開かれてる可能性があります。",3)
'
'
'    ' wb_SHUKEIを開く
'    Dim wb_SHUKEI As Workbook    ' ちゃんと宣言しないと、関数SheetExistsの引数が異なると怒られる
'    Set wb_SHUKEI = OpenBook(BNAME_SHUKEI) ' フルパスを指定
'    If wb_SHUKEI Is Nothing Then Call Fin("ブックが開けませんでした。パスの異なる同じ名前のブックが既に開かれてる可能性があります。",3)
'
'    Dim result As Boolean
'    Dim macroName As String: macroName = "Cleanup" ' とりあえず、Sub Cleanupが含まれるモジュールを削除する。要検討改修
'    result = sourceWorkbookからtargetWorkbookにmoduleNameを流し込む(sourceWorkbook, wb_SHUKEI, "Module8", "Cleanup", False)
'    If result Then
'        MsgBox "成功 「マクロいろいろxlsmからSACLA運転状況集計BLxlsmにマクロを流し込むだけ」", Buttons:=vbInformation
'    Else
'        MsgBox "失敗 「マクロいろいろxlsmからSACLA運転状況集計BLxlsmにマクロを流し込むだけ」", Buttons:=vbInformation
'    End If
    
    
'    Dim folderPath As String
'    folderPath = GetWorkbookFolder()
'    If folderPath = "" Then
'        MsgBox "このブックはまだ保存されていません。"
'    Else
'        MsgBox "このブックが開かれているフォルダのパス: " & folderPath
'    End If
    
    Call ToggleButton
    
End Sub



' ========================================================================================================================
Sub PDF_output_Click()
    Debug.Print "PDF_output_Click"
   
    Dim myArray(2) As Variant
    Dim i As Integer
    Dim pdfPath As String
    Dim OutDir As String
    OutDir = "PDF作成"

    Dim wb_MATOME As Workbook    ' ちゃんと宣言しないと、関数SheetExistsの引数が異なると怒られる
    Set wb_MATOME = OpenBook(BNAME_MATOME, False) ' フルパスを指定
    If wb_MATOME Is Nothing Then Call Fin("ブックが開けませんでした。パスの異なる同じ名前のブックが既に開かれてる可能性があります。", 3)
        
    Dim sheet As Worksheet
    myArray(0) = "まとめ " 'シート名　まとめ シートには半角スペースがあるので注意
    myArray(1) = "Fault集計"
    myArray(2) = ThisWorkbook.sheetS("手順").Range(UNITNAME & UNITROW)
    
    For i = LBound(myArray) To UBound(myArray)
'        MsgBox "要素 " & i & ": " & myArray(i)
        Set sheet = wb_MATOME.Worksheets(myArray(i))
'        sheet.PrintPreview
        
        pdfPath = CPATH & WHICH & "\" & OutDir & "\" & WHICH & "運転状況集計(" & Replace(myArray(i), " ", "") & ").pdf"
        Debug.Print "pdfPath:   " & pdfPath
        ' シートをPDFとしてエクスポート
        sheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfPath, Quality:=xlQualityStandard, _
                              IncludeDocProperties:=True, IgnorePrintAreas:=False, _
                              OpenAfterPublish:=False
                              '  IgnorePrintAreas: Falseの場合、設定された印刷エリアのみがPDFにエクスポートされます。

        ' PDFを開く
        pdfPath = CPATH & WHICH & "\" & OutDir & "\" & WHICH & "運転状況集計(" & Replace(myArray(i), " ", "") & ").pdf"
        shell """" & edgePath & """ --new-window """ & pdfPath & """", vbNormalFocus
'         shell """" & edgePath & """ --start-maximized """ & pdfPath & """", vbNormalFocus      [--start-maximized]オプションつけても最大化されず
        MsgBox "運転状況集計(" & myArray(i) & ").pdf" & vbCrLf & "を出力しました。", vbInformation
'        Exit For
    Next i
    
End Sub



