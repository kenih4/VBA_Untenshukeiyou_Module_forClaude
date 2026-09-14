Attribute VB_Name = "Module15"
Option Explicit


Sub SACLA運転集計記録の確認()
'    If Check_checkbox_status("CheckBox_Untenshukeikiroku") = True Then
'        Debug.Print "チェックが入ってたので終了"
'        End
'    Else
'        Debug.Print "チェックが入なかったよ"
'    End If
'    MsgBox "GO"
    Call 運転集計記録_Check("SACLA", "停止時間")
    Call 運転集計記録_Check("SACLA", "調整時間")
End Sub


Sub check_Initial_Check_BL2_Click()
   Call Initial_Check(2)
End Sub

Sub check_Initial_Check_BL3_Click()
   Call Initial_Check(3)
End Sub

Sub Middle_Check_BL2_Click()
   Call Middle_Check(2)
End Sub

Sub Middle_Check_BL3_Click()
   Call Middle_Check(3)
End Sub


Sub セル内容がセルの幅を越えてたら折り返す_BL2_Click()
    セル内容がセルの幅を越えてたら折り返す_BL (2)
End Sub

Sub セル内容がセルの幅を越えてたら折り返す_BL3_Click()
    セル内容がセルの幅を越えてたら折り返す_BL (3)
End Sub






Sub 計画時間xlsxを出力_Click()
    On Error GoTo ErrorHandler
    
    Dim fileNum As Integer

    If MsgBox("出力するユニットは" & vbCrLf & "   「 " & ThisWorkbook.sheetS("手順").Range(UNITNAME & UNITROW) & " 」" & vbCrLf & ThisWorkbook.sheetS("手順").Range(BEGIN_COL & UNITROW) & "    ～　　" & vbCrLf & ThisWorkbook.sheetS("手順").Range(END_COL & UNITROW) & vbCrLf & "開始しますか？", vbYesNo) = vbNo Then Exit Sub

    fileNum = FreeFile
    Open OperationSummaryDir & "\dt_beg.txt" For Output As #fileNum
    Print #fileNum, Format(ThisWorkbook.sheetS("手順").Cells(UNITROW, 5).Value, "yyyy/mm/dd hh:nn");
    Close #fileNum
    
    fileNum = FreeFile
    Open OperationSummaryDir & "\dt_end.txt" For Output As #fileNum
    Print #fileNum, Format(ThisWorkbook.sheetS("手順").Cells(UNITROW, 7).Value, "yyyy/mm/dd hh:nn");
    Close #fileNum
       
    If ThisWorkbook.sheetS("手順").Cells(UNITROW, 7).Value > Now Then
        MsgBox "ユニットの終了日時が未来です。" & vbCrLf & "pythonスクリプトgetGunHvOffTime_LOCALTEST.pyを実行してもエラーになるだけなので、これは実行しません。" & vbCrLf & "icalから取得するスクリプトだけ実行します。"
    Else
        If RunPythonScript("getGunHvOffTime_LOCALTEST.py", OperationSummaryDir) = False Then
            MsgBox "pythonでエラー発生の模様", Buttons:=vbCritical
        End If
    End If
    
    If RunPythonScript("ical.py -u " & IcalDir & "\ical_setting.xlsx " & IcalDir & "\ical_SHISETUCHOUSEI.xlsx", IcalDir) = False Then
        MsgBox "pythonでエラー発生の模様", Buttons:=vbCritical
    End If
    
    Exit Sub ' 通常の処理が完了したらエラーハンドラをスキップ
ErrorHandler:
    MsgBox "エラーです。内容は　 " & Err.Description, Buttons:=vbCritical
    
End Sub

Sub 計画時間xlsx_Check_BL2_Click()
    Call 計画時間xlsx_Check(2)
    Call 計画時間xlsx_GUN_HV_OFF_Check(2)
End Sub

Sub 計画時間xlsx_Check_BL3_Click()
    Call 計画時間xlsx_Check(3)
    Call 計画時間xlsx_GUN_HV_OFF_Check(3)
End Sub








Sub cp_paste_KEIKAKUZIKAN_UNTENZYOKYOSYUKEI_BL2_Click()
    Call cp_paste_KEIKAKUZIKAN_UNTENZYOKYOSYUKEI(2)
End Sub

Sub cp_paste_KEIKAKUZIKAN_UNTENZYOKYOSYUKEI_BL3_Click()
    Call cp_paste_KEIKAKUZIKAN_UNTENZYOKYOSYUKEI(3)
End Sub



Sub faulttxtを出力_BL2_Click()
    If RunPythonScript("getBlFaultSummary_LOCALTEST.py bl2", OperationSummaryDir) = False Then
        MsgBox "pythonでエラー発生の模様", Buttons:=vbCritical
    End If
End Sub

Sub faulttxtを出力_BL3_Click()
    If RunPythonScript("getBlFaultSummary_LOCALTEST.py bl3", OperationSummaryDir) = False Then
        MsgBox "pythonでエラー発生の模様", Buttons:=vbCritical
    End If
End Sub





Sub 利用時間Userに手動入力_BL2_Click()
    Call 利用時間Userに手動入力(2)
End Sub

Sub 利用時間Userに手動入力_BL3_Click()
    Call 利用時間Userに手動入力(3)
End Sub








Sub Fault集計m_BL2_Click()
    Call マクロいろいろxlsmからSACLA運転状況集計BLxlsmにマクロを流し込んで実行(2, "Fault集計m")
    MsgBox "年度初め、1シフト分だけのユーザー運転があった場合、Fault集計(BL*)にFault回数の合計が空欄になるバグがあります。詳細はパワポにまとめた。実害はないのでそのまま。", Buttons:=vbInformation
End Sub

Sub Fault集計m_BL3_Click()
    Call マクロいろいろxlsmからSACLA運転状況集計BLxlsmにマクロを流し込んで実行(3, "Fault集計m")
    MsgBox "年度初め、1シフト分だけのユーザー運転があった場合、Fault集計(BL*)にFault回数の合計が空欄になるバグがあります。詳細はパワポにまとめた。実害はないのでそのまま。", Buttons:=vbInformation
End Sub



Sub 運転集計_形式処理m_BL2_Click()
    Call マクロいろいろxlsmからSACLA運転状況集計BLxlsmにマクロを流し込んで実行(2, "運転集計_形式処理m")
End Sub

Sub 運転集計_形式処理m_BL3_Click()
    Call マクロいろいろxlsmからSACLA運転状況集計BLxlsmにマクロを流し込んで実行(3, "運転集計_形式処理m")
End Sub





Sub 運転状況集計まとめの追編集_BL2_Click()
    Call 運転状況集計まとめの追編集(2)
End Sub

Sub 運転状況集計まとめの追編集_BL3_Click()
    Call 運転状況集計まとめの追編集(3)
End Sub

Sub check_Final_Check_BL2_Click()
    Call Final_Check(2)
End Sub

Sub check_Final_Check_BL3_Click()
    Call Final_Check(3)
End Sub



Sub Clear_Click()
    Dim chk As Shape
    For Each chk In ActiveSheet.Shapes
        Debug.Print chk.Name
        If chk.Type = msoFormControl Then
            If chk.FormControlType = xlCheckBox Then
                chk.OLEFormat.Object.Value = xlOff    ' チェックを外す
            End If
        End If
    Next chk
End Sub



Sub 作成前のバックアップ取得()
    Call バックアップ取得("作成前")
End Sub

Sub 完了のバックアップ取得()
    Call バックアップ取得("作成完了")
End Sub

Sub バックアップ取得(Comment As String)
    Dim result As Boolean
    Dim createdPath As String
    
    If Not CheckServerAccess_FSO("\\saclaopr18.spring8.or.jp\ses-users\jkenichi\BU\") Then
        Exit Sub
    End If
'    createdPath = CreateFolderWithAutoRename("C:\Users\kenic\OneDrive\Desktop\集計のBK", ThisWorkbook.sheetS("手順").Range(UNITNAME & UNITROW) & "_" & Comment)
    createdPath = CreateFolderWithAutoRename("\\saclaopr18.spring8.or.jp\ses-users\jkenichi\BU\集計のBK", ThisWorkbook.sheetS("手順").Range(UNITNAME & UNITROW) & "_" & Comment)
    If createdPath <> "" Then
'        MsgBox "フォルダ作成成功：" & createdPath
        result = CopyFileSafely(CPATH & "\計画時間.xlsx", createdPath & "\計画時間.xlsx")
        result = CopyFileSafely(CPATH & WHICH & "\SACLA運転集計記録.xlsm", createdPath & "\SACLA運転集計記録.xlsm")
        result = CopyFileSafely(CPATH & WHICH & "\SACLA運転状況集計BL2.xlsm", createdPath & "\SACLA運転状況集計BL2.xlsm")
        result = CopyFileSafely(CPATH & WHICH & "\SACLA運転状況集計BL3.xlsm", createdPath & "\SACLA運転状況集計BL3.xlsm")
        result = CopyFileSafely(CPATH & WHICH & "\SACLA運転状況集計まとめ.xlsm", createdPath & "\SACLA運転状況集計まとめ.xlsm")
        
        result = CopyFileSafely(CPATH & "SCSS" & "\SCSS運転集計記録.xlsm", createdPath & "\SCSS運転集計記録.xlsm")
        result = CopyFileSafely(CPATH & "SCSS" & "\SCSS運転集計記録.xlsm", createdPath & "\SCSS運転集計記録BL1.xlsm")
        
        result = CopyFileSafely("C:\me\unten\マクロいろいろ.xlsm", createdPath & "\マクロいろいろ.xlsm")
        
        If Not result Then
            MsgBox "コピー失敗…", vbCritical
        End If
        
'        result = SetFileReadOnly(createdPath & "\SACLA運転集計記録.xlsm")
'        result = SetFileReadOnly(createdPath & "\SACLA運転状況集計BL2.xlsm")
'        result = SetFileReadOnly(createdPath & "\SACLA運転状況集計BL3.xlsm")
'        result = SetFileReadOnly(createdPath & "\SACLA運転状況集計まとめ.xlsm")
    Else
        MsgBox "フォルダ作成に失敗しました。", vbCritical
    End If
    MsgBox Comment & "のバックアップ取得、完了", vbInformation
    shell "explorer.exe " & createdPath, vbNormalFocus
End Sub





Sub ログノートをHTML出力と調整時間がログノートに記載されてるか確認_ユニット開始月_Click()
    Call ログノートをHTML出力と調整時間がログノートに記載されてるか確認_ユニット月(Year(ThisWorkbook.sheetS("手順").Range(BEGIN_COL & UNITROW)), Month(ThisWorkbook.sheetS("手順").Range(BEGIN_COL & UNITROW)))
End Sub

Sub ログノートをHTML出力と調整時間がログノートに記載されてるか確認_ユニット終了月_Click()
    Call ログノートをHTML出力と調整時間がログノートに記載されてるか確認_ユニット月(Year(ThisWorkbook.sheetS("手順").Range(END_COL & UNITROW)), Month(ThisWorkbook.sheetS("手順").Range(END_COL & UNITROW)))
End Sub





