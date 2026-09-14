Attribute VB_Name = "Module17"
' ///////////////////////////////////////////////////////////////////////////
' // 定期的なファイルアクセス確認用モジュール          定期的にサーバーに置いてるファイルがアクセス可能か確認                    //
' ///////////////////////////////////////////////////////////////////////////

' 公開変数：監視対象のファイルパス
Public Const TARGET_FILE_PATH As String = BNAME_MATOME ' ★★★ ここを監視したいファイルのパスに書き換えてください ★★★

' 公開変数：次回の実行時刻を格納
Public NextRunTime As Date

' 公開変数：前回のファイルアクセス状態を記憶 (True:アクセス可能, False:アクセス不可)
Private previousAccessStatus As Boolean

' ///////////////////////////////////////////////////////////////////////////
' // 関数：ファイルのアクセス可否をチェックする  ネットワーク上のファイルにアクセスできるか確認　「Microsoft Scripting Runtime」が必要======================================================
' ///////////////////////////////////////////////////////////////////////////
' VBAエディタで「ツール」>「参照設定」>「Microsoft Scripting Runtime」にチェックを入れてください。
Function CheckServerAccess_FSO(ByVal fullNetworkFilePath As String) As Boolean
    Dim fso As Object
    
    CheckServerAccess_FSO = False ' デフォルトでは失敗と設定
    
    On Error GoTo ErrorHandler ' エラーハンドリングを設定
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fullNetworkFilePath Like "*\" Then
'        MsgBox "末尾は \ です"
        If fso.FolderExists(fullNetworkFilePath) Then
            CheckServerAccess_FSO = True ' アクセス成功
            Debug.Print "アクセス成功==="
        Else
            MsgBox "Err@CheckServerAccess_FSO  " & fullNetworkFilePath & "' にアクセスできません。ネットワーク接続の問題か、フォルダが存在しないか、アクセス権がありません。", vbCritical
        End If
    Else
        If fso.FileExists(fullNetworkFilePath) Then
            CheckServerAccess_FSO = True ' アクセス成功
            'ThisWorkbook.sheetS("手順").Range("B2").Value = "Connect" '定期的な監視はやめたのでコメントアウト
            'MsgBox "OK@CheckServerAccess_FSO  " & fullNetworkFilePath & "' にアクセスOK", vbInformation
            Debug.Print "アクセス成功==="
        Else
            MsgBox "Err@CheckServerAccess_FSO  " & fullNetworkFilePath & "' にアクセスできません。ネットワーク接続の問題か、ファイルが存在しないか、アクセス権がありません。", vbCritical
        End If
    End If

    
    Set fso = Nothing
    Exit Function ' 正常終了時はエラーハンドラをスキップ
    
ErrorHandler:
    ' エラーが発生した場合（例：パスが不正、権限がないなど）
    Debug.Print "エラー発生 (CheckServerAccess_FSO): " & Err.Description ' デバッグ用にエラーメッセージを表示
    CheckServerAccess_FSO = False ' エラー時はアクセス失敗
    Set fso = Nothing
End Function

' ///////////////////////////////////////////////////////////////////////////
' // サブプロシージャ：定期的に実行される監視処理                          //
' ///////////////////////////////////////////////////////////////////////////
Sub MonitorFileAccess()
    Dim currentAccessStatus As Boolean
    
    ' ファイルの現在のアクセス状態を確認
    currentAccessStatus = CheckServerAccess_FSO(TARGET_FILE_PATH)
    
    ' Debug.Print Now & " - アクセス状態: " & currentAccessStatus & " (前回: " & previousAccessStatus & ")" ' デバッグ用
    
    ' 初回実行時、またはアクセス状態が前回から変化した場合
    If Not IsEmpty(previousAccessStatus) Then ' 初回実行時以外
        If currentAccessStatus <> previousAccessStatus Then
            If currentAccessStatus = False Then
                ' アクセス可能だった状態からアクセス不可になった場合
                ThisWorkbook.sheetS("手順").Range("B2").Value = "Not Connect"
                MsgBox "接続が切れた可能性があります: ファイル「" & TARGET_FILE_PATH & "」にアクセスできません。", vbCritical + vbOKOnly, "サーバーファイルアクセス警告"
            Else
                ' アクセス不可だった状態からアクセス可能になった場合
                ThisWorkbook.sheetS("手順").Range("B2").Value = "Connect"
                ' MsgBox "通知: ファイル「" & TARGET_FILE_PATH & "」へのアクセスが回復しました。", vbInformation + vbOKOnly, "サーバーファイルアクセス回復"
            End If
        End If
    End If
    
    ' 現在の状態を前回の状態として記憶
    previousAccessStatus = currentAccessStatus
    
    ' 次回の実行時刻を設定 (例: 5分後)
    NextRunTime = Now + TimeValue("00:05:00") ' ★★★ 監視間隔をここで調整してください ★★★
    Application.OnTime NextRunTime, "MonitorFileAccess"
End Sub

' ///////////////////////////////////////////////////////////////////////////
' // サブプロシージャ：監視を開始する                                      //
' ///////////////////////////////////////////////////////////////////////////
Sub StartMonitoring()
    ' 初回実行時は previousAccessStatus を設定しない (IsEmpty)
    ' 既にタイマーが設定されている場合は、既存のタイマーをキャンセルしてから開始
    Call StopMonitoring ' 念のため、既存のタイマーをキャンセル
    
    ' 初回のファイルアクセスチェックを行い、previousAccessStatus を設定
    previousAccessStatus = CheckServerAccess_FSO(TARGET_FILE_PATH)
    
    MsgBox "ファイルアクセス監視を開始します。" & vbCrLf & _
           "監視対象: " & TARGET_FILE_PATH & vbCrLf & _
           "監視間隔: 5分ごと" & vbCrLf & _
           "初回アクセス状態: " & IIf(previousAccessStatus, "可能", "不可"), vbInformation
    
    ' 最初の監視をすぐに実行
    Call MonitorFileAccess
End Sub

' ///////////////////////////////////////////////////////////////////////////
' // サブプロシージャ：監視を停止する                                      //
' ///////////////////////////////////////////////////////////////////////////
Sub StopMonitoring()
    On Error Resume Next ' 実行中のタイマーがない場合のエラーを無視
    Application.OnTime NextRunTime, "MonitorFileAccess", , False
    On Error GoTo 0 ' エラーハンドリングを元に戻す
    
    ' 状態をリセット
    previousAccessStatus = Empty ' previousAccessStatus をリセット
    NextRunTime = 0 ' NextRunTime をリセット
    
    MsgBox "ファイルアクセス監視を停止しました。", vbInformation
End Sub

