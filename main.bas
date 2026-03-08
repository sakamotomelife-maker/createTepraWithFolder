Attribute VB_Name = "main"
'グローバル変数
Public flag As Long
Public MAX_LINE_COUNT As Long       'tepraで使用


'実行マクロ
Sub CreateTepraWithFolder_RUN()

      Application.ScreenUpdating = False
      Dim response As String
      Dim makeFolderValue As Boolean
      Dim makeLabelValue As Boolean
      
      
      flag = 0    '1：中止
      makeLabelValue = Worksheets(1).OLEObjects("makeLabel").Object.Value
      makeFolderValue = Worksheets(1).OLEObjects("makeFolder").Object.Value

      
      '実行確認
      If makeLabelValue = True And makeFolderValue = True Then
            response = MsgBox("ラベルとフォルダを一括で作成します。" & vbCrLf & _
                                        "処理を続行しますか？", _
                                         vbYesNo)
      ElseIf makeLabelValue = False And makeFolderValue = True Then
            response = MsgBox("フォルダのみ作成します。" & vbCrLf & _
                             "処理を続行しますか？", _
                             vbYesNo)
      ElseIf makeLabelValue = True And makeFolderValue = False Then
            response = MsgBox("ラベルのみ作成します。" & vbCrLf & _
                             "処理を続行しますか？", _
                             vbYesNo)
      Else
            MsgBox "チェックボックスをONにしてください。"
            Exit Sub
      End If
      
      If response = vbNo Then
            Application.ScreenUpdating = True
            Exit Sub
      End If
      
      
      '処理実行
      If makeLabelValue = True Then Call Sheet1.ラベル作成
      If flag = 1 Then Exit Sub
      
      If makeFolderValue = True Then Call フォルダ一括作成
      If flag = 1 Then Exit Sub
      
      MsgBox "処理が完了しました。"
      Application.ScreenUpdating = True

End Sub


