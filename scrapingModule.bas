Attribute VB_Name = "scrapingModule"

'　コピペ関係
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long

'Edge画面最小化用
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
'

'クリップボードをリセット
Sub reset_clipboard()
                OpenClipboard (0&)
                EmptyClipboard
                CloseClipboard
End Sub

'全選択コピー
Sub allcopy()
    SendKeys "^a", True
    SendKeys "^c", True
End Sub


'ペーストモジュール ("sheet名","A1")
Sub allpaste(ByVal sheet_name As String, ByVal cell_name As String)

    ' クリップボードにデータがあるか確認
    If Application.CutCopyMode = False Then
        MsgBox "クリップボードにデータを格納できませんでした。"
        Exit Sub
    End If

    ' 指定されたセルにペースト
    ThisWorkbook.Sheets(sheet_name).Range(cell_name).PasteSpecial Paste:=xlPasteValues
    
End Sub


'Edgeウインドウ最小化
Sub minimize()

    Dim hwnd As LongPtr
    Dim edgeTitle As String: edgeTitle = "XXXXXXXXXXX" ' Edgeウィンドウのタイトルを指定

    ' Edgeウィンドウのハンドルを取得
    hwnd = FindWindow(vbNullString, edgeTitle)
    
    If hwnd <> 0 Then
        ' ウィンドウを最小化
        ShowWindow hwnd, 6      '6 ---> SW_MINIMIZE
    Else
        MsgBox "Edgeウィンドウが見つかりませんでした。"
    End If
    
End Sub


'待機時間の指定（例：1秒待機 ---> call waiting(1) ）
Sub waiting(ByVal sec As Integer)

    Dim time As String: time = "00:00:" & Format(sec, "00")
    Application.Wait Now + TimeValue(time)
    
End Sub

