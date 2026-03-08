Attribute VB_Name = "scrapingSystem"

'コピペ関係
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long


'ID検索のメインプロシージャ
Sub inputInfo()

    MsgBox "管理システムから情報を抽出します。" & vbCrLf & _
            "メッセージが出るまでPCを操作しないでください。"
    
    '描画を止めて高速化
    Application.ScreenUpdating = False

    Dim last As String
    Dim i As Long
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets(1)
    
    i = 5
    last = ws.Cells(Rows.Count, "B").End(xlUp).Row
    
    
    Dim id As String
    Dim productName As String: productName = ""
    
    
    'Edgeの保存場所
    Dim edge As String
    edge = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    'edge = "C:\Program Files\Microsoft\Edge\Application\msedge.exe"

    'ダイレクトリンク                                    　★★★URLを入力★★★
    Dim id_search As String: id_search = "http://www.XXXXXXXXXXXXXXXXXXXXX="
    
    
    '処理
    For i = 5 To last
            
            'システム画面を開く
            id = ThisWorkbook.Worksheets(1).Cells(i, "B").Value
            shell edge & " " & id_search & id, vbNormalFocus
            
            Call getInfo(productName, i)
            
           ' ws.Cells(i, "D") = productName
            
    Next i

    Call minimize
    
    MsgBox "処理が完了しました"
    
    Application.ScreenUpdating = True
    
End Sub


'管理画面のコピペ処理
Sub getInfo(ByRef productName As String, ByVal i As Long)

                Dim ws As Worksheet
                Set ws = ThisWorkbook.Worksheets(1)
      
                'productNameに移動し、内容をclipboradへコピー
                Call waiting(3)
                SendKeys "{TAB 1}", True
                SendKeys "+{end}", True
                Call waiting(1)
                SendKeys "^(c)", True
                Call waiting(1)
                Call bufInfo(productName)
                
                ws.Cells(i, "C") = productName
                
                '--------------------------------------
                
                '商品バーコードへ移動し、内容をclipboardへコピー
                SendKeys "{TAB 1}", True
                SendKeys "+{end}", True
                Call waiting(1)
                SendKeys "^(c)", True
                Call waiting(1)
                Call bufInfo(productName)
                
                ws.Cells(i, "D") = productName
                
                SendKeys "^(w)", True
End Sub


'getinfoで呼び出し
Sub bufInfo(ByRef buf As String)

    Dim CB As New DataObject
    Dim objCb As New DataObject

    '入力がありコピーした場合は格納
    If Application.ClipboardFormats(1) <> -1 Then
        objCb.GetFromClipboard
        With CB
            .GetFromClipboard             'クリップボードからDataObjectにデータを格納する
            buf = .GetText                  'DataObjectのデータをselect_telに格納する
        End With
        
        'クリップボードをリセットしておく
        OpenClipboard (0&)
        EmptyClipboard
        CloseClipboard
    End If
End Sub

