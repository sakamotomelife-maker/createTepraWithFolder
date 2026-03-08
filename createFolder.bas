Attribute VB_Name = "folder"

'フォルダ作成のメインプロシージャ
Sub フォルダ一括作成()

      Dim ws As Worksheet
      Dim i As Long
      Dim last As Long
      Dim id As String
      Dim name As String
           
      Dim folder_path As String
      Dim cell As Range
          
      Set ws = ThisWorkbook.Worksheets(1)
      last = ws.Cells(ws.Rows.Count, "b").End(xlUp).Row
      

      '入力がない場合のエラー処理
      If last <= 4 Then
            MsgBox "入力されていません"
            flag = 1    '1：中止（グローバル関数）
            Exit Sub
      End If
      
      'Path作成したいフォルダ名を指定
      folder_path = "XXXXXXXXXXXXXXXXXXXXXXXX\"

      '-----------------------------------------------------------------------------
      
      '空白エラー処理
       Dim targetRange As String
        
        For j = 2 To 4
                For i = 5 To last
                        targetRange = Right(ThisWorkbook.Worksheets(1).Cells(i, j), 1)
                        
                        If targetRange = "　" Then
                            MsgBox "入力セルの末尾に空白あるため、処理を中止しました。"
                            flag = 1
                            Exit Sub
                        End If
                Next i
        Next j
         
      '
      'フォルダ一括作成
      Call set_md(ws, last)
      Call cmd_newfolder(ws, folder_path, last)
      
      '次回のIDをセット
      If ws.Cells(last, "B").Value > ws.Range("M1") Or ws.Cells(last, "B").Value = 9999 Then
            ws.Range("M1").Value = CDec(ws.Cells(last, "B").Value) + 1
      End If
      
      '入力内容リセット
      Call reset(last)

      ThisWorkbook.Save
                                                                                                                                                                  
      'エクスプローラーでフォルダをアクティブにする
      Call open_folder(folder_path)
      
End Sub
      

'E列にIDと商品名を結合し書き出し用の名前を作成する
Sub set_md(ByRef ws As Worksheet, ByRef last As Long)

    Dim i As Long
    Dim id As String
    Dim name As String
    
    For i = 5 To last
            id = Format(ws.Cells(i, "B").Value, "00000")    'Formatを使いリーディングゼロ問題解決(24/01/15)
            name = ws.Cells(i, "C").Value
            
            'idが空白じゃなければE列に結合
            If id <> "" Then
                   ws.Cells(i, "E").Value = id & "-" & name
            Else
                ws.Cells(i, "E").Value = ""
            End If
    Next i
    
End Sub


'新しいフォルダを作成して開く
Sub create_folder(ByRef folder_path As String)
  
    ' フォルダーが存在しない場合は作成
    If Dir(folder_path, vbDirectory) = "" Then
        MkDir folder_path
    End If

End Sub
    
    
'フォルダーを開きアクティブにする
Sub open_folder(ByRef folder_path As String)
 
   Dim shell As Object
 
    Set shell = CreateObject("WScript.Shell")
    shell.Run "explorer.exe " & folder_path
End Sub


'E列の名前で新しいフォルダを一括作成
Sub cmd_newfolder(ByRef ws As Worksheet, ByRef folder_path As String, ByRef last As Long)

      Dim cell As Range

      For Each cell In ws.Range("E5:E" & last)
      
            folder_name = cell.Value
                
            If folder_name <> "" Then
                  ' フォルダを作成
                  If Dir(folder_path & "\" & folder_name, vbDirectory) = "" Then
                        MkDir folder_path & "\" & folder_name
                  End If
            End If
      Next cell

End Sub

'B列からD列のクリア
Sub reset(last)
    Range("B5:E" & last).ClearContents
End Sub


'入力セルの末尾にスペースがある場合は削除
Sub rightSpaceDel()

    Dim targetRange As String
    Dim i As Long
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets(1)
    
    last = ws.Cells(ws.Rows.Count, "b").End(xlUp).Row
    
    For j = 2 To 4
            For i = 5 To last
                    targetRange = Right(ThisWorkbook.Worksheets(1).Cells(i, j), 1)
                    
                    If targetRange = "　" Then
                        MsgBox "入力セルの末尾に空白あるため、処理を中止しました。"
                        flag = 1
                        Exit Sub
                    End If
            Next i
    Next j
End Sub

