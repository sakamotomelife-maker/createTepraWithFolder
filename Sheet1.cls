VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' エラーメッセージの定義
Const ERROR_MESSAGE_NO_PRINT_JOB = "印刷する項目が入力されていないため､ラベルを印刷できません｡"
Const ERROR_MESSAGE_GET_TAPE_WIDTH = "テープ幅が取得できません。"
Const ERROR_MESSAGE_TPE_FILE_NOT_FOUND = "テプラに18mmのテープをセットしてください。"
Const ERROR_MESSAGE_RUN_PRINT = """SPC10.exe""が指定した場所に存在しません。インストール先を確認してください。"


'==============================================================================
' CSVファイルを作成し、印刷実行関数を呼び出す_OK
'==============================================================================

'ラベル作成のメインプロシージャ
Public Sub 受講生ラベル作成()

    Dim strExePathName As String
    Dim strTextPathName As String
    Dim strCsvPathName As String
    Dim strPrintLogPathName As String
    Dim strTpePathName As String
    Dim strOption As String
    Dim dblRetValue As Double

    MAX_LINE_COUNT = ThisWorkbook.Worksheets(1).Cells(Rows.Count, "B").End(xlUp).Row
    
    'テプラのファイルパス
    Dim tepraFullPath As String
    tepraFullName = ThisWorkbook.Worksheets(2).Range("C5").Value
    
    flag = 0

    '-----------------------------------------------------------------------------
    ' SPC10のEXEファイルをパス付きで指定する
    '-----------------------------------------------------------------------------
    If IsWow64() Then
        ' OSが64ビット環境
        strExePathName = tepraFullName
    Else
        ' OSが32ビット環境
        strExePathName = tepraFullName
    End If

    '-----------------------------------------------------------------------------
    ' テープ幅の出力ファイル、CSVファイル、印刷結果ファイル、LW1ファイルの指定
    '-----------------------------------------------------------------------------
    strTextPathName = ThisWorkbook.Path & "\" & "TapeWidth.txt"                               ' テープ幅の出力ファイルを指定する
    strCsvPathName = ThisWorkbook.Path & "\" & "data.csv"                                       ' CSVファイルを指定する
    strPrintLogPathName = ThisWorkbook.Path & "\" & "PrintResult.txt"                         ' 印刷結果ファイルを指定する
    strTpePathName = ThisWorkbook.Path & "\CreateTepraWithFolder_18.lw1"              ' LW1ファイルを指定

    '-----------------------------------------------------------------------------
    ' 設定情報の取得_OK
    '-----------------------------------------------------------------------------
    ' ハーフカット設定_ON
    Dim blnHalfcut As Boolean
    blnHalfcut = True

    ' テープ幅確認メッセージ設定_OFF
    Dim blnConfirmTapeWidth As Boolean
        blnConfirmTapeWidth = False
        
    ' 印刷結果をファイルに出力する設定
    Dim strPrintLog As String
        strPrintLog = strPrintLogPathName

    '-----------------------------------------------------------------------------
    ' 印刷対象の確認_OK（起動しないはず）
    '-----------------------------------------------------------------------------
    If (getPrintJobCount() = 0) Then
        MsgBox ERROR_MESSAGE_NO_PRINT_JOB
        
        Application.ScreenUpdating = True
        flag = 1
        Exit Sub
    End If

    '-----------------------------------------------------------------------------
    ' テープ幅のファイル出力関数の呼び出し_★要確認
    '-----------------------------------------------------------------------------
    strOption = createPrintOption(strTpePathName, strCsvPathName, 1, blnHalfcut, blnConfirmTapeWidth, strPrintLog, strTextPathName)
    dblRetValue = PrtSpc10Api(strExePathName, strOption, "")
    
    If (dblRetValue = 0) Then
        ' API実行エラー
        MsgBox ERROR_MESSAGE_RUN_PRINT
        Application.ScreenUpdating = True
        flag = 1
        Exit Sub
    End If

    ' テープ幅の出力ファイルが保存されるまで少し待つ
    Application.Wait (Now + TimeValue("00:00:03"))
    
    '-----------------------------------------------------------------------------
    ' テープ幅の出力ファイルの存在確認_★要確認
    '-----------------------------------------------------------------------------
    If Dir(strTextPathName) = "" Then
        ' テープ幅の出力ファイルが存在しない場合
        MsgBox ERROR_MESSAGE_GET_TAPE_WIDTH
        Application.ScreenUpdating = True
        flag = 1
        Exit Sub
    End If

    '-----------------------------------------------------------------------------
    ' LW1ファイルの存在確認_OK
    '-----------------------------------------------------------------------------
    Dim strTapeWidth As String
    Dim strTapeType As String
    
    ' テープ幅の出力ファイルからテープ幅（テープ種類）を取得
    strTapeType = ""
    strTapeWidth = getTapeWidth(strTextPathName, strTapeType)
    
    ' テープ幅の確認
    If StrComp(strTapeWidth, "0") = 0 Then
        ' テープ未装着の場合
        Application.ScreenUpdating = True
        flag = 1
        Exit Sub
    End If
    
    ' テープ種類の確認
    If StrComp(strTapeType, "0x00") Then
        ' Standard type以外の場合
        MsgBox ERROR_MESSAGE_TPE_FILE_NOT_FOUND
        Application.ScreenUpdating = True
        flag = 1
        Exit Sub
    End If

    ' LW1ファイルを指定
    strTpePathName = ThisWorkbook.Path & "\CreateTepraWithFolder_" & strTapeWidth & ".lw1"
    
    If Dir(strTpePathName) = "" Then
        ' TPEファイルが存在しない場合
        MsgBox ERROR_MESSAGE_TPE_FILE_NOT_FOUND
        Application.ScreenUpdating = True
        flag = 1
        Exit Sub
    End If
    
    '-----------------------------------------------------------------------------
    ' CSVファイルの作成
    '-----------------------------------------------------------------------------
      Dim ws As Worksheet
     
      Dim id As String
      Dim name As String
          
     Set ws = ThisWorkbook.Worksheets(1)
     last = ws.Cells(ws.Rows.Count, "b").End(xlUp).Row
     
     'リーディングゼロ対策１_★動いていない？
      For i = 5 To last
            id = Format(ws.Cells(i, "B").Value, "00000")
      Next i
       
     Dim csvStr1        As String    'ID
     Dim csvStr2        As String    '氏名
     Dim csvStr3        As String    'ふりがな


     On Error Resume Next
    
     ' 初期値設定
     Dim fileNo As Integer
     fileNo = FreeFile
    
     Open strCsvPathName For Output As #fileNo
    
    
     For i = 1 To MAX_LINE_COUNT
        
        If Range("B" & (i + LINE_OFFSET)).Value <> "" Then
            
            'リーディングゼロ対策２
            If Left(Range("B" & (i + LINE_OFFSET)).Value, 1) <> 0 Then
                csvStr1 = "0" & Range("B" & (i + LINE_OFFSET)).Value
            Else
                csvStr1 = Range("B" & (i + LINE_OFFSET)).Value
            End If
                
            csvStr2 = Range("C" & (i + LINE_OFFSET)).Value
            csvStr3 = Range("D" & (i + LINE_OFFSET)).Value
            
            ' データの存在確認
            If (Len(csvStr1) = 0 And Len(csvStr2) = 0 And Len(csvStr3) = 0) Then
                ' 空データの行は、チェックされていても無視する
            Else
                ' 空でなければCSV化
                csvStrAll = strCompanyName & _
                            csvStr1 & "," & _
                            csvStr2 & "," & _
                            csvStr3
                
                ' CSVファイル出力
                Print #fileNo, csvStrAll
            End If
        End If
    Next i
    
    Close #fileNo

    '-----------------------------------------------------------------------------
    ' 印刷関数の呼び出し
    '-----------------------------------------------------------------------------
    ' 印刷実行
    strOption = createPrintOption(strTpePathName, strCsvPathName, 1, blnHalfcut, blnConfirmTapeWidth, strPrintLog, "")
    dblRetValue = PrtSpc10Api(strExePathName, strOption, "")
    If (dblRetValue = 0) Then
        ' API実行エラー
        MsgBox ERROR_MESSAGE_RUN_PRINT
        flag = 1
        Application.ScreenUpdating = True
    End If

End Sub

'==============================================================================
' 印刷ジョブ数を取得する
'==============================================================================
Private Function getPrintJobCount() As Integer

    Dim csvStr1 As String
    Dim csvStr2 As String
    Dim csvStr3 As String
    Dim intPrintJob As Integer ' 印刷JOB数
    Dim i As Integer

    ' 初期値設定
    intPrintJob = 0
    MAX_LINE_COUNT = ThisWorkbook.Worksheets(1).Cells(Rows.Count, "B").End(xlUp).Row
    
    For i = 1 To MAX_LINE_COUNT
    
            csvStr1 = Range("B" & (i + LINE_OFFSET)).Value
            csvStr2 = Range("C" & (i + LINE_OFFSET)).Value
            csvStr3 = Range("D" & (i + LINE_OFFSET)).Value
           
            ' データの存在確認
            If (Len(csvStr1) = 0 And Len(csvStr2) = 0 And Len(csvStr3) = 0) Then
                ' 空データの行は、チェックされていても無視する
            Else
                intPrintJob = intPrintJob + 1
            End If
        
    Next i
    
    getPrintJobCount = intPrintJob
    
End Function
