Attribute VB_Name = "tepraAPI"
'★★★テプラ作成のメインプロシージャはsheet1に記述★★★


' Windows APIの関数を呼び出すVBA関数を定義
Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" _
                    (ByVal lpModuleName As String) As Long
                       
Declare PtrSafe Function GetProcAddress Lib "kernel32" _
                    (ByVal hModule As Long, _
                     ByVal lpProcName As String) As Long
                         
Declare PtrSafe Function GetCurrentProcess Lib "kernel32" () As Long

Declare PtrSafe Function IsWow64Process Lib "kernel32" _
                    (ByVal hProcess As Long, _
                     ByRef Wow64Process As Long) As Long

' データ開始行のオフセット
Public Const LINE_OFFSET = 4


'==============================================================================
' OSが64ビット環境かどうかを判別する関数の定義_OK
'==============================================================================
Function IsWow64() As Boolean

    Dim bIsWow64 As Long
    Dim fnIsWow64Process As Long
    
    ' 初期化
    bIsWow64 = False
    
    ' IsWow64Process関数が存在するかどうかを確認する
    fnIsWow64Process = GetProcAddress(GetModuleHandle("kernel32"), "IsWow64Process")

    If (0 <> fnIsWow64Process) Then
        ' IsWow64ProcessはWindows XP SP2から導入された関数
        If 0 = IsWow64Process(GetCurrentProcess(), bIsWow64) Then
            ' Windows XP
        End If
    End If

    IsWow64 = (bIsWow64 <> 0)

End Function


'==============================================================================
' テープ幅取得関数の定義_★要確認
'==============================================================================
Function getTapeWidth(ByVal strFileName As String, ByRef strTapeType As String) As String
    
    Dim retValue As String
    
    ' ファイルを読み込むための配列
    Dim Arr()
    ReDim Preserve Arr(0)

    ' オブジェクトを作成
    Dim obj As Object
    Set obj = CreateObject("ADODB.Stream")

    ' オブジェクトに保存するデータの種類を文字列型に指定する
    obj.Type = adTypeText
    
    ' 文字列型のオブジェクトの文字コードを指定する
    obj.Charset = "UTF-16"

    ' オブジェクトのインスタンスを作成
    obj.Open

    ' ファイルからデータを読み込む
    obj.LoadFromFile (strFileName)

    ' 最終行までループする
    Do While Not obj.EOS
    
        ' 次の行を読み取る
        Arr(UBound(Arr)) = obj.ReadText(adReadLine)
        ReDim Preserve Arr(UBound(Arr) + 1)
    Loop

    ' オブジェクトを閉じる
    obj.Close

    ' メモリからオブジェクトを削除する
    Set obj = Nothing
    
    
    '-----------------------------------------------------------------------------
    ' テープ幅の取得_★要確認（18mmのみ使用）
    '-----------------------------------------------------------------------------
    ' 読み込んだ1行目の文字列を分割（例：0x04 18mm）
    Dim strData As Variant
    strData = Split(Arr(0), " ")

    ' テープ幅を取得
    Dim strTapeWidth As String
    strTapeWidth = strData(0)

    ' テープ幅の設定
    Select Case strTapeWidth
        Case "0x00"
            retValue = "0"
        Case "0x01"
            retValue = "6"
        Case "0x02"
            retValue = "9"
        Case "0x03"
            retValue = "12"
        Case "0x04"
            retValue = "18"
        Case "0x05"
            retValue = "24"
        Case "0x06"
            retValue = "36"
        Case "0x07"
            retValue = "50"
        Case "0x0B"
            retValue = "4"
        Case "0x21"
            retValue = "50"         ' WR1000用 50mm
        Case "0x23"
            retValue = "100"        ' WR1000用 100mm
        Case "0xFF"
            retValue = ""
        Case Else
           retValue = ""
    End Select

    '-----------------------------------------------------------------------------
    ' テープ種類の取得_OK
    '-----------------------------------------------------------------------------
    ' 読み込んだ2行目の文字列を分割（例：0x00 Standard tape）
    Dim strTypeData As Variant
    strTypeData = Split(Arr(1), " ")

    ' テープ種類を取得
    strTapeType = strTypeData(0)
    getTapeWidth = retValue
    
End Function


'==============================================================================
' オプション文字列生成関数の定義_OK
'==============================================================================
Function createPrintOption(pathTempl As String, pathCsv As String, printNum As Integer, blnHalfcut As Boolean, blnConfirmTapeWidth As Boolean, strPrintLog As String, strTapeWidth As String) As String

    Dim comStrg   As String
    Dim retValue  As Double
    Dim strOption As String
    
    ' TPEファイルのフルパス名,CSVファイルのフルパス名,印刷部数
    strOption = pathTempl & "," & pathCsv & "," & printNum & ","
    
    ' テープ幅のファイル出力
    If (Len(strTapeWidth) > 0) Then
        strOption = strOption + "," + "/GT " + strTapeWidth
    End If
    
    ' カット設定
    If (blnHalfcut) Then
        strOption = strOption + "," + "/C -f -h"
    Else
        strOption = strOption + "," + "/C -f -hn"
    End If
    
    ' テープ幅確認メッセージのon/off設定
    If (blnConfirmTapeWidth) Then
        strOption = strOption + "," + "/TW -on"
    Else
        strOption = strOption + "," + "/TW -off"
    End If
    
    ' 印刷結果のファイル出力
    If (Len(strPrintLog) > 0) Then
        strOption = strOption + "," + "/L " + strPrintLog
    End If
   
    createPrintOption = strOption

End Function


'==============================================================================
' 印刷実行関数の定義_OK
'==============================================================================
Function PrtSpc10Api(pathCmd As String, strOption As String, strPrinterName As String) As Double
    
    Dim comStrg  As String
    Dim retValue As Double
    
    ' 印刷コマンド
    If (Len(strPrinterName) > 0) Then
        ' /ptオプション
        comStrg = pathCmd & " " & "/pt " & Chr(34) & strOption & Chr(34) & " " & Chr(34) & strPrinterName & Chr(34)
    Else
        ' /pオプション
        comStrg = pathCmd & " " & "/p " & Chr(34) & strOption & Chr(34)
    End If
    
    On Error Resume Next
    
    retValue = shell(comStrg, vbHide)
    
    PrtSpc10Api = retValue
    
End Function

