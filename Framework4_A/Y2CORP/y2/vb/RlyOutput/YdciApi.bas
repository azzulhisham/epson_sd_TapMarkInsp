Attribute VB_Name = "YdciApi"
'************************************ 定義 ************************************

Public Const YDCI_OPEN_NORMAL = 0
Public Const YDCI_OPEN_OUT_NOT_INIT = &H1


'************************************ API *************************************

Declare Function YdciOpen Lib "Ydci.dll" (ByVal wBoardSw As Integer, _
                                          ByVal lpszModelName As String, _
                                          ByRef pwID As Integer, _
                                          Optional ByVal wMode As Integer = YDCI_OPEN_NORMAL) _
                                          As Long
Declare Function YdciClose Lib "Ydci.dll" (ByVal wID As Integer) _
                                           As Boolean
Declare Function YdciDioInput Lib "Ydci.dll" (ByVal wID As Integer, _
                                              ByRef pbyData As Byte, _
                                              ByVal wStart As Integer, _
                                              ByVal wCount As Integer) _
                                              As Long
Declare Function YdciDioOutput Lib "Ydci.dll" (ByVal wID As Integer, _
                                               ByRef pbyData As Byte, _
                                               ByVal wStart As Integer, _
                                               ByVal wCount As Integer) _
                                               As Long
Declare Function YdciDioOutputStatus Lib "Ydci.dll" (ByVal wID As Integer, _
                                               ByRef pbyData As Byte, _
                                               ByVal wStart As Integer, _
                                               ByVal wCount As Integer) _
                                               As Long
Declare Function YdciRlyOutput Lib "Ydci.dll" (ByVal wID As Integer, _
                                               ByRef pbyData As Byte, _
                                               ByVal wStart As Integer, _
                                               ByVal wCount As Integer) _
                                               As Long
Declare Function YdciRlyOutputStatus Lib "Ydci.dll" (ByVal wID As Integer, _
                                               ByRef pbyData As Byte, _
                                               ByVal wStart As Integer, _
                                               ByVal wCount As Integer) _
                                               As Long


'******************************** エラーコード ********************************
Public Const YDCI_RESULT_SUCCESS = 0                    ' 正常終了
Public Const YDCI_RESULT_ERROR = &HCE000001             ' エラー
Public Const YDCI_RESULT_NOT_OPEN = &HCE000002          ' オープンされていない
Public Const YDCI_RESULT_ALREADY_OPEN = &HCE000003      ' オープン済み
Public Const YDCI_RESULT_INVALID_ID = &HCE000004        ' IDが不正
Public Const YDCI_RESULT_INVALID_EPNO = &HCE000005      ' EP指定が不正(内部エラー)
Public Const YDCI_RESULT_CANNOT_OPEN = &HCE000006       ' デバイスがオープンできなかった
Public Const YDCI_RESULT_PARAMETER_ERROR = &HCE000008   ' 引数不正
Public Const YDCI_RESULT_MEM_ALLOC_ERROR = &HCE000009   ' メモリ確保エラー
Public Const YDCI_RESULT_MODELNAME_ERROR = &HCE00000A   ' 型名指定が不正
Public Const YDCI_RESULT_NOT_SUPPORTED = &HCE00000C     ' サポートされていない
Public Const YDCI_RESULT_INVALID_BOARD_SW = &HCE000013  ' ボード識別スイッチが不正

Public Const YDCI_RESULT_FATAL_ERROR = &HCEFFFFFF       ' 致命的エラー

