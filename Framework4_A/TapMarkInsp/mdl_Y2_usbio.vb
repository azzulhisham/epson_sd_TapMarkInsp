Module mdl_Y2_usbio

    Public Const YDCI_OPEN_NORMAL As Short = 0
    Public Const YDCI_OPEN_OUT_NOT_INIT As Integer = &H1

    Public Const YDCI_RESULT_SUCCESS As Short = 0 ' 正常終了
    Public Const YDCI_RESULT_ERROR As Integer = &HCE000001 ' エラー
    Public Const YDCI_RESULT_NOT_OPEN As Integer = &HCE000002 ' オープンされていない
    Public Const YDCI_RESULT_ALREADY_OPEN As Integer = &HCE000003 ' オープン済み
    Public Const YDCI_RESULT_INVALID_ID As Integer = &HCE000004 ' IDが不正
    Public Const YDCI_RESULT_INVALID_EPNO As Integer = &HCE000005 ' EP指定が不正(内部エラー)
    Public Const YDCI_RESULT_CANNOT_OPEN As Integer = &HCE000006 ' デバイスがオープンできなかった
    Public Const YDCI_RESULT_PARAMETER_ERROR As Integer = &HCE000008 ' 引数不正
    Public Const YDCI_RESULT_MEM_ALLOC_ERROR As Integer = &HCE000009 ' メモリ確保エラー
    Public Const YDCI_RESULT_MODELNAME_ERROR As Integer = &HCE00000A ' 型名指定が不正
    Public Const YDCI_RESULT_NOT_SUPPORTED As Integer = &HCE00000C ' サポートされていない
    Public Const YDCI_RESULT_INVALID_BOARD_SW As Integer = &HCE000013 ' ボード識別スイッチが不正

    Public Const YDCI_RESULT_FATAL_ERROR As Integer = &HCEFFFFFF ' 致命的エラー


    Public Const Bit_OFF = 0
    Public Const Bit_ON = 1

    Public Declare Function YdciOpen Lib "Ydci.dll" (ByVal wBoardSw As Short, ByVal lpszModelName As String, ByRef pwID As Short, Optional ByVal wMode As Short = YDCI_OPEN_NORMAL) As Integer
    Public Declare Function YdciClose Lib "Ydci.dll" (ByVal wID As Short) As Boolean
    Public Declare Function YdciDioInput Lib "Ydci.dll" (ByVal wID As Short, ByRef pbyData As Byte, ByVal wStart As Short, ByVal wCount As Short) As Integer
    Public Declare Function YdciDioOutput Lib "Ydci.dll" (ByVal wID As Short, ByRef pbyData As Byte, ByVal wStart As Short, ByVal wCount As Short) As Integer
    Public Declare Function YdciDioOutputStatus Lib "Ydci.dll" (ByVal wID As Short, ByRef pbyData As Byte, ByVal wStart As Short, ByVal wCount As Short) As Integer
    Public Declare Function YdciRlyOutput Lib "Ydci.dll" (ByVal wID As Short, ByRef pbyData As Byte, ByVal wStart As Short, ByVal wCount As Short) As Integer
    Public Declare Function YdciRlyOutputStatus Lib "Ydci.dll" (ByVal wID As Short, ByRef pbyData As Byte, ByVal wStart As Short, ByVal wCount As Short) As Integer

    Public Enum BitsState
        eBit_OFF = Bit_OFF
        eBit_ON = Bit_ON
    End Enum

End Module
