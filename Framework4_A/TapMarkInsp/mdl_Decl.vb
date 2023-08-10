
'-----------------------------------------------------
'   Taping  : Marking Inspection System Integration
'=====================================================
'   Name    : TapMarkInsp
'   Written : Zulhisham Tan - S1609
'   Date    : 18th May '2011
'-----------------------------------------------------

Imports System.Math
Imports System.Threading
Imports System.IO
Imports System.IO.Ports
Imports System.Management
Imports System.Runtime.InteropServices
Imports Microsoft.Win32


Module mdl_Decl

    Public Const Authentication As String = "Sd00"

    Public Const sqlServer As String = "20.10.30.2\SQLEXPRESS"
    Public Const sqlName As String = "Marking"
    Public Const sqluid As String = "VB-SQL"
    Public Const sqlpwd As String = "Anyn0m0us"


    Public Const Func_Ret_Success = 0
    Public Const Func_Ret_Fail = -1

    Public Const fg_OFF = 0
    Public Const fg_ON = 1

    Public Const MarkingLine1Char As Integer = 5
    Public Const CameraSceneNoStartAt As Integer = 14

    Public Enum SuccessError
        retError = -1
        Terminated = 0
        retSuccess = 1
    End Enum

    Public Structure DB
        Public FileName As String
        Public Path As String
    End Structure

    Public Structure InspData
        Public IMISpec As String
        Public LotNo As String
        Public Freq As String

        Public Opt As String
        Public Plant As String

        Public MarkingDate As Date
        Public MarkingTime As Date

        Public MarkData1 As String
        Public MarkData2 As String
        Public MC_no As String
    End Structure

    Structure LotDetails
        Public LotNo As String
        Public LotNo_ As String
        Public SpecNo As String
        Public OperatorNo As String
        Public InputQty As String
        Public InspDirection As String
        Public M_LotNo As String
        Public ReviseLotNo As String

        Public dbData() As InspData
        Public WeekCdChk As Integer
    End Structure

    Structure DIO_Board
        Public sBoardType As String
        Public sBoardName As String
        Public sBoardMaker As String
        Public iBoardNo As Integer
        Public iHnd As Integer
    End Structure

    Public Structure CommPortData
        Public PortName As String
        Public DataBits As Integer
        Public BaudRate As Integer
        Public StopBits As System.IO.Ports.StopBits
        Public Parity As System.IO.Ports.Parity
    End Structure

    Public Structure Rec
        Public Lot_No As String
        Public IMI_No As String
        Public FreqVal As String
        Public Opt As String
        Public RecDate As String
        Public Profile As String
        Public CtrlNo As String
        Public MacNo As String
        Public MData1 As String
        Public MData2 As String
        Public MData3 As String
        Public MData4 As String
        Public MData5 As String
        Public MData6 As String
    End Structure

    Structure SystData
        Public SettingLot As LotDetails
        Public OAI_CAM_1 As CommPortData
        Public USB_IO() As DIO_Board
        Public dbSource As DB

        Public BootState As Integer
        Public RunningState As Integer

        Public MC_No As String
        Public DataPath As String
        Public LotDataPath As String
    End Structure


    Public GrnLED_OnOff(1) As Color
    Public RedLED_OnOff(1) As Color
    Public OrgLED_OnOff(1) As Color
    Public BlueLED_OnOff(1) As Color
    Public CrayLED_OnOff(1) As Color

    Public regKey As RegistryKey = Registry.CurrentUser
    Public regSubKey As RegistryKey = regKey.CreateSubKey("Software\az_Logics\TapMarkInsp")


    Public TapMarkInsp As SystData

End Module
