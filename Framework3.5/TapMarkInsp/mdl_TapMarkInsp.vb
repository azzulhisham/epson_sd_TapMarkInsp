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
Imports System.Data
Imports System.Data.Odbc



Module mdl_TapMarkInsp

    Public Function CreateDBConnection(ByVal dbFilePath As String) As OdbcConnection

        ' Connect string.
        Dim sConnStr As String = _
            "driver=Microsoft Access Driver (*.mdb); uid=admin; UserCommitSync=Yes; " & _
                    "Threads=3; SafeTransactions=0; PageTimeout=5; MaxScanRows=8; MaxBufferSize=2048; FIL=MS Access; DriverId=25; " & _
                    "DefaultDir=" & dbFilePath.Substring(0, dbFilePath.LastIndexOf("\")) & "; " & _
                    "DBQ=" & dbFilePath

        Try
            ' Open Connection.
            Dim oConn As New OdbcConnection(sConnStr)
            oConn.Open()
            ' Return Object.
            Return oConn
        Catch ex As Exception

        End Try

        Return Nothing

    End Function

    Public Function Read_DataBase(ByVal DBSource As DB, ByVal LotNo As String, ByRef Data() As InspData) As Integer

        'F:\az_Logics\AI\MyVBProject\VB_Project\ML-7111A\Gyro\Data\Record\record.mdb
        Dim oConn As OdbcConnection = CreateDBConnection(DBSource.Path & "\" & DBSource.FileName)

        Dim ch As Char = ChrW(39)
        Dim sGUID As String = ""
        Dim lRetVal As Long = -1


        If IsNothing(oConn) Then
            Return -1
        End If

        Dim sSQLcmd As String = String.Empty

        Dim str_pLotNo As String = LotNo
        sSQLcmd = "SELECT * FROM Record WHERE sLotNo ='"

        Do Until str_pLotNo = ""
            Application.DoEvents()

            If str_pLotNo.IndexOf(",") < 0 Then
                sSQLcmd &= str_pLotNo.Trim & "'"
                str_pLotNo = ""
            Else
                sSQLcmd &= str_pLotNo.Substring(0, str_pLotNo.IndexOf(",")) & "'"
                str_pLotNo = str_pLotNo.Substring(str_pLotNo.IndexOf(",") + 1).Trim

                sSQLcmd &= " OR sLotNo='"
            End If
        Loop


        Dim OdbcCmd As New OdbcCommand(sSQLcmd, oConn)
        Dim oDBReader As OdbcDataReader = OdbcCmd.ExecuteReader()

        With oDBReader
            Dim iFieldCnt As Integer = .FieldCount
            Dim iRecNo As Integer = 0


            If .HasRows Then
                Dim iLp As Integer = 0
                Dim sRetData(iFieldCnt - 1) As String

                ReDim Data(iRecNo)
                Application.DoEvents()

                Do While .Read()
                    Application.DoEvents()

                    ReDim Preserve Data(iRecNo)

                    With Data(iRecNo)
                        .IMISpec = oDBReader.GetString(0)
                        .LotNo = oDBReader.GetString(1)
                        .Freq = oDBReader.GetString(2)
                        .Opt = oDBReader.GetString(3)
                        .Plant = oDBReader.GetString(4)
                        .MarkingDate = oDBReader.GetDateTime(5)
                        .MarkingTime = oDBReader.GetDateTime(6)
                        .MarkData1 = oDBReader.GetString(7)
                        .MarkData2 = oDBReader.GetString(8)
                        .MC_no = oDBReader.GetString(9)
                    End With

                    iRecNo += 1
                Loop

                lRetVal = iRecNo - 1
            Else
                lRetVal = -1
            End If
        End With

        oConn.Close()
        oConn.Dispose()

        Return lRetVal

    End Function

    Public Function DIO_Open(ByVal BoardIdx As Integer) As Integer

        Dim iRet As Integer = 0


        With TapMarkInsp.USB_IO(BoardIdx)
            Try
                iRet = YdciOpen(.iBoardNo, .sBoardName, .iHnd)
            Catch
                MessageBox.Show("Open DIO error. Please Check the hardware and make sure the driver is installed as well as the system configuration is being done properly.", "Hardware Initialize Error...", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return -1
            End Try
        End With

        If iRet = YDCI_RESULT_SUCCESS Then
            Return 0
        Else
            Return -1
        End If

    End Function

    Public Function DIO_Close(ByVal BoardIdx As Integer) As Integer

        With TapMarkInsp.USB_IO(BoardIdx)
            If .iHnd <= 0 Then
                Return 0
            End If

            Dim iRet As Integer = (YdciClose(.iHnd))

            If iRet = YDCI_RESULT_SUCCESS Then
                TapMarkInsp.USB_IO(BoardIdx).iHnd = -1
            End If

            Return iRet
        End With

    End Function

    Public Function InputPoint(ByVal BrdIdx As Integer, ByVal shBitNo As Integer) As Integer

        Dim iRetVal As Integer = -1
        Dim PointData As Integer = 0


        If TapMarkInsp.USB_IO(BrdIdx).iHnd > 0 Then
            iRetVal = YdciDioInput(TapMarkInsp.USB_IO(BrdIdx).iHnd, PointData, shBitNo, 1)

            If Not iRetVal = YDCI_RESULT_SUCCESS Then
                Return -1
            End If
        End If

        Return PointData

    End Function

    Public Function InputByte(ByVal BrdIdx As Integer) As Integer

        Dim iRetVal As Integer = -1
        Dim PointData As Integer = 0


        If TapMarkInsp.USB_IO(BrdIdx).iHnd > 0 Then
            iRetVal = YdciDioInput(TapMarkInsp.USB_IO(BrdIdx).iHnd, PointData, 0, 8)

            If Not iRetVal = YDCI_RESULT_SUCCESS Then
                Return -1
            End If
        End If

        Return PointData

    End Function

    Public Function OutputPoint(ByVal BrdIdx As Integer, ByVal shBitNo As Integer, ByVal OnOff As BitsState) As BitsState

        Dim iRetVal As Integer = -1
        Dim PointData As Byte = OnOff


        If TapMarkInsp.USB_IO(BrdIdx).iHnd > 0 Then
            iRetVal = YdciDioOutput(TapMarkInsp.USB_IO(BrdIdx).iHnd, PointData, shBitNo, 1)

            If Not iRetVal = YDCI_RESULT_SUCCESS Then
                Return -1
            End If
        End If

        Return PointData

    End Function

    Public Sub ReadRegData()

        Dim regSubKeyComm_1 As RegistryKey = regKey.CreateSubKey("Software\az_Logics\TapMarkInsp\CAM_1")


        With TapMarkInsp
            .MC_No = regSubKey.GetValue("MachineNo", "1")
            .DataPath = regSubKey.GetValue("DataPath", "C:\Data")
            .LotDataPath = regSubKey.GetValue("LotDataPath", "C:\LotEndData")

            With .dbSource
                .FileName = regSubKey.GetValue("DBFileName", "Record.mdb")
                .Path = regSubKey.GetValue("DBOurcePath", "\\172.16.59.2\epmmn\SD\LaserMarking\Record")
            End With

            With .OAI_CAM_1
                .PortName = regSubKeyComm_1.GetValue("CommPortName", "COM7")

                .BaudRate = CType(regSubKeyComm_1.GetValue("CommBaudRate"), Integer)
                .BaudRate = IIf((.BaudRate = 0), 38400, .BaudRate)

                .DataBits = CType(regSubKeyComm_1.GetValue("CommDataBits"), Integer)
                .DataBits = IIf((.DataBits = 0), 8, .DataBits)

                .StopBits = CType(regSubKeyComm_1.GetValue("CommStopBits"), System.IO.Ports.StopBits)
                .StopBits = IIf((.StopBits = 0), Ports.StopBits.One, .StopBits)

                .Parity = CType(regSubKeyComm_1.GetValue("CommParity"), System.IO.Ports.Parity)
            End With
        End With

    End Sub

End Module
