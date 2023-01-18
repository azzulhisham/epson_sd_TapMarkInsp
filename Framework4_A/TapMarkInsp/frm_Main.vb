'-----------------------------------------------------
'   Taping  : Marking Inspection System Integration
'=====================================================
'   Name    : TapMarkInsp
'   Written : Zulhisham Tan - S1609
'   Date    : 18th May '2011
'-----------------------------------------------------

Imports System.Drawing
Imports System.IO
Imports System.IO.Ports

Imports System.Math
Imports Microsoft.Win32
Imports System.Threading


Public Class frm_Main

    Delegate Sub DispMsg(ByVal FormControl As Label, ByVal ThreadName As String, ByVal Text As String)

    Private Declare Function GetAsyncKeyState Lib "user32" Alias "GetAsyncKeyState" (ByVal vKey As Integer) As Integer
    Public WithEvents TapMarkVS As SerialPort = New SerialPort

    Dim pic_BitState(15) As PictureBox


    Private Function KeyDownMon() As String

        '&H1 to ensure the key is pressed - &H1 means least significant bit. If it is ON, then the key is down.

        If GetAsyncKeyState(Keys.A) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.A) And &H1) = 0
            Loop

            Return "A"
        End If

        If GetAsyncKeyState(Keys.B) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.B) And &H1) = 0
            Loop

            Return "B"
        End If

        If GetAsyncKeyState(Keys.C) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.C) And &H1) = 0
            Loop

            Return "C"
        End If

        If GetAsyncKeyState(Keys.D) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.D) And &H1) = 0
            Loop

            Return "D"
        End If

        If GetAsyncKeyState(Keys.E) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.E) And &H1) = 0
            Loop

            Return "E"
        End If

        If GetAsyncKeyState(Keys.F) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.F) And &H1) = 0
            Loop

            Return "F"
        End If

        If GetAsyncKeyState(Keys.G) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.G) And &H1) = 0
            Loop

            Return "G"
        End If

        If GetAsyncKeyState(Keys.H) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.H) And &H1) = 0
            Loop

            Return "H"
        End If

        If GetAsyncKeyState(Keys.I) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.I) And &H1) = 0
            Loop

            Return "I"
        End If

        If GetAsyncKeyState(Keys.J) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.J) And &H1) = 0
            Loop

            Return "J"
        End If

        If GetAsyncKeyState(Keys.K) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.K) And &H1) = 0
            Loop

            Return "K"
        End If

        If GetAsyncKeyState(Keys.L) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.L) And &H1) = 0
            Loop

            Return "L"
        End If

        If GetAsyncKeyState(Keys.M) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.M) And &H1) = 0
            Loop

            Return "M"
        End If

        If GetAsyncKeyState(Keys.N) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.N) And &H1) = 0
            Loop

            Return "N"
        End If

        If GetAsyncKeyState(Keys.O) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.O) And &H1) = 0
            Loop

            Return "O"
        End If

        If GetAsyncKeyState(Keys.P) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.P) And &H1) = 0
            Loop

            Return "P"
        End If

        If GetAsyncKeyState(Keys.Q) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.Q) And &H1) = 0
            Loop

            Return "Q"
        End If

        If GetAsyncKeyState(Keys.R) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.R) And &H1) = 0
            Loop

            Return "R"
        End If

        If GetAsyncKeyState(Keys.S) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.S) And &H1) = 0
            Loop

            Return "S"
        End If

        If GetAsyncKeyState(Keys.T) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.T) And &H1) = 0
            Loop

            Return "T"
        End If

        If GetAsyncKeyState(Keys.U) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.U) And &H1) = 0
            Loop

            Return "U"
        End If

        If GetAsyncKeyState(Keys.V) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.V) And &H1) = 0
            Loop

            Return "V"
        End If

        If GetAsyncKeyState(Keys.W) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.W) And &H1) = 0
            Loop

            Return "W"
        End If

        If GetAsyncKeyState(Keys.X) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.X) And &H1) = 0
            Loop

            Return "X"
        End If

        If GetAsyncKeyState(Keys.Y) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.Y) And &H1) = 0
            Loop

            Return "Y"
        End If

        If GetAsyncKeyState(Keys.Z) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.Z) And &H1) = 0
            Loop

            Return "Z"
        End If

        If GetAsyncKeyState(Keys.Enter) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.Enter) And &H1) = 0
            Loop

            Return vbCrLf
        End If

        If GetAsyncKeyState(Keys.Tab) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.Tab) And &H1) = 0
            Loop

            Return vbCrLf
        End If

        If GetAsyncKeyState(Keys.OemMinus) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.OemMinus) And &H1) = 0
            Loop

            Return "-"
        End If

        If GetAsyncKeyState(Keys.Subtract) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.Subtract) And &H1) = 0
            Loop

            Return "-"
        End If

        If GetAsyncKeyState(Keys.Divide) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.Divide) And &H1) = 0
            Loop

            Return "/"
        End If

        If GetAsyncKeyState(Keys.OemPeriod) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.OemPeriod) And &H1) = 0
            Loop

            Return "."
        End If

        If GetAsyncKeyState(Keys.Space) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.Space) And &H1) = 0
            Loop

            Return " "
        End If

        If GetAsyncKeyState(Keys.Oemplus) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.Oemplus) And &H1) = 0
            Loop

            Return "+"
        End If

        If GetAsyncKeyState(Keys.Back) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.Back) And &H1) = 0
            Loop

            Return "---"
        End If

        If GetAsyncKeyState(Keys.LButton) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.LButton) And &H1) = 0
            Loop

            Return vbCrLf
            'Return "|1|"
        End If


        'If GetAsyncKeyState(Keys.RButton) And &H1 Then
        '    Do Until (GetAsyncKeyState(Keys.RButton) And &H1) = 0
        '    Loop

        '    Return "|2|"
        'End If

        If GetAsyncKeyState(Keys.D0) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.D0) And &H1) = 0
            Loop

            Return "0"
        End If

        If GetAsyncKeyState(Keys.D1) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.D1) And &H1) = 0
            Loop

            Return "1"
        End If

        If GetAsyncKeyState(Keys.D2) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.D2) And &H1) = 0
            Loop

            Return "2"
        End If

        If GetAsyncKeyState(Keys.D3) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.D3) And &H1) = 0
            Loop

            Return "3"
        End If

        If GetAsyncKeyState(Keys.D4) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.D4) And &H1) = 0
            Loop

            Return "4"
        End If

        If GetAsyncKeyState(Keys.D5) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.D5) And &H1) = 0
            Loop

            Return "5"
        End If

        If GetAsyncKeyState(Keys.D6) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.D6) And &H1) = 0
            Loop

            Return "6"
        End If

        If GetAsyncKeyState(Keys.D7) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.D7) And &H1) = 0
            Loop

            Return "7"
        End If

        If GetAsyncKeyState(Keys.D8) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.D8) And &H1) = 0
            Loop

            Return "8"
        End If

        If GetAsyncKeyState(Keys.D9) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.D9) And &H1) = 0
            Loop

            Return "9"
        End If

        If GetAsyncKeyState(Keys.NumPad0) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.NumPad0) And &H1) = 0
            Loop

            Return "0"
        End If

        If GetAsyncKeyState(Keys.NumPad1) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.NumPad1) And &H1) = 0
            Loop

            Return "1"
        End If

        If GetAsyncKeyState(Keys.NumPad2) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.NumPad2) And &H1) = 0
            Loop

            Return "2"
        End If

        If GetAsyncKeyState(Keys.NumPad3) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.NumPad3) And &H1) = 0
            Loop

            Return "3"
        End If

        If GetAsyncKeyState(Keys.NumPad4) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.NumPad4) And &H1) = 0
            Loop

            Return "4"
        End If

        If GetAsyncKeyState(Keys.NumPad5) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.NumPad5) And &H1) = 0
            Loop

            Return "5"
        End If

        If GetAsyncKeyState(Keys.NumPad6) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.NumPad6) And &H1) = 0
            Loop

            Return "6"
        End If

        If GetAsyncKeyState(Keys.NumPad7) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.NumPad7) And &H1) = 0
            Loop

            Return "7"
        End If

        If GetAsyncKeyState(Keys.NumPad8) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.NumPad8) And &H1) = 0
            Loop

            Return "8"
        End If

        If GetAsyncKeyState(Keys.NumPad9) And &H1 Then
            Do Until (GetAsyncKeyState(Keys.NumPad9) And &H1) = 0
            Loop

            Return "9"
        End If

        If GetAsyncKeyState(Asc("0")) And &H1 Then
            Do Until (GetAsyncKeyState(Asc("0")) And &H1) = 0
            Loop

            Return "0"
        End If

        If GetAsyncKeyState(Asc("1")) And &H1 Then
            Do Until (GetAsyncKeyState(Asc("1")) And &H1) = 0
            Loop

            Return "1"
        End If

        If GetAsyncKeyState(Asc("2")) And &H1 Then
            Do Until (GetAsyncKeyState(Asc("2")) And &H1) = 0
            Loop

            Return "2"
        End If

        If GetAsyncKeyState(Asc("3")) And &H1 Then
            Do Until (GetAsyncKeyState(Asc("3")) And &H1) = 0
            Loop

            Return "3"
        End If

        If GetAsyncKeyState(Asc("4")) And &H1 Then
            Do Until (GetAsyncKeyState(Asc("4")) And &H1) = 0
            Loop

            Return "4"
        End If

        If GetAsyncKeyState(Asc("5")) And &H1 Then
            Do Until (GetAsyncKeyState(Asc("5")) And &H1) = 0
            Loop

            Return "5"
        End If

        If GetAsyncKeyState(Asc("6")) And &H1 Then
            Do Until (GetAsyncKeyState(Asc("6")) And &H1) = 0
            Loop

            Return "6"
        End If

        If GetAsyncKeyState(Asc("7")) And &H1 Then
            Do Until (GetAsyncKeyState(Asc("7")) And &H1) = 0
            Loop

            Return "7"
        End If

        If GetAsyncKeyState(Asc("8")) And &H1 Then
            Do Until (GetAsyncKeyState(Asc("8")) And &H1) = 0
            Loop

            Return "8"
        End If

        If GetAsyncKeyState(Asc("9")) And &H1 Then
            Do Until (GetAsyncKeyState(Asc("9")) And &H1) = 0
            Loop

            Return "9"
        End If

        Return ""

    End Function

    Private Sub frm_Main_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        With Me
            .tmr_Update.Enabled = False
            .tmr_KeyScanner.Enabled = False
            .tmr_IOScanner.Enabled = False

            If .TapMarkVS.IsOpen Then .TapMarkVS.Close()

#If dbg = 0 Then
            If TapMarkInsp.USB_IO(0).iHnd > 0 Then DIO_Close(0)
#End If

        End With

    End Sub

    Private Sub frm_Main_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim pt As Point
        TapMarkInsp.BootState = Func_Ret_Fail

        With pt
            .X = 670
            .Y = 18
        End With

        With Me
            'Set Outlook
            .Location = pt
            .Width = 460
            .Height = 190
            .TopMost = True

            'Set Child Control
#If dbg = 0 Then
            .btn_Meas.Visible = False
#End If

            .ToolTip1.SetToolTip(Me.PictureBox2, String.Format("Version No. : {0:D4}-{1:D4}-{2:D3}-{3:D3}", My.Application.Info.Version.Major, My.Application.Info.Version.Minor, My.Application.Info.Version.MajorRevision, My.Application.Info.Version.Revision))

            'Initialize Control
            For ilp = 0 To .pic_BitState.GetUpperBound(0)
                Application.DoEvents()
                .pic_BitState(ilp) = New PictureBox
            Next

            .pic_BitState(0) = .PictureBox3
            .pic_BitState(1) = .PictureBox4
            .pic_BitState(2) = .PictureBox5
            .pic_BitState(3) = .PictureBox6
            .pic_BitState(4) = .PictureBox7
            .pic_BitState(5) = .PictureBox8
            .pic_BitState(6) = .PictureBox9
            .pic_BitState(7) = .PictureBox10
            .pic_BitState(8) = .PictureBox11
            .pic_BitState(9) = .PictureBox12
            .pic_BitState(10) = .PictureBox13
            .pic_BitState(11) = .PictureBox14
            .pic_BitState(12) = .PictureBox15
            .pic_BitState(13) = .PictureBox16
            .pic_BitState(14) = .PictureBox17
            .pic_BitState(15) = .PictureBox18

            GrnLED_OnOff(0) = .PictureBox3.BackColor
            GrnLED_OnOff(1) = .PictureBox20.BackColor
            RedLED_OnOff(0) = .PictureBox11.BackColor
            RedLED_OnOff(1) = .PictureBox19.BackColor

            For iLp As Integer = 0 To .pic_BitState.GetUpperBound(0)
                Application.DoEvents()
                .pic_BitState(iLp).BackColor = IIf(iLp < 8, GrnLED_OnOff(0), RedLED_OnOff(0))
            Next

            With .lbl_DataInput
                .Text = ""
            End With

            With TapMarkInsp.SettingLot
                ReDim .dbData(0)

                With .dbData(0)
                    .MarkData1 = "X3500"
                    .MarkData2 = "ABCDEF"
                End With
            End With


            ReadRegData()

#If dbg = 0 Then
            'Initialize Serial Port
            If InitSerialPort() = Func_Ret_Fail Then
                MessageBox.Show("The serial communication is fail to initialize. The program can not be function and will close immediately. Please re-confirm the communication hardware is working properly or consult the system engineer for further advise.", My.Application.Info.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            'Initialize USB-IO
            ReDim TapMarkInsp.USB_IO(0)

            With TapMarkInsp
                With .USB_IO(0)
                    .sBoardType = "USB Digital I/O"
                    .sBoardName = "DIO-8/8B-UBC"
                    .sBoardMaker = "Y2C"
                    .iBoardNo = 0
                    .iHnd = -1
                End With
            End With

            If DIO_Open(0) < 0 Then
                MessageBox.Show("The Digital I/O is fail to initialize. The program can not be function and will close immediately. Please re-confirm the DIO hardware is working properly or consult the system engineer for further advise.", My.Application.Info.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            .sub_Reset.PerformClick()
#End If

            With .tmr_KeyScanner
                .Interval = 1
                .Enabled = True
            End With

#If dbg = 0 Then
            With .tmr_IOScanner
                .Interval = 30
                .Enabled = True
            End With
#End If
        End With

        TapMarkInsp.BootState = Func_Ret_Success

    End Sub

    Private Sub tmr_KeyScanner_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tmr_KeyScanner.Tick

        'TapMarkInsp.RunningState = fg_OFF               'For Debug - Should be remarked when deploy!

#If dbg = 0 Then
        If TapMarkInsp.RunningState = fg_ON Then Exit Sub
#End If
        'Application.DoEvents()

        Dim KeyStroke As String = KeyDownMon()

        With Me.lbl_DataInput
            If Not KeyStroke = "" Then
                If KeyStroke = "---" Then
                    If .Text.Length > 0 Then
                        .Text = .Text.Substring(0, .Text.Length - 1)
                    End If
                Else
                    .Text &= KeyStroke
                End If
            End If
        End With

    End Sub

    Private Sub lbl_DataInput_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbl_DataInput.TextChanged

        With Me
            If (.lbl_DataInput.Text.EndsWith(vbCrLf) Or .lbl_DataInput.Text.EndsWith("|1|")) And .lbl_DataInput.Text.Trim <> "" Then
                If .lbl_DataInput.Text.Trim.Contains("|1|") Then
                    .lbl_DataInput.Text = .lbl_DataInput.Text.Trim.Replace("|1|", "")
                End If

                .lbl_DataInput.Text = .lbl_DataInput.Text.Replace(vbLf, "")
                .lbl_DataInput.Text = .lbl_DataInput.Text.Replace(vbCr, "")

                If .lbl_DataInput.Text <> "" Then
                    If .lbl_DataInput.Text.Trim.ToUpper.StartsWith("-FX3") AndAlso (.lbl_DataInput.Text.IndexOf("-", 1) = 6 Or .lbl_DataInput.Text.Trim.Substring(2).Length = 10) Then
                        .lbl_DataInput.Text = .lbl_DataInput.Text.Substring(2) & .lbl_DataInput.Text.Substring(0, 2)
                    End If

                    If .lbl_DataInput.Text.ToUpper.StartsWith("CAL-") Or (.lbl_DataInput.Text.ToUpper.StartsWith("V") AndAlso .lbl_DataInput.Text.ToUpper.EndsWith("A") AndAlso Not IsNumeric(.lbl_DataInput.Text)) Then
                        If Not lbl_DataInput.Text.ToUpper.StartsWith("CAL") Or Not TapMarkInsp.SettingLot.LotNo = "" Then
                            .sub_Reset.PerformClick()
                        End If

                        .lbl_LotNo.Text = .lbl_DataInput.Text.Trim.ToUpper
                        TapMarkInsp.SettingLot.LotNo_ = .lbl_DataInput.Text.Trim.ToUpper
                        TapMarkInsp.SettingLot.LotNo = TapMarkInsp.SettingLot.LotNo_.Substring(0, TapMarkInsp.SettingLot.LotNo_.Length - 1)

                        If lbl_DataInput.Text.ToUpper.StartsWith("CAL") Then
                            TapMarkInsp.SettingLot.LotNo = .lbl_DataInput.Text.Trim.ToUpper

                            With TapMarkInsp.SettingLot
                                ReDim .dbData(0)

                                With .dbData(0)
                                    .Freq = "35"
                                    .IMISpec = "U0000003"
                                    .MarkData1 = "X3500"
                                    .MarkData2 = "ABCDEF"
                                End With

                                'Me.ChgScene(31)

                                OutputPoint(0, 2, fg_ON)
                                Me.pic_BitState(10).BackColor = RedLED_OnOff(fg_ON)
                                Me.ToolTip1.SetToolTip(Me.lbl_MarkingData, "<" & .dbData(0).MarkData1 & "> <" & .dbData(0).MarkData2 & ">")
                            End With
                        End If
                    ElseIf .lbl_DataInput.Text.ToUpper.StartsWith("CAL-") Or (.lbl_DataInput.Text.ToUpper.StartsWith("V") AndAlso Not .lbl_DataInput.Text.ToUpper.EndsWith("A") AndAlso Not IsNumeric(.lbl_DataInput.Text)) Then
                        .lbl_LotNo.Text = .lbl_DataInput.Text.Trim.ToUpper
                        TapMarkInsp.SettingLot.LotNo = .lbl_DataInput.Text.Trim.ToUpper

                        If lbl_DataInput.Text.ToUpper.StartsWith("CAL") Then
                            TapMarkInsp.SettingLot.LotNo_ = .lbl_DataInput.Text.Trim.ToUpper
                        End If
                    ElseIf (.lbl_DataInput.Text.ToUpper.StartsWith("CAL") And .lbl_DataInput.Text.ToUpper.Length = 3) Or (.lbl_DataInput.Text.ToUpper.StartsWith("U") AndAlso .lbl_DataInput.Text.Length >= 8 AndAlso Not IsNumeric(.lbl_DataInput.Text)) Then
                        .lbl_SpecNo.Text = .lbl_DataInput.Text.Trim.ToUpper
                        TapMarkInsp.SettingLot.SpecNo = .lbl_DataInput.Text.Trim.ToUpper

                        If lbl_DataInput.Text.ToUpper.StartsWith("CAL") Then
                            Me.ChgScene(31)
                        Else
                            'Me.ChgScene(14)
                        End If
                    ElseIf (((.lbl_DataInput.Text.ToUpper.StartsWith("X") Or .lbl_DataInput.Text.ToUpper.StartsWith("TR")) AndAlso .lbl_DataInput.Text.Length >= 7) Or _
                        ((.lbl_DataInput.Text.ToUpper.StartsWith("E")) AndAlso .lbl_DataInput.Text.Length = 10)) _
                        AndAlso Not IsNumeric(.lbl_DataInput.Text) Then

                        If .lbl_DataInput.Text.IndexOf("-") < 0 And .lbl_DataInput.Text.Trim.Length <> 10 Then
                            If .lbl_DataInput.Text.ToUpper.StartsWith("X") Then
                                If Not .lbl_DataInput.Text.ToUpper.Contains("N") And .lbl_DataInput.Text.ToUpper.IndexOf("TEST") < 0 And .lbl_DataInput.Text.ToUpper.IndexOf("SAM") < 0 Then
                                    .lbl_DataInput.Text = .lbl_DataInput.Text.Substring(0, 4) & "-" & .lbl_DataInput.Text.Substring(4)
                                End If
                            End If
                        End If

                        If Not .lbl_DataInput.Text.ToUpper.IndexOf("TEST") < 0 Or Not .lbl_DataInput.Text.ToUpper.IndexOf("SAM") < 0 Then
                            If Not .lbl_DataInput.Text.ToUpper.Contains(" ") Then
                                If Not .lbl_DataInput.Text.ToUpper.IndexOf("TEST") < 0 And Not .lbl_DataInput.Text.Contains("-") Then
                                    .lbl_DataInput.Text = .lbl_DataInput.Text.ToUpper.Replace("TEST", " TEST")
                                End If

                                If Not .lbl_DataInput.Text.ToUpper.IndexOf("SAM") < 0 And Not .lbl_DataInput.Text.Contains("-") Then
                                    .lbl_DataInput.Text = .lbl_DataInput.Text.ToUpper.Replace("SAM", " SAM")
                                End If
                            End If
                        End If


                        Dim NewLotData As Boolean = IIf(Not TapMarkInsp.SettingLot.dbData(0).LotNo = Nothing, False, True)

                        .lbl_PLotNo.Text = .lbl_DataInput.Text.Trim.ToUpper
                        TapMarkInsp.SettingLot.M_LotNo = .lbl_DataInput.Text.Trim.ToUpper   '.Substring(0, 9)

                        With TapMarkInsp
                            With .SettingLot
                                If Not String.IsNullOrEmpty(.LotNo_) AndAlso Not String.IsNullOrEmpty(.LotNo) AndAlso Not .LotNo_.IndexOf(.LotNo) < 0 AndAlso .LotNo_.IndexOf("---") < 0 AndAlso .LotNo.IndexOf("---") < 0 Then
                                    TapMarkInsp.SettingLot.InspDirection = "-L"

                                    If lbl_DataInput.Text.ToUpper.StartsWith("CAL") Then
                                        ReDim .dbData(0)

                                        With .dbData(0)
                                            .Freq = "35"
                                            .IMISpec = "U0000003"
                                            .MarkData1 = "X3500"
                                            .MarkData2 = "ABCDEF"
                                        End With

                                        'Me.ChgScene(31)

                                        OutputPoint(0, 2, fg_ON)
                                        Me.pic_BitState(10).BackColor = RedLED_OnOff(fg_ON)
                                        Me.ToolTip1.SetToolTip(Me.lbl_MarkingData, "<" & .dbData(0).MarkData1 & "> <" & .dbData(0).MarkData2 & ">")
                                    Else
                                        Dim iRetRec As Integer = -1

                                        If NewLotData = True Then
                                            iRetRec = Read_DataBase(TapMarkInsp.dbSource, TapMarkInsp.SettingLot.M_LotNo, .dbData)

                                            If Not iRetRec < 0 Then
                                                Try
                                                    Dim dbMarkingData As String = "<" & .dbData(.dbData.GetUpperBound(0)).MarkData1 & "> <" & .dbData(.dbData.GetUpperBound(0)).MarkData2 & ">"

                                                    Me.ToolTip1.SetToolTip(Me.lbl_MarkingData, dbMarkingData)
                                                    Me.lbl_DBData.Text = dbMarkingData
                                                Catch ex As Exception

                                                End Try
                                            End If

                                        Else
                                            Dim addData() As InspData = Nothing
                                            iRetRec = Read_DataBase(TapMarkInsp.dbSource, TapMarkInsp.SettingLot.M_LotNo, addData)

                                            If Not iRetRec < 0 Then
                                                iRetRec = -1

                                                For ilp As Integer = 0 To addData.GetUpperBound(0)
                                                    If .SpecNo.Trim.Contains(addData(ilp).IMISpec) OrElse .dbData(.dbData.GetUpperBound(0)).IMISpec.Contains(addData(ilp).IMISpec) Then
                                                        ReDim Preserve .dbData(.dbData.Length)

                                                        With .dbData(.dbData.GetUpperBound(0))
                                                            .Freq = addData(ilp).Freq
                                                            .IMISpec = addData(ilp).IMISpec
                                                            .LotNo = addData(ilp).LotNo
                                                            .MarkData1 = addData(ilp).MarkData1
                                                            .MarkData2 = addData(ilp).MarkData2
                                                            .MarkingDate = addData(ilp).MarkingDate
                                                            .MarkingTime = addData(ilp).MarkingTime
                                                            .MC_no = addData(ilp).MC_no
                                                            .Opt = addData(ilp).Opt
                                                            .Plant = addData(ilp).Plant
                                                        End With

                                                        iRetRec = ilp
                                                        Exit For
                                                    End If
                                                Next

                                                Try
                                                    Dim dbMarkingData As String = "<" & .dbData(.dbData.GetUpperBound(0)).MarkData1 & "> <" & .dbData(.dbData.GetUpperBound(0)).MarkData2 & ">"

                                                    Me.ToolTip1.SetToolTip(Me.lbl_MarkingData, dbMarkingData)
                                                    Me.lbl_DBData.Text = dbMarkingData
                                                Catch ex As Exception

                                                End Try
                                            End If
                                        End If

#If dbg = 0 Then
                                        If iRetRec < 0 Then     'No Data -> Not Rdy For Auto
                                            OutputPoint(0, 2, fg_OFF)
                                            Me.pic_BitState(10).BackColor = RedLED_OnOff(fg_OFF)
                                            Me.ToolTip1.SetToolTip(Me.lbl_MarkingData, "No Data !")
                                        Else                    'Data OK -> Rdy For Auto
                                            'Set Camera Scene No.
                                            Me.ChgScene(CameraSceneNoStartAt + .dbData(.dbData.GetUpperBound(0)).MarkData1.Trim.Length - MarkingLine1Char)

                                            OutputPoint(0, 2, fg_ON)
                                            Me.pic_BitState(10).BackColor = RedLED_OnOff(fg_ON)

                                            Try
                                                Dim dbMarkingData As String = "<" & .dbData(.dbData.GetUpperBound(0)).MarkData1 & "> <" & .dbData(.dbData.GetUpperBound(0)).MarkData2 & ">"

                                                Me.ToolTip1.SetToolTip(Me.lbl_MarkingData, dbMarkingData)
                                                Me.lbl_DBData.Text = dbMarkingData
                                            Catch ex As Exception

                                            End Try
                                        End If
#End If

                                    End If
                                End If
                            End With
                        End With
                    ElseIf IsNumeric(.lbl_DataInput.Text) AndAlso .lbl_DataInput.Text.Length = 4 AndAlso Not .lbl_DataInput.Text.StartsWith("0") Then
                        .lbl_Qty.Text = .lbl_DataInput.Text.Trim
                        TapMarkInsp.SettingLot.InputQty = .lbl_DataInput.Text.Trim
                    ElseIf .lbl_DataInput.Text.ToUpper.StartsWith("-") AndAlso Not IsNumeric(.lbl_DataInput.Text) Then
                        If .lbl_DataInput.Text.ToUpper.Substring(1, 1) = "L" Or .lbl_DataInput.Text.ToUpper.Substring(1, 1) = "R" Then
                            .lbl_Dir.Text = .lbl_DataInput.Text.Trim.ToUpper.Substring(0, 2)
                            TapMarkInsp.SettingLot.InspDirection = .lbl_DataInput.Text.Trim.ToUpper.Substring(0, 2)
                        End If
                    Else
                        .lbl_Opt.Text = .lbl_DataInput.Text.Trim.ToUpper
                        TapMarkInsp.SettingLot.OperatorNo = .lbl_DataInput.Text.Trim.ToUpper
                    End If
                End If

                .lbl_DataInput.Text = ""
            End If
        End With

    End Sub

    Private Sub sub_Exit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles sub_Exit.Click

        If Not TapMarkInsp.BootState < 0 Then
            If MessageBox.Show("To be reminded that the system is not design to terminate in the method as this is. The system reliability will no longer accurate. Are you sure you want to close the application now?", "TapMarkInsp...", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If
        End If

        With Me
            .tmr_KeyScanner.Enabled = False
            .tmr_IOScanner.Enabled = False

            If .TapMarkVS.IsOpen Then .TapMarkVS.Close()

#If dbg = 0 Then
            If TapMarkInsp.USB_IO(0).iHnd > 0 Then DIO_Close(0)
#End If

            With .tmr_Fade
                .Interval = 30
                .Enabled = True
            End With
        End With

    End Sub

    Private Sub sub_Reset_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles sub_Reset.Click

        OutputPoint(0, 2, fg_OFF)
        Me.pic_BitState(10).BackColor = RedLED_OnOff(fg_OFF)

        With Me
            .lbl_Dir.Text = "---"
            .lbl_LotNo.Text = "---"
            .lbl_Opt.Text = "---"
            .lbl_PLotNo.Text = "---"
            .lbl_Qty.Text = "---"
            .lbl_SpecNo.Text = "---"
        End With

        With TapMarkInsp.SettingLot
            .InputQty = ""
            .InspDirection = ""
            .LotNo = ""
            .LotNo_ = ""
            .M_LotNo = ""
            .OperatorNo = ""
            .SpecNo = ""
            .WeekCdChk = -1

            ReDim .dbData(0)
        End With

    End Sub

    Private Function InitSerialPort() As Integer

        Dim lng_RetVal As Long = Func_Ret_Fail


        Try
            With TapMarkVS
                If .IsOpen = True Then
                    .Close()
                End If

                .PortName = TapMarkInsp.OAI_CAM_1.PortName
                .Parity = TapMarkInsp.OAI_CAM_1.Parity
                .BaudRate = TapMarkInsp.OAI_CAM_1.BaudRate
                .StopBits = TapMarkInsp.OAI_CAM_1.StopBits
                .DataBits = TapMarkInsp.OAI_CAM_1.DataBits
                .ReceivedBytesThreshold = 1
                .ReadBufferSize = 1024

                '.Open()
            End With
        Catch
            Return lng_RetVal
        End Try

        Return Func_Ret_Success

    End Function

    Private Sub frm_Main_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown

        If TapMarkInsp.BootState < 0 Then
            Me.Close()
            Exit Sub
        End If

        With Me
            With .tmr_Update
                .Interval = 3000
                .Enabled = True
            End With
        End With

    End Sub

    Private Sub tmr_Fade_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tmr_Fade.Tick

        With Me
            If .Opacity <= 0 Then
                .tmr_Fade.Enabled = False
                .Close()
            End If

            .Opacity -= 0.025
        End With

    End Sub

    Private Sub tmr_IOScanner_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tmr_IOScanner.Tick

        With Me
            If Not TapMarkInsp.USB_IO(0).iHnd > 0 Then Exit Sub
            .tmr_IOScanner.Enabled = False


            For ilp As Integer = 0 To 7
                Application.DoEvents()

                Dim ReadInPort As Integer = InputPoint(0, ilp)
                'Dim IO_State As Integer = (ReadInPort And (2 ^ ilp)) / (2 ^ ilp)
                .pic_BitState(ilp).BackColor = GrnLED_OnOff(ReadInPort)

                Select Case ilp
                    Case Is = 0             'Machine Operation State
                        TapMarkInsp.RunningState = ReadInPort
                    Case Is = 1             'Triggering Status
                        '.btn_Meas.PerformClick()
                        If ReadInPort <> 0 Then .TrgVS()
                    Case Is = 2, 4         'Vision Control Status
                        OutputPoint(0, ilp + 1, ReadInPort)
                        .pic_BitState(ilp + 9).BackColor = RedLED_OnOff(ReadInPort)
                    Case Is = 5
                        OutputPoint(0, 7, ReadInPort)
                        .pic_BitState(15).BackColor = RedLED_OnOff(ReadInPort)
                    Case Is = 6
                        If ReadInPort <> 0 Then
                            If TapMarkInsp.SettingLot.SpecNo.ToUpper.StartsWith("CAL") Then
                                Me.ChgScene(31)
                            Else
                                'Set Camera Scene No.
                                'Me.ChgScene(14)
                                Me.ChgScene(CameraSceneNoStartAt + TapMarkInsp.SettingLot.dbData(0).MarkData1.Trim.Length - MarkingLine1Char)
                            End If
                        End If
                End Select
            Next

            .tmr_IOScanner.Enabled = True
        End With

    End Sub

    Private Sub TrgMeasThread()

        'New Thread To Update Records
        Dim tsk_Measure As Thread = New Thread(AddressOf InspMarkingChar)

        With tsk_Measure
            .Name = "tsk_Measure"
            .Start()
        End With

    End Sub

    Private Sub InspMarkingChar()

        Dim iFlg As Integer = 0
        Dim VS_Data As String = String.Empty


        With Me
            With .lbl_MarkingData
                If .InvokeRequired Then
                    .Invoke(New frm_Main.DispMsg(AddressOf Me.LabelCtrlUpDate), New Object() {Me.lbl_MarkingData, Thread.CurrentThread.Name, VS_Data})
                Else
                    .Text = VS_Data
                End If
            End With
        End With

    End Sub

    Public Sub LabelCtrlUpDate(ByVal FormControl As Label, ByVal ThreadName As String, ByVal Text As String)

        Try
            FormControl.Text = Text
        Catch ex As Exception

        End Try

    End Sub

    Public Function FZ_SerCmd(ByVal strCmd As String, Optional ByRef strMeasDat As String = "") As Integer

        Dim str_Buffer As String = String.Empty
        Dim Buffer() As Char = ""


        With Me.TapMarkVS
            If .IsOpen = False Then
                .Open()
                Thread.Sleep(100)
            End If

            .DiscardInBuffer()
            str_Buffer = strCmd & vbCrLf
            .Write(str_Buffer)


            'Simply Delay Operation
            Thread.Sleep(80)

            Dim WaitReplyTimer As Integer = My.Computer.Clock.TickCount
            Dim fgData As Integer = 0

            Do Until Not .BytesToRead = 0
                Application.DoEvents()

                If My.Computer.Clock.TickCount > WaitReplyTimer + 800 Then
                    fgData = -1 : Exit Do
                End If
            Loop

            If fgData < 0 Then Return Func_Ret_Fail

            Dim ReadByteSize As Integer = .BytesToRead
            WaitReplyTimer = My.Computer.Clock.TickCount
            str_Buffer = ""


            If Not strCmd.IndexOf("MEASURE") < 0 Then
                If Not strCmd.IndexOf("/E") < 0 Or Not strCmd.IndexOf("/C") < 0 Then
                    Do Until ReadByteSize = 0
                        ReDim Buffer(ReadByteSize)
                        .Read(Buffer, 0, ReadByteSize)

                        For int_Dmy = 0 To Buffer.GetUpperBound(0)
                            'Application.DoEvents()

                            If Not Buffer(int_Dmy) = Nothing Then
                                If Buffer(int_Dmy) = vbCr Then
                                    str_Buffer &= " "
                                ElseIf Buffer(int_Dmy) = vbLf Then
                                    str_Buffer &= ""
                                Else
                                    str_Buffer &= Buffer(int_Dmy)
                                End If
                            End If
                        Next

                        ReadByteSize = .BytesToRead
                    Loop
                Else
                    Dim WaitVC_Respond As Integer = 300
                    Dim TimeOver As Integer = 0


                    Do Until Not str_Buffer.IndexOf("OK") < 0 Or Not str_Buffer.IndexOf("ER") < 0
                        If My.Computer.Clock.TickCount > WaitReplyTimer + WaitVC_Respond And .BytesToRead = 0 Then TimeOver = 1 : Exit Do

                        If Not ReadByteSize = 0 Then
                            ReDim Buffer(ReadByteSize)
                            .Read(Buffer, 0, ReadByteSize)

                            For int_Dmy = 0 To Buffer.GetUpperBound(0)
                                'Application.DoEvents()

                                If Not Buffer(int_Dmy) = Nothing Then
                                    If Buffer(int_Dmy) = vbCr Then
                                        str_Buffer &= " "
                                    ElseIf Buffer(int_Dmy) = vbLf Then
                                        str_Buffer &= ""
                                    Else
                                        str_Buffer &= Buffer(int_Dmy)
                                    End If
                                End If
                            Next
                        End If

                        ReadByteSize = .BytesToRead
                    Loop

                    'Respond String
                    '"    0,   -1,    0,   69,   88,   51,   53,   48,   48,   65,   57,   56,   76,   51,   67,   66,   77,   56,   57,   56,   56,   82,   81,   83,   85,   56,   71,    0 OK ER"

                    If Not TimeOver = 0 Then
                        Return -1
                    End If
                End If
            Else
                Do Until ReadByteSize = 0
                    ReDim Buffer(ReadByteSize)
                    .Read(Buffer, 0, ReadByteSize)

                    For int_Dmy = 0 To Buffer.GetUpperBound(0)
                        'Application.DoEvents()

                        If Not Buffer(int_Dmy) = Nothing Then
                            If Buffer(int_Dmy) = vbCr Then
                                str_Buffer &= " "
                            ElseIf Buffer(int_Dmy) = vbLf Then
                                str_Buffer &= ""
                            Else
                                str_Buffer &= Buffer(int_Dmy)
                            End If
                        End If
                    Next

                    ReadByteSize = .BytesToRead
                Loop
            End If

            str_Buffer = str_Buffer.Trim
            strMeasDat = str_Buffer

            Return str_Buffer.Length
        End With

    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Meas.Click

        'Dim VS_Data As String = String.Empty
        'Dim FuncRet As Integer = FZ_SerCmd("MEASURE", VS_Data)

        TrgVS()

        'Me.lbl_DataInput.Text = "X3B8-0378" & vbCrLf

    End Sub

    Private Sub ChgScene(Optional ByVal SceneNo As Integer = 14)

        Dim VS_Data As String = String.Empty


        With Me
            OutputPoint(0, 4, fg_ON)
            .pic_BitState(12).BackColor = RedLED_OnOff(fg_ON)

            Dim FuncRet As Integer = FZ_SerCmd("SCENE " & SceneNo.ToString.Trim, VS_Data)

            OutputPoint(0, 4, fg_OFF)
            .pic_BitState(12).BackColor = RedLED_OnOff(fg_OFF)
        End With

        Do Until InputPoint(0, 6) = 0
            Application.DoEvents()
        Loop

    End Sub


    Private Sub TrgVS()

        Dim VS_Data As String = String.Empty


        With Me
            OutputPoint(0, 4, fg_ON)
            .pic_BitState(12).BackColor = RedLED_OnOff(fg_ON)

            OutputPoint(0, 0, fg_OFF)
            .pic_BitState(8).BackColor = RedLED_OnOff(fg_OFF)

            OutputPoint(0, 1, fg_OFF)
            .pic_BitState(9).BackColor = RedLED_OnOff(fg_OFF)


            If Not TapMarkInsp.SettingLot.LotNo.ToUpper.IndexOf("CAL") < 0 Then
                Thread.Sleep(180)

                OutputPoint(0, 0, fg_ON)
                .pic_BitState(8).BackColor = RedLED_OnOff(fg_ON)

                OutputPoint(0, 4, fg_OFF)
                .pic_BitState(12).BackColor = RedLED_OnOff(fg_OFF)

                Do Until InputPoint(0, 1) = 0
                    Application.DoEvents()
                Loop

                Exit Sub
            End If


            '--- Trigger Vision System ---
            Dim FuncRet As Integer = FZ_SerCmd("MEASURE", VS_Data)
            'VS_Data = "    0,   -1,    0,   69,   88,   51,   53,   48,   48,   65,   57,   56,   76,   51,   67,   66,   77,   56,   57,   56,   56,   82,   81,   83,   85,   56,   71,    0 OK ER"

            If FuncRet < 0 Then
                VS_Data = "Error"
            End If

            If VS_Data.IndexOf(",") < 0 Then
                OutputPoint(0, 0, fg_OFF)
                .pic_BitState(8).BackColor = RedLED_OnOff(fg_OFF)

                OutputPoint(0, 1, fg_OFF)
                .pic_BitState(9).BackColor = RedLED_OnOff(fg_OFF)

                OutputPoint(0, 4, fg_OFF)
                .pic_BitState(12).BackColor = RedLED_OnOff(fg_OFF)

                .lbl_MarkingData.Text = VS_Data

                Do Until InputPoint(0, 1) = 0
                    Application.DoEvents()
                Loop

                Exit Sub
            End If


            Dim VS_Data_Chk() As String = VS_Data.Split(","c)
            Dim vsRealData As String = String.Empty


            If VS_Data_Chk Is Nothing Or VS_Data_Chk.Length < 3 Then
                OutputPoint(0, 0, fg_OFF)
                .pic_BitState(8).BackColor = RedLED_OnOff(fg_OFF)

                OutputPoint(0, 1, fg_OFF)
                .pic_BitState(9).BackColor = RedLED_OnOff(fg_OFF)

                OutputPoint(0, 4, fg_OFF)
                .pic_BitState(12).BackColor = RedLED_OnOff(fg_OFF)

                .lbl_MarkingData.Text = VS_Data

                Do Until InputPoint(0, 1) = 0
                    Application.DoEvents()
                Loop

                Exit Sub
            End If


            For iLp As Integer = 3 To VS_Data_Chk.GetUpperBound(0)
                Application.DoEvents()

                Try
                    If IsNumeric(VS_Data_Chk(iLp)) Then
                        vsRealData &= Chr(VS_Data_Chk(iLp))
                    End If
                Catch ex As Exception

                End Try
            Next


            'Display Inspection Data
            .lbl_MarkingData.Text = vsRealData


            'Data Judgment
            'vsRealData.Length = 24 And VS_Data_Chk(2).Trim = "0" And (vsRealData.Substring(0, 1) = "E" Or vsRealData.Substring(12, 1) = "E")

            If vsRealData.Length >= 3 Then
                'Chr. Str. OK
                Dim chk0 As Integer = -1

                If TapMarkInsp.SettingLot.InspDirection = "-L" Then
                    If VS_Data_Chk(0).Trim = "0" Then chk0 = 0
                ElseIf TapMarkInsp.SettingLot.InspDirection = "-R" Then
                    If VS_Data_Chk(1).Trim = "0" Then chk0 = 0
                End If


                Dim chk1 As Integer = -1
                Dim chk2 As Integer = -1

                For iLp As Integer = 0 To TapMarkInsp.SettingLot.dbData.GetUpperBound(0)
                    Application.DoEvents()

                    If chk1 < 0 Then chk1 = vsRealData.IndexOf(TapMarkInsp.SettingLot.dbData(iLp).MarkData1)

                    If chk2 < 0 Then
                        chk2 = vsRealData.IndexOf(TapMarkInsp.SettingLot.dbData(iLp).MarkData2)

                        If Not chk2 < 0 Then
                            If TapMarkInsp.SettingLot.WeekCdChk < 0 Then
                                TapMarkInsp.SettingLot.WeekCdChk = iLp
                            Else
                                If TapMarkInsp.SettingLot.WeekCdChk <= iLp Then
                                    TapMarkInsp.SettingLot.WeekCdChk = iLp
                                    .ToolTip1.SetToolTip(.lbl_PLotNo, "Week Code OK!!!")
                                Else
                                    chk2 = -1
                                    .ToolTip1.SetToolTip(.lbl_PLotNo, "Week Code Jump!!!")
                                End If
                            End If
                        End If
                    End If
                Next

                If chk0 < 0 Or chk1 < 0 Or chk2 < 0 Then
                    'NG
                    OutputPoint(0, 0, fg_OFF)
                    .pic_BitState(8).BackColor = RedLED_OnOff(fg_OFF)

                    OutputPoint(0, 1, fg_OFF)
                    .pic_BitState(9).BackColor = RedLED_OnOff(fg_OFF)

                    If chk0 < 0 Then
                        .ToolTip1.SetToolTip(.lbl_PLotNo, "Direction NG!!!")
                    End If

                    If chk1 < 0 Then
                        .ToolTip1.SetToolTip(.lbl_PLotNo, "Freq. NG!!!")
                    End If

                    If chk2 < 0 Then
                        .ToolTip1.SetToolTip(.lbl_PLotNo, "Week Code Jump!!!")
                    End If
                Else
                    'OK
                    If TapMarkInsp.SettingLot.InspDirection = "-L" Then
                        OutputPoint(0, 0, fg_ON)
                        .pic_BitState(8).BackColor = RedLED_OnOff(fg_ON)
                    ElseIf TapMarkInsp.SettingLot.InspDirection = "-R" Then
                        OutputPoint(0, 1, fg_ON)
                        .pic_BitState(9).BackColor = RedLED_OnOff(fg_ON)
                    End If

                    .ToolTip1.SetToolTip(.lbl_PLotNo, "Marking OK...")
                End If
            Else
                'Chr. Str. NG
                OutputPoint(0, 0, fg_OFF)
                .pic_BitState(8).BackColor = RedLED_OnOff(fg_OFF)

                OutputPoint(0, 1, fg_OFF)
                .pic_BitState(9).BackColor = RedLED_OnOff(fg_OFF)
            End If

            OutputPoint(0, 4, fg_OFF)
            .pic_BitState(12).BackColor = RedLED_OnOff(fg_OFF)

            Do Until InputPoint(0, 1) = 0
                Application.DoEvents()
            Loop
        End With

    End Sub

    Private Sub sub_ReviseLotNo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles sub_ReviseLotNo.Click

        If TapMarkInsp.RunningState = fg_ON Then Exit Sub
        frm_ReviseLN.ShowDialog()

        With TapMarkInsp.SettingLot
            If .ReviseLotNo <> "" Then
                Me.lbl_PLotNo.Text = .ReviseLotNo
            End If
        End With

    End Sub

    Private Sub tmr_Update_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmr_Update.Tick

        If Not Me.lbl_LotNo.Text.IndexOf("---") < 0 Then
            If My.Computer.Network.IsAvailable Then
                Try
                    If My.Computer.Network.Ping("172.16.59.254") Then
                        If System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed And System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CheckForUpdate Then
                            If Me.Lnk_Update.Visible = False Then
                                Me.Lnk_Update.Visible = True
                            End If
                        End If
                    End If
                Catch ex As Exception

                End Try
            End If
        End If

    End Sub

    Private Sub AppReStart()

        'RemoveHandler System.Deployment.Application.ApplicationDeployment.CurrentDeployment.UpdateCompleted, AddressOf Me.AppReStart
        'MessageBox.Show("The Application has completely update.", "Network Deployment...", MessageBoxButtons.OK, MessageBoxIcon.Information)

        With Me
            .Lnk_Update.Visible = False
        End With


        Application.Restart()

    End Sub

    Private Sub Lnk_Update_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Lnk_Update.Click

        Try
            If TapMarkInsp.SettingLot.dbData(0).LotNo = Nothing Then
                With Me
                    .tmr_KeyScanner.Enabled = False
                    .tmr_IOScanner.Enabled = False

                    If .TapMarkVS.IsOpen Then .TapMarkVS.Close()
                End With

                AddHandler System.Deployment.Application.ApplicationDeployment.CurrentDeployment.UpdateCompleted, AddressOf AppReStart
                System.Deployment.Application.ApplicationDeployment.CurrentDeployment.UpdateAsync()
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        TrgVS()

    End Sub
End Class
