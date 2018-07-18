Imports System.IO
Imports System.Net

Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        WebLoadData()
        WebBrowser7.Tag = 1
        WebBrowser7.Navigate("http://nodejs1.aero.und.edu/king-air-server")

        'display version
        lblVersion.Text = "Version: " & Application.ProductVersion

    End Sub

#Region " Variables "

    Public Shared TakeoffAltitude As Double = 999999
    Public Shared TakeoffAltimeterSetting As Double = 999999
    Public Shared TakeoffTemperature As Double = 999999
    Public Shared TakeoffWeight As Double = 999999
    Public Shared TakeoffWind As Double = 999999
    Public Shared TakeoffDistance As Double = 999999
    Public Shared UnitTakeoffAltitude As Integer = 999999
    Public Shared UnitTakeoffTemperature As Integer = 999999
    Public Shared UnitTakeoffWeight As Integer = 999999
    Public Shared UnitTakeoffWind As Integer = 999999
    Public Shared AccelerateStopDistance As Double = 999999
    Public Shared AccelerateGoDistance As Double = 999999
    Public Shared TakeoffPower As Double = 999999
    Public Shared V1 As Double = 999999
    Public Shared V2 As Double = 999999
    Public Shared Veref As Double = 999999
    Public Shared TakeoffPerformanceFormStatus As Integer = 0
    Public Shared ClimbTemperature12 As Double = 999999
    Public Shared ClimbTemperature18 As Double = 999999
    Public Shared ClimbTemperature24 As Double = 999999
    Public Shared ClimbPerformanceFormStatus As Integer = 0
    Public Shared UnitClimbTemperature12 As Integer = 999999
    Public Shared UnitClimbTemperature18 As Integer = 999999
    Public Shared UnitClimbTemperature24 As Integer = 999999
    Public Shared ClimbRate As Double = 999999
    Public Shared ClimbServiceCeiling As Double = 999999
    Public Shared LandingAltitude As Double = 999999
    Public Shared LandingAltimeterSetting As Double = 999999
    Public Shared LandingTemperature As Double = 999999
    Public Shared LandingWeight As Double = 999999
    Public Shared LandingWind As Double = 999999
    Public Shared LandingDistance As Double = 999999
    Public Shared UnitLandingAltitude As Integer = 999999
    Public Shared UnitLandingTemperature As Integer = 999999
    Public Shared UnitLandingWeight As Integer = 999999
    Public Shared UnitLandingWind As Integer = 999999
    Public Shared Vref As Double = 999999
    Public Shared LandingPerformanceFormStatus As Integer = 0
    Public Shared CruiseAltitude As Double = 999999
    Public Shared CruiseTemperature As Double = 999999
    Public Shared CruiseMode As Integer = 999999
    Public Shared CruisePower As Double = 999999
    Public Shared CruiseTorque As Double = 999999
    Public Shared CruiseFuelFlow As Double = 999999
    Public Shared CruiseIndicatedAirspeed As Double = 999999
    Public Shared CruiseTrueAirspeed As Double = 999999
    Public Shared CruisePerformanceFormStatus As Integer = 0
    Public Shared WBPilot As Double = 999999
    Public Shared WBCopilot As Double = 999999
    Public Shared WBFWDCabinet As Double = 999999
    Public Shared WBPax1 As Double = 999999
    Public Shared WBPax2 As Double = 999999
    Public Shared WBPax3 As Double = 999999
    Public Shared WBPax4 As Double = 999999
    Public Shared WBPax5 As Double = 999999
    Public Shared WBPax6 As Double = 999999
    Public Shared WBAFTCabinet1 As Double = 999999
    Public Shared WBAFTCabinet2 As Double = 999999
    Public Shared WBBaggage As Double = 999999
    Public Shared WBFuelLoad As Double = 999999
    Public Shared WBFuelUsed As Double = 999999
    Public Shared WBZeroFuelWeight As Double = 999999
    Public Shared WBAircraft As Integer = 999999
    Public Shared WBFormStatus As Integer = 0
    Public Shared WBRampWeight As Double = 999999
    Public Shared WBMZeroFuelWeight As Double = 999999
    Public Shared WBMRampWeight As Double = 999999
    Public Shared WBMTakeoffWeight As Double = 999999
    Public Shared WBMLandingWeight As Double = 999999
    Public Shared CGZeroFuelWeight As Double = 999999
    Public Shared CGRampWeight As Double = 999999
    Public Shared CGTakeoffWeight As Double = 999999
    Public Shared CGLandingWeight As Double = 999999
    Public Shared FPDep As String = "999999"
    Public Shared FPDes As String = "999999"
    Public Shared FPETD As Double = 999999
    Public Shared FPETA As Double = 999999
    Public Shared FPDWindDirection As Integer = 999999
    Public Shared FPDWindSpeed As Integer = 999999
    Public Shared FPDWindGust As Integer = 999999
    Public Shared FPAWindDirection As Integer = 999999
    Public Shared FPAWindSpeed As Integer = 999999
    Public Shared FPAWindGust As Integer = 999999
    Public Shared WAAirports(175) As String
    Public Shared WA12(175) As Integer
    Public Shared WA18(175) As Integer
    Public Shared WA24(175) As Integer
    Public Shared foundWindsAloft As Boolean = False
    Public Shared DepRunways As String = ""
    Public Shared ArrRunways As String = ""
    Public Shared RunwayFormStatus As Integer = 1
    Public Shared FltPlanFormStatus As Integer = 0
    Public Shared CompleteEverything As Integer = 0
    Public Shared OrderToComplete As Boolean = False
    Public Shared FPUsername As String = "999999"
    Public Shared FPPassword As String = "999999"
    Public Shared FPDate As String = "999999"
    Public Shared CurrentForm As Integer = 0
    Public Shared CurrentBackgroundIsHome As Boolean = True
    Public Shared CurrentToolbar As Integer = 0
    Public Shared WBBasicEmptyWeights() As Double
    Public Shared WBBasicEmptyMoments() As Double
    Public Shared TakeoffPressureAltitude As Double = 999999
    Public Shared LandingPressureAltitude As Double = 999999
    Public Shared FPInstructorNames() As String = ({"", "Jacob Boyd", "Ryan Wilson"})
    Public Shared FPUsernames() As String = ({"", "kingboyd", "WILR"})
    Public Shared FPPasswords() As String = ({"", "kingair", "77985"})
    Public Shared FPUserSelection As Integer = 999999
    Public Shared FPLoadHomepage As Boolean
    Public Shared WSData() As String
    Public Shared FPMetars(2) As String
    Public Shared Airplanes() As String
    Public Shared EditedAirplanes() As String
    Public Shared EditedWeights() As String
    Public Shared EditedMoments() As String
    Public Shared NextFlightPlanAvailable As Boolean = False
    Public Shared FPSelected As Integer = 999999
    Public Shared FPFlightplans As Integer = 999999
    Public Shared SkipLogin As Boolean = False
#End Region
#Region " Move Form "

    Public MoveForm As Boolean
    Public MoveForm_MousePosition As Point

    Public Sub MoveForm_MouseDown(sender As Object, e As MouseEventArgs) Handles _
    PictureBox1.MouseDown

        If e.Button = MouseButtons.Left Then
            MoveForm = True
            Me.Cursor = Cursors.NoMove2D
            MoveForm_MousePosition = e.Location
        End If

    End Sub

    Public Sub MoveForm_MouseMove(sender As Object, e As MouseEventArgs) Handles _
    PictureBox1.MouseMove

        If MoveForm Then
            Me.Location = Me.Location + (e.Location - MoveForm_MousePosition)
        End If

    End Sub

    Public Sub MoveForm_MouseUp(sender As Object, e As MouseEventArgs) Handles _
    PictureBox1.MouseUp

        If e.Button = MouseButtons.Left Then
            MoveForm = False
            Me.Cursor = Cursors.Default
        End If

    End Sub

#End Region
#Region " Main Bar "

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        Me.Close()
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

#End Region
#Region " Layout Configuration "

    Public Sub MainSelectionChange(ByVal selection As Integer)
        CurrentForm = selection
        ResetMainButtons()
        ResetPanels()
        If CurrentBackgroundIsHome Then
            BackgroundImage = My.Resources.Background
            CurrentBackgroundIsHome = False
        End If
        Select Case CurrentForm
            Case 0
                BackgroundImage = My.Resources.Background_home
                CurrentBackgroundIsHome = True
            Case 1
                MLabel1.Image = My.Resources.buttonYes
                PanelWB.Visible = True
            Case 2
                MLabel2.Image = My.Resources.buttonYes
                PanelTakeoff.Visible = True
            Case 3
                MLabel3.Image = My.Resources.buttonYes
                PanelClimb.Visible = True
            Case 4
                MLabel4.Image = My.Resources.buttonYes
                PanelCruise.Visible = True
            Case 5
                MLabel5.Image = My.Resources.buttonYes
                PanelLanding.Visible = True
            Case 6
                MLabel6.Image = My.Resources.buttonYes
                PanelSettingsLogin.Visible = True
                LLogin2.Text = ""
            Case 7
                PanelTOLD.Visible = True
                GLoadData()
            Case 8
                PanelLogin.Visible = True
                HLoadData()
            Case 9
                PanelFP.Visible = True
                ILoadData()
            Case 10
                PanelRunways.Visible = True
                JLoadData()
            Case 11
                PanelWBGraph.Visible = True
                KLoadData()
        End Select
        ResetLabelColors()

    End Sub

    Public Sub ResetMainButtons()
        MLabel1.Image = My.Resources.buttonNo
        MLabel2.Image = My.Resources.buttonNo
        MLabel3.Image = My.Resources.buttonNo
        MLabel4.Image = My.Resources.buttonNo
        MLabel5.Image = My.Resources.buttonNo
        MLabel6.Image = My.Resources.buttonNo

    End Sub

    Public Sub ResetLabelColors()
        If WBFormStatus = 0 Then
            MLabel1.ForeColor = Color.White
        Else
            MLabel1.ForeColor = Color.YellowGreen
        End If
        If TakeoffPerformanceFormStatus = 0 Then
            MLabel2.ForeColor = Color.White
        Else
            MLabel2.ForeColor = Color.YellowGreen
        End If
        If ClimbPerformanceFormStatus = 0 Then
            MLabel3.ForeColor = Color.White
        Else
            MLabel3.ForeColor = Color.YellowGreen
        End If
        If CruisePerformanceFormStatus = 0 Then
            MLabel4.ForeColor = Color.White
        Else
            MLabel4.ForeColor = Color.YellowGreen
        End If
        If LandingPerformanceFormStatus = 0 Then
            MLabel5.ForeColor = Color.White
        Else
            MLabel5.ForeColor = Color.YellowGreen
        End If
        CheckStatus()
    End Sub

    Public Sub ResetPanels()
        PanelWB.Visible = False
        PanelTakeoff.Visible = False
        PanelClimb.Visible = False
        PanelCruise.Visible = False
        PanelLanding.Visible = False
        PanelTOLD.Visible = False
        PanelLogin.Visible = False
        PanelFP.Visible = False
        PanelRunways.Visible = False
        PanelWBGraph.Visible = False
        PanelSettings.Visible = False
        PanelSettingsLogin.Visible = False
    End Sub

    Public Sub CheckStatus()
        If FltPlanFormStatus = 1 And WBFormStatus = 1 And CompleteEverything = 0 Then
            BCalculateData()
            CCalculateData()
            DCalculateData()
            ECalculateData()
            CompleteEverything = 1
            ResetLabelColors()
        End If
        If (WBFormStatus + TakeoffPerformanceFormStatus + ClimbPerformanceFormStatus + CruisePerformanceFormStatus + LandingPerformanceFormStatus) = 5 Then
            If CurrentToolbar <> 2 Then ChangeToolbar(2)
        End If
    End Sub

    Private Sub MLabel1_Click(sender As Object, e As EventArgs) Handles MLabel1.Click
        MainSelectionChange(1)
        ALoadData()

    End Sub

    Private Sub MLabel2_Click(sender As Object, e As EventArgs) Handles MLabel2.Click
        MainSelectionChange(2)
        BLoadData()

    End Sub

    Private Sub MLabel3_Click(sender As Object, e As EventArgs) Handles MLabel3.Click
        MainSelectionChange(3)
        CLoadData()

    End Sub

    Private Sub MLabel4_Click(sender As Object, e As EventArgs) Handles MLabel4.Click
        MainSelectionChange(4)
        DLoadData()

    End Sub

    Private Sub MLabel5_Click(sender As Object, e As EventArgs) Handles MLabel5.Click
        MainSelectionChange(5)
        ELoadData()

    End Sub

    Private Sub MLabel6_Click(sender As Object, e As EventArgs) Handles MLabel6.Click
        MainSelectionChange(6)

    End Sub

#End Region
#Region " Toolbar "

    Public Sub ChangeToolbar(ByVal toolbarChange As Integer)
        CurrentToolbar = toolbarChange
        ResetToolbar()
        Select Case CurrentToolbar
            Case 0
                TLabel2.Text = "HOME"
            Case 1
                TLabel3.Visible = True
                TBar2.Visible = True
                TLabel2.Text = "CLEAR ALL"
                TLabel3.Text = "HOME"
            Case 2
                TLabel3.Visible = True
                TLabel4.Visible = True
                TLabel5.Visible = True
                TBar2.Visible = True
                TBar3.Visible = True
                TBar4.Visible = True
                TLabel2.Text = "CLEAR ALL"
                TLabel3.Text = "PRINT"
                TLabel4.Text = "TOLD CARD"
                TLabel5.Text = "HOME"
                If NextFlightPlanAvailable Then
                    TLabel6.Visible = True
                    TBar5.Visible = True
                End If
        End Select
        CheckStatus()

    End Sub

    Public Sub ResetToolbar()
        TLabel3.Visible = False
        TLabel4.Visible = False
        TLabel5.Visible = False
        TLabel6.Visible = False
        TBar2.Visible = False
        TBar3.Visible = False
        TBar4.Visible = False
        TBar5.Visible = False
    End Sub

    Public Sub ClearAllData()
        MainSelectionChange(0)
        ResetPanelLayouts()
        TakeoffAltitude = 999999
        TakeoffAltimeterSetting = 999999
        TakeoffTemperature = 999999
        TakeoffWeight = 999999
        TakeoffWind = 999999
        TakeoffDistance = 999999
        UnitTakeoffAltitude = 999999
        UnitTakeoffTemperature = 999999
        UnitTakeoffWeight = 999999
        UnitTakeoffWind = 999999
        AccelerateStopDistance = 999999
        AccelerateGoDistance = 999999
        TakeoffPower = 999999
        V1 = 999999
        V2 = 999999
        Veref = 999999
        TakeoffPerformanceFormStatus = 0
        ClimbTemperature12 = 999999
        ClimbTemperature18 = 999999
        ClimbTemperature24 = 999999
        ClimbPerformanceFormStatus = 0
        UnitClimbTemperature12 = 999999
        UnitClimbTemperature18 = 999999
        UnitClimbTemperature24 = 999999
        ClimbRate = 999999
        ClimbServiceCeiling = 999999
        LandingAltitude = 999999
        LandingAltimeterSetting = 999999
        LandingTemperature = 999999
        LandingWeight = 999999
        LandingWind = 999999
        LandingDistance = 999999
        UnitLandingAltitude = 999999
        UnitLandingTemperature = 999999
        UnitLandingWeight = 999999
        UnitLandingWind = 999999
        Vref = 999999
        LandingPerformanceFormStatus = 0
        CruiseAltitude = 999999
        CruiseTemperature = 999999
        CruiseMode = 999999
        CruisePower = 999999
        CruiseTorque = 999999
        CruiseFuelFlow = 999999
        CruiseIndicatedAirspeed = 999999
        CruiseTrueAirspeed = 999999
        CruisePerformanceFormStatus = 0
        WBPilot = 999999
        WBCopilot = 999999
        WBFWDCabinet = 999999
        WBPax1 = 999999
        WBPax2 = 999999
        WBPax3 = 999999
        WBPax4 = 999999
        WBPax5 = 999999
        WBPax6 = 999999
        WBAFTCabinet1 = 999999
        WBAFTCabinet2 = 999999
        WBBaggage = 999999
        WBFuelLoad = 999999
        WBFuelUsed = 999999
        WBZeroFuelWeight = 999999
        WBAircraft = 999999
        WBFormStatus = 0
        WBRampWeight = 999999
        WBMZeroFuelWeight = 999999
        WBMRampWeight = 999999
        WBMTakeoffWeight = 999999
        WBMLandingWeight = 999999
        CGZeroFuelWeight = 999999
        CGRampWeight = 999999
        CGTakeoffWeight = 999999
        CGLandingWeight = 999999
        FPDep = "999999"
        FPDes = "999999"
        FPETD = 999999
        FPETA = 999999
        FPDWindDirection = 999999
        FPDWindSpeed = 999999
        FPDWindGust = 999999
        FPAWindDirection = 999999
        FPAWindSpeed = 999999
        FPAWindGust = 999999
        foundWindsAloft = False
        DepRunways = ""
        ArrRunways = ""
        RunwayFormStatus = 1
        FltPlanFormStatus = 0
        CompleteEverything = 0
        OrderToComplete = False
        FPUsername = "999999"
        FPPassword = "999999"
        FPDate = "999999"
        TakeoffPressureAltitude = 999999
        LandingPressureAltitude = 999999
        FPUserSelection = 999999
        ChangeToolbar(0)
        ResetLabelColors()
        NextFlightPlanAvailable = False
        FPSelected = 999999
        FPFlightplans = 999999
        SkipLogin = False
    End Sub

    Public Sub ResetPanelLayouts()
        If WBFormStatus = 1 Then
            AResultLabel1.Visible = False
            AResultLabel2.Visible = False
            AResultLabel3.Visible = False
            AResultLabel4.Visible = False
            AResultLabel5.Visible = False
            AResultLabel6.Visible = False
            AResultLabel7.Visible = False
            AResultLabel8.Visible = False
            AResultLabel9.Visible = False
            AResultValue1.Visible = False
            AResultValue2.Visible = False
            AResultValue3.Visible = False
            AResultValue4.Visible = False
            AResultValue5.Visible = False
            AResultValue6.Visible = False
            AResultValue7.Visible = False
            AResultValue8.Visible = False
            AResultValue9.Visible = False
            AButton2.Visible = False
            ATextBox1.Text = ""
            ATextBox2.Text = ""
            ATextBox3.Text = ""
            ATextBox4.Text = ""
            ATextBox5.Text = ""
            ATextBox6.Text = ""
            ATextBox7.Text = ""
            ATextBox8.Text = ""
            ATextBox9.Text = ""
            ATextBox10.Text = ""
            ATextBox11.Text = ""
            ATextBox12.Text = ""
            ATextBox13.Text = ""
            ATextBox14.Text = ""
            AComboBox1.SelectedIndex = -1
        End If
        If TakeoffPerformanceFormStatus = 1 Then
            BResultLabel1.Visible = False
            BResultLabel2.Visible = False
            BResultLabel3.Visible = False
            BResultLabel4.Visible = False
            BResultLabel5.Visible = False
            BResultLabel6.Visible = False
            BResultLabel7.Visible = False
            BResultLabel8.Visible = False
            BResultLabel9.Visible = False
            BResultValue1.Visible = False
            BResultValue2.Visible = False
            BResultValue3.Visible = False
            BResultValue4.Visible = False
            BResultValue5.Visible = False
            BResultValue6.Visible = False
            BResultValue7.Visible = False
            BResultValue8.Visible = False
            BResultValue9.Visible = False
            BTextBox1.Text = ""
            BTextBox2.Text = ""
            BTextBox3.Text = ""
            BTextBox4.Text = ""
            BTextBox5.Text = ""
            BRadioButton1.Checked = True
            BRadioButton3.Checked = True
            BRadioButton5.Checked = True
            BRadioButton7.Checked = True
        End If
        If LandingPerformanceFormStatus = 1 Then
            EResultLabel1.Visible = False
            EResultLabel2.Visible = False
            EResultValue1.Visible = False
            EResultValue2.Visible = False
            ETextBox1.Text = ""
            ETextBox2.Text = ""
            ETextBox3.Text = ""
            ETextBox4.Text = ""
            ETextBox5.Text = ""
            ERadioButton1.Checked = True
            ERadioButton3.Checked = True
            ERadioButton5.Checked = True
            ERadioButton7.Checked = True
        End If
        If ClimbPerformanceFormStatus = 1 Then
            CResultLabel1.Visible = False
            CResultLabel2.Visible = False
            CResultValue1.Visible = False
            CResultValue2.Visible = False
            CTextBox1.Text = ""
            CTextBox2.Text = ""
            CTextBox3.Text = ""
            CTextBox4.Text = ""
            CTextBox5.Text = ""
            CTextBox6.Text = ""
            CTextBox7.Text = ""
            CRadioButton1.Checked = True
            CRadioButton3.Checked = True
            CRadioButton5.Checked = True
        End If
        If CruisePerformanceFormStatus = 1 Then
            DResultLabel1.Visible = False
            DResultLabel2.Visible = False
            DResultLabel3.Visible = False
            DResultLabel4.Visible = False
            DResultValue1.Visible = False
            DResultValue2.Visible = False
            DResultValue3.Visible = False
            DResultValue4.Visible = False
            DTextBox1.Text = ""
            DTextBox2.Text = ""
            DTextBox3.Text = ""
            DRadioButton1.Checked = True
        End If
    End Sub

    Public Sub ToolbarButtonClicked(ByVal buttonLabel As String)
        Select Case buttonLabel
            Case "HOME"
                MainSelectionChange(0)
            Case "CLEAR ALL"
                If MsgBox("Do you want to delete all the data?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    ClearAllData()
                End If
            Case "TOLD CARD"
                MainSelectionChange(7)
            Case "FLTPLAN.COM"
                MainSelectionChange(8)
            Case "PRINT"
                'PrintPreviewDialog1.ShowDialog()
                If PrintDialog1.ShowDialog() = DialogResult.OK Then
                    PrintDocument1.Print()
                End If
        End Select
    End Sub

    Private Sub TLabel1_Click(sender As Object, e As EventArgs) Handles TLabel1.Click, TLabel2.Click, TLabel3.Click, TLabel4.Click, TLabel5.Click
        ToolbarButtonClicked(sender.Text)
    End Sub

#End Region
#Region " Conversions "

    Public Function ConvertAltitude(ByVal altitude As Double, ByVal altimeterSetting As Double) As Double
        Return (((29.92 - altimeterSetting) * 1000) + altitude)
    End Function

    Public Function ConvertTemperature(ByVal temperature As Double) As Double
        Return ((temperature - 32) * (5 / 7))
    End Function

    Public Function ConvertWeight(ByVal weight As Double) As Double
        Return (weight * 2.20462)
    End Function

#End Region

#Region " W&B "

    Public Sub ALoadData()
        AComboBox1.Items.Clear()
        AComboBox1.Items.AddRange(Airplanes)
        AComboBox1.Focus()
        If WBFormStatus = 1 Then
            ATextBox1.Text = WBPilot
            ATextBox2.Text = WBCopilot
            ATextBox3.Text = WBFWDCabinet
            ATextBox4.Text = WBPax1
            ATextBox5.Text = WBPax2
            ATextBox6.Text = WBPax3
            ATextBox7.Text = WBPax4
            ATextBox8.Text = WBPax5
            ATextBox9.Text = WBPax6
            ATextBox10.Text = WBAFTCabinet1
            ATextBox11.Text = WBAFTCabinet2
            ATextBox12.Text = WBBaggage
            ATextBox13.Text = WBFuelLoad
            ATextBox14.Text = WBFuelUsed
            AShowData()
        Else
            If WBPilot <> 999999 Then ATextBox1.Text = WBPilot
            If WBCopilot <> 999999 Then ATextBox2.Text = WBCopilot
            If WBFWDCabinet <> 999999 Then ATextBox3.Text = WBFWDCabinet
            If WBPax1 <> 999999 Then ATextBox4.Text = WBPax1
            If WBPax2 <> 999999 Then ATextBox5.Text = WBPax2
            If WBPax3 <> 999999 Then ATextBox6.Text = WBPax3
            If WBPax4 <> 999999 Then ATextBox7.Text = WBPax4
            If WBPax5 <> 999999 Then ATextBox8.Text = WBPax5
            If WBPax6 <> 999999 Then ATextBox9.Text = WBPax6
            If WBAFTCabinet1 <> 999999 Then ATextBox10.Text = WBAFTCabinet1
            If WBAFTCabinet2 <> 999999 Then ATextBox11.Text = WBAFTCabinet2
            If WBBaggage <> 999999 Then ATextBox12.Text = WBBaggage
            If WBFuelLoad <> 999999 Then ATextBox13.Text = WBFuelLoad
            If WBFuelUsed <> 999999 Then ATextBox14.Text = WBFuelUsed
        End If
        If WBAircraft <> 999999 Then
            AComboBox1.SelectedItem = AComboBox1.Items.Item(WBAircraft)
        End If
        MaximumFuel()
        If SkipLogin Then
            If ACheckData() Then
                If CurrentToolbar = 0 Then ChangeToolbar(1)
                ASaveData()
                ACalculateData()
                AShowData()
                ResetLabelColors()
            End If
        End If
    End Sub

    Public Function ACheckData() As Boolean
        Dim returnValue As Boolean = True
        If ATextBox2.Text = "" Then ATextBox2.Text = 0
        If ATextBox3.Text = "" Then ATextBox3.Text = 0
        If ATextBox4.Text = "" Then ATextBox4.Text = 0
        If ATextBox5.Text = "" Then ATextBox5.Text = 0
        If ATextBox6.Text = "" Then ATextBox6.Text = 0
        If ATextBox7.Text = "" Then ATextBox7.Text = 0
        If ATextBox8.Text = "" Then ATextBox8.Text = 0
        If ATextBox9.Text = "" Then ATextBox9.Text = 0
        If ATextBox10.Text = "" Then ATextBox10.Text = 0
        If ATextBox11.Text = "" Then ATextBox11.Text = 0
        If ATextBox12.Text = "" Then ATextBox12.Text = 0
        If ATextBox1.Text = "" Or ATextBox13.Text = "" Or ATextBox14.Text = "" Or AComboBox1.SelectedIndex = -1 Then
            MsgBox("Missing Information!")
            returnValue = False
        ElseIf (Not IsNumeric(ATextBox1.Text)) Or (Not IsNumeric(ATextBox2.Text)) Or (Not IsNumeric(ATextBox3.Text)) Or (Not IsNumeric(ATextBox4.Text)) Or (Not IsNumeric(ATextBox5.Text)) Or (Not IsNumeric(ATextBox6.Text)) Or (Not IsNumeric(ATextBox7.Text)) Or (Not IsNumeric(ATextBox8.Text)) Or (Not IsNumeric(ATextBox9.Text)) Or (Not IsNumeric(ATextBox10.Text)) Or (Not IsNumeric(ATextBox11.Text)) Or (Not IsNumeric(ATextBox12.Text)) Or (Not IsNumeric(ATextBox13.Text)) Or (Not IsNumeric(ATextBox14.Text)) Then
            MsgBox("Invalid Input!")
            returnValue = False
        ElseIf ATextBox13.Text > 2573 Then
            MsgBox("Fuel is too high")
            returnValue = False
        ElseIf CDbl(ATextBox3.Text) > 28 Then
            MsgBox("FWD Cabinet load is too high")
            returnValue = False
        ElseIf CDbl(ATextBox10.Text) > 6 Then
            MsgBox("AFT Cabinet 1 load is too high")
            returnValue = False
        ElseIf CDbl(ATextBox11.Text) > 30 Then
            MsgBox("AFT Cabinet 2 load is too high")
            returnValue = False
        ElseIf CDbl(ATextBox12.Text) > 350 Then
            MsgBox("Total Baggage load is too high")
            returnValue = False
        End If
        Return returnValue
    End Function

    Public Sub ASaveData()
        WBPilot = CDbl(ATextBox1.Text)
        WBCopilot = CDbl(ATextBox2.Text)
        WBFWDCabinet = CDbl(ATextBox3.Text)
        WBPax1 = CDbl(ATextBox4.Text)
        WBPax2 = CDbl(ATextBox5.Text)
        WBPax3 = CDbl(ATextBox6.Text)
        WBPax4 = CDbl(ATextBox7.Text)
        WBPax5 = CDbl(ATextBox8.Text)
        WBPax6 = CDbl(ATextBox9.Text)
        WBAFTCabinet1 = CDbl(ATextBox10.Text)
        WBAFTCabinet2 = CDbl(ATextBox11.Text)
        WBBaggage = CDbl(ATextBox12.Text)
        WBFuelLoad = CDbl(ATextBox13.Text)
        WBFuelUsed = CDbl(ATextBox14.Text)
        WBAircraft = AComboBox1.SelectedIndex

    End Sub

    Public Sub ACalculateData()
        WBZeroFuelWeight = WBBasicEmptyWeights(WBAircraft) + WBPilot + WBCopilot + WBFWDCabinet + WBPax1 + WBPax2 + WBPax3 + WBPax4 + WBPax5 + WBPax6 + WBAFTCabinet1 + WBAFTCabinet2 + WBBaggage
        WBRampWeight = WBZeroFuelWeight + WBFuelLoad
        TakeoffWeight = WBRampWeight - 60
        LandingWeight = TakeoffWeight - WBFuelUsed

        WBMZeroFuelWeight = WBBasicEmptyMoments(WBAircraft) * 100 + WBPilot * 129 + WBCopilot * 129 + WBFWDCabinet * 153 + WBPax1 * 168 + WBPax2 * 168 + WBPax3 * 218 + WBPax4 * 215 + WBPax5 * 243 + WBPax6 * 285 + WBAFTCabinet1 * 226 + WBAFTCabinet2 * 243 + WBBaggage * 277
        WBMRampWeight = WBMZeroFuelWeight + (FuelMoment(WBFuelLoad) * 100)
        WBMTakeoffWeight = WBMRampWeight - 9300
        WBMLandingWeight = WBMZeroFuelWeight + (FuelMoment(WBFuelLoad - WBFuelUsed) * 100)

        CGZeroFuelWeight = WBMZeroFuelWeight / WBZeroFuelWeight
        CGRampWeight = (WBMRampWeight / WBRampWeight) - 0.1
        CGTakeoffWeight = (WBMTakeoffWeight / TakeoffWeight) - 0.1
        CGLandingWeight = WBMLandingWeight / LandingWeight

        WBFormStatus = 1
    End Sub

    Public Sub AShowData()
        AResultLabel1.Visible = True
        AResultLabel2.Visible = True
        AResultLabel3.Visible = True
        AResultLabel4.Visible = True
        AResultLabel5.Visible = True
        AResultLabel6.Visible = True
        AResultLabel7.Visible = True
        AResultLabel8.Visible = True
        AResultLabel9.Visible = True
        AResultValue1.Visible = True
        AResultValue2.Visible = True
        AResultValue3.Visible = True
        AResultValue4.Visible = True
        AResultValue5.Visible = True
        AResultValue6.Visible = True
        AResultValue7.Visible = True
        AResultValue8.Visible = True
        AResultValue9.Visible = True
        AButton2.Visible = True
        AResultValue1.Text = Math.Round(WBZeroFuelWeight).ToString + "  Lbs"
        AResultValue2.Text = Math.Round(WBRampWeight).ToString + "  Lbs"
        AResultValue3.Text = Math.Round(TakeoffWeight).ToString + "  Lbs"
        AResultValue4.Text = Math.Round(LandingWeight).ToString + "   Lbs"
        AResultValue5.Text = Math.Round(CGZeroFuelWeight, 1).ToString + "  in"
        AResultValue6.Text = Math.Round(CGRampWeight, 1).ToString + "  in"
        AResultValue7.Text = Math.Round(CGTakeoffWeight, 1).ToString + "  in"
        AResultValue8.Text = Math.Round(CGLandingWeight, 1).ToString + "  in"
        AResultValue9.Text = (WBFuelLoad - WBFuelUsed).ToString + "  Lbs"
        If WBRampWeight > 10160 Then
            AResultValue2.ForeColor = Color.Red
        Else
            AResultValue2.ForeColor = Color.Black
        End If
        If TakeoffWeight > 10100 Then
            AResultValue3.ForeColor = Color.Red
        Else
            AResultValue3.ForeColor = Color.Black
        End If

        If LandingWeight > 9600 Then
            AResultValue4.ForeColor = Color.Red
        Else
            AResultValue4.ForeColor = Color.Black
        End If
        If CGZeroFuelWeight > 162.2 Or CGZeroFuelWeight < 144.6 Then
            AResultValue5.ForeColor = Color.Red
        Else
            AResultValue5.ForeColor = Color.Black
        End If
        If CGRampWeight > 162.2 Or CGRampWeight < 144.6 Then
            AResultValue6.ForeColor = Color.Red
        Else
            AResultValue6.ForeColor = Color.Black
        End If
        If CGTakeoffWeight > 162.2 Or CGTakeoffWeight < 144.6 Then
            AResultValue7.ForeColor = Color.Red
        Else
            AResultValue7.ForeColor = Color.Black
        End If
        If CGLandingWeight > 162.2 Or CGLandingWeight < 144.6 Then
            AResultValue8.ForeColor = Color.Red
        Else
            AResultValue8.ForeColor = Color.Black
        End If

    End Sub

    Public Function FuelMoment(ByVal fuelLoad As Double)
        Return (-8.39381658484718E-17 * fuelLoad ^ 6 + 0.000000000000687910769184429 * fuelLoad ^ 5 - 0.0000000021280003261337 * fuelLoad ^ 4 + 0.00000298683212340053 * fuelLoad ^ 3 - 0.0016653259159681 * fuelLoad ^ 2 + 1.63511539015189 * fuelLoad - 22.6488167205946)
    End Function

    Private Sub AButton1_Click(sender As Object, e As EventArgs) Handles AButton1.Click
        If ACheckData() Then
            If CurrentToolbar = 0 Then ChangeToolbar(1)
            If SkipLogin Then CompleteEverything = 0
            ASaveData()
            ACalculateData()
            AShowData()
            ResetLabelColors()
        End If
    End Sub

    Private Sub AButton2_Click(sender As Object, e As EventArgs) Handles AButton2.Click
        MainSelectionChange(11)
    End Sub

    Private Sub ATextBox_KeyDown(sender As Object, e As KeyEventArgs) Handles ATextBox1.KeyDown, ATextBox2.KeyDown, ATextBox3.KeyDown, ATextBox4.KeyDown, ATextBox5.KeyDown, ATextBox6.KeyDown, ATextBox7.KeyDown, ATextBox8.KeyDown, ATextBox9.KeyDown, ATextBox10.KeyDown, ATextBox11.KeyDown, ATextBox12.KeyDown, ATextBox13.KeyDown, ATextBox14.KeyDown, AComboBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            AButton1.PerformClick()
        End If
    End Sub

    Public Sub MaximumFuel() Handles ATextBox1.TextChanged, ATextBox2.TextChanged, ATextBox3.TextChanged, ATextBox4.TextChanged, ATextBox5.TextChanged, ATextBox6.TextChanged, ATextBox7.TextChanged, ATextBox8.TextChanged, ATextBox9.TextChanged, ATextBox10.TextChanged, ATextBox11.TextChanged, ATextBox12.TextChanged, ATextBox14.TextChanged, AComboBox1.SelectedIndexChanged
        Dim maxFuel As Double
        Dim totalWeight As Double
        Dim limitTakeoff, limitLanding As Double
        Dim text1, text2, text3, text4, text5, text6, text7, text8, text9, text10, text11, text12, text13 As Double
        If IsNumeric(ATextBox1.Text) Then
            text1 = ATextBox1.Text
        Else
            text1 = 0
        End If
        If IsNumeric(ATextBox2.Text) Then
            text2 = ATextBox2.Text
        Else
            text2 = 0
        End If
        If IsNumeric(ATextBox3.Text) Then
            text3 = ATextBox3.Text
        Else
            text3 = 0
        End If
        If IsNumeric(ATextBox4.Text) Then
            text4 = ATextBox4.Text
        Else
            text4 = 0
        End If
        If IsNumeric(ATextBox5.Text) Then
            text5 = ATextBox5.Text
        Else
            text5 = 0
        End If
        If IsNumeric(ATextBox6.Text) Then
            text6 = ATextBox6.Text
        Else
            text6 = 0
        End If
        If IsNumeric(ATextBox7.Text) Then
            text7 = ATextBox7.Text
        Else
            text7 = 0
        End If
        If IsNumeric(ATextBox8.Text) Then
            text8 = ATextBox8.Text
        Else
            text8 = 0
        End If
        If IsNumeric(ATextBox9.Text) Then
            text9 = ATextBox9.Text
        Else
            text9 = 0
        End If
        If IsNumeric(ATextBox10.Text) Then
            text10 = ATextBox10.Text
        Else
            text10 = 0
        End If
        If IsNumeric(ATextBox11.Text) Then
            text11 = ATextBox11.Text
        Else
            text11 = 0
        End If
        If IsNumeric(ATextBox12.Text) Then
            text12 = ATextBox12.Text
        Else
            text12 = 0
        End If
        If IsNumeric(ATextBox14.Text) Then
            text13 = ATextBox14.Text
        Else
            text13 = 0
        End If

        If AComboBox1.SelectedIndex = -1 Then
            ALabel16.Visible = False
        Else
            ALabel16.Visible = True
            totalWeight = WBBasicEmptyWeights(AComboBox1.SelectedIndex) + text1 + text2 + text3 + text4 + text5 + text6 + text7 + text8 + text9 + text10 + text11 + text12
            limitTakeoff = 10160 - totalWeight
            limitLanding = 9660 - totalWeight + text13
            If limitTakeoff > limitLanding Then
                If limitLanding > 2573 Then
                    maxFuel = 2573
                Else
                    maxFuel = limitLanding
                End If
            Else
                If limitTakeoff > 2573 Then
                    maxFuel = 2573
                Else
                    maxFuel = limitTakeoff
                End If
            End If
            ALabel16.Text = "Max: " + CInt(maxFuel).ToString + " Lbs"
        End If

    End Sub

#End Region
#Region " Takeoff "

    Public Sub BLoadData()
        BTextBox1.Focus()
        If TakeoffPerformanceFormStatus = 1 Then
            BTextBox1.Text = TakeoffAltitude
            BTextBox2.Text = TakeoffTemperature
            BTextBox3.Text = TakeoffWeight
            BTextBox4.Text = TakeoffWind
            BShowData()
        Else
            If TakeoffAltitude <> 999999 Then BTextBox1.Text = TakeoffAltitude
            If TakeoffTemperature <> 999999 Then BTextBox2.Text = TakeoffTemperature
            If TakeoffWeight <> 999999 Then BTextBox3.Text = TakeoffWeight
            If TakeoffWind <> 999999 Then BTextBox4.Text = TakeoffWind
        End If
        If UnitTakeoffAltitude = 2 Then
            BRadioButton2.Checked = True
            BTextBox5.Text = TakeoffAltimeterSetting
        Else
            BRadioButton1.Checked = True
            BTextBox5.Text = ""
        End If
        If UnitTakeoffTemperature = 2 Then BRadioButton4.Checked = True Else BRadioButton3.Checked = True
        If UnitTakeoffWeight = 2 Then BRadioButton6.Checked = True Else BRadioButton5.Checked = True
        If UnitTakeoffWind = 2 Then BRadioButton8.Checked = True Else BRadioButton7.Checked = True
    End Sub

    Public Function BCheckData() As Boolean
        Dim returnValue As Boolean = True
        If BTextBox1.Text = "" Then BTextBox1.Text = "0"
        If BTextBox2.Text = "" Then BTextBox2.Text = "15"
        If BTextBox3.Text = "" Then BTextBox3.Text = "9500"
        If BTextBox4.Text = "" Then BTextBox4.Text = "0"
        If (Not IsNumeric(BTextBox1.Text)) Or (Not IsNumeric(BTextBox2.Text)) Or (Not IsNumeric(BTextBox3.Text)) Or (Not IsNumeric(BTextBox4.Text)) Then
            MsgBox("Invalid Input!")
            returnValue = False
        End If
        If BRadioButton2.Checked Then
            If BTextBox5.Text = "" Then BTextBox5.Text = "29.92"
            If (Not IsNumeric(BTextBox5.Text)) Then
                MsgBox("Invalid Input!")
                returnValue = False
            End If
            If ConvertAltitude(CDbl(BTextBox1.Text), CDbl(BTextBox5.Text)) > 10000 Then
                MsgBox("Pressure Altitude is too high")
                returnValue = False
            ElseIf ConvertAltitude(CDbl(BTextBox1.Text), CDbl(BTextBox5.Text)) < 0 Then
                MsgBox("Pressure Altitude is too low")
                returnValue = False
            End If
        Else
            If CDbl(BTextBox1.Text) > 10000 Then
                MsgBox("Pressure Altitude is too high")
                returnValue = False
            ElseIf CDbl(BTextBox1.Text) < 0 Then
                MsgBox("Pressure Altitude is too low")
                returnValue = False
            End If
        End If
        If BRadioButton3.Checked Then
            If CDbl(BTextBox2.Text) > 50 Then
                MsgBox("Temperature is too high")
                returnValue = False
            ElseIf CDbl(BTextBox2.Text) < -40 Then
                MsgBox("Temperature is too low")
                returnValue = False
            End If
        Else
            If ConvertTemperature(CDbl(BTextBox2.Text)) > 50 Then
                MsgBox("Temperature is too high")
                returnValue = False
            ElseIf ConvertTemperature(CDbl(BTextBox2.Text)) < -40 Then
                MsgBox("Temperature is too low")
                returnValue = False
            End If
        End If
        If BRadioButton5.Checked Then
            If CDbl(BTextBox3.Text) < 7000 Then
                MsgBox("Weight is too low")
                returnValue = False
            ElseIf CDbl(BTextBox3.Text) > 10100 Then
                MsgBox("Weight is too high")
                returnValue = False
            End If
        Else
            If ConvertWeight(CDbl(BTextBox3.Text)) < 7000 Then
                MsgBox("Weight is too low")
                returnValue = False
            ElseIf ConvertWeight(CDbl(BTextBox3.Text)) > 10100 Then
                MsgBox("Weight is too high")
                returnValue = False
            End If
        End If
        If BRadioButton7.Checked Then
            If CDbl(BTextBox4.Text) < 0 Then
                MsgBox("Wind is too low")
                returnValue = False
            ElseIf CDbl(BTextBox4.Text) > 30 Then
                MsgBox("Wind is too high")
                returnValue = False
            End If
        Else
            If CDbl(BTextBox4.Text) < 0 Then
                MsgBox("Wind is too low")
                returnValue = False
            ElseIf CDbl(BTextBox4.Text) > 10 Then
                MsgBox("Wind is too high")
                returnValue = False
            End If
        End If
        Return returnValue
    End Function

    Public Sub BSaveData()
        TakeoffAltitude = CDbl(BTextBox1.Text)
        TakeoffTemperature = CDbl(BTextBox2.Text)
        TakeoffWeight = CDbl(BTextBox3.Text)
        TakeoffWind = CDbl(BTextBox4.Text)
        If BRadioButton2.Checked Then
            UnitTakeoffAltitude = 2
            TakeoffAltimeterSetting = CDbl(BTextBox5.Text)
            TakeoffPressureAltitude = ConvertAltitude(TakeoffAltitude, TakeoffAltimeterSetting)
        Else
            TakeoffPressureAltitude = TakeoffAltitude
            UnitTakeoffAltitude = 1
        End If
        If BRadioButton3.Checked Then UnitTakeoffTemperature = 1
        If BRadioButton4.Checked Then UnitTakeoffTemperature = 2
        If BRadioButton5.Checked Then UnitTakeoffWeight = 1
        If BRadioButton6.Checked Then UnitTakeoffWeight = 2
        If BRadioButton7.Checked Then UnitTakeoffWind = 1
        If BRadioButton8.Checked Then UnitTakeoffWind = 2
    End Sub

    Public Sub BCalculateData()
        Dim By1, By2, By3, ASBy1, ASBy2, ASBy3, AGBy1, AGBy2, AGBy3 As Double
        Dim altitude As Double = TakeoffPressureAltitude
        Dim temperature As Double = TakeoffTemperature
        Dim weight As Double = TakeoffWeight
        Dim wind As Double = TakeoffWind
        Dim PActual As Double

        If altitude <= 2000 Then
            By1 = BAltitude0(temperature) + (altitude / 2000) * (BAltitude2000(temperature) - BAltitude0(temperature))
            ASBy1 = ASBAltitude0(temperature) + (altitude / 2000) * (ASBAltitude2000(temperature) - ASBAltitude0(temperature))
            AGBy1 = AGBAltitude0(temperature) + (altitude / 2000) * (AGBAltitude2000(temperature) - AGBAltitude0(temperature))
        ElseIf altitude <= 4000 Then
            By1 = BAltitude2000(temperature) + ((altitude - 2000) / 2000) * (BAltitude4000(temperature) - BAltitude2000(temperature))
            ASBy1 = ASBAltitude2000(temperature) + ((altitude - 2000) / 2000) * (ASBAltitude4000(temperature) - ASBAltitude2000(temperature))
            AGBy1 = AGBAltitude2000(temperature) + ((altitude - 2000) / 2000) * (AGBAltitude4000(temperature) - AGBAltitude2000(temperature))
        ElseIf altitude <= 6000 Then
            By1 = BAltitude4000(temperature) + ((altitude - 4000) / 2000) * (BAltitude6000(temperature) - BAltitude4000(temperature))
            ASBy1 = ASBAltitude4000(temperature) + ((altitude - 4000) / 2000) * (ASBAltitude6000(temperature) - ASBAltitude4000(temperature))
            AGBy1 = AGBAltitude4000(temperature) + ((altitude - 4000) / 2000) * (AGBAltitude6000(temperature) - AGBAltitude4000(temperature))
        ElseIf altitude <= 8000 Then
            By1 = BAltitude6000(temperature) + ((altitude - 6000) / 2000) * (BAltitude8000(temperature) - BAltitude6000(temperature))
            ASBy1 = ASBAltitude6000(temperature) + ((altitude - 6000) / 2000) * (ASBAltitude8000(temperature) - ASBAltitude6000(temperature))
            AGBy1 = AGBAltitude6000(temperature) + ((altitude - 6000) / 2000) * (AGBAltitude8000(temperature) - AGBAltitude6000(temperature))
        ElseIf altitude <= 10000 Then
            By1 = BAltitude8000(temperature) + ((altitude - 8000) / 2000) * (BAltitude10000(temperature) - BAltitude8000(temperature))
            ASBy1 = ASBAltitude8000(temperature) + ((altitude - 8000) / 2000) * (ASBAltitude10000(temperature) - ASBAltitude8000(temperature))
            AGBy1 = AGBAltitude8000(temperature) + ((altitude - 8000) / 2000) * (AGBAltitude10000(temperature) - AGBAltitude8000(temperature))
        End If

        If By1 <= 30 Then
            By2 = BWeight20(weight) + ((By1 - 20) / 10) * (BWeight30(weight) - BWeight20(weight))
        ElseIf By1 <= 40 Then
            By2 = BWeight30(weight) + ((By1 - 30) / 10) * (BWeight40(weight) - BWeight30(weight))
        ElseIf By1 <= 50 Then
            By2 = BWeight40(weight) + ((By1 - 40) / 10) * (BWeight50(weight) - BWeight40(weight))
        ElseIf By1 <= 60 Then
            By2 = BWeight50(weight) + ((By1 - 50) / 10) * (BWeight60(weight) - BWeight50(weight))
        ElseIf By1 <= 70 Then
            By2 = BWeight60(weight) + ((By1 - 60) / 10) * (BWeight70(weight) - BWeight60(weight))
        ElseIf By1 <= 80 Then
            By2 = BWeight70(weight) + ((By1 - 70) / 10) * (BWeight80(weight) - BWeight70(weight))
        End If

        If ASBy1 <= 30 Then
            ASBy2 = ASBWeight20(weight) + ((ASBy1 - 20) / 10) * (ASBWeight30(weight) - ASBWeight20(weight))
        ElseIf ASBy1 <= 40 Then
            ASBy2 = ASBWeight30(weight) + ((ASBy1 - 30) / 10) * (ASBWeight40(weight) - ASBWeight30(weight))
        ElseIf ASBy1 <= 50 Then
            ASBy2 = ASBWeight40(weight) + ((ASBy1 - 40) / 10) * (ASBWeight50(weight) - ASBWeight40(weight))
        ElseIf ASBy1 <= 60 Then
            ASBy2 = ASBWeight50(weight) + ((ASBy1 - 50) / 10) * (ASBWeight60(weight) - ASBWeight50(weight))
        ElseIf ASBy1 <= 70 Then
            ASBy2 = ASBWeight60(weight) + ((ASBy1 - 60) / 10) * (ASBWeight70(weight) - ASBWeight60(weight))
        End If

        If AGBy1 <= 10 Then
            AGBy2 = AGBWeight5(weight) + ((AGBy1 - 5) / 5) * (AGBWeight10(weight) - AGBWeight5(weight))
        ElseIf AGBy1 <= 20 Then
            AGBy2 = AGBWeight10(weight) + ((AGBy1 - 10) / 10) * (AGBWeight20(weight) - AGBWeight10(weight))
        ElseIf AGBy1 <= 30 Then
            AGBy2 = AGBWeight20(weight) + ((AGBy1 - 20) / 10) * (AGBWeight30(weight) - AGBWeight20(weight))
        ElseIf AGBy1 <= 40 Then
            AGBy2 = AGBWeight30(weight) + ((AGBy1 - 30) / 10) * (AGBWeight40(weight) - AGBWeight30(weight))
        ElseIf AGBy1 <= 50 Then
            AGBy2 = AGBWeight40(weight) + ((AGBy1 - 40) / 10) * (AGBWeight50(weight) - AGBWeight40(weight))
        End If

        If altitude < 2000 Then
            PActual = 1520
        ElseIf altitude < 4000 Then
            PActual = Power2000(temperature) + ((altitude - 2000) / 2000) * (Power4000(temperature) - Power2000(temperature))
        ElseIf altitude < 6000 Then
            PActual = Power4000(temperature) + ((altitude - 4000) / 2000) * (Power6000(temperature) - Power4000(temperature))
        ElseIf altitude < 8000 Then
            PActual = Power6000(temperature) + ((altitude - 6000) / 2000) * (Power8000(temperature) - Power6000(temperature))
        ElseIf altitude <= 10000 Then
            PActual = Power8000(temperature) + ((altitude - 8000) / 2000) * (Power10000(temperature) - Power8000(temperature))
        End If

        If PActual < PowerMinimum(temperature) Then MsgBox("Power is not enough for takeoff")
        If PActual > 1520 Then PActual = 1520

        If weight < 8000 Then
            V1 = 85
            V2 = 97
        ElseIf weight < 9000 Then
            V1 = 85 + ((weight - 8000) / 1000) * (88 - 85)
            V2 = 97 + ((weight - 8000) / 1000) * (99 - 97)
        ElseIf weight <= 10100 Then
            V1 = 88 + ((weight - 9000) / 1100) * (93 - 88)
            V2 = 99 + ((weight - 9000) / 1100) * (102 - 99)
        End If
        If weight <= 10100 Then Veref = 102
        If weight < 9700 Then Veref = 101
        If weight < 9200 Then Veref = 100
        If weight < 8400 Then Veref = 99

        If UnitTakeoffWind = 1 Then
            If By2 <= 20 Then
                By3 = BHWind10(wind) + ((By2 - 10) / 10) * (BHWind20(wind) - BHWind10(wind))
            ElseIf By2 <= 30 Then
                By3 = BHWind20(wind) + ((By2 - 20) / 10) * (BHWind30(wind) - BHWind20(wind))
            ElseIf By2 <= 40 Then
                By3 = BHWind30(wind) + ((By2 - 30) / 10) * (BHWind40(wind) - BHWind30(wind))
            ElseIf By2 <= 50 Then
                By3 = BHWind40(wind) + ((By2 - 40) / 10) * (BHWind50(wind) - BHWind40(wind))
            ElseIf By2 <= 60 Then
                By3 = BHWind50(wind) + ((By2 - 50) / 10) * (BHWind60(wind) - BHWind50(wind))
            ElseIf By2 <= 70 Then
                By3 = BHWind60(wind) + ((By2 - 60) / 10) * (BHWind70(wind) - BHWind60(wind))
            ElseIf By2 <= 80 Then
                By3 = BHWind70(wind) + ((By2 - 70) / 10) * (BHWind80(wind) - BHWind70(wind))
            End If

            If ASBy2 <= 20 Then
                ASBy3 = ASBHWind10(wind) + ((ASBy2 - 10) / 10) * (ASBHWind20(wind) - ASBHWind10(wind))
            ElseIf ASBy2 <= 30 Then
                ASBy3 = ASBHWind20(wind) + ((ASBy2 - 20) / 10) * (ASBHWind30(wind) - ASBHWind20(wind))
            ElseIf ASBy2 <= 40 Then
                ASBy3 = ASBHWind30(wind) + ((ASBy2 - 30) / 10) * (ASBHWind40(wind) - ASBHWind30(wind))
            ElseIf ASBy2 <= 50 Then
                ASBy3 = ASBHWind40(wind) + ((ASBy2 - 40) / 10) * (ASBHWind50(wind) - ASBHWind40(wind))
            ElseIf ASBy2 <= 60 Then
                ASBy3 = ASBHWind50(wind) + ((ASBy2 - 50) / 10) * (ASBHWind60(wind) - ASBHWind50(wind))
            ElseIf ASBy2 <= 70 Then
                ASBy3 = ASBHWind60(wind) + ((ASBy2 - 60) / 10) * (ASBHWind70(wind) - ASBHWind60(wind))
            End If

            If AGBy2 <= 10 Then
                AGBy3 = AGBHWind5(wind) + ((AGBy2 - 5) / 5) * (AGBHWind10(wind) - AGBHWind5(wind))
            ElseIf AGBy2 <= 20 Then
                AGBy3 = AGBHWind10(wind) + ((AGBy2 - 10) / 10) * (AGBHWind20(wind) - AGBHWind10(wind))
            ElseIf AGBy2 <= 30 Then
                AGBy3 = AGBHWind20(wind) + ((AGBy2 - 20) / 10) * (AGBHWind30(wind) - AGBHWind20(wind))
            ElseIf AGBy2 <= 40 Then
                AGBy3 = AGBHWind30(wind) + ((AGBy2 - 30) / 10) * (AGBHWind40(wind) - AGBHWind30(wind))
            ElseIf AGBy2 <= 50 Then
                AGBy3 = AGBHWind40(wind) + ((AGBy2 - 40) / 10) * (AGBHWind50(wind) - AGBHWind40(wind))
            ElseIf AGBy2 <= 60 Then
                AGBy3 = AGBHWind50(wind) + ((AGBy2 - 50) / 10) * (AGBHWind60(wind) - AGBHWind50(wind))
            ElseIf AGBy2 <= 70 Then
                AGBy3 = AGBHWind60(wind) + ((AGBy2 - 60) / 10) * (AGBHWind70(wind) - AGBHWind60(wind))
            ElseIf AGBy2 <= 80 Then
                AGBy3 = AGBHWind70(wind) + ((AGBy2 - 70) / 10) * (AGBHWind80(wind) - AGBHWind70(wind))
            ElseIf AGBy2 <= 90 Then
                AGBy3 = AGBHWind80(wind) + ((AGBy2 - 80) / 10) * (AGBHWind90(wind) - AGBHWind80(wind))
            End If

            If AGBy2 > 90 Then AGBy3 = 90

        Else
            If By2 <= 20 Then
                By3 = BTWind10(wind) + ((By2 - 10) / 10) * (BTWind20(wind) - BTWind10(wind))
            ElseIf By2 <= 30 Then
                By3 = BTWind20(wind) + ((By2 - 20) / 10) * (BTWind30(wind) - BTWind20(wind))
            ElseIf By2 <= 40 Then
                By3 = BTWind30(wind) + ((By2 - 30) / 10) * (BTWind40(wind) - BTWind30(wind))
            ElseIf By2 <= 50 Then
                By3 = BTWind40(wind) + ((By2 - 40) / 10) * (BTWind50(wind) - BTWind40(wind))
            ElseIf By2 <= 60 Then
                By3 = BTWind50(wind) + ((By2 - 50) / 10) * (BTWind60(wind) - BTWind50(wind))
            ElseIf By2 <= 70 Then
                By3 = BTWind60(wind) + ((By2 - 60) / 10) * (BTWind70(wind) - BTWind60(wind))
            ElseIf By2 <= 80 Then
                By3 = BTWind70(wind) + ((By2 - 70) / 10) * (BTWind80(wind) - BTWind70(wind))
            End If

            If ASBy2 <= 20 Then
                ASBy3 = ASBTWind10(wind) + ((ASBy2 - 10) / 10) * (ASBTWind20(wind) - ASBTWind10(wind))
            ElseIf ASBy2 <= 30 Then
                ASBy3 = ASBTWind20(wind) + ((ASBy2 - 20) / 10) * (ASBTWind30(wind) - ASBTWind20(wind))
            ElseIf ASBy2 <= 40 Then
                ASBy3 = ASBTWind30(wind) + ((ASBy2 - 30) / 10) * (ASBTWind40(wind) - ASBTWind30(wind))
            ElseIf ASBy2 <= 50 Then
                ASBy3 = ASBTWind40(wind) + ((ASBy2 - 40) / 10) * (ASBTWind50(wind) - ASBTWind40(wind))
            ElseIf ASBy2 <= 60 Then
                ASBy3 = ASBTWind50(wind) + ((ASBy2 - 50) / 10) * (ASBTWind60(wind) - ASBTWind50(wind))
            ElseIf ASBy2 <= 70 Then
                ASBy3 = ASBTWind60(wind) + ((ASBy2 - 60) / 10) * (ASBTWind70(wind) - ASBTWind60(wind))
            End If

            If AGBy2 <= 10 Then
                AGBy3 = AGBTWind5(wind) + ((AGBy2 - 5) / 5) * (AGBTWind10(wind) - AGBTWind5(wind))
            ElseIf AGBy2 <= 20 Then
                AGBy3 = AGBTWind10(wind) + ((AGBy2 - 10) / 10) * (AGBTWind20(wind) - AGBTWind10(wind))
            ElseIf AGBy2 <= 30 Then
                AGBy3 = AGBTWind20(wind) + ((AGBy2 - 20) / 10) * (AGBTWind30(wind) - AGBTWind20(wind))
            ElseIf AGBy2 <= 40 Then
                AGBy3 = AGBTWind30(wind) + ((AGBy2 - 30) / 10) * (AGBTWind40(wind) - AGBTWind30(wind))
            ElseIf AGBy2 <= 50 Then
                AGBy3 = AGBTWind40(wind) + ((AGBy2 - 40) / 10) * (AGBTWind50(wind) - AGBTWind40(wind))
            ElseIf AGBy2 <= 60 Then
                AGBy3 = AGBTWind50(wind) + ((AGBy2 - 50) / 10) * (AGBTWind60(wind) - AGBTWind50(wind))
            ElseIf AGBy2 <= 70 Then
                AGBy3 = AGBTWind60(wind) + ((AGBy2 - 60) / 10) * (AGBTWind70(wind) - AGBTWind60(wind))
            End If

            If AGBy2 > 90 Then AGBy3 = 90

        End If
        TakeoffDistance = CInt(By3 * 100)
        AccelerateStopDistance = CInt((ASBy3 * 100) + 1000)
        AccelerateGoDistance = CInt((AGBy3 * 100) + 1000)
        TakeoffPower = PActual
        TakeoffPerformanceFormStatus = 1
    End Sub

    Public Sub BShowData()
        BResultLabel1.Visible = True
        BResultLabel2.Visible = True
        BResultLabel3.Visible = True
        BResultLabel4.Visible = True
        BResultLabel5.Visible = True
        BResultLabel6.Visible = True
        BResultLabel7.Visible = True
        BResultLabel8.Visible = True
        BResultLabel9.Visible = True
        BResultValue1.Visible = True
        BResultValue2.Visible = True
        BResultValue3.Visible = True
        BResultValue4.Visible = True
        BResultValue5.Visible = True
        BResultValue6.Visible = True
        BResultValue7.Visible = True
        BResultValue8.Visible = True
        BResultValue9.Visible = True
        BResultValue1.Text = CInt(TakeoffDistance).ToString + "  Feet"
        BResultValue2.Text = CInt(AccelerateStopDistance).ToString + "  Feet"
        BResultValue3.Text = CInt(AccelerateGoDistance).ToString + "  Feet"
        BResultValue4.Text = CInt(TakeoffPower).ToString + "  Ft-Lbs"
        BResultValue5.Text = CInt(Veref).ToString + "  Knots"
        BResultValue6.Text = CInt(V1).ToString + "  Knots"
        BResultValue7.Text = CInt(V1).ToString + "  Knots"
        BResultValue8.Text = CInt(V2).ToString + "  Knots"
        BResultValue9.Text = "108  Knots"

    End Sub

    Private Sub BButton1_Click(sender As Object, e As EventArgs) Handles BButton1.Click
        If BCheckData() Then
            If CurrentToolbar = 0 Then ChangeToolbar(1)
            BSaveData()
            BCalculateData()
            BShowData()
            ResetLabelColors()
        End If
    End Sub

    Private Sub BRadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles BRadioButton2.CheckedChanged
        If BRadioButton2.Checked Then
            BTextBox5.Enabled = True
        Else
            BTextBox5.Enabled = False
        End If
    End Sub

    Private Sub BTextBox_KeyDown(sender As Object, e As KeyEventArgs) Handles BTextBox1.KeyDown, BTextBox2.KeyDown, BTextBox3.KeyDown, BTextBox4.KeyDown, BTextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            BButton1.PerformClick()
        End If
    End Sub
#End Region
#Region " Takeoff Equations "

    Function BAltitude0(ByVal x As Double) As Double
        Return (0.16 * x + 27.6)
    End Function

    Function BAltitude2000(ByVal x As Double) As Double
        Return ((7.2 / 40) * x + 31)
    End Function

    Function BAltitude4000(ByVal x As Double) As Double
        If x <= 36 Then
            Return ((10 / 49) * x + (244 / 7))
        Else
            Return (0.075 * x ^ 2 - 5.05 * x + 126.8)
        End If
    End Function

    Function BAltitude6000(ByVal x As Double) As Double
        If x <= 28 Then
            Return ((8 / 35) * x + (196 / 5))
        Else
            Return (0.0143 * x ^ 2 - 0.1443 * x + 38.454)
        End If
    End Function

    Function BAltitude8000(ByVal x As Double) As Double
        If x <= 19 Then
            Return ((11 / 42) * x + (310 / 7))
        Else
            Return (0.0089 * x ^ 2 + 0.3813 * x + 38.862)
        End If
    End Function

    Function BAltitude10000(ByVal x As Double) As Double
        If x <= 10 Then
            Return (0.3 * x + 50.3)
        Else
            Return (0.0173 * x ^ 2 + 0.3383 * x + 48.285)
        End If
    End Function

    Function BWeight20(ByVal x As Double) As Double
        If x >= 9000 Then
            Return ((21 / 5500) * x - (1021 / 55))
        ElseIf x >= 8300 Then
            Return ((23 / 7000) * x - (482 / 35))
        Else
            Return ((23 / 13000) * x - (77 / 65))
        End If
    End Function

    Function BWeight30(ByVal x As Double) As Double
        If x >= 9000 Then
            Return ((3 / 550) * x - (276 / 11))
        ElseIf x >= 8300 Then
            Return ((9 / 1750) * x - (156 / 7))
        Else
            Return ((37 / 13000) * x - (419 / 130))
        End If
    End Function

    Function BWeight40(ByVal x As Double) As Double
        If x >= 9000 Then
            Return ((43 / 5500) * x - (2143 / 55))
        ElseIf x >= 8300 Then
            Return ((11 / 1750) * x - (881 / 35))
        Else
            Return ((1 / 260) * x - (64 / 13))
        End If
    End Function

    Function BWeight50(ByVal x As Double) As Double
        If x >= 9000 Then
            Return ((27 / 2750) * x - (2704 / 55))
        ElseIf x >= 8300 Then
            Return ((57 / 7000) * x - (1193 / 35))
        Else
            Return ((1 / 200) * x - 8)
        End If
    End Function

    Function BWeight60(ByVal x As Double) As Double
        If x >= 9000 Then
            Return ((67 / 5500) * x - (3467 / 55))
        ElseIf x >= 8300 Then
            Return ((71 / 7000) * x - (1564 / 35))
        Else
            Return ((3 / 520) * x - (109 / 13))
        End If
    End Function

    Function BWeight70(ByVal x As Double) As Double
        If x >= 9000 Then
            Return ((83 / 5500) * x - (4533 / 55))
        ElseIf x >= 8300 Then
            Return ((3 / 250) * x - (273 / 5))
        Else
            Return ((9 / 1300) * x - (162 / 13))
        End If
    End Function

    Function BWeight80(ByVal x As Double) As Double
        If x >= 9000 Then
            Return ((49 / 2750) * x - (5498 / 55))
        ElseIf x >= 8300 Then
            Return ((27 / 1750) * x - (2746 / 35))
        Else
            Return ((12 / 1625) * x - (152 / 13))
        End If
    End Function

    Function BHWind10(ByVal x As Double) As Double
        BHWind10 = -0.0666666666666667 * x + 10
    End Function

    Function BHWind20(ByVal x As Double) As Double
        BHWind20 = -0.13333333333 * x + 20
    End Function

    Function BHWind30(ByVal x As Double) As Double
        BHWind30 = -0.2 * x + 30
    End Function

    Function BHWind40(ByVal x As Double) As Double
        BHWind40 = -0.26666666666666666 * x + 40
    End Function

    Function BHWind50(ByVal x As Double) As Double
        BHWind50 = -0.32 * x + 50
    End Function

    Function BHWind60(ByVal x As Double) As Double
        BHWind60 = -0.36666666666666664 * x + 60
    End Function

    Function BHWind70(ByVal x As Double) As Double
        BHWind70 = -0.4 * x + 70
    End Function

    Function BHWind80(ByVal x As Double) As Double
        BHWind80 = -0.46666666666666667 * x + 80
    End Function

    Function BTWind10(ByVal x As Double) As Double
        BTWind10 = 0.38 * x + 10
    End Function

    Function BTWind20(ByVal x As Double) As Double
        BTWind20 = 0.6 * x + 20
    End Function
    Function BTWind30(ByVal x As Double) As Double
        BTWind30 = 0.8 * x + 30
    End Function
    Function BTWind40(ByVal x As Double) As Double
        BTWind40 = 0.9 * x + 40
    End Function
    Function BTWind50(ByVal x As Double) As Double
        BTWind50 = 1.1 * x + 50
    End Function
    Function BTWind60(ByVal x As Double) As Double
        BTWind60 = 1.25 * x + 60
    End Function
    Function BTWind70(ByVal x As Double) As Double
        BTWind70 = 1.4 * x + 70
    End Function
    Function BTWind80(ByVal x As Double) As Double
        BTWind80 = 1.3 * x + 80
    End Function

    Function ASBAltitude0(ByVal x As Double) As Double
        Return ((4 / 25) * x + (138 / 5))
    End Function

    Function ASBAltitude2000(ByVal x As Double) As Double
        If x <= 40 Then
            Return ((9 / 50) * x + (154 / 5))
        Else
            Return (0.0125 * x ^ 2 - 0.675 * x + 45)
        End If
    End Function

    Function ASBAltitude4000(ByVal x As Double) As Double
        If x <= 32 Then
            Return ((10 / 51) * x + (1771 / 51))
        Else
            Return (0.0083 * x ^ 2 - 0.225 * x + 39.667)
        End If
    End Function

    Function ASBAltitude6000(ByVal x As Double) As Double
        If x <= 28 Then
            Return ((57 / 250) * x + (4877 / 125))
        Else
            Return (0.0125 * x ^ 2 - 0.325 * x + 44.7)
        End If
    End Function

    Function ASBAltitude8000(ByVal x As Double) As Double
        If x <= 18 Then
            Return ((123 / 500) * x + (5484 / 125))
        Else
            Return (0.0155 * x ^ 2 - 0.2381 * x + 47.571)
        End If
    End Function

    Function ASBAltitude10000(ByVal x As Double) As Double
        If x <= 6 Then
            Return ((11 / 40) * x + (987 / 20))
        Else
            Return (0.0111 * x ^ 2 + 0.1667 * x + 49.6)
        End If
    End Function

    Function ASBWeight20(ByVal x As Double) As Double
        If x >= 9000 Then
            Return ((1 / 275) * x - (184 / 11))
        ElseIf x >= 8300 Then
            Return ((23 / 7000) * x - (95 / 7))
        Else
            Return ((17 / 13000) * x + (37 / 13))
        End If
    End Function

    Function ASBWeight30(ByVal x As Double) As Double
        If x >= 9000 Then
            Return ((59 / 11000) * x - (2659 / 110))
        ElseIf x >= 8300 Then
            Return ((31 / 7000) * x - (1103 / 70))
        Else
            Return ((1 / 650) * x + (107 / 13))
        End If
    End Function

    Function ASBWeight40(ByVal x As Double) As Double
        If x >= 9000 Then
            Return ((7 / 1000) * x - (307 / 10))
        ElseIf x >= 8300 Then
            Return ((41 / 7000) * x - (1429 / 70))
        Else
            Return ((1 / 520) * x + (1591 / 130))
        End If
    End Function

    Function ASBWeight50(ByVal x As Double) As Double
        If x >= 9000 Then
            Return ((97 / 11000) * x - (4297 / 110))
        ElseIf x >= 8300 Then
            Return ((47 / 7000) * x - (1409 / 70))
        Else
            Return ((9 / 3250) * x + (164 / 13))
        End If
    End Function

    Function ASBWeight60(ByVal x As Double) As Double
        If x >= 9000 Then
            Return ((3 / 275) * x - (552 / 11))
        ElseIf x >= 8300 Then
            Return ((3 / 350) * x - (204 / 7))
        Else
            Return ((1 / 325) * x + (214 / 13))
        End If
    End Function

    Function ASBWeight70(ByVal x As Double) As Double
        If x >= 9000 Then
            Return ((73 / 5500) * x - (3523 / 55))
        ElseIf x >= 8300 Then
            Return ((9 / 875) * x - (1301 / 35))
        Else
            Return ((1 / 250) * x + 15)
        End If
    End Function

    Function ASBHWind10(ByVal x As Double) As Double
        ASBHWind10 = (-0.2 * x) + 10
    End Function

    Function ASBHWind20(ByVal x As Double) As Double
        ASBHWind20 = (-0.266666666666667 * x) + 20
    End Function

    Function ASBHWind30(ByVal x As Double) As Double
        ASBHWind30 = (-0.333333333333333 * x) + 30
    End Function

    Function ASBHWind40(ByVal x As Double) As Double
        ASBHWind40 = (-0.4 * x) + 40
    End Function

    Function ASBHWind50(ByVal x As Double) As Double
        ASBHWind50 = (-0.466666666666667 * x) + 50
    End Function

    Function ASBHWind60(ByVal x As Double) As Double
        ASBHWind60 = (-0.5 * x) + 60
    End Function

    Function ASBHWind70(ByVal x As Double) As Double
        ASBHWind70 = (-0.566666666666667 * x) + 70
    End Function

    Function ASBTWind10(ByVal x As Double) As Double
        ASBTWind10 = (0.9 * x) + 10
    End Function

    Function ASBTWind20(ByVal x As Double) As Double
        ASBTWind20 = (1 * x) + 20
    End Function

    Function ASBTWind30(ByVal x As Double) As Double
        ASBTWind30 = (1.2 * x) + 30
    End Function
    Function ASBTWind40(ByVal x As Double) As Double
        ASBTWind40 = (1.4 * x) + 40
    End Function
    Function ASBTWind50(ByVal x As Double) As Double
        ASBTWind50 = (1.6 * x) + 50
    End Function
    Function ASBTWind60(ByVal x As Double) As Double
        ASBTWind60 = (1.8 * x) + 60
    End Function
    Function ASBTWind70(ByVal x As Double) As Double
        ASBTWind70 = (1.8 * x) + 70
    End Function

    Function AGBAltitude0(ByVal x As Double) As Double
        AGBAltitude0 = (1 / 8) * x + (97 / 8)
    End Function

    Function AGBAltitude2000(ByVal x As Double) As Double
        If x <= 41 Then
            Return ((41 / 305) * x + (896 / 61))
        Else
            Return (0.0167 * x ^ 2 - 0.9833 * x + 32.5)
        End If
    End Function

    Function AGBAltitude4000(ByVal x As Double) As Double
        If x <= 32 Then
            Return ((83 / 540) * x + (472 / 27))
        Else
            Return (0.0167 * x ^ 2 - 0.75 * x + 29.333)
        End If
    End Function

    Function AGBAltitude6000(ByVal x As Double) As Double
        If x <= 28 Then
            Return ((9 / 50) * x + (524 / 25))
        Else
            Return (0.0042 * x ^ 2 + 0.4083 * x + 11.3)
        End If
    End Function

    Function AGBAltitude8000(ByVal x As Double) As Double
        If x <= 19 Then
            Return ((33 / 160) * x + (3997 / 160))
        Else
            Return (0.0148 * x ^ 2 - 0.0167 * x + 23.872)
        End If
    End Function

    Function AGBAltitude10000(ByVal x As Double) As Double
        If x <= 10 Then
            Return ((5 / 21) * x + (622 / 21))
        Else
            Return (0.0125 * x ^ 2 + 0.3417 * x + 27.333)
        End If
    End Function

    Function AGBWeight20(ByVal x As Double) As Double
        AGBWeight20 = 0.0000015251 * x ^ 2 - 0.0186207037 * x + 75.6120947485
    End Function

    Function AGBWeight30(ByVal x As Double) As Double
        AGBWeight30 = 0.0000021671 * x ^ 2 - 0.0263658862 * x + 108.4472543968
    End Function

    Function AGBWeight40(ByVal x As Double) As Double
        AGBWeight40 = 1.1142857143 * (1 + (x - 7000) / 500) ^ 2 - 0.6428571429 * (1 + (x - 7000) / 500) + 39.8
    End Function

    Function AGBWeight50(ByVal x As Double) As Double
        AGBWeight50 = 1.55 * (1 + (x - 7000) / 500) ^ 2 - 1.05 * (1 + (x - 7000) / 500) + 49.88
    End Function

    Function AGBWeight10(ByVal x As Double) As Double
        AGBWeight10 = 0.0000008738 * x ^ 2 - 0.0105138755 * x + 40.6633478438
    End Function

    Function AGBWeight5(ByVal x As Double) As Double
        If x <= 9100 Then
            Return ((3 / 1400) * x - 10)
        Else
            Return ((7 / 2000) * x - (447 / 20))
        End If
    End Function

    Function AGBHWind5(ByVal x As Double) As Double
        AGBHWind5 = (-0.133333333333333 * x) + 5
    End Function

    Function AGBHWind10(ByVal x As Double) As Double
        AGBHWind10 = (-0.133333333333333 * x) + 10
    End Function

    Function AGBHWind20(ByVal x As Double) As Double
        AGBHWind20 = (-0.166666666666667 * x) + 20
    End Function

    Function AGBHWind30(ByVal x As Double) As Double
        AGBHWind30 = (-0.233333333333333 * x) + 30
    End Function

    Function AGBHWind40(ByVal x As Double) As Double
        AGBHWind40 = (-0.3 * x) + 40
    End Function

    Function AGBHWind50(ByVal x As Double) As Double
        AGBHWind50 = (-0.366666666666667 * x) + 50
    End Function

    Function AGBHWind60(ByVal x As Double) As Double
        AGBHWind60 = (-0.4 * x) + 60
    End Function

    Function AGBHWind70(ByVal x As Double) As Double
        AGBHWind70 = (-0.433333333333333 * x) + 70
    End Function

    Function AGBHWind80(ByVal x As Double) As Double
        AGBHWind80 = (-0.5 * x) + 80
    End Function

    Function AGBHWind90(ByVal x As Double) As Double
        AGBHWind90 = (-0.533333333333333 * x) + 90
    End Function

    Function AGBTWind5(ByVal x As Double) As Double
        AGBTWind5 = (0.5 * x) + 5
    End Function
    Function AGBTWind10(ByVal x As Double) As Double
        AGBTWind10 = (0.6 * x) + 10
    End Function
    Function AGBTWind20(ByVal x As Double) As Double
        AGBTWind20 = (0.8 * x) + 20
    End Function
    Function AGBTWind30(ByVal x As Double) As Double
        AGBTWind30 = (1 * x) + 30
    End Function
    Function AGBTWind40(ByVal x As Double) As Double
        AGBTWind40 = (1.1 * x) + 40
    End Function
    Function AGBTWind50(ByVal x As Double) As Double
        AGBTWind50 = (1.3 * x) + 50
    End Function
    Function AGBTWind60(ByVal x As Double) As Double
        AGBTWind60 = (1.4 * x) + 60
    End Function
    Function AGBTWind70(ByVal x As Double) As Double
        AGBTWind70 = (1.6 * x) + 70
    End Function

    Function PowerMinimum(ByVal x As Double) As Double
        PowerMinimum = (16.6666666666667 * x) + 783.333333333333
    End Function

    Function Power10000(ByVal x As Double) As Double
        Power10000 = (-13.75 * x) + 1657.5
    End Function

    Function Power8000(ByVal x As Double) As Double
        Power8000 = (-15 * x) + 1810
    End Function

    Function Power6000(ByVal x As Double) As Double
        Power6000 = (-16.25 * x) + 1975
    End Function

    Function Power4000(ByVal x As Double) As Double
        Power4000 = (-17.5 * x) + 2150
    End Function

    Function Power2000(ByVal x As Double) As Double
        Power2000 = (-20 * x) + 2390
    End Function

#End Region
#Region " Landing "

    Public Sub ELoadData()
        ETextBox1.Focus()
        If LandingPerformanceFormStatus = 1 Then
            ETextBox1.Text = LandingAltitude
            ETextBox2.Text = LandingTemperature
            ETextBox3.Text = LandingWeight
            ETextBox4.Text = LandingWind
            EShowData()
        Else
            If LandingAltitude <> 999999 Then ETextBox1.Text = LandingAltitude
            If LandingTemperature <> 999999 Then ETextBox2.Text = LandingTemperature
            If LandingWeight <> 999999 Then ETextBox3.Text = LandingWeight
            If LandingWind <> 999999 Then ETextBox4.Text = LandingWind
        End If
        If UnitLandingAltitude = 2 Then
            ERadioButton2.Checked = True
            ETextBox5.Text = LandingAltimeterSetting
        Else
            ERadioButton1.Checked = True
            ETextBox5.Text = ""
        End If
        If UnitLandingTemperature = 2 Then ERadioButton4.Checked = True Else ERadioButton3.Checked = True
        If UnitLandingWeight = 2 Then ERadioButton6.Checked = True Else ERadioButton5.Checked = True
        If UnitLandingWind = 2 Then ERadioButton8.Checked = True Else ERadioButton7.Checked = True
    End Sub

    Public Function ECheckData() As Boolean
        Dim returnValue As Boolean = True
        If ETextBox1.Text = "" Then ETextBox1.Text = "0"
        If ETextBox2.Text = "" Then ETextBox2.Text = "15"
        If ETextBox3.Text = "" Then ETextBox3.Text = "9500"
        If ETextBox4.Text = "" Then ETextBox4.Text = "0"
        If (Not IsNumeric(ETextBox1.Text)) Or (Not IsNumeric(ETextBox2.Text)) Or (Not IsNumeric(ETextBox3.Text)) Or (Not IsNumeric(ETextBox4.Text)) Then
            MsgBox("Invalid Input!")
            returnValue = False
        End If
        If ERadioButton2.Checked Then
            If ETextBox5.Text = "" Then ETextBox5.Text = "29.92"
            If (Not IsNumeric(ETextBox5.Text)) Then
                MsgBox("Invalid Input!")
                returnValue = False
            End If
            If ConvertAltitude(CDbl(ETextBox1.Text), CDbl(ETextBox5.Text)) > 10000 Then
                MsgBox("Pressure Altitude is too high")
                returnValue = False
            ElseIf ConvertAltitude(CDbl(ETextBox1.Text), CDbl(ETextBox5.Text)) < 0 Then
                MsgBox("Pressure Altitude is too low")
                returnValue = False
            End If
        Else
            If CDbl(ETextBox1.Text) > 10000 Then
                MsgBox("Pressure Altitude is too high")
                returnValue = False
            ElseIf CDbl(ETextBox1.Text) < 0 Then
                MsgBox("Pressure Altitude is too low")
                returnValue = False
            End If
        End If
        If ERadioButton3.Checked Then
            If CDbl(ETextBox2.Text) > 50 Then
                MsgBox("Temperature is too high")
                returnValue = False
            ElseIf CDbl(ETextBox2.Text) < -40 Then
                MsgBox("Temperature is too low")
                returnValue = False
            End If
        Else
            If ConvertTemperature(CDbl(ETextBox2.Text)) > 50 Then
                MsgBox("Temperature is too high")
                returnValue = False
            ElseIf ConvertTemperature(CDbl(ETextBox2.Text)) < -40 Then
                MsgBox("Temperature is too low")
                returnValue = False
            End If
        End If
        If ERadioButton5.Checked Then
            If CDbl(ETextBox3.Text) < 7000 Then
                MsgBox("Weight is too low")
                returnValue = False
            ElseIf CDbl(ETextBox3.Text) > 10100 Then
                MsgBox("Weight is too high")
                returnValue = False
            End If
        Else
            If ConvertWeight(CDbl(ETextBox3.Text)) < 7000 Then
                MsgBox("Weight is too low")
                returnValue = False
            ElseIf ConvertWeight(CDbl(ETextBox3.Text)) > 10100 Then
                MsgBox("Weight is too high")
                returnValue = False
            End If
        End If
        If ERadioButton7.Checked Then
            If CDbl(ETextBox4.Text) < 0 Then
                MsgBox("Wind is too low")
                returnValue = False
            ElseIf CDbl(ETextBox4.Text) > 30 Then
                MsgBox("Wind is too high")
                returnValue = False
            End If
        Else
            If CDbl(ETextBox4.Text) < 0 Then
                MsgBox("Wind is too low")
                returnValue = False
            ElseIf CDbl(ETextBox4.Text) > 10 Then
                MsgBox("Wind is too high")
                returnValue = False
            End If
        End If
        Return returnValue
    End Function

    Public Sub ESaveData()
        LandingAltitude = CDbl(ETextBox1.Text)
        LandingTemperature = CDbl(ETextBox2.Text)
        LandingWeight = CDbl(ETextBox3.Text)
        LandingWind = CDbl(ETextBox4.Text)
        If ERadioButton2.Checked Then
            UnitLandingAltitude = 2
            LandingAltimeterSetting = CDbl(ETextBox5.Text)
            LandingPressureAltitude = ConvertAltitude(LandingAltitude, LandingAltimeterSetting)
        Else
            LandingPressureAltitude = LandingAltitude
            UnitLandingAltitude = 1
        End If
        If ERadioButton3.Checked Then UnitLandingTemperature = 1
        If ERadioButton4.Checked Then UnitLandingTemperature = 2
        If ERadioButton5.Checked Then UnitLandingWeight = 1
        If ERadioButton6.Checked Then UnitLandingWeight = 2
        If ERadioButton7.Checked Then UnitLandingWind = 1
        If ERadioButton8.Checked Then UnitLandingWind = 2
    End Sub

    Public Sub ECalculateData()
        Dim Ey1, Ey2, Ey3, Ey4 As Double
        Dim altitude As Double = LandingPressureAltitude
        Dim temperature As Double = LandingTemperature
        Dim weight As Double = LandingWeight
        Dim wind As Double = LandingWind

        If altitude <= 2000 Then
            Ey1 = EAltitude0(temperature) + (altitude / 2000) * (EAltitude2000(temperature) - EAltitude0(temperature))
        ElseIf altitude <= 4000 Then
            Ey1 = EAltitude2000(temperature) + ((altitude - 2000) / 2000) * (EAltitude4000(temperature) - EAltitude2000(temperature))
        ElseIf altitude <= 6000 Then
            Ey1 = EAltitude4000(temperature) + ((altitude - 4000) / 2000) * (EAltitude6000(temperature) - EAltitude4000(temperature))
        ElseIf altitude <= 8000 Then
            Ey1 = EAltitude6000(temperature) + ((altitude - 6000) / 2000) * (EAltitude8000(temperature) - EAltitude6000(temperature))
        ElseIf altitude <= 10000 Then
            Ey1 = EAltitude8000(temperature) + ((altitude - 8000) / 2000) * (EAltitude10000(temperature) - EAltitude8000(temperature))
        End If

        If Ey1 <= 1340 Then
            Ey2 = EWeight1040(weight) + ((Ey1 - 1040) / (1340 - 1040)) * (EWeight1340(weight) - EWeight1040(weight))
        ElseIf Ey1 <= 1650 Then
            Ey2 = EWeight1340(weight) + ((Ey1 - 1340) / (1650 - 1340)) * (EWeight1650(weight) - EWeight1340(weight))
        ElseIf Ey1 <= 1970 Then
            Ey2 = EWeight1650(weight) + ((Ey1 - 1650) / (1970 - 1650)) * (EWeight1970(weight) - EWeight1650(weight))
        ElseIf Ey1 <= 2280 Then
            Ey2 = EWeight1970(weight) + ((Ey1 - 1970) / (2280 - 1970)) * (EWeight2280(weight) - EWeight1970(weight))
        End If

        If UnitLandingWind = 1 Then
            If Ey2 <= 1300 Then
                Ey3 = EHWind1000(wind) + ((Ey2 - 1000) / 300) * (EHWind1300(wind) - EHWind1000(wind))
            ElseIf Ey2 <= 1600 Then
                Ey3 = EHWind1300(wind) + ((Ey2 - 1300) / 300) * (EHWind1600(wind) - EHWind1300(wind))
            ElseIf Ey2 <= 1900 Then
                Ey3 = EHWind1600(wind) + ((Ey2 - 1600) / 300) * (EHWind1900(wind) - EHWind1600(wind))
            ElseIf Ey2 <= 2200 Then
                Ey3 = EHWind1900(wind) + ((Ey2 - 1900) / 300) * (EHWind2200(wind) - EHWind1900(wind))
            End If
        Else
            If Ey2 <= 1300 Then
                Ey3 = ETWind1000(wind) + ((Ey2 - 1000) / 300) * (ETWind1300(wind) - ETWind1000(wind))
            ElseIf Ey2 <= 1600 Then
                Ey3 = ETWind1300(wind) + ((Ey2 - 1300) / 300) * (ETWind1600(wind) - ETWind1300(wind))
            ElseIf Ey2 <= 1900 Then
                Ey3 = ETWind1600(wind) + ((Ey2 - 1600) / 300) * (ETWind1900(wind) - ETWind1600(wind))
            ElseIf Ey2 <= 2200 Then
                Ey3 = ETWind1900(wind) + ((Ey2 - 1900) / 300) * (ETWind2200(wind) - ETWind1900(wind))
            End If
        End If

        If Ey3 <= 1000 Then
            Ey4 = 1300 + ((Ey3 - 500) / 500) * (1930 - 1300)
        ElseIf Ey3 <= 1500 Then
            Ey4 = 1930 + ((Ey3 - 1000) / 500) * (2600 - 1930)
        ElseIf Ey3 <= 2000 Then
            Ey4 = 2600 + ((Ey3 - 1500) / 500) * (3220 - 2600)
        ElseIf Ey3 <= 2500 Then
            Ey4 = 3220 + ((Ey3 - 2000) / 500) * (3840 - 3220)
        ElseIf Ey3 <= 3000 Then
            Ey4 = 3840 + ((Ey3 - 2500) / 500) * (4450 - 3840)
        End If

        If weight <= 9600 Then Vref = 101
        If weight <= 9100 Then Vref = 100
        If weight <= 8300 Then Vref = 99

        LandingDistance = Ey4
        LandingPerformanceFormStatus = 1

    End Sub

    Public Sub EShowData()
        EResultLabel1.Visible = True
        EResultLabel2.Visible = True
        EResultValue1.Visible = True
        EResultValue2.Visible = True
        EResultValue1.Text = CInt(LandingDistance).ToString + "  Feet"
        EResultValue2.Text = CInt(Vref).ToString + "  Knots"
    End Sub

    Private Sub EButton1_Click(sender As Object, e As EventArgs) Handles EButton1.Click
        If ECheckData() Then
            If CurrentToolbar = 0 Then ChangeToolbar(1)
            ESaveData()
            ECalculateData()
            EShowData()
            ResetLabelColors()
        End If
    End Sub

    Private Sub ERadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles ERadioButton2.CheckedChanged
        If ERadioButton2.Checked Then
            ETextBox5.Enabled = True
        Else
            ETextBox5.Enabled = False
        End If
    End Sub

    Private Sub ETextBox_KeyDown(sender As Object, e As KeyEventArgs) Handles ETextBox1.KeyDown, ETextBox2.KeyDown, ETextBox3.KeyDown, ETextBox4.KeyDown, ETextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            EButton1.PerformClick()
        End If
    End Sub

#End Region
#Region " Landing Equations "

    Function EAltitude0(ByVal x As Double) As Double
        EAltitude0 = (4.76190476190476 * x) + 1328.57142857143
    End Function

    Function EAltitude2000(ByVal x As Double) As Double
        EAltitude2000 = (4.76190476190476 * x) + 1428.57142857143
    End Function

    Function EAltitude4000(ByVal x As Double) As Double
        EAltitude4000 = (5.26315789473684 * x) + 1531.57894736842
    End Function

    Function EAltitude6000(ByVal x As Double) As Double
        EAltitude6000 = (5.88235294117647 * x) + 1647.05882352941
    End Function

    Function EAltitude8000(ByVal x As Double) As Double
        EAltitude8000 = (6.06060606060606 * x) + 1769.69696969697
    End Function

    Function EAltitude10000(ByVal x As Double) As Double
        EAltitude10000 = (6.81818181818182 * x) + 1904.54545454545
    End Function

    Function EWeight1040(ByVal x As Double) As Double
        EWeight1040 = (-1.82095012549839E-19 * x ^ 6) + (0.00000000000000702427636827506 * x ^ 5) + (-0.000000000104824386825171 * x ^ 4) + (0.00000073432848049193 * x ^ 3) + (-0.00214730312058489 * x ^ 2) + (0 * x) + (9352.71368734908)
    End Function

    Function EWeight1340(ByVal x As Double) As Double
        EWeight1340 = (-7.14205371859896E-19 * x ^ 6) + (0.0000000000000290702548737743 * x ^ 5) + (-0.000000000459694866376331 * x ^ 4) + (0.00000342826537345696 * x ^ 3) + (-0.0107322347962743 * x ^ 2) + (0 * x) + (50423.6992591721)
    End Function

    Function EWeight1650(ByVal x As Double) As Double
        EWeight1650 = (-4.8356852457647E-20 * x ^ 6) + (0.00000000000000163727586413821 * x ^ 5) + (-0.0000000000202791121499604 * x ^ 4) + (0.000000107782243001699 * x ^ 3) + (-0.000199259390995171 * x ^ 2) + (0 * x) + (1205.92817588475)
    End Function

    Function EWeight1970(ByVal x As Double) As Double
        EWeight1970 = (-5.26915939347567E-19 * x ^ 6) + (0.0000000000000213879243035322 * x ^ 5) + (-0.000000000336901172032484 * x ^ 4) + (0.00000250072292250972 * x ^ 3) + (-0.00778733665556164 * x ^ 2) + (0 * x) + (37105.5600213103)
    End Function

    Function EWeight2280(ByVal x As Double) As Double
        EWeight2280 = (-4.32266879649579E-19 * x ^ 6) + (0.000000000000017063986313192 * x ^ 5) + (-0.000000000260462868711313 * x ^ 4) + (0.00000186607569949921 * x ^ 3) + (-0.00558429433773547 * x ^ 2) + (0 * x) + (25119.1809444609)
    End Function

    Function EHWind1000(ByVal x As Double) As Double
        EHWind1000 = (-11.3333333333333 * x) + 1000
    End Function

    Function EHWind1300(ByVal x As Double) As Double
        EHWind1300 = (-13.3333333333333 * x) + 1300
    End Function

    Function EHWind1600(ByVal x As Double) As Double
        EHWind1600 = (-15 * x) + 1600
    End Function

    Function EHWind1900(ByVal x As Double) As Double
        EHWind1900 = (-16.6666666666667 * x) + 1900
    End Function

    Function EHWind2200(ByVal x As Double) As Double
        EHWind2200 = (-17.3333333333333 * x) + 2200
    End Function

    Function ETWind1000(ByVal x As Double) As Double
        ETWind1000 = (42 * x) + 1000
    End Function

    Function ETWind1300(ByVal x As Double) As Double
        ETWind1300 = (50 * x) + 1300
    End Function

    Function ETWind1600(ByVal x As Double) As Double
        ETWind1600 = (54 * x) + 1600
    End Function

    Function ETWind1900(ByVal x As Double) As Double
        ETWind1900 = (58 * x) + 1900
    End Function

    Function ETWind2200(ByVal x As Double) As Double
        ETWind2200 = (61 * x) + 2200
    End Function

#End Region
#Region " Climb "

    Public Sub CLoadData()
        CTextBox1.Focus()
        If ClimbPerformanceFormStatus = 1 Then
            CTextBox1.Text = TakeoffAltitude
            CTextBox2.Text = TakeoffTemperature
            CTextBox3.Text = TakeoffWeight
            CTextBox5.Text = ClimbTemperature12
            CTextBox6.Text = ClimbTemperature18
            CTextBox7.Text = ClimbTemperature24
            CShowData()
        Else
            If TakeoffAltitude <> 999999 Then CTextBox1.Text = TakeoffAltitude
            If TakeoffTemperature <> 999999 Then CTextBox2.Text = TakeoffTemperature
            If TakeoffWeight <> 999999 Then CTextBox3.Text = TakeoffWeight
            If ClimbTemperature12 <> 999999 Then CTextBox5.Text = ClimbTemperature12
            If ClimbTemperature18 <> 999999 Then CTextBox6.Text = ClimbTemperature18
            If ClimbTemperature24 <> 999999 Then CTextBox7.Text = ClimbTemperature24
        End If
        If UnitTakeoffAltitude = 2 Then
            CRadioButton2.Checked = True
            CTextBox4.Text = TakeoffAltimeterSetting
        Else
            CRadioButton1.Checked = True
            CTextBox4.Text = ""
        End If
        If UnitTakeoffTemperature = 2 Then CRadioButton4.Checked = True Else CRadioButton3.Checked = True
        If UnitTakeoffWeight = 2 Then CRadioButton6.Checked = True Else CRadioButton5.Checked = True
    End Sub

    Public Function CCheckData() As Boolean
        Dim returnValue As Boolean = True
        If CTextBox1.Text = "" Then CTextBox1.Text = "0"
        If CTextBox2.Text = "" Then CTextBox2.Text = "15"
        If CTextBox3.Text = "" Then CTextBox3.Text = "9500"
        If CTextBox5.Text = "" Then CTextBox5.Text = "-9"
        If CTextBox6.Text = "" Then CTextBox6.Text = "-21"
        If CTextBox7.Text = "" Then CTextBox7.Text = "-33"
        If (Not IsNumeric(CTextBox1.Text)) Or (Not IsNumeric(CTextBox2.Text)) Or (Not IsNumeric(CTextBox3.Text)) Or (Not IsNumeric(CTextBox5.Text)) Or (Not IsNumeric(CTextBox6.Text)) Or (Not IsNumeric(CTextBox7.Text)) Then
            MsgBox("Invalid Input!")
            returnValue = False
        End If
        If CRadioButton2.Checked Then
            If CTextBox4.Text = "" Then CTextBox4.Text = "29.92"
            If (Not IsNumeric(CTextBox4.Text)) Then
                MsgBox("Invalid Input!")
                returnValue = False
            End If
            If ConvertAltitude(CDbl(CTextBox1.Text), CDbl(CTextBox4.Text)) > 10000 Then
                MsgBox("Pressure Altitude is too high")
                returnValue = False
            ElseIf ConvertAltitude(CDbl(CTextBox1.Text), CDbl(CTextBox4.Text)) < 0 Then
                MsgBox("Pressure Altitude is too low")
                returnValue = False
            End If
        Else
            If CDbl(CTextBox1.Text) > 10000 Then
                MsgBox("Pressure Altitude is too high")
                returnValue = False
            ElseIf CDbl(CTextBox1.Text) < 0 Then
                MsgBox("Pressure Altitude is too low")
                returnValue = False
            End If
        End If
        If CRadioButton3.Checked Then
            If CDbl(CTextBox2.Text) > 50 Then
                MsgBox("Temperature is too high")
                returnValue = False
            ElseIf CDbl(CTextBox2.Text) < -40 Then
                MsgBox("Temperature is too low")
                returnValue = False
            End If
        Else
            If ConvertTemperature(CDbl(CTextBox2.Text)) > 50 Then
                MsgBox("Temperature is too high")
                returnValue = False
            ElseIf ConvertTemperature(CDbl(CTextBox2.Text)) < -40 Then
                MsgBox("Temperature is too low")
                returnValue = False
            End If
        End If
        If CRadioButton5.Checked Then
            If CDbl(CTextBox3.Text) < 7000 Then
                MsgBox("Weight is too low")
                returnValue = False
            ElseIf CDbl(CTextBox3.Text) > 10100 Then
                MsgBox("Weight is too high")
                returnValue = False
            End If
        Else
            If ConvertWeight(CDbl(CTextBox3.Text)) < 7000 Then
                MsgBox("Weight is too low")
                returnValue = False
            ElseIf ConvertWeight(CDbl(CTextBox3.Text)) > 10100 Then
                MsgBox("Weight is too high")
                returnValue = False
            End If
        End If
        Return returnValue
    End Function

    Public Sub CSaveData()
        TakeoffAltitude = CDbl(CTextBox1.Text)
        TakeoffTemperature = CDbl(CTextBox2.Text)
        TakeoffWeight = CDbl(CTextBox3.Text)
        ClimbTemperature12 = CDbl(CTextBox5.Text)
        ClimbTemperature18 = CDbl(CTextBox6.Text)
        ClimbTemperature24 = CDbl(CTextBox7.Text)
        If CRadioButton2.Checked Then
            UnitTakeoffAltitude = 2
            TakeoffAltimeterSetting = CDbl(CTextBox4.Text)
            TakeoffPressureAltitude = ConvertAltitude(TakeoffAltitude, TakeoffAltimeterSetting)
        Else
            TakeoffPressureAltitude = TakeoffAltitude
            UnitTakeoffAltitude = 1
        End If
        If CRadioButton3.Checked Then UnitTakeoffTemperature = 1
        If CRadioButton4.Checked Then UnitTakeoffTemperature = 2
        If CRadioButton5.Checked Then UnitTakeoffWeight = 1
        If CRadioButton6.Checked Then UnitTakeoffWeight = 2
    End Sub

    Public Sub CCalculateData()
        Dim Cy1, Cy2 As Double
        Dim altitude As Double = TakeoffPressureAltitude
        Dim temperature As Double = TakeoffTemperature
        Dim weight As Double = TakeoffWeight
        If altitude <= 2000 Then
            Cy1 = CAltitude0(temperature) - (altitude / 2000) * (-CAltitude2000(temperature) + CAltitude0(temperature))
        ElseIf altitude <= 4000 Then
            Cy1 = CAltitude2000(temperature) - ((altitude - 2000) / 2000) * (-CAltitude4000(temperature) + CAltitude2000(temperature))
        ElseIf altitude <= 6000 Then
            Cy1 = CAltitude4000(temperature) - ((altitude - 4000) / 2000) * (-CAltitude6000(temperature) + CAltitude4000(temperature))
        ElseIf altitude <= 8000 Then
            Cy1 = CAltitude6000(temperature) - ((altitude - 6000) / 2000) * (-CAltitude8000(temperature) + CAltitude6000(temperature))
        ElseIf altitude <= 10000 Then
            Cy1 = CAltitude8000(temperature) - ((altitude - 8000) / 2000) * (-CAltitude10000(temperature) + CAltitude8000(temperature))
        ElseIf altitude <= 12000 Then
            Cy1 = CAltitude10000(temperature) - ((altitude - 10000) / 2000) * (-CAltitude12000(temperature) + CAltitude10000(temperature))
        ElseIf altitude <= 14000 Then
            Cy1 = CAltitude12000(temperature) - ((altitude - 12000) / 2000) * (-CAltitude14000(temperature) + CAltitude12000(temperature))
        ElseIf altitude <= 16000 Then
            Cy1 = CAltitude14000(temperature) - ((altitude - 14000) / 2000) * (-CAltitude16000(temperature) + CAltitude14000(temperature))
        ElseIf altitude <= 18000 Then
            Cy1 = CAltitude16000(temperature) - ((altitude - 16000) / 2000) * (-CAltitude18000(temperature) + CAltitude16000(temperature))
        ElseIf altitude <= 20000 Then
            Cy1 = CAltitude18000(temperature) - ((altitude - 18000) / 2000) * (-CAltitude20000(temperature) + CAltitude18000(temperature))
        ElseIf altitude <= 22000 Then
            Cy1 = CAltitude20000(temperature) - ((altitude - 20000) / 2000) * (-CAltitude22000(temperature) + CAltitude20000(temperature))
        ElseIf altitude <= 24000 Then
            Cy1 = CAltitude22000(temperature) - ((altitude - 22000) / 2000) * (-CAltitude24000(temperature) + CAltitude22000(temperature))
        ElseIf altitude <= 26000 Then
            Cy1 = CAltitude24000(temperature) - ((altitude - 24000) / 2000) * (-CAltitude26000(temperature) + CAltitude24000(temperature))
        ElseIf altitude <= 28000 Then
            Cy1 = CAltitude26000(temperature) - ((altitude - 26000) / 2000) * (-CAltitude28000(temperature) + CAltitude26000(temperature))
        End If

        If Cy1 <= -200 Then
            Cy2 = CWeightn400(weight) + ((Cy1 + 400) / 200) * (CWeightn200(weight) - CWeightn400(weight))
        ElseIf Cy1 <= 0 Then
            Cy2 = CWeightn200(weight) + ((Cy1 + 200) / 200) * (CWeight0(weight) - CWeightn200(weight))
        ElseIf Cy1 <= 200 Then
            Cy2 = CWeight0(weight) + ((Cy1 - 0) / 200) * (CWeight200(weight) - CWeight0(weight))
        ElseIf Cy1 <= 400 Then
            Cy2 = CWeight200(weight) + ((Cy1 - 200) / 200) * (CWeight400(weight) - CWeight200(weight))
        ElseIf Cy1 <= 600 Then
            Cy2 = CWeight400(weight) + ((Cy1 - 400) / 200) * (CWeight600(weight) - CWeight400(weight))
        End If

        ClimbRate = Cy2
        ClimbServiceCeiling = CServiceCeiling(ClimbTemperature12, ClimbTemperature18, ClimbTemperature24, TakeoffWeight)
        ClimbPerformanceFormStatus = 1
    End Sub

    Public Sub CShowData()
        CResultLabel1.Visible = True
        CResultLabel2.Visible = True
        CResultValue1.Visible = True
        CResultValue2.Visible = True
        CResultValue1.Text = CInt(ClimbRate).ToString + "  Ft/min"
        CResultValue2.Text = CInt(ClimbServiceCeiling).ToString + "  Feet"
    End Sub

    Private Sub CButton1_Click(sender As Object, e As EventArgs) Handles CButton1.Click
        If CCheckData() Then
            If CurrentToolbar = 0 Then ChangeToolbar(1)
            CSaveData()
            CCalculateData()
            CShowData()
            ResetLabelColors()
        End If
    End Sub

    Private Sub CRadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles CRadioButton2.CheckedChanged
        If CRadioButton2.Checked Then
            CTextBox4.Enabled = True
        Else
            CTextBox4.Enabled = False
        End If
    End Sub

    Private Sub CTextBox_KeyDown(sender As Object, e As KeyEventArgs) Handles CTextBox1.KeyDown, CTextBox2.KeyDown, CTextBox3.KeyDown, CTextBox4.KeyDown, CTextBox5.KeyDown, CTextBox6.KeyDown, CTextBox7.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            CButton1.PerformClick()
        End If
    End Sub

#End Region
#Region " Climb Equations "

    Function CAltitude0(ByVal x As Double) As Double
        CAltitude0 = -0.8065 * x + 476.61
    End Function

    Function CAltitude2000(ByVal x As Double) As Double
        CAltitude2000 = -0.0004 * x ^ 3 + 0.0107 * x ^ 2 - 1.3145 * x + 463.79
    End Function

    Function CAltitude4000(ByVal x As Double) As Double
        CAltitude4000 = -0.0026 * x ^ 3 + 0.0723 * x ^ 2 - 1.2009 * x + 428.69
    End Function

    Function CAltitude6000(ByVal x As Double) As Double
        CAltitude6000 = -0.00004 * x ^ 4 - 0.0023 * x ^ 3 + 0.0387 * x ^ 2 - 0.8859 * x + 400.1
    End Function

    Function CAltitude8000(ByVal x As Double) As Double
        CAltitude8000 = -0.0002 * x ^ 4 - 0.0008 * x ^ 3 + 0.0392 * x ^ 2 - 1.8088 * x + 367.65
    End Function

    Function CAltitude10000(ByVal x As Double) As Double
        CAltitude10000 = -0.0024 * x ^ 3 - 0.1067 * x ^ 2 - 3.0651 * x + 328.66
    End Function

    Function CAltitude12000(ByVal x As Double) As Double
        CAltitude12000 = 0.000003 * x ^ 5 + 0.0001 * x ^ 4 - 0.0037 * x ^ 3 - 0.2256 * x ^ 2 - 4.7814 * x + 279.78
    End Function

    Function CAltitude14000(ByVal x As Double) As Double
        CAltitude14000 = -0.0000002 * x ^ 6 - 0.00001 * x ^ 5 - 0.00002 * x ^ 4 + 0.0077 * x ^ 3 - 0.075 * x ^ 2 - 8.757 * x + 179.44
    End Function

    Function CAltitude16000(ByVal x As Double) As Double
        CAltitude16000 = -0.0000004 * x ^ 5 + 0.000008 * x ^ 4 + 0.0026 * x ^ 3 - 0.0248 * x ^ 2 - 9.0226 * x + 65.402
    End Function

    Function CAltitude18000(ByVal x As Double) As Double
        CAltitude18000 = -0.0000003 * x ^ 5 - 0.00007 * x ^ 4 - 0.003 * x ^ 3 - 0.0383 * x ^ 2 - 7.869 * x - 54.753
    End Function

    Function CAltitude20000(ByVal x As Double) As Double
        CAltitude20000 = -0.00002 * x ^ 4 - 0.0012 * x ^ 3 - 0.0307 * x ^ 2 - 7.9967 * x - 175.29
    End Function

    Function CAltitude22000(ByVal x As Double) As Double
        CAltitude22000 = -0.00000003 * x ^ 6 - 0.000006 * x ^ 5 - 0.0005 * x ^ 4 - 0.0203 * x ^ 3 - 0.3592 * x ^ 2 - 9.8745 * x - 300.01
    End Function

    Function CAltitude24000(ByVal x As Double) As Double
        CAltitude24000 = -0.00002 * x ^ 4 - 0.0005 * x ^ 3 + 0.0151 * x ^ 2 - 6.8539 * x - 420.88
    End Function

    Function CAltitude26000(ByVal x As Double) As Double
        CAltitude26000 = -0.0932 * x ^ 2 - 12.565 * x - 617.15
    End Function

    Function CAltitude28000(ByVal x As Double) As Double
        CAltitude28000 = -0.0348 * x ^ 2 - 6.6551 * x - 596.26
    End Function

    Function CWeightn400(ByVal x As Double) As Double
        CWeightn400 = (1.98704092307177E-19 * x ^ 6) + (-0.00000000000000778295454393822 * x ^ 5) + (0.000000000117694309857271 * x ^ 4) + (-0.000000831923034602321 * x ^ 3) + (0.00243201332598426 * x ^ 2) + (0 * x) + (-9022.30488206638)
    End Function

    Function CWeightn200(ByVal x As Double) As Double
        CWeightn200 = (2.59627243902836E-19 * x ^ 6) + (-0.0000000000000107240794281611 * x ^ 5) + (0.000000000172270669261917 * x ^ 4) + (-0.000001304611610239 * x ^ 3) + (0.00413213641088455 * x ^ 2) + (0 * x) + (-18720.1205469247)
    End Function

    Function CWeight0(ByVal x As Double) As Double
        CWeight0 = (7.62613513315362E-20 * x ^ 6) + (-0.00000000000000329867967112623 * x ^ 5) + (0.0000000000553756061070876 * x ^ 4) + (-0.000000436507088824029 * x ^ 3) + (0.00142568339173214 * x ^ 2) + (0 * x) + (-6199.43622326048)
    End Function

    Function CWeight200(ByVal x As Double) As Double
        CWeight200 = (3.77151067792658E-19 * x ^ 6) + (-0.0000000000000159001047019964 * x ^ 5) + (0.000000000260405176522944 * x ^ 4) + (-0.00000200865564196399 * x ^ 3) + (0.00647819196491849 * x ^ 2) + (0 * x) + (-30143.5955517426)
    End Function

    Function CWeight400(ByVal x As Double) As Double
        CWeight400 = (-5.87699210524353E-19 * x ^ 6) + (0.0000000000000236394454707762 * x ^ 5) + (-0.000000000369577592371145 * x ^ 4) + (0.00000272737055745436 * x ^ 3) + (-0.00847171089862244 * x ^ 2) + (0 * x) + (39751.0481955255)
    End Function

    Function CWeight600(ByVal x As Double) As Double
        CWeight600 = (-2.84481688570485E-19 * x ^ 6) + (0.0000000000000113017603235269 * x ^ 5) + (-0.000000000174525330348439 * x ^ 4) + (0.00000127356646991142 * x ^ 3) + (-0.00392680538131984 * x ^ 2) + (0 * x) + (19290.9000441687)
    End Function

    Function SCWeight7000(ByVal x As Double) As Double
        SCWeight7000 = (0.00000866666666666696 * x ^ 6) + (0.0013666666666667 * x ^ 5) + (0.0803333333333346 * x ^ 4) + (2.18333333333335 * x ^ 3) + (27.0800000000001 * x ^ 2) + (0 * x) + (22400)
    End Function

    Function SCWeight7500(ByVal x As Double) As Double
        SCWeight7500 = (0.00000987500000000026 * x ^ 6) + (0.0014979166666667 * x ^ 5) + (0.085187500000001 * x ^ 4) + (2.25520833333334 * x ^ 3) + (27.4325 * x ^ 2) + (0 * x) + (21500)
    End Function

    Function SCWeight8000(ByVal x As Double) As Double
        SCWeight8000 = (0.000000277777777777781 * x ^ 6) + (0.0000250000000000005 * x ^ 5) + (-0.00013888888888887 * x ^ 4) + (-0.0458333333333331 * x ^ 3) + (-0.988888888888889 * x ^ 2) + (-125.666666666667 * x) + (20600)
    End Function

    Function SCWeight8500(ByVal x As Double) As Double
        SCWeight8500 = (-0.00000055555555555551 * x ^ 6) + (-0.000066666666666661 * x ^ 5) + (-0.00347222222222199 * x ^ 4) + (-0.091666666666664 * x ^ 3) + (-1.14722222222224 * x ^ 2) + (-125.166666666667 * x) + (19600)
    End Function

    Function SCWeight9000(ByVal x As Double) As Double
        SCWeight9000 = (-0.000000555555555555583 * x ^ 6) + (-0.0000416666666666698 * x ^ 5) + (-0.000972222222222331 * x ^ 4) + (-0.0208333333333343 * x ^ 3) + (-0.897222222222211 * x ^ 2) + (-137.5 * x) + (18500)
    End Function

    Function SCWeight9500(ByVal x As Double) As Double
        SCWeight9500 = (-0.000000555555555555583 * x ^ 6) + (-0.0000416666666666698 * x ^ 5) + (-0.000972222222222331 * x ^ 4) + (-0.0208333333333343 * x ^ 3) + (-0.897222222222211 * x ^ 2) + (-137.5 * x) + (17500)
    End Function

    Function SCWeight10100(ByVal x As Double) As Double
        SCWeight10100 = (-0.000000416666666666681 * x ^ 6) + (-0.0000201923076923089 * x ^ 5) + (0.0000480769230769097 * x ^ 4) + (-0.00569638694638627 * x ^ 3) + (-0.951689976689971 * x ^ 2) + (-139.391608391608 * x) + (16297.5524475524)
    End Function

    Function TemperatureLine(ByVal x As Double, ByVal m As Double, ByVal b As Double)
        TemperatureLine = (m * x) + b
    End Function

    Function CServiceCeiling(ByVal x12 As Double, ByVal x18 As Double, ByVal x24 As Double, ByVal weight As Double)
        Dim slope As Double
        Dim b As Double
        Dim z As Double
        Dim i As Double = 10

        slope = 6000 / (x24 - x18)
        b = 24000 - (slope * x24)

        If weight <= 7500 Then
            z = SCWeight7000(i) - ((weight - 7000) / 500) * (SCWeight7000(i) - SCWeight7500(i))
            While z > TemperatureLine(i, slope, b)
                i -= 0.1
                z = SCWeight7000(i) - ((weight - 7000) / 500) * (SCWeight7000(i) - SCWeight7500(i))
            End While
            i += 0.1
            z = SCWeight7000(i) - ((weight - 7000) / 500) * (SCWeight7000(i) - SCWeight7500(i))
        ElseIf weight <= 8000 Then
            z = SCWeight7500(i) - ((weight - 7500) / 500) * (SCWeight7500(i) - SCWeight8000(i))
            While z > TemperatureLine(i, slope, b)
                i -= 0.1
                z = SCWeight7500(i) - ((weight - 7500) / 500) * (SCWeight7500(i) - SCWeight8000(i))
            End While
            i += 0.1
            z = SCWeight7500(i) - ((weight - 7500) / 500) * (SCWeight7500(i) - SCWeight8000(i))
        ElseIf weight <= 8500 Then
            z = SCWeight8000(i) - ((weight - 8000) / 500) * (SCWeight8000(i) - SCWeight8500(i))
            While z > TemperatureLine(i, slope, b)
                i -= 0.1
                z = SCWeight8000(i) - ((weight - 8000) / 500) * (SCWeight8000(i) - SCWeight8500(i))
            End While
            i += 0.1
            z = SCWeight8000(i) - ((weight - 8000) / 500) * (SCWeight8000(i) - SCWeight8500(i))
        ElseIf weight <= 9000 Then
            z = SCWeight8500(i) - ((weight - 8500) / 500) * (SCWeight8500(i) - SCWeight9000(i))
            While z > TemperatureLine(i, slope, b)
                i -= 0.1
                z = SCWeight8500(i) - ((weight - 8500) / 500) * (SCWeight8500(i) - SCWeight9000(i))
            End While
            i += 0.1
            z = SCWeight8500(i) - ((weight - 8500) / 500) * (SCWeight8500(i) - SCWeight9000(i))
        ElseIf weight <= 9500 Then
            z = SCWeight9000(i) - ((weight - 9000) / 500) * (SCWeight9000(i) - SCWeight9500(i))
            While z > TemperatureLine(i, slope, b)
                i -= 0.1
                z = SCWeight9000(i) - ((weight - 9000) / 500) * (SCWeight9000(i) - SCWeight9500(i))
            End While
            i += 0.1
            z = SCWeight9000(i) - ((weight - 9000) / 500) * (SCWeight9000(i) - SCWeight9500(i))
        ElseIf weight <= 10100 Then
            z = SCWeight9500(i) - ((weight - 9500) / 600) * (SCWeight9500(i) - SCWeight10100(i))
            While z > TemperatureLine(i, slope, b)
                i -= 0.1
                z = SCWeight9500(i) - ((weight - 9500) / 600) * (SCWeight9500(i) - SCWeight10100(i))
            End While
            i += 0.1
            z = SCWeight9500(i) - ((weight - 9500) / 600) * (SCWeight9500(i) - SCWeight10100(i))
        End If
        Return z
    End Function
#End Region
#Region " Cruise "

    Public Sub DLoadData()
        DTextBox1.Focus()
        If CruisePerformanceFormStatus = 1 Then
            DTextBox1.Text = CruiseAltitude
            DTextBox2.Text = CruiseTemperature
            DTextBox3.Text = TakeoffWeight
            DShowData()
        Else
            If CruiseAltitude <> 999999 Then DTextBox1.Text = CruiseAltitude
            If CruiseTemperature <> 999999 Then DTextBox2.Text = CruiseTemperature
            If TakeoffWeight <> 999999 Then DTextBox3.Text = TakeoffWeight
        End If
        If CruiseMode = 2 Then DRadioButton2.Checked = True Else DRadioButton1.Checked = True
    End Sub

    Public Function DCheckData() As Boolean
        Dim returnValue As Boolean = True
        If DTextBox1.Text = "" Then DTextBox1.Text = "10000"
        If DTextBox2.Text = "" Then DTextBox2.Text = "0"
        If DTextBox3.Text = "" Then DTextBox3.Text = "9500"
        If (Not IsNumeric(DTextBox1.Text)) Or (Not IsNumeric(DTextBox2.Text)) Or (Not IsNumeric(DTextBox3.Text)) Then
            MsgBox("Invalid Input!")
            returnValue = False
        End If
        If CDbl(DTextBox1.Text) > 29000 Then
            MsgBox("Altitude is too high")
            returnValue = False
        ElseIf CDbl(DTextBox1.Text) < 0 Then
            MsgBox("Altitude is too low")
            returnValue = False
        End If
        If CDbl(DTextBox2.Text) > 33 Then
            MsgBox("Temperature is too high")
            returnValue = False
        ElseIf CDbl(DTextBox2.Text) < -30 Then
            MsgBox("Temperature is too low")
            returnValue = False
        End If
        If CDbl(DTextBox3.Text) > 10100 Then
            MsgBox("Weight is too high")
            returnValue = False
        ElseIf CDbl(DTextBox3.Text) < 7000 Then
            MsgBox("Weight is too low")
            returnValue = False
        End If
        Return returnValue
    End Function

    Public Sub DSaveData()
        CruiseAltitude = CDbl(DTextBox1.Text)
        CruiseTemperature = CDbl(DTextBox2.Text)
        TakeoffWeight = CDbl(DTextBox3.Text)
        If DRadioButton1.Checked Then CruiseMode = 1
        If DRadioButton2.Checked Then CruiseMode = 2
    End Sub

    Public Sub DCalculateData()
        Dim altitude As Double = CruiseAltitude
        Dim temperature As Double = CruiseTemperature
        Dim weight As Double = TakeoffWeight
        Dim torque, fuelFlow, IAS, TAS As Double
        Dim IASlower, IAShigher, TASlower, TAShigher As Double
        If DRadioButton1.Checked Then
            If temperature <= -20 Then
                torque = CTorquen20(altitude)
                fuelFlow = CFuelFlown20(altitude)
                If weight >= 9500 Then
                    IAS = CIAS9500n20(altitude)
                    TAS = CTAS9500n20(altitude)
                ElseIf weight >= 8500 Then
                    IAS = CIAS8500n20(altitude) + ((weight - 8500) / 1000) * (CIAS9500n20(altitude) - CIAS8500n20(altitude))
                    TAS = CTAS8500n20(altitude) + ((weight - 8500) / 1000) * (CTAS9500n20(altitude) - CTAS8500n20(altitude))
                ElseIf weight >= 7500 Then
                    IAS = CIAS7500n20(altitude) + ((weight - 7500) / 1000) * (CIAS8500n20(altitude) - CIAS7500n20(altitude))
                    TAS = CTAS7500n20(altitude) + ((weight - 7500) / 1000) * (CTAS8500n20(altitude) - CTAS7500n20(altitude))
                Else
                    IAS = CIAS7500n20(altitude)
                    TAS = CTAS7500n20(altitude)
                End If
            ElseIf temperature <= -10 Then
                torque = CTorquen20(altitude) + ((temperature + 20) / 10) * (CTorquen10(altitude) - CTorquen20(altitude))
                fuelFlow = CFuelFlown20(altitude) + ((temperature + 20) / 10) * (CFuelFlown10(altitude) - CFuelFlown20(altitude))
                If weight >= 9500 Then
                    IAS = CIAS9500n20(altitude) + ((temperature + 20) / 10) * (CIAS9500n10(altitude) - CIAS9500n20(altitude))
                    TAS = CTAS9500n20(altitude) + ((temperature + 20) / 10) * (CTAS9500n10(altitude) - CTAS9500n20(altitude))
                ElseIf weight >= 8500 Then
                    IASlower = CIAS8500n20(altitude) + ((temperature + 20) / 10) * (CIAS8500n10(altitude) - CIAS8500n20(altitude))
                    TASlower = CTAS8500n20(altitude) + ((temperature + 20) / 10) * (CTAS8500n10(altitude) - CTAS8500n20(altitude))
                    IAShigher = CIAS9500n20(altitude) + ((temperature + 20) / 10) * (CIAS9500n10(altitude) - CIAS9500n20(altitude))
                    TAShigher = CTAS9500n20(altitude) + ((temperature + 20) / 10) * (CTAS9500n10(altitude) - CTAS9500n20(altitude))
                    IAS = IASlower + ((weight - 8500) / 1000) * (IAShigher - IASlower)
                    TAS = TASlower + ((weight - 8500) / 1000) * (TAShigher - TASlower)
                ElseIf weight >= 7500 Then
                    IAShigher = CIAS8500n20(altitude) + ((temperature + 20) / 10) * (CIAS8500n10(altitude) - CIAS8500n20(altitude))
                    TAShigher = CTAS8500n20(altitude) + ((temperature + 20) / 10) * (CTAS8500n10(altitude) - CTAS8500n20(altitude))
                    IASlower = CIAS7500n20(altitude) + ((temperature + 20) / 10) * (CIAS7500n10(altitude) - CIAS7500n20(altitude))
                    TASlower = CTAS7500n20(altitude) + ((temperature + 20) / 10) * (CTAS7500n10(altitude) - CTAS7500n20(altitude))
                    IAS = IASlower + ((weight - 7500) / 1000) * (IAShigher - IASlower)
                    TAS = TASlower + ((weight - 7500) / 1000) * (TAShigher - TASlower)
                Else
                    IAS = CIAS7500n20(altitude) + ((temperature + 20) / 10) * (CIAS7500n10(altitude) - CIAS7500n20(altitude))
                    TAS = CTAS7500n20(altitude) + ((temperature + 20) / 10) * (CTAS7500n10(altitude) - CTAS7500n20(altitude))
                End If
            ElseIf temperature <= 0 Then
                torque = CTorquen10(altitude) + ((temperature + 10) / 10) * (CTorque0(altitude) - CTorquen10(altitude))
                fuelFlow = CFuelFlown10(altitude) + ((temperature + 10) / 10) * (CFuelFlow0(altitude) - CFuelFlown10(altitude))
                If weight >= 9500 Then
                    IAS = CIAS9500n10(altitude) + ((temperature + 10) / 10) * (CIAS95000(altitude) - CIAS9500n10(altitude))
                    TAS = CTAS9500n10(altitude) + ((temperature + 10) / 10) * (CTAS95000(altitude) - CTAS9500n10(altitude))
                ElseIf weight >= 8500 Then
                    IASlower = CIAS8500n10(altitude) + ((temperature + 10) / 10) * (CIAS85000(altitude) - CIAS8500n10(altitude))
                    TASlower = CTAS8500n10(altitude) + ((temperature + 10) / 10) * (CTAS85000(altitude) - CTAS8500n10(altitude))
                    IAShigher = CIAS9500n10(altitude) + ((temperature + 10) / 10) * (CIAS95000(altitude) - CIAS9500n10(altitude))
                    TAShigher = CTAS9500n10(altitude) + ((temperature + 10) / 10) * (CTAS95000(altitude) - CTAS9500n10(altitude))
                    IAS = IASlower + ((weight - 8500) / 1000) * (IAShigher - IASlower)
                    TAS = TASlower + ((weight - 8500) / 1000) * (TAShigher - TASlower)
                ElseIf weight >= 7500 Then
                    IAShigher = CIAS8500n10(altitude) + ((temperature + 10) / 10) * (CIAS85000(altitude) - CIAS8500n10(altitude))
                    TAShigher = CTAS8500n10(altitude) + ((temperature + 10) / 10) * (CTAS85000(altitude) - CTAS8500n10(altitude))
                    IASlower = CIAS7500n10(altitude) + ((temperature + 10) / 10) * (CIAS75000(altitude) - CIAS7500n10(altitude))
                    TASlower = CTAS7500n10(altitude) + ((temperature + 10) / 10) * (CTAS75000(altitude) - CTAS7500n10(altitude))
                    IAS = IASlower + ((weight - 7500) / 1000) * (IAShigher - IASlower)
                    TAS = TASlower + ((weight - 7500) / 1000) * (TAShigher - TASlower)
                Else
                    IAS = CIAS7500n10(altitude) + ((temperature + 10) / 10) * (CIAS75000(altitude) - CIAS7500n10(altitude))
                    TAS = CTAS7500n10(altitude) + ((temperature + 10) / 10) * (CTAS75000(altitude) - CTAS7500n10(altitude))
                End If
            ElseIf temperature <= 10 Then
                torque = CTorque0(altitude) + ((temperature + 0) / 10) * (CTorque10(altitude) - CTorque0(altitude))
                fuelFlow = CFuelFlow0(altitude) + ((temperature + 0) / 10) * (CFuelFlow10(altitude) - CFuelFlow0(altitude))
                If weight >= 9500 Then
                    IAS = CIAS95000(altitude) + ((temperature + 0) / 10) * (CIAS950010(altitude) - CIAS95000(altitude))
                    TAS = CTAS95000(altitude) + ((temperature + 0) / 10) * (CTAS950010(altitude) - CTAS95000(altitude))
                ElseIf weight >= 8500 Then
                    IASlower = CIAS85000(altitude) + ((temperature + 0) / 10) * (CIAS850010(altitude) - CIAS85000(altitude))
                    TASlower = CTAS85000(altitude) + ((temperature + 0) / 10) * (CTAS850010(altitude) - CTAS85000(altitude))
                    IAShigher = CIAS95000(altitude) + ((temperature + 0) / 10) * (CIAS950010(altitude) - CIAS95000(altitude))
                    TAShigher = CTAS95000(altitude) + ((temperature + 0) / 10) * (CTAS950010(altitude) - CTAS95000(altitude))
                    IAS = IASlower + ((weight - 8500) / 1000) * (IAShigher - IASlower)
                    TAS = TASlower + ((weight - 8500) / 1000) * (TAShigher - TASlower)
                ElseIf weight >= 7500 Then
                    IAShigher = CIAS85000(altitude) + ((temperature + 0) / 10) * (CIAS850010(altitude) - CIAS85000(altitude))
                    TAShigher = CTAS85000(altitude) + ((temperature + 0) / 10) * (CTAS850010(altitude) - CTAS85000(altitude))
                    IASlower = CIAS75000(altitude) + ((temperature + 0) / 10) * (CIAS750010(altitude) - CIAS75000(altitude))
                    TASlower = CTAS75000(altitude) + ((temperature + 0) / 10) * (CTAS750010(altitude) - CTAS75000(altitude))
                    IAS = IASlower + ((weight - 7500) / 1000) * (IAShigher - IASlower)
                    TAS = TASlower + ((weight - 7500) / 1000) * (TAShigher - TASlower)
                Else
                    IAS = CIAS75000(altitude) + ((temperature + 0) / 10) * (CIAS750010(altitude) - CIAS75000(altitude))
                    TAS = CTAS75000(altitude) + ((temperature + 0) / 10) * (CTAS750010(altitude) - CTAS75000(altitude))
                End If
            ElseIf temperature <= 20 Then
                torque = CTorque10(altitude) + ((temperature - 10) / 10) * (CTorque20(altitude) - CTorque10(altitude))
                fuelFlow = CFuelFlow10(altitude) + ((temperature - 10) / 10) * (CFuelFlow20(altitude) - CFuelFlow10(altitude))
                If weight >= 9500 Then
                    IAS = CIAS950010(altitude) + ((temperature - 10) / 10) * (CIAS950020(altitude) - CIAS950010(altitude))
                    TAS = CTAS950010(altitude) + ((temperature - 10) / 10) * (CTAS950020(altitude) - CTAS950010(altitude))
                ElseIf weight >= 8500 Then
                    IASlower = CIAS850010(altitude) + ((temperature - 10) / 10) * (CIAS850020(altitude) - CIAS850010(altitude))
                    TASlower = CTAS850010(altitude) + ((temperature - 10) / 10) * (CTAS850020(altitude) - CTAS850010(altitude))
                    IAShigher = CIAS950010(altitude) + ((temperature - 10) / 10) * (CIAS950020(altitude) - CIAS950010(altitude))
                    TAShigher = CTAS950010(altitude) + ((temperature - 10) / 10) * (CTAS950020(altitude) - CTAS950010(altitude))
                    IAS = IASlower + ((weight - 8500) / 1000) * (IAShigher - IASlower)
                    TAS = TASlower + ((weight - 8500) / 1000) * (TAShigher - TASlower)
                ElseIf weight >= 7500 Then
                    IAShigher = CIAS850010(altitude) + ((temperature - 10) / 10) * (CIAS850020(altitude) - CIAS850010(altitude))
                    TAShigher = CTAS850010(altitude) + ((temperature - 10) / 10) * (CTAS850020(altitude) - CTAS850010(altitude))
                    IASlower = CIAS750010(altitude) + ((temperature - 10) / 10) * (CIAS750020(altitude) - CIAS750010(altitude))
                    TASlower = CTAS750010(altitude) + ((temperature - 10) / 10) * (CTAS750020(altitude) - CTAS750010(altitude))
                    IAS = IASlower + ((weight - 7500) / 1000) * (IAShigher - IASlower)
                    TAS = TASlower + ((weight - 7500) / 1000) * (TAShigher - TASlower)
                Else
                    IAS = CIAS750010(altitude) + ((temperature - 10) / 10) * (CIAS750020(altitude) - CIAS750010(altitude))
                    TAS = CTAS750010(altitude) + ((temperature - 10) / 10) * (CTAS750020(altitude) - CTAS750010(altitude))
                End If
            Else
                torque = CTorque20(altitude)
                fuelFlow = CFuelFlow20(altitude)
                If weight >= 9500 Then
                    IAS = CIAS950020(altitude)
                    TAS = CTAS950020(altitude)
                ElseIf weight >= 8500 Then
                    IAS = CIAS850020(altitude) + ((weight - 8500) / 1000) * (CIAS950020(altitude) - CIAS850020(altitude))
                    TAS = CTAS850020(altitude) + ((weight - 8500) / 1000) * (CTAS950020(altitude) - CTAS850020(altitude))
                ElseIf weight >= 7500 Then
                    IAS = CIAS750020(altitude) + ((weight - 7500) / 1000) * (CIAS850020(altitude) - CIAS750020(altitude))
                    TAS = CTAS750020(altitude) + ((weight - 7500) / 1000) * (CTAS850020(altitude) - CTAS750020(altitude))
                Else
                    IAS = CIAS750020(altitude)
                    TAS = CTAS750020(altitude)
                End If
            End If
        Else
            IAS = 888888
            If temperature <= -20 Then
                torque = RTorquen20(altitude)
                fuelFlow = RFuelFlown20(altitude)
                TAS = RTASn20(altitude)
            ElseIf temperature <= 10 Then
                torque = RTorquen20(altitude) + ((temperature + 20) / 10) * (RTorquen10(altitude) - RTorquen20(altitude))
                fuelFlow = RFuelFlown20(altitude) + ((temperature + 20) / 10) * (RFuelFlown10(altitude) - RFuelFlown20(altitude))
                TAS = RTASn20(altitude) + ((temperature + 20) / 10) * (RTASn10(altitude) - RTASn20(altitude))
            ElseIf temperature <= 0 Then
                torque = RTorquen10(altitude) + ((temperature + 10) / 10) * (RTorque0(altitude) - RTorquen10(altitude))
                fuelFlow = RFuelFlown10(altitude) + ((temperature + 10) / 10) * (RFuelFlow0(altitude) - RFuelFlown10(altitude))
                TAS = RTASn10(altitude) + ((temperature + 10) / 10) * (RTAS0(altitude) - RTASn10(altitude))
            ElseIf temperature <= 10 Then
                torque = RTorque0(altitude) + ((temperature + 0) / 10) * (RTorque10(altitude) - RTorque0(altitude))
                fuelFlow = RFuelFlow0(altitude) + ((temperature + 0) / 10) * (RFuelFlow10(altitude) - RFuelFlow0(altitude))
                TAS = RTAS0(altitude) + ((temperature + 0) / 10) * (RTAS10(altitude) - RTAS0(altitude))
            ElseIf temperature <= 20 Then
                torque = RTorque10(altitude) + ((temperature - 10) / 10) * (RTorque20(altitude) - RTorque10(altitude))
                fuelFlow = RFuelFlow10(altitude) + ((temperature - 10) / 10) * (RFuelFlow20(altitude) - RFuelFlow10(altitude))
                TAS = RTAS10(altitude) + ((temperature - 10) / 10) * (RTAS20(altitude) - RTAS10(altitude))
            Else
                torque = RTorque20(altitude)
                fuelFlow = RFuelFlow20(altitude)
                TAS = RTAS20(altitude)
            End If
        End If

        CruiseTorque = torque
        CruiseFuelFlow = fuelFlow
        CruiseIndicatedAirspeed = IAS
        CruiseTrueAirspeed = TAS
        CruisePerformanceFormStatus = 1
    End Sub

    Public Sub DShowData()
        DResultLabel1.Visible = True
        DResultLabel2.Visible = True
        DResultLabel3.Visible = True
        DResultLabel4.Visible = True
        DResultValue1.Visible = True
        DResultValue2.Visible = True
        DResultValue3.Visible = True
        DResultValue4.Visible = True
        DResultValue1.Text = CInt(CruiseTorque).ToString + "  Ft-Lbs"
        DResultValue2.Text = CInt(CruiseFuelFlow).ToString + "  Lb/hour"
        If CruiseIndicatedAirspeed = 888888 Then DResultValue3.Text = "NA"
        If CruiseIndicatedAirspeed <> 888888 Then DResultValue3.Text = CInt(CruiseIndicatedAirspeed).ToString + "  Knots"
        DResultValue4.Text = CInt(CruiseTrueAirspeed).ToString + "  Knots"
    End Sub

    Private Sub DButton1_Click(sender As Object, e As EventArgs) Handles DButton1.Click
        If DCheckData() Then
            If CurrentToolbar = 0 Then ChangeToolbar(1)
            DSaveData()
            DCalculateData()
            DShowData()
            ResetLabelColors()
        End If
    End Sub

    Private Sub DTextBox_KeyDown(sender As Object, e As KeyEventArgs) Handles DTextBox1.KeyDown, DTextBox2.KeyDown, DTextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            DButton1.PerformClick()
        End If
    End Sub

#End Region
#Region " Cruise Equations "

    Function CTorquen20(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {1444, 1477, 1513, 1520, 1520, 1520, 1520, 1520, 1520, 1520, 1520, 1456, 1336, 1217, 1102, 1045}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function CTorquen10(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {1458, 1495, 1520, 1520, 1520, 1520, 1520, 1520, 1520, 1520, 1520, 1434, 1328, 1222, 1113, 1057}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function CTorque0(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {1490, 1520, 1520, 1520, 1520, 1520, 1520, 1520, 1520, 1520, 1449, 1355, 1262, 1170, 1082, 1039}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function CTorque10(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {1515, 1520, 1520, 1520, 1520, 1520, 1520, 1520, 1520, 1443, 1358, 1272, 1186, 1102, 1022, 983}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function CTorque20(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {1520, 1520, 1520, 1520, 1520, 1520, 1520, 1477, 1415, 1348, 1269, 1189, 1109, 1030, 956, 919}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function RTorquen20(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {787, 773, 763, 755, 746, 736, 729, 722, 725, 733, 747, 758, 766, 772, 744, 725}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function RTorquen10(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {751, 740, 735, 731, 726, 715, 714, 715, 720, 737, 754, 758, 740, 711, 699, 708}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function RTorque0(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {1155, 1163, 1142, 1075, 996, 933, 872, 771, 745, 740, 735, 720, 719, 724, 734, 739}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function RTorque10(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {1008, 1013, 1023, 1015, 1000, 978, 956, 922, 875, 836, 824, 803, 778, 754, 737, 741}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function RTorque20(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {1306, 1185, 1074, 975, 930, 892, 860, 852, 846, 822, 796, 761, 741, 737, 759, 768}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function CFuelFlown20(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {363, 357, 351, 342, 333, 326, 321, 318, 317, 316, 317, 306, 283, 260, 237, 227}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function CFuelFlown10(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {366, 360, 353, 343, 335, 328, 323, 319, 318, 318, 319, 303, 283, 262, 241, 229}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function CFuelFlow0(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {375, 367, 355, 345, 336, 329, 325, 322, 321, 320, 306, 288, 270, 252, 235, 226}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function CFuelFlow10(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {381, 370, 358, 347, 339, 332, 327, 324, 322, 307, 290, 273, 256, 240, 223, 215}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function CFuelFlow20(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {385, 373, 361, 351, 342, 335, 329, 318, 304, 289, 273, 258, 242, 227, 212, 204}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function RFuelFlown20(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {261, 247, 234, 222, 211, 200, 192, 184, 180, 178, 176, 175, 174, 172, 166, 162}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function RFuelFlown10(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {270, 253, 239, 226, 215, 203, 194, 187, 181, 179, 177, 174, 169, 162, 159, 160}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function RFuelFlow0(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {328, 316, 301, 279, 258, 239, 221, 200, 189, 183, 177, 171, 168, 166, 166, 166}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function RFuelFlow10(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {313, 299, 287, 273, 260, 246, 234, 222, 209, 198, 191, 184, 178, 172, 168, 168}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function RFuelFlow20(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {355, 325, 297, 271, 253, 237, 223, 214, 207, 198, 189, 180, 174, 171, 173, 174}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function CIAS9500n20(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {226, 226, 226, 223, 221, 219, 217, 215, 212, 210, 207, 201, 192, 182, 172, 167}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function CIAS8500n20(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {226, 226, 226, 224, 222, 220, 218, 216, 213, 211, 209, 203, 194, 185, 175, 170}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function CIAS7500n20(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {226, 226, 226, 225, 223, 221, 219, 217, 214, 213, 210, 205, 196, 187, 178, 173}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function CTAS9500n20(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {217, 223, 230, 234, 238, 244, 248, 253, 258, 263, 268, 269, 266, 261, 256, 252}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function CTAS8500n20(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {217, 223, 230, 235, 240, 245, 250, 254, 259, 265, 270, 271, 268, 265, 260, 257}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function CTAS7500n20(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {217, 223, 230, 236, 241, 246, 251, 256, 261, 267, 272, 273, 271, 267, 263, 261}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function CIAS9500n10(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {226, 226, 224, 222, 220, 217, 215, 213, 210, 208, 205, 198, 190, 181, 171, 166}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function CTAS9500n10(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {221, 228, 233, 237, 242, 246, 251, 256, 261, 266, 272, 217, 268, 265, 260, 256}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function CIAS8500n10(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {226, 226, 225, 223, 221, 219, 216, 214, 212, 209, 207, 200, 192, 183, 174, 169}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function CTAS8500n10(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {221, 228, 234, 238, 243, 248, 252, 257, 263, 268, 273, 273, 217, 268, 264, 261}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function CIAS7500n10(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {226, 226, 226, 224, 222, 220, 217, 215, 213, 211, 208, 202, 194, 158, 176, 172}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function CTAS7500n10(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {221, 228, 235, 239, 244, 249, 254, 259, 264, 270, 275, 275, 274, 271, 268, 265}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function CIAS95000(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {226, 225, 222, 220, 218, 215, 214, 211, 209, 206, 200, 192, 184, 175, 167, 162}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function CTAS95000(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {225, 231, 235, 240, 244, 249, 254, 259, 264, 269, 270, 268, 266, 263, 259, 257}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function CIAS85000(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {226, 226, 224, 221, 219, 217, 215, 212, 210, 207, 201, 194, 186, 178, 170, 166}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function CTAS85000(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {225, 232, 236, 241, 264, 250, 256, 261, 266, 271, 272, 271, 269, 267, 264, 262}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function CIAS75000(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {226, 226, 225, 222, 220, 218, 216, 214, 211, 209, 203, 196, 188, 180, 173, 169}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function CTAS75000(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {225, 232, 237, 242, 247, 251, 257, 262, 268, 273, 274, 273, 272, 270, 268, 267}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function CIAS950010(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {225, 223, 221, 219, 217, 214, 212, 209, 207, 200, 193, 185, 177, 169, 161, 156}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function CTAS950010(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {229, 233, 237, 243, 247, 252, 257, 262, 267, 267, 266, 265, 262, 259, 256, 253}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function CIAS850010(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {226, 224, 222, 220, 218, 216, 213, 211, 208, 202, 195, 187, 180, 172, 164, 160}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function CTAS850010(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {229, 234, 239, 244, 249, 254, 258, 264, 269, 269, 269, 268, 266, 264, 261, 259}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function CIAS750010(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {226, 225, 223, 221, 219, 217, 214, 212, 209, 203, 196, 189, 182, 175, 167, 163}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function CTAS750010(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {229, 235, 240, 245, 250, 255, 260, 265, 270, 271, 271, 270, 269, 267, 265, 264}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function CIAS950020(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {224, 222, 220, 217, 215, 213, 210, 205, 199, 193, 186, 178, 171, 162, 154, 149}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function CTAS950020(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {232, 236, 240, 245, 250, 255, 259, 262, 263, 263, 262, 260, 258, 254, 250, 247}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function CIAS850020(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {225, 223, 221, 219, 216, 214, 211, 207, 201, 195, 188, 181, 173, 166, 158, 154}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function CTAS850020(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {233, 237, 242, 246, 251, 256, 261, 263, 265, 265, 265, 264, 262, 259, 256, 254}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function CIAS750020(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {226, 224, 222, 220, 217, 215, 212, 208, 203, 197, 190, 183, 176, 168, 161, 157}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function CTAS750020(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {234, 238, 243, 248, 252, 257, 262, 265, 267, 268, 267, 267, 265, 264, 261, 260}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function RTASn20(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {170, 172, 174, 176, 178, 180, 183, 185, 188, 193, 198, 203, 208, 213, 211, 210}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function RTASn10(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {168, 170, 173, 175, 178, 180, 183, 187, 190, 196, 202, 206, 207, 205, 205, 208}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function RTAS0(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {203, 208, 211, 209, 207, 205, 203, 195, 195, 198, 201, 202, 205, 208, 213, 216}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function RTAS10(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {195, 199, 204, 207, 210, 212, 214, 214, 213, 213, 215, 216, 216, 215, 215, 217}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

    Function RTAS20(ByVal altitude As Double) As Double
        Dim matrix(15) As Integer
        Dim lower As Integer
        Dim higher As Integer
        Dim output As Double

        matrix = {218, 214, 210, 205, 205, 205, 206, 209, 212, 213, 213, 212, 212, 214, 221, 224}

        If altitude > 29000 Then
            output = matrix(15)
        ElseIf altitude > 28000 Then
            output = matrix(14) + ((altitude - 28000) / 1000) * (matrix(15) - matrix(14))
        Else
            lower = Math.Floor(altitude / 2000)
            higher = lower + 1

            If lower = (altitude / 2000) Then
                output = matrix(lower)
            Else
                output = matrix(lower) + ((altitude - (lower * 2000)) / 2000) * (matrix(higher) - matrix(lower))
            End If

        End If

        Return output
    End Function

#End Region
#Region " Settings "

    Public Sub LLogin() Handles LLogin3.Click
        If LLogin2.Text = WSData(0) Then
            PanelSettings.Visible = True
            PanelSettingsLogin.Visible = False
            LLoadData()
        Else
            MsgBox("Password is incorrect")
        End If
    End Sub

    Private Sub LTextBox_KeyDown(sender As Object, e As KeyEventArgs) Handles LLogin2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            LLogin3.PerformClick()
        End If
    End Sub

    Public Sub LLoadData()
        LListBox1.Items.Clear()
        LListBox1.Items.AddRange(Airplanes)
        LListBox1.SelectedIndex = 0
        LTextBox4.Text = WSData(0)
        LButton1.Text = "Save"
        ReDim EditedAirplanes(Airplanes.Length - 1)
        ReDim EditedWeights(Airplanes.Length - 1)
        ReDim EditedMoments(Airplanes.Length - 1)
        EditedAirplanes = Airplanes
        For i As Integer = 0 To Airplanes.Length - 1
            EditedWeights(i) = WBBasicEmptyWeights(i)
            EditedMoments(i) = WBBasicEmptyMoments(i)
        Next
        LTextBox1.Text = EditedAirplanes(0)
        LTextBox2.Text = EditedWeights(0)
        LTextBox3.Text = EditedMoments(0)
    End Sub

    Public Sub LChangeAirplane() Handles LListBox1.SelectedIndexChanged
        If LListBox1.Tag = 1 Then
            LTextBox1.Text = EditedAirplanes(LListBox1.SelectedIndex)
            LTextBox2.Text = EditedWeights(LListBox1.SelectedIndex)
            LTextBox3.Text = EditedMoments(LListBox1.SelectedIndex)
        End If
        LListBox1.Tag = 1
    End Sub

    Public Sub LChangeTextBox1() Handles LTextBox1.TextChanged
        EditedAirplanes(LListBox1.SelectedIndex) = LTextBox1.Text
        LListBox1.Tag = 2
        LListBox1.Items.Item(LListBox1.SelectedIndex) = LTextBox1.Text
    End Sub

    Public Sub LChangeTextBox2() Handles LTextBox2.TextChanged
        EditedWeights(LListBox1.SelectedIndex) = LTextBox2.Text
    End Sub

    Public Sub LChangeTextBox3() Handles LTextBox3.TextChanged
        EditedMoments(LListBox1.SelectedIndex) = LTextBox3.Text
    End Sub

    Public Sub LCheckData() Handles LButton1.Click
        Dim errorDetected = False
        For i As Integer = 0 To EditedAirplanes.Length - 1
            If (Not IsNumeric(EditedWeights(i))) Or (Not IsNumeric(EditedMoments(i))) Then errorDetected = True
        Next
        If errorDetected Then
            MsgBox("Invalid Input!")
        Else
            LSaveData()
        End If
    End Sub

    Public Sub LSaveData()
        Dim WebSiteText As String = WebBrowser7.DocumentText
        Dim theElementCollection As HtmlElementCollection = WebBrowser7.Document.GetElementsByTagName("input")
        Dim newText As String = LTextBox4.Text
        WebBrowser7.Tag = 2
        LButton1.Text = "Saving..."
        For i As Integer = 0 To EditedAirplanes.Length - 1
            newText = newText + "," + EditedAirplanes(i) + "," + EditedWeights(i) + "," + EditedMoments(i)
        Next
        For Each planelement As HtmlElement In theElementCollection
            If planelement.GetAttribute("type") = "text" Then
                planelement.SetAttribute("value", newText)
            End If
        Next
        For Each planeElement As HtmlElement In theElementCollection
            If planeElement.GetAttribute("type") = "submit" Then
                planeElement.InvokeMember("click")
            End If
        Next
    End Sub

    Public Sub LAdd() Handles LButton2.Click
        Dim oldLength As Integer = EditedAirplanes.Length
        ReDim Preserve EditedAirplanes(oldLength)
        ReDim Preserve EditedWeights(oldLength)
        ReDim Preserve EditedMoments(oldLength)
        EditedAirplanes(oldLength) = ""
        EditedWeights(oldLength) = ""
        EditedMoments(oldLength) = ""
        LListBox1.Items.Add(EditedAirplanes(oldLength))
        LListBox1.SelectedIndex = oldLength
        LTextBox1.Text = ""
        LTextBox2.Text = ""
        LTextBox3.Text = ""
    End Sub

    Public Sub LDelete() Handles LButton3.Click
        If MsgBox("Do you want to delete selected airplane?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            Dim deleteIndex As Integer = LListBox1.SelectedIndex
            Dim ArraysLength As Integer = EditedAirplanes.Length - 2
            LListBox1.Tag = 2
            LListBox1.Items.RemoveAt(deleteIndex)
            EditedAirplanes = RemoveAt(EditedAirplanes, deleteIndex)
            EditedWeights = RemoveAt(EditedWeights, deleteIndex)
            EditedMoments = RemoveAt(EditedMoments, deleteIndex)
            ReDim Preserve EditedAirplanes(ArraysLength)
            ReDim Preserve EditedWeights(ArraysLength)
            ReDim Preserve EditedMoments(ArraysLength)
        End If
    End Sub

    Public Function RemoveAt(ByVal oldArray() As String, ByVal deletedIndex As Integer) As Array
        Dim newArray(oldArray.Length - 1) As String
        Dim n As Integer = 0
        For i As Integer = 0 To oldArray.Length - 1
            If n = i Then
                If deletedIndex <> i Then
                    newArray(i) = oldArray(i)
                    n = n + 1
                End If
            Else
                newArray(n) = oldArray(i)
                n = n + 1
            End If
        Next
        Return newArray
    End Function

    Private Sub WebBrowser7_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles WebBrowser7.DocumentCompleted
        If WebBrowser7.Tag = 2 Then
            MainSelectionChange(0)
            MyBase.OnLoad(e)
        End If
    End Sub

#End Region

#Region " TOLD Card "

    Public Sub GLoadData()
        GLabel1.Text = CInt(TakeoffPressureAltitude)
        GLabel2.Text = Math.Ceiling(TakeoffPower / 10) * 10
        GLabel3.Text = Math.Ceiling(TakeoffDistance / 100) * 100
        GLabel4.Text = Math.Ceiling(AccelerateStopDistance / 100) * 100
        GLabel5.Text = Math.Ceiling(AccelerateGoDistance / 100) * 100
        GLabel6.Text = Math.Ceiling(ClimbRate / 10) * 10
        GLabel7.Text = Math.Ceiling(ClimbServiceCeiling / 500) * 500
        GLabel8.Text = CInt(LandingPressureAltitude)
        GLabel9.Text = Math.Ceiling(LandingDistance / 100) * 100
        GLabel10.Text = CInt(V1)
        GLabel11.Text = CInt(V1)
        GLabel12.Text = CInt(V2)
        GLabel13.Text = "108"
        GLabel14.Text = CInt(Veref)
        GLabel15.Text = CInt(Vref)
        GLabel16.Text = Math.Round(WBZeroFuelWeight / 10) * 10
        GLabel17.Text = CInt(TakeoffWeight)
        GLabel18.Text = CInt(LandingWeight)
        If CruiseAltitude >= 18000 Then
            GLabel19.Text = "FL" + CInt(CruiseAltitude / 100).ToString
        Else
            GLabel19.Text = CInt(CruiseAltitude)
        End If
        GLabel20.Text = Math.Ceiling(CruiseTorque / 10) * 10
        GLabel21.Text = Math.Ceiling(CruiseFuelFlow / 10) * 10
        If CruiseIndicatedAirspeed = 888888 Then GLabel22.Text = "NA"
        If CruiseIndicatedAirspeed <> 888888 Then GLabel22.Text = CInt(CruiseIndicatedAirspeed)
        GLabel23.Text = CInt(CruiseTrueAirspeed)
        If FPDep = "999999" Then
            GLabel24.Visible = False
        Else
            GLabel24.Text = Microsoft.VisualBasic.Right(FPDep, 3) + "-" + Microsoft.VisualBasic.Right(FPDes, 3)
            GLabel24.Visible = True
        End If
        If FPDate = "999999" Then
            GLabel25.Visible = False
        Else
            GLabel25.Text = FPDate
            GLabel25.Visible = True
        End If
    End Sub

#End Region
#Region " Login "

    Public Sub HLoadData()
        HProgressBar1.Visible = False
        FPLoadHomepage = False
        WebBrowser1.Tag = 1
        'WebBrowser1.Navigate("http://fltplan.com")
        WebBrowser1.Navigate("http://fltplan2.com")
        'WebBrowser1.Navigate("http://12.132.107.203/fltplan4.htm")
        If FPUsername <> "999999" Then HTextBox1.Text = FPUsername
        If FPPassword <> "999999" Then HTextBox2.Text = FPPassword
        'Console.WriteLine("Hello")
    End Sub

    Private Sub HButton1_Click(sender As Object, e As EventArgs) Handles HButton1.Click
        If HTextBox1.Text = "" Or HTextBox2.Text = "" Then
            MsgBox("Please enter Username and Password")
        Else
            FPUsername = HTextBox1.Text
            FPPassword = HTextBox2.Text
            HProgressBar1.Visible = True
            If FPLoadHomepage = False Then
                WebBrowser1.Tag = 2
            Else
                WebBrowser1.Tag = 3
                HLogin()
            End If
        End If
    End Sub

    Private Sub HTextBox_KeyDown(sender As Object, e As KeyEventArgs) Handles HTextBox1.KeyDown, HTextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            HButton1.PerformClick()
        End If
    End Sub

    Private Sub WebBrowser1_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted
        If WebBrowser1.Tag = 1 Then
            FPLoadHomepage = True
        ElseIf WebBrowser1.Tag = 2 Then
            WebBrowser1.Tag = 3
            HLogin()
        ElseIf WebBrowser1.Tag = 3 Then
            HProgressBar1.Visible = False
            If HCheckLogin() Then
                MainSelectionChange(9)
            Else
                HLoadData()
            End If
        ElseIf WebBrowser1.Tag = 4 Then

            Dim planText As String = WebBrowser1.DocumentText
            Dim altitude As String
            Dim temperature() As String
            Dim elevation As String
            Dim elevations() As String
            Dim fuel As String = ""
            Dim Wfuel() As String
            Dim plane() As String = Microsoft.VisualBasic.Split(planText, "<TR><TD><center><font color=black size=-2>")
            Dim airplane As String
            FPDep = Mid(planText, InStr(planText, "Dep: ", CompareMethod.Binary) + 5, 4)
            FPDes = Mid(planText, InStr(planText, "Dest: ", CompareMethod.Binary) + 6, 4)
            FPETD = CInt(Mid(planText, InStr(planText, "Z</center></TD><TD><center>", CompareMethod.Binary) - 4, 4))
            FPETA = (CInt(Mid(planText, InStr(planText, "Arr: ", CompareMethod.Binary) + 5, 4)) - CInt(Mid(planText, InStr(planText, "Dept: ", CompareMethod.Binary) + 6, 4))) + FPETD
            If plane(1)(5) = "<" Then
                airplane = Mid(plane(1), 1, 5)
            Else
                airplane = Mid(plane(1), 1, 6)
            End If
            WBAircraft = Array.IndexOf(Airplanes, airplane)

            Dim altitudeCheck As Integer
            altitudeCheck = InStr(planText, "<TD><center><font color=black size=-2>FL", CompareMethod.Binary)
            If altitudeCheck <> 0 Then
                altitude = Mid(planText, InStr(planText, "<TD><center><font color=black size=-2>FL", CompareMethod.Binary) + 38, 5)
                If Microsoft.VisualBasic.Right(altitude, 1) = "<" Then
                    CruiseAltitude = 100 * CInt(Mid(altitude, 3, 2))
                Else
                    CruiseAltitude = 100 * CInt(Mid(altitude, 3, 3))
                End If
            Else
                altitude = Mid(planText, InStr(planText, "0</center></TD><TD>&nbsp;", CompareMethod.Binary) - 5, 6)
                If Microsoft.VisualBasic.Left(altitude, 1) = ">" Then
                    If Microsoft.VisualBasic.Left(altitude, 2) = ">F" Then
                        CruiseAltitude = 100 * CInt(Microsoft.VisualBasic.Right(altitude, 3))
                    Else
                        CruiseAltitude = 1000 * CInt(Mid(altitude, 2, 1))
                    End If
                Else
                    CruiseAltitude = 1000 * CInt(Mid(altitude, 1, 2))
                End If
            End If

            temperature = Split(planText, " &nbsp; ")
            If temperature(5) <> "N/A" Then
                CruiseTemperature = CInt(Microsoft.VisualBasic.Left(temperature(5), 3))
            ElseIf temperature(5) = "N/A" Then
                CruiseTemperature = CInt(Microsoft.VisualBasic.Left(temperature(3), 3))
            End If
            Console.WriteLine("done")
            UnitTakeoffAltitude = 2

            elevations = Microsoft.VisualBasic.Split(planText, "Elev:")
            TakeoffAltitude = CInt(Mid(elevations(1), 1, 4))
            UnitLandingAltitude = 2
            elevation = Mid(elevations(2), 1, 4)

            While IsNumeric(elevation) = False
                elevation = Mid(elevation, 1, Len(elevation) - 1)
            End While
            LandingAltitude = CInt(elevation)

            Wfuel = Split(planText, " Lbs")
            If Wfuel.Length = 4 Then
                fuel = Mid(planText, InStr(planText, " Lbs", CompareMethod.Binary) - 5, 5)
            ElseIf Wfuel.Length = 5 Then
                fuel = Microsoft.VisualBasic.Right(Wfuel(1), 5)
            End If
            If fuel.Chars(1) = ">" Then
                fuel = Microsoft.VisualBasic.Right(fuel, 3)
            End If
            WBFuelUsed = CInt(fuel)

            WebBrowser2.Navigate("http://aviationweather.gov/metar/data?ids=" + FPDep + "&format=raw&hours=0&taf=on&layout=off&date=0")
            WebBrowser3.Navigate("http://aviationweather.gov/metar/data?ids=" + FPDes + "&format=raw&hours=0&taf=on&layout=off&date=0")
            WebBrowser4.Navigate("http://aviationweather.gov/windtemp/data?level=low&fcst=06&region=all&layout=on&date=")
            WebBrowser5.Navigate("http://www.airnav.com/airport/" + FPDep)
            WebBrowser6.Navigate("http://www.airnav.com/airport/" + FPDes)

            'IWeb()
            'IProgressBar1.Visible = False
            'MainSelectionChange(10)

        End If
    End Sub

    Public Sub HLogin()
        Dim theElementCollection As HtmlElementCollection
        theElementCollection = WebBrowser1.Document.GetElementsByTagName("input")
        For Each curElement As HtmlElement In theElementCollection
            Dim controlName As String = curElement.GetAttribute("name").ToString
            If controlName = "USERNAME" Then
                curElement.SetAttribute("VALUE", FPUsername)
            ElseIf controlName = "PASSWORD" Then
                curElement.SetAttribute("VALUE", FPPassword)
            End If
        Next
        For Each curElement As HtmlElement In theElementCollection
            If curElement.GetAttribute("VALUE").Equals("ENTER") Then
                curElement.InvokeMember("CLICK")
                Exit For
            End If
        Next
    End Sub

    Public Function HCheckLogin() As Boolean
        If Split(WebBrowser1.DocumentText, "Please Verify and Try Again").Length = 2 Then
            MsgBox("Incorrect Username or Password!")
            Return False
        Else
            Return True
        End If
    End Function

#End Region
#Region " Flight Plans "

    Public Sub ILoadData()
        IProgressBar1.Visible = False
        IButton1.Text = "Load"
        IListView1.Clear()
        IListBox1.Items.Clear()
        WebBrowser2.Tag = 0
        WebBrowser3.Tag = 0
        WebBrowser4.Tag = 0
        WebBrowser5.Tag = 0
        WebBrowser6.Tag = 0
        Dim webText As String = WebBrowser1.DocumentText
        Dim numberOfPlans() As String
        Dim webArray() As String = Split(webText)
        Dim detailsOfPlans() As String
        Dim planDate As String = ""
        Dim planDTime As String = ""
        Dim planATime As String = ""
        Dim planPlane As String = ""
        Dim x As Integer
        Dim y As Integer
        Dim i As Integer = 0
        Dim n As Integer = 0
        Dim planDateSplitCheck() As String

        numberOfPlans = Filter(webArray, "QC", True, CompareMethod.Binary)
        detailsOfPlans = Filter(webArray, "<u>K", True, CompareMethod.Binary)
        If numberOfPlans.Length = 0 Then
            IListBox1.Items.Add("No flight plans were found")
        End If

        y = InStr(webText, "</thead>", CompareMethod.Binary)
        For Each planElement As String In detailsOfPlans
            x = InStr(planElement, "<u>K", CompareMethod.Binary)
            IListView1.Items.Add(Mid(planElement, x + 3, 4))

            i += 1
            If i = 2 Then
                If Microsoft.VisualBasic.Left(planDate, 3) = "<b>" Then
                    IListBox1.Items.Add("    " + (IListView1.Items.Item(IListView1.Items.Count - 2)).Text + "  ->  " + (IListView1.Items.Item(IListView1.Items.Count - 1).Text) + "       " + "   -   " + planPlane.ToString)
                Else
                    IListBox1.Items.Add(planDate + ":    " + (IListView1.Items.Item(IListView1.Items.Count - 2)).Text + "  ->  " + (IListView1.Items.Item(IListView1.Items.Count - 1).Text) + "       (" + planDTime.ToString + "-" + planATime.ToString + ")   -   " + planPlane.ToString)
                End If
                i = 0

                y = InStr(y, webText, "</tr><tr class=""fpc5", CompareMethod.Binary)
                'Mid(webText, y, 2) = "xx"

            Else
                y = InStr(y, webText, "><center>", CompareMethod.Binary)
                planDate = Mid(webText, y + 9, 10)

                Mid(webText, y, 2) = "xx"
                If Microsoft.VisualBasic.Right(planDate, 1) = "<" Then
                    planDate = Mid(planDate, 1, 9)
                End If

                planDateSplitCheck = Microsoft.VisualBasic.Split(planDate)
                If (planDateSplitCheck.Length < 2) Then
                    Exit For
                End If

                FPDate = planDateSplitCheck(1)

                y = InStr(y, webText, "<TD><center>", CompareMethod.Binary)
                planDTime = Mid(webText, y + 12, 4)
                Mid(webText, y, 2) = "xx"

                y = InStr(y, webText, "<TD><center>", CompareMethod.Binary)
                planATime = Mid(webText, y + 12, 4)
                Mid(webText, y, 2) = "xx"

                y = InStr(y, webText, "<TD width=50> &nbsp;", CompareMethod.Binary)
                planPlane = Mid(webText, y + 20, 6)
                Mid(webText, y, 2) = "xx"
                If Microsoft.VisualBasic.Right(planPlane, 1) = "<" Then
                    planPlane = Mid(planPlane, 1, 5)
                End If

            End If
        Next
    End Sub

    Private Sub IButton1_Click(sender As Object, e As EventArgs) Handles IButton1.Click
        If IListBox1.Items.Item(0) <> "No flight plans were found" Then
            If IListBox1.SelectedIndex = -1 Then
                MsgBox("Please select a flight plan")
            Else
                IButton1.Text = "Loading..."
                IProgressBar1.Value = 0
                IProgressBar1.Visible = True
                FPDate = Microsoft.VisualBasic.Left(Split(IListBox1.SelectedItem)(1), Split(IListBox1.SelectedItem)(1).Length - 1)
                Dim i As Integer = 1
                Dim n As Integer = 1
                Dim x As Integer = IListBox1.SelectedIndex + 1
                Dim theElementCollection As HtmlElementCollection
                theElementCollection = WebBrowser1.Document.GetElementsByTagName("input")
                IProgressBar1.Visible = True
                TakeoffPerformanceFormStatus = 0
                ClimbPerformanceFormStatus = 0
                CruisePerformanceFormStatus = 0
                LandingPerformanceFormStatus = 0
                WBFormStatus = 0
                SkipLogin = False
                ChangeToolbar(1)
                FPSelected = IListBox1.SelectedIndex
                FPFlightplans = IListBox1.Items.Count
                If IListBox1.SelectedIndex = IListBox1.Items.Count - 1 Then
                    NextFlightPlanAvailable = False
                Else
                    NextFlightPlanAvailable = True
                End If
                For Each planElement As HtmlElement In theElementCollection
                    If (planElement.GetAttribute("type") = "radio" Or planElement.GetAttribute("type") = "RADIO") Then
                        If i = 2 Then
                            If n = x Then
                                WebBrowser1.Tag = 4
                                planElement.InvokeMember("click")
                                n += 1
                            Else
                                n += 1
                            End If
                            i += 1
                        ElseIf i = 6 Then
                            i = 1
                        Else
                            i += 1
                        End If
                    End If
                Next

            End If
        End If
    End Sub

    Public Sub ICheckComplete()
        If WebBrowser2.Tag + WebBrowser3.Tag + WebBrowser4.Tag + WebBrowser5.Tag + WebBrowser6.Tag = 5 Then
            If SkipLogin = False Then
                IProgressBar1.PerformStep()
                IProgressBar1.Visible = False
            Else
                TLabel6.Text = "LOAD NEXT LEG"
                Me.Cursor = Cursors.Default
            End If
            CompleteEverything = 0
            MainSelectionChange(10)
        Else
            If SkipLogin = False Then
                IProgressBar1.PerformStep()
            End If
        End If
    End Sub

    Private Sub WebBrowser2_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles WebBrowser2.DocumentCompleted
        Dim weatherText As String = WebBrowser2.DocumentText
        weatherText = Microsoft.VisualBasic.Split(weatherText, "<!-- Data starts here -->")(1)
        weatherText = Microsoft.VisualBasic.Split(weatherText, "<!-- Data ends here -->")(0)
        Dim data() As String = Microsoft.VisualBasic.Split(weatherText, "FM")
        Dim extracted() As String
        Dim needed As String

        For i As Integer = 0 To data.Length - 1
            data(i) = Mid(data(i), 1, InStr(data(i), "<", CompareMethod.Binary) - 1)
        Next

        'altimeter setting
        extracted = Microsoft.VisualBasic.Split(data(0))
        For i As Integer = 0 To extracted.Length - 1
            If Microsoft.VisualBasic.Left(extracted(i), 1) = "A" And extracted(i) <> "AUTO" Then
                needed = Microsoft.VisualBasic.Right(extracted(i), 4)
                TakeoffAltimeterSetting = needed / 100
                Exit For
            End If
        Next

        'temperature
        For i As Integer = 0 To extracted.Length - 1
            If extracted(i).Length > 2 Then
                If extracted(i).Chars(2) = "/" Then
                    TakeoffTemperature = CInt(Microsoft.VisualBasic.Left(extracted(i), 2))
                    Exit For
                ElseIf extracted(i).Chars(0) = "M" Then
                    needed = Microsoft.VisualBasic.Left(extracted(i), 3)
                    TakeoffTemperature = -1 * CInt(Microsoft.VisualBasic.Right(needed, 2))
                    Exit For
                End If
            End If
        Next

        'wind
        For i As Integer = 0 To extracted.Length - 1
            If Microsoft.VisualBasic.Right(extracted(i), 2) = "KT" Then
                If Microsoft.VisualBasic.Left(extracted(i), 3) = "VRB" Then
                    FPDWindDirection = 9999
                Else
                    FPDWindDirection = CInt(Microsoft.VisualBasic.Left(extracted(i), 3))
                End If
                FPDWindSpeed = CInt(Mid(extracted(i), 4, 2))
                FPDWindGust = CInt(Mid(extracted(i), extracted(i).Length - 3, 2) - FPDWindSpeed)
                Exit For
            End If
        Next

        If data.Length = 1 Then
            'no TAF
        Else
            'There is TAF
        End If

        FPMetars(0) = Split(weatherText, "<br")(0)

        TakeoffPressureAltitude = ConvertAltitude(TakeoffAltitude, TakeoffAltimeterSetting)
        WebBrowser2.Tag = 1
        ICheckComplete()
    End Sub

    Private Sub WebBrowser3_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles WebBrowser3.DocumentCompleted
        Dim weatherText As String = WebBrowser3.DocumentText
        weatherText = Microsoft.VisualBasic.Split(weatherText, "<!-- Data starts here -->")(1)
        weatherText = Microsoft.VisualBasic.Split(weatherText, "<!-- Data ends here -->")(0)
        Dim data() As String = Microsoft.VisualBasic.Split(weatherText, "FM")
        Dim extracted() As String
        Dim needed As String

        For i As Integer = 0 To data.Length - 1
            data(i) = Mid(data(i), 1, InStr(data(i), "<", CompareMethod.Binary) - 1)
        Next

        'altimeter setting
        extracted = Microsoft.VisualBasic.Split(data(0))
        For i As Integer = 0 To extracted.Length - 1
            If Microsoft.VisualBasic.Left(extracted(i), 1) = "A" And extracted(i) <> "AUTO" Then
                needed = Microsoft.VisualBasic.Right(extracted(i), 4)
                LandingAltimeterSetting = needed / 100
                Exit For
            End If
        Next

        'temperature
        For i As Integer = 0 To extracted.Length - 1
            If extracted(i).Length > 2 Then
                If extracted(i).Chars(2) = "/" Then
                    LandingTemperature = CInt(Microsoft.VisualBasic.Left(extracted(i), 2))
                    Exit For
                ElseIf extracted(i).Chars(0) = "M" Then
                    needed = Microsoft.VisualBasic.Left(extracted(i), 3)
                    LandingTemperature = -1 * CInt(Microsoft.VisualBasic.Right(needed, 2))
                    Exit For
                End If
            End If
        Next

        'wind
        For i As Integer = 0 To extracted.Length - 1
            If Microsoft.VisualBasic.Right(extracted(i), 2) = "KT" Then
                If Microsoft.VisualBasic.Left(extracted(i), 3) = "VRB" Then
                    FPAWindDirection = 9999
                Else
                    FPAWindDirection = CInt(Microsoft.VisualBasic.Left(extracted(i), 3))
                End If
                FPAWindSpeed = CInt(Mid(extracted(i), 4, 2))
                FPAWindGust = CInt(Mid(extracted(i), extracted(i).Length - 3, 2) - FPAWindSpeed)
                Exit For
            End If
        Next

        If data.Length = 1 Then
            'no TAF
        Else
            'There is TAF
        End If

        FPMetars(1) = Split(weatherText, "<br")(0)

        LandingPressureAltitude = ConvertAltitude(LandingAltitude, LandingAltimeterSetting)
        WebBrowser3.Tag = 1
        ICheckComplete()
    End Sub

    Private Sub WebBrowser4_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles WebBrowser4.DocumentCompleted
        Dim windText As String = WebBrowser4.DocumentText
        Dim n As Integer = 0
        windText = Microsoft.VisualBasic.Split(windText, "FT  3000    6000    9000   12000   18000   24000  30000  34000  39000")(1)
        windText = Microsoft.VisualBasic.Split(windText, "</pre>")(0)
        For i As Integer = 2 To windText.Length
            WAAirports(n) = Mid(windText, i, 3)
            WA12(n) = CInt(Mid(windText, i + 29, 3))
            WA18(n) = CInt(Mid(windText, i + 37, 3))
            WA24(n) = CInt(Mid(windText, i + 45, 3))
            n += 1
            i += 69
        Next

        foundWindsAloft = False
        For i = 0 To 174
            If Microsoft.VisualBasic.Right(FPDep, 3) = WAAirports(i) Then
                ClimbTemperature12 = WA12(i)
                ClimbTemperature18 = WA18(i)
                ClimbTemperature24 = WA24(i)
                foundWindsAloft = True
            End If
        Next

        WebBrowser4.Tag = 1
        ICheckComplete()
    End Sub

    Private Sub WebBrowser5_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles WebBrowser5.DocumentCompleted
        Dim airportText As String = WebBrowser5.DocumentText
        Dim runways() As String = Microsoft.VisualBasic.Split(airportText, "<H4>Runway ")
        DepRunways = ""

        For i As Integer = 1 To runways.Length - 1
            DepRunways = DepRunways + Microsoft.VisualBasic.Left(runways(i), InStr(runways(i), "</H4>", CompareMethod.Binary) - 1) + "/"

        Next
        DepRunways = Mid(DepRunways, 1, DepRunways.Length - 1)

        WebBrowser5.Tag = 1
        ICheckComplete()
    End Sub

    Private Sub WebBrowser6_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles WebBrowser6.DocumentCompleted
        Dim airportTextDes As String = WebBrowser6.DocumentText
        Dim runwaysDes() As String = Microsoft.VisualBasic.Split(airportTextDes, "<H4>Runway ")
        ArrRunways = ""
        For j As Integer = 1 To runwaysDes.Length - 1
                ArrRunways = ArrRunways + Microsoft.VisualBasic.Left(runwaysDes(j), InStr(runwaysDes(j), "</H4>", CompareMethod.Binary) - 1) + "/"
            Next
        ArrRunways = Mid(ArrRunways, 1, ArrRunways.Length - 1)

        WebBrowser6.Tag = 1
        ICheckComplete()
    End Sub

    Public Sub IWeb()
        Dim request As WebRequest = WebRequest.Create("http://aviationweather.gov/metar/data?ids=" + FPDep + "&format=raw&hours=0&taf=on&layout=off&date=0")
        request.Credentials = CredentialCache.DefaultCredentials
        Dim response As WebResponse = request.GetResponse()
        Dim dataStream As Stream = response.GetResponseStream()
        Dim reader As New StreamReader(dataStream)
        Dim weatherText As String = reader.ReadToEnd
        weatherText = Microsoft.VisualBasic.Split(weatherText, "<!-- Data starts here -->")(1)
        weatherText = Microsoft.VisualBasic.Split(weatherText, "<!-- Data ends here -->")(0)
        Dim data() As String = Microsoft.VisualBasic.Split(weatherText, "FM")
        Dim extracted() As String
        Dim needed As String

        For i As Integer = 0 To data.Length - 1
            data(i) = Mid(data(i), 1, InStr(data(i), "<", CompareMethod.Binary) - 1)
        Next

        'altimeter setting
        extracted = Microsoft.VisualBasic.Split(data(0))
        For i As Integer = 0 To extracted.Length - 1
            If Microsoft.VisualBasic.Left(extracted(i), 1) = "A" And extracted(i) <> "AUTO" Then
                needed = Microsoft.VisualBasic.Right(extracted(i), 4)
                TakeoffAltimeterSetting = needed / 100
                Exit For
            End If
        Next

        'temperature
        For i As Integer = 0 To extracted.Length - 1
            If extracted(i).Chars(2) = "/" Then
                TakeoffTemperature = CInt(Microsoft.VisualBasic.Left(extracted(i), 2))
                Exit For
            ElseIf extracted(i).Chars(0) = "M" Then
                needed = Microsoft.VisualBasic.Left(extracted(i), 3)
                TakeoffTemperature = -1 * CInt(Microsoft.VisualBasic.Right(needed, 2))
                Exit For
            End If
        Next

        'wind
        For i As Integer = 0 To extracted.Length - 1
            If Microsoft.VisualBasic.Right(extracted(i), 2) = "KT" Then
                If Microsoft.VisualBasic.Left(extracted(i), 3) = "VRB" Then
                    FPDWindDirection = 9999
                Else
                    FPDWindDirection = CInt(Microsoft.VisualBasic.Left(extracted(i), 3))
                End If
                FPDWindSpeed = CInt(Mid(extracted(i), 4, 2))
                FPDWindGust = CInt(Mid(extracted(i), extracted(i).Length - 3, 2) - FPDWindSpeed)
                Exit For
            End If
        Next

        If data.Length = 1 Then
            'no TAF
        Else
            'There is TAF
        End If
        TakeoffPressureAltitude = ConvertAltitude(TakeoffAltitude, TakeoffAltimeterSetting)
        reader.Close()
        response.Close()
        IProgressBar1.PerformStep()
        '---------------------------------------------------------------------------------------
        request = WebRequest.Create("http://aviationweather.gov/metar/data?ids=" + FPDes + "&format=raw&hours=0&taf=on&layout=off&date=0")
        response = request.GetResponse()
        dataStream = response.GetResponseStream()
        reader = New StreamReader(dataStream)
        weatherText = reader.ReadToEnd
        weatherText = Microsoft.VisualBasic.Split(weatherText, "<!-- Data starts here -->")(1)
        weatherText = Microsoft.VisualBasic.Split(weatherText, "<!-- Data ends here -->")(0)
        data = Microsoft.VisualBasic.Split(weatherText, "FM")

        For i As Integer = 0 To data.Length - 1
            data(i) = Mid(data(i), 1, InStr(data(i), "<", CompareMethod.Binary) - 1)
        Next

        'altimeter setting
        extracted = Microsoft.VisualBasic.Split(data(0))
        For i As Integer = 0 To extracted.Length - 1
            If Microsoft.VisualBasic.Left(extracted(i), 1) = "A" And extracted(i) <> "AUTO" Then
                needed = Microsoft.VisualBasic.Right(extracted(i), 4)
                LandingAltimeterSetting = needed / 100
                Exit For
            End If
        Next

        'temperature
        For i As Integer = 0 To extracted.Length - 1
            If extracted(i).Chars(2) = "/" Then
                LandingTemperature = CInt(Microsoft.VisualBasic.Left(extracted(i), 2))
                Exit For
            ElseIf extracted(i).Chars(0) = "M" Then
                needed = Microsoft.VisualBasic.Left(extracted(i), 3)
                LandingTemperature = -1 * CInt(Microsoft.VisualBasic.Right(needed, 2))
                Exit For
            End If
        Next

        'wind
        For i As Integer = 0 To extracted.Length - 1
            If Microsoft.VisualBasic.Right(extracted(i), 2) = "KT" Then
                If Microsoft.VisualBasic.Left(extracted(i), 3) = "VRB" Then
                    FPAWindDirection = 9999
                Else
                    FPAWindDirection = CInt(Microsoft.VisualBasic.Left(extracted(i), 3))
                End If
                FPAWindSpeed = CInt(Mid(extracted(i), 4, 2))
                FPAWindGust = CInt(Mid(extracted(i), extracted(i).Length - 3, 2) - FPAWindSpeed)
                Exit For
            End If
        Next

        If data.Length = 1 Then
            'no TAF
        Else
            'There is TAF
        End If

        LandingPressureAltitude = ConvertAltitude(LandingAltitude, LandingAltimeterSetting)
        reader.Close()
        response.Close()
        IProgressBar1.PerformStep()
        '---------------------------------------------------------------------------------------
        request = WebRequest.Create("http://aviationweather.gov/windtemp/data?level=low&fcst=06&region=all&layout=on&date=")
        response = request.GetResponse()
        dataStream = response.GetResponseStream()
        reader = New StreamReader(dataStream)
        Dim windText As String = reader.ReadToEnd
        Dim n As Integer = 0
        windText = Microsoft.VisualBasic.Split(windText, "FT  3000    6000    9000   12000   18000   24000  30000  34000  39000")(1)
        windText = Microsoft.VisualBasic.Split(windText, "</pre>")(0)

        For i As Integer = 2 To windText.Length
            WAAirports(n) = Mid(windText, i, 3)
            WA12(n) = CInt(Mid(windText, i + 29, 3))
            WA18(n) = CInt(Mid(windText, i + 37, 3))
            WA24(n) = CInt(Mid(windText, i + 45, 3))
            n += 1
            i += 69
        Next

        foundWindsAloft = False
        For i = 0 To 174
            If Microsoft.VisualBasic.Right(FPDep, 3) = WAAirports(i) Then
                ClimbTemperature12 = WA12(i)
                ClimbTemperature18 = WA18(i)
                ClimbTemperature24 = WA24(i)
                foundWindsAloft = True
            End If
        Next
        reader.Close()
        response.Close()
        IProgressBar1.PerformStep()
        '-----------------------------------------------------------------------------------------
        request = WebRequest.Create("http://www.airnav.com/airport/" + FPDep)
        response = request.GetResponse()
        dataStream = response.GetResponseStream()
        reader = New StreamReader(dataStream)
        Dim airportText As String = reader.ReadToEnd
        Dim runways() As String = Microsoft.VisualBasic.Split(airportText, "<H4>Runway ")
        DepRunways = ""

        For i As Integer = 1 To runways.Length - 1
            DepRunways = DepRunways + Microsoft.VisualBasic.Left(runways(i), InStr(runways(i), "</H4>", CompareMethod.Binary) - 1) + "/"

        Next
        DepRunways = Mid(DepRunways, 1, DepRunways.Length - 1)
        reader.Close()
        response.Close()
        IProgressBar1.PerformStep()
        '----------------------------------------------------------------------------------------
        request = WebRequest.Create("http://www.airnav.com/airport/" + FPDes)
        response = request.GetResponse()
        dataStream = response.GetResponseStream()
        reader = New StreamReader(dataStream)
        Dim airportTextDes As String = reader.ReadToEnd
        Dim runwaysDes() As String = Microsoft.VisualBasic.Split(airportTextDes, "<H4>Runway ")
        ArrRunways = ""

        For j As Integer = 1 To runwaysDes.Length - 1
            ArrRunways = ArrRunways + Microsoft.VisualBasic.Left(runwaysDes(j), InStr(runwaysDes(j), "</H4>", CompareMethod.Binary) - 1) + "/"

        Next
        ArrRunways = Mid(ArrRunways, 1, ArrRunways.Length - 1)
        reader.Close()
        response.Close()
        IProgressBar1.PerformStep()
    End Sub

#End Region
#Region " Runways "

    Public Sub JLoadData()
        JComboBox1.Focus()
        JComboBox1.Items.Clear()
        JComboBox2.Items.Clear()
        JComboBox3.Items.Clear()
        Dim RunwayListDep() As String = Microsoft.VisualBasic.Split(DepRunways, "/")
        Dim RunwayListDes() As String = Microsoft.VisualBasic.Split(ArrRunways, "/")
        If FPDWindDirection = 9999 Then
            JLabel4.Text = "Variable"
        Else
            JLabel4.Text = FPDWindDirection.ToString
        End If
        If FPAWindDirection = 9999 Then
            JLabel11.Text = "Variable"
        Else
            JLabel11.Text = FPAWindDirection.ToString
        End If
        JLabel6.Text = FPDWindSpeed.ToString + " Knots"
        JLabel13.Text = FPAWindSpeed.ToString + " Knots"
        JLabel2.Text = FPDep
        JLabel9.Text = FPDes
        Array.Sort(RunwayListDep)
        Array.Sort(RunwayListDes)
        JComboBox1.Items.AddRange(RunwayListDep)
        JComboBox2.Items.AddRange(RunwayListDes)
        If foundWindsAloft = True Then
            JLabel15.Visible = False
            JComboBox3.Visible = False
        Else
            JLabel15.Visible = True
            JComboBox3.Visible = True
            JComboBox3.Items.AddRange(WAAirports)
        End If

    End Sub

    Public Function JCheckData() As Boolean
        Dim returnValue As Boolean = True
        If JComboBox1.SelectedIndex = -1 Then
            MsgBox("Select Takeoff Runway")
            returnValue = False
        End If
        If JComboBox2.SelectedIndex = -1 Then
            MsgBox("Select Landing Runway")
            returnValue = False
        End If
        If (Not foundWindsAloft) Then
            If JComboBox3.SelectedIndex = -1 Then
                MsgBox("Select Winds Aloft Station")
                returnValue = False
            End If
        End If
        Return returnValue
    End Function

    Public Sub JCalculateData()
        Dim headingDifferenceDep As Integer
        Dim headingDifferenceDes As Integer
        Dim runwayDirectionDep As String = ""
        Dim runwayDirectionDes As String = ""

        runwayDirectionDep = JComboBox1.SelectedItem
        runwayDirectionDes = JComboBox2.SelectedItem
        If IsNumeric(runwayDirectionDep.Chars(runwayDirectionDep.Length - 1)) Then
            runwayDirectionDep = runwayDirectionDep * 10
        Else
            runwayDirectionDep = Microsoft.VisualBasic.Left(runwayDirectionDep, runwayDirectionDep.Length - 1) * 10
        End If
        If IsNumeric(runwayDirectionDes.Chars(runwayDirectionDes.Length - 1)) Then
            runwayDirectionDes = runwayDirectionDes * 10
        Else
            runwayDirectionDes = Microsoft.VisualBasic.Left(runwayDirectionDes, runwayDirectionDes.Length - 1) * 10
        End If

        If FPDWindDirection = 9999 Then
            headingDifferenceDep = 0
        Else
            If CInt(runwayDirectionDep) > FPDWindDirection Then
                headingDifferenceDep = CInt(runwayDirectionDep) - FPDWindDirection
            Else
                headingDifferenceDep = FPDWindDirection - CInt(runwayDirectionDep)
            End If
            If headingDifferenceDep > 180 Then
                headingDifferenceDep = 360 - headingDifferenceDep
            End If
        End If
        If Math.Cos(headingDifferenceDep * Math.PI / 180) > 0 Then
            UnitTakeoffWind = 1
            TakeoffWind = Math.Round(FPDWindSpeed * Math.Cos(headingDifferenceDep * Math.PI / 180))
        Else
            UnitTakeoffWind = 2
            TakeoffWind = Math.Abs(Math.Round(FPDWindSpeed * Math.Cos(headingDifferenceDep * Math.PI / 180)))
        End If

        If FPAWindDirection = 9999 Then
            headingDifferenceDes = 0
        Else
            If CInt(runwayDirectionDes) > FPAWindDirection Then
                headingDifferenceDes = CInt(runwayDirectionDes) - FPAWindDirection
            Else
                headingDifferenceDes = FPAWindDirection - CInt(runwayDirectionDes)
            End If
            If headingDifferenceDes > 180 Then
                headingDifferenceDes = 360 - headingDifferenceDes
            End If
        End If

        If Math.Cos(headingDifferenceDes * Math.PI / 180) > 0 Then
            UnitLandingWind = 1
            LandingWind = Math.Round(FPAWindSpeed * Math.Cos(headingDifferenceDes * Math.PI / 180))
        Else
            UnitLandingWind = 2
            LandingWind = Math.Abs(Math.Round(FPAWindSpeed * Math.Cos(headingDifferenceDes * Math.PI / 180)))
        End If
        If (Not foundWindsAloft) Then
            ClimbTemperature12 = WA12(JComboBox3.SelectedIndex)
            ClimbTemperature18 = WA18(JComboBox3.SelectedIndex)
            ClimbTemperature24 = WA24(JComboBox3.SelectedIndex)
            foundWindsAloft = True
        End If
        FltPlanFormStatus = 1

    End Sub

    Private Sub JButton1_Click(sender As Object, e As EventArgs) Handles JButton1.Click
        If JCheckData() Then
            If CurrentToolbar = 0 Then ChangeToolbar(1)
            JCalculateData()
            MainSelectionChange(1)
            ALoadData()
        End If
    End Sub

#End Region
#Region " WB Graph "

    Public Sub KLoadData()
        Dim minY As Integer = 304
        Dim maxY As Integer = 9
        Dim minX As Integer = 128
        Dim maxX As Integer = 448
        Dim ZFWY As Integer = minY - ((WBZeroFuelWeight - 5800) / (10200 - 5800)) * (minY - maxY)
        Dim ZFWX As Integer = minX + ((CGZeroFuelWeight - 141) / (167 - 141)) * (maxX - minX)
        Dim RWY As Integer = minY - ((WBRampWeight - 5800) / (10200 - 5800)) * (minY - maxY)
        Dim RWX As Integer = minX + ((CGRampWeight - 141) / (167 - 141)) * (maxX - minX)
        Dim TWY As Integer = minY - ((TakeoffWeight - 5800) / (10200 - 5800)) * (minY - maxY)
        Dim TWX As Integer = minX + ((CGTakeoffWeight - 141) / (167 - 141)) * (maxX - minX)
        Dim LWY As Integer = minY - ((LandingWeight - 5800) / (10200 - 5800)) * (minY - maxY)
        Dim LWX As Integer = minX + ((CGLandingWeight - 141) / (167 - 141)) * (maxX - minX)

        KPictureBox1.Location = New Point(ZFWX, ZFWY)
        KPictureBox2.Location = New Point(RWX, RWY)
        KPictureBox3.Location = New Point(TWX, TWY)
        KPictureBox4.Location = New Point(LWX, LWY)
    End Sub

    Private Sub KButton1_Click(sender As Object, e As EventArgs) Handles KButton1.Click
        MainSelectionChange(1)
    End Sub

#End Region
#Region " Print "

    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        'TOLD Card
        e.Graphics.DrawImage(My.Resources.TOLD, 0, 0, 324, 374)
        e.Graphics.DrawString(CInt(TakeoffPressureAltitude), New Font("Comic Sans MS", 10, FontStyle.Regular), Brushes.Black, 130, 42)
        e.Graphics.DrawString(Math.Ceiling(TakeoffPower / 10) * 10, New Font("Comic Sans MS", 10, FontStyle.Regular), Brushes.Black, 130, 65)
        e.Graphics.DrawString(Math.Ceiling(TakeoffDistance / 100) * 100, New Font("Comic Sans MS", 10, FontStyle.Regular), Brushes.Black, 130, 89)
        e.Graphics.DrawString(Math.Ceiling(AccelerateStopDistance / 100) * 100, New Font("Comic Sans MS", 10, FontStyle.Regular), Brushes.Black, 130, 113)
        e.Graphics.DrawString(Math.Ceiling(AccelerateGoDistance / 100) * 100, New Font("Comic Sans MS", 10, FontStyle.Regular), Brushes.Black, 130, 137)
        e.Graphics.DrawString(Math.Ceiling(ClimbRate / 10) * 10, New Font("Comic Sans MS", 10, FontStyle.Regular), Brushes.Black, 130, 160)
        e.Graphics.DrawString(Math.Ceiling(ClimbServiceCeiling / 500) * 500, New Font("Comic Sans MS", 10, FontStyle.Regular), Brushes.Black, 125, 184)
        e.Graphics.DrawString(CInt(LandingPressureAltitude), New Font("Comic Sans MS", 10, FontStyle.Regular), Brushes.Black, 130, 208)
        e.Graphics.DrawString(Math.Ceiling(LandingDistance / 100) * 100, New Font("Comic Sans MS", 10, FontStyle.Regular), Brushes.Black, 130, 231)
        e.Graphics.DrawString(CInt(V1), New Font("Comic Sans MS", 10, FontStyle.Regular), Brushes.Black, 270, 42)
        e.Graphics.DrawString(CInt(V1), New Font("Comic Sans MS", 10, FontStyle.Regular), Brushes.Black, 270, 65)
        e.Graphics.DrawString(CInt(V2), New Font("Comic Sans MS", 10, FontStyle.Regular), Brushes.Black, 270, 89)
        e.Graphics.DrawString("108", New Font("Comic Sans MS", 10, FontStyle.Regular), Brushes.Black, 270, 113)
        e.Graphics.DrawString(CInt(Veref), New Font("Comic Sans MS", 10, FontStyle.Regular), Brushes.Black, 270, 137)
        e.Graphics.DrawString(CInt(Vref), New Font("Comic Sans MS", 10, FontStyle.Regular), Brushes.Black, 270, 160)
        e.Graphics.DrawString(CInt(WBZeroFuelWeight), New Font("Comic Sans MS", 10, FontStyle.Regular), Brushes.Black, 265, 184)
        e.Graphics.DrawString(CInt(TakeoffWeight), New Font("Comic Sans MS", 10, FontStyle.Regular), Brushes.Black, 265, 208)
        e.Graphics.DrawString(CInt(LandingWeight), New Font("Comic Sans MS", 10, FontStyle.Regular), Brushes.Black, 265, 231)
        If CruiseAltitude >= 18000 Then
            e.Graphics.DrawString("FL" + CInt(CruiseAltitude / 100).ToString, New Font("Comic Sans MS", 9, FontStyle.Regular), Brushes.Black, 52, 258)
        Else
            e.Graphics.DrawString(CInt(CruiseAltitude), New Font("Comic Sans MS", 9, FontStyle.Regular), Brushes.Black, 52, 258)

        End If
        e.Graphics.DrawString(Math.Ceiling(CruiseTorque / 10) * 10, New Font("Comic Sans MS", 9, FontStyle.Regular), Brushes.Black, 119, 258)
        e.Graphics.DrawString(Math.Ceiling(CruiseFuelFlow / 10) * 10, New Font("Comic Sans MS", 9, FontStyle.Regular), Brushes.Black, 169, 258)
        If CruiseIndicatedAirspeed <> 888888 Then e.Graphics.DrawString(CInt(CruiseIndicatedAirspeed), New Font("Comic Sans MS", 9, FontStyle.Regular), Brushes.Black, 225, 258)
        If CruiseIndicatedAirspeed = 888888 Then e.Graphics.DrawString("NA", New Font("Comic Sans MS", 9, FontStyle.Regular), Brushes.Black, 225, 258)
        e.Graphics.DrawString(CInt(CruiseTrueAirspeed), New Font("Comic Sans MS", 9, FontStyle.Regular), Brushes.Black, 275, 258)
        If FPDep <> "999999" Then e.Graphics.DrawString(Microsoft.VisualBasic.Right(FPDep, 3) + "-" + Microsoft.VisualBasic.Right(FPDes, 3), New Font("Comic Sans MS", 10, FontStyle.Regular), Brushes.Black, 25, 7)
        If FPDate <> "999999" Then e.Graphics.DrawString(FPDate, New Font("Comic Sans MS", 10, FontStyle.Regular), Brushes.Black, 122, 7)
        'Graph
        Dim minY As Integer = 802
        Dim maxY As Integer = 413
        Dim minX As Integer = 76
        Dim maxX As Integer = 500
        Dim ZFWY As Integer = minY - ((WBZeroFuelWeight - 5800) / (10200 - 5800)) * (minY - maxY)
        Dim ZFWX As Integer = minX + ((CGZeroFuelWeight - 141) / (167 - 141)) * (maxX - minX)
        Dim RWY As Integer = minY - ((WBRampWeight - 5800) / (10200 - 5800)) * (minY - maxY)
        Dim RWX As Integer = minX + ((CGRampWeight - 141) / (167 - 141)) * (maxX - minX)
        Dim TWY As Integer = minY - ((TakeoffWeight - 5800) / (10200 - 5800)) * (minY - maxY)
        Dim TWX As Integer = minX + ((CGTakeoffWeight - 141) / (167 - 141)) * (maxX - minX)
        Dim LWY As Integer = minY - ((LandingWeight - 5800) / (10200 - 5800)) * (minY - maxY)
        Dim LWX As Integer = minX + ((CGLandingWeight - 141) / (167 - 141)) * (maxX - minX)
        e.Graphics.DrawImage(My.Resources.graph, 0, 400)
        e.Graphics.DrawImage(My.Resources.moyn, ZFWX, ZFWY)
        e.Graphics.DrawImage(My.Resources.circle, RWX, RWY)
        e.Graphics.DrawImage(My.Resources.square, TWX, TWY)
        e.Graphics.DrawImage(My.Resources.triangle, LWX, LWY)

        e.Graphics.DrawImage(My.Resources.moyn, 550, 550)
        e.Graphics.DrawImage(My.Resources.circle, 550, 580)
        e.Graphics.DrawImage(My.Resources.square, 550, 610)
        e.Graphics.DrawImage(My.Resources.triangle, 550, 640)
        e.Graphics.DrawString("Zero Fuel Weight", New Font("Comic Sans MS", 10, FontStyle.Regular), Brushes.Black, 570, 545)
        e.Graphics.DrawString("Ramp Weight", New Font("Comic Sans MS", 10, FontStyle.Regular), Brushes.Black, 570, 575)
        e.Graphics.DrawString("Takeoff Weight", New Font("Comic Sans MS", 10, FontStyle.Regular), Brushes.Black, 570, 605)
        e.Graphics.DrawString("Landing Weight", New Font("Comic Sans MS", 10, FontStyle.Regular), Brushes.Black, 570, 635)
        'Data
        If FltPlanFormStatus = 1 Then
            e.Graphics.DrawString("Departure Metar:", New Font("Comic Sans MS", 10, FontStyle.Bold), Brushes.Black, 330, 5)
            e.Graphics.DrawString(FPMetars(0), New Font("Comic Sans MS", 9, FontStyle.Regular), Brushes.Black, New RectangleF(330, 15, 380, 100))
            e.Graphics.DrawString("Arrival Metar:", New Font("Comic Sans MS", 10, FontStyle.Bold), Brushes.Black, 330, 115)
            e.Graphics.DrawString(FPMetars(1), New Font("Comic Sans MS", 9, FontStyle.Regular), Brushes.Black, New RectangleF(330, 125, 380, 100))
        End If
        e.Graphics.DrawString("Fuel Load:", New Font("Comic Sans MS", 10, FontStyle.Bold), Brushes.Black, 330, 225)
        e.Graphics.DrawString(CInt(WBFuelLoad).ToString + " Lbs", New Font("Comic Sans MS", 10, FontStyle.Regular), Brushes.Black, 410, 225)
        e.Graphics.DrawString("Fuel Use:", New Font("Comic Sans MS", 10, FontStyle.Bold), Brushes.Black, 330, 250)
        e.Graphics.DrawString(CInt(WBFuelUsed).ToString + " Lbs", New Font("Comic Sans MS", 10, FontStyle.Regular), Brushes.Black, 400, 250)
        e.Graphics.DrawString("Fuel Remaining:", New Font("Comic Sans MS", 10, FontStyle.Bold), Brushes.Black, 330, 275)
        e.Graphics.DrawString(CInt(CDbl(WBFuelLoad) - CDbl(WBFuelUsed)).ToString + " Lbs", New Font("Comic Sans MS", 10, FontStyle.Regular), Brushes.Black, 440, 275)

    End Sub

#End Region
#Region " Load Next Leg "

    Private Sub TLabel6_Click(sender As Object, e As EventArgs) Handles TLabel6.Click
        WBFuelLoad = WBFuelLoad - WBFuelUsed
        SkipLogin = True
        WebBrowser8.Tag = 1
        TLabel6.Text = "LOADING..."
        Me.Cursor = Cursors.WaitCursor
        WebBrowser8.Navigate("http://fltplan2.com")
    End Sub


#End Region

#Region " Web Data "

    Public Sub WebLoadData()
        Dim request As WebRequest = WebRequest.Create("http://nodejs1.aero.und.edu/king-air-server")
        request.Credentials = CredentialCache.DefaultCredentials
        Dim response As WebResponse = request.GetResponse()
        Dim dataStream As Stream = response.GetResponseStream()
        Dim reader As New StreamReader(dataStream)
        Dim responseFromServer As String = reader.ReadToEnd
        WSData = Split(Microsoft.VisualBasic.Left(responseFromServer, InStr(responseFromServer, "<form") - 1), ",")
        reader.Close()
        response.Close()

        ReDim WBBasicEmptyWeights(((WSData.Length - 1) / 3) - 1)
        ReDim WBBasicEmptyMoments(((WSData.Length - 1) / 3) - 1)
        ReDim Airplanes(((WSData.Length - 1) / 3) - 1)
        Dim n As Integer = 1
        Dim r As Integer = 0
        For i As Integer = 1 To WSData.Length - 1
            If n = 1 Then
                Airplanes(r) = WSData(i)
                n = 2
            ElseIf n = 2 Then
                WBBasicEmptyWeights(r) = WSData(i)
                n = 3
            ElseIf n = 3 Then
                WBBasicEmptyMoments(r) = WSData(i)
                n = 1
                r = r + 1
            End If
        Next
    End Sub

    Private Sub WebBrowser8_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles WebBrowser8.DocumentCompleted
        If WebBrowser8.Tag = 1 Then
            Dim theElementCollection As HtmlElementCollection
            theElementCollection = WebBrowser8.Document.GetElementsByTagName("input")
            For Each curElement As HtmlElement In theElementCollection
                Dim controlName As String = curElement.GetAttribute("name").ToString
                If controlName = "USERNAME" Then
                    curElement.SetAttribute("VALUE", FPUsername)
                ElseIf controlName = "PASSWORD" Then
                    curElement.SetAttribute("VALUE", FPPassword)
                End If
            Next
            For Each curElement As HtmlElement In theElementCollection
                If curElement.GetAttribute("VALUE").Equals("ENTER") Then
                    curElement.InvokeMember("CLICK")
                    Exit For
                End If
            Next
            WebBrowser8.Tag = 2
        ElseIf WebBrowser8.Tag = 2 Then
            Dim i As Integer = 1
            Dim n As Integer = 1
            Dim x As Integer = FPSelected + 2
            Dim theElementCollection As HtmlElementCollection
            theElementCollection = WebBrowser8.Document.GetElementsByTagName("input")
            TakeoffPerformanceFormStatus = 0
            ClimbPerformanceFormStatus = 0
            CruisePerformanceFormStatus = 0
            LandingPerformanceFormStatus = 0
            WBFormStatus = 0
            ChangeToolbar(1)
            FPSelected = FPSelected + 1
            If FPSelected = FPFlightplans - 1 Then
                NextFlightPlanAvailable = False
            Else
                NextFlightPlanAvailable = True
            End If
            For Each planElement As HtmlElement In theElementCollection
                If (planElement.GetAttribute("type") = "radio") Or (planElement.GetAttribute("type") = "RADIO") Then
                    If i = 2 Then
                        If n = x Then
                            WebBrowser8.Tag = 3
                            planElement.InvokeMember("click")
                            n += 1
                        Else
                            n += 1
                        End If
                        i += 1
                    ElseIf i = 6 Then
                        i = 1
                    Else
                        i += 1
                    End If
                End If
            Next
        ElseIf WebBrowser8.Tag = 3 Then
            Dim planText As String = WebBrowser8.DocumentText
            Dim altitude As String
            Dim temperature() As String
            Dim elevation As String
            Dim elevations() As String
            Dim fuel As String = ""
            Dim Wfuel() As String
            Dim plane() As String = Microsoft.VisualBasic.Split(planText, "<TR><TD><center><font color=black size=-2>")
            Dim airplane As String
            FPDep = Mid(planText, InStr(planText, "Dep: ", CompareMethod.Binary) + 5, 4)
            FPDes = Mid(planText, InStr(planText, "Dest: ", CompareMethod.Binary) + 6, 4)
            FPETD = CInt(Mid(planText, InStr(planText, "Z</center></TD><TD><center>", CompareMethod.Binary) - 4, 4))
            FPETA = (CInt(Mid(planText, InStr(planText, "Arr: ", CompareMethod.Binary) + 5, 4)) - CInt(Mid(planText, InStr(planText, "Dept: ", CompareMethod.Binary) + 6, 4))) + FPETD
            If plane(1)(5) = "<" Then
                airplane = Mid(plane(1), 1, 5)
            Else
                airplane = Mid(plane(1), 1, 6)
            End If
            WBAircraft = Array.IndexOf(Airplanes, airplane)

            Dim altitudeCheck As Integer
            altitudeCheck = InStr(planText, "<TD><center><font color=black size=-2>FL", CompareMethod.Binary)
            If altitudeCheck <> 0 Then
                altitude = Mid(planText, InStr(planText, "<TD><center><font color=black size=-2>FL", CompareMethod.Binary) + 38, 5)
                If Microsoft.VisualBasic.Right(altitude, 1) = "<" Then
                    CruiseAltitude = 100 * CInt(Mid(altitude, 3, 2))
                Else
                    CruiseAltitude = 100 * CInt(Mid(altitude, 3, 3))
                End If
            Else
                altitude = Mid(planText, InStr(planText, "0</center></TD><TD>&nbsp;", CompareMethod.Binary) - 5, 6)
                If Microsoft.VisualBasic.Left(altitude, 1) = ">" Then
                    If Microsoft.VisualBasic.Left(altitude, 2) = ">F" Then
                        CruiseAltitude = 100 * CInt(Microsoft.VisualBasic.Right(altitude, 3))
                    Else
                        CruiseAltitude = 1000 * CInt(Mid(altitude, 2, 1))
                    End If
                Else
                    CruiseAltitude = 1000 * CInt(Mid(altitude, 1, 2))
                End If
            End If

            temperature = Split(planText, " &nbsp; ")
            If temperature(5) <> "N/A" Then
                CruiseTemperature = CInt(Microsoft.VisualBasic.Left(temperature(5), 3))
            ElseIf temperature(5) = "N/A" Then
                CruiseTemperature = CInt(Microsoft.VisualBasic.Left(temperature(3), 3))
            End If
            Console.WriteLine("done")
            UnitTakeoffAltitude = 2

            elevations = Microsoft.VisualBasic.Split(planText, "Elev:")
            TakeoffAltitude = CInt(Mid(elevations(1), 1, 4))
            UnitLandingAltitude = 2
            elevation = Mid(elevations(2), 1, 4)

            While IsNumeric(elevation) = False
                elevation = Mid(elevation, 1, Len(elevation) - 1)
            End While
            LandingAltitude = CInt(elevation)

            Wfuel = Split(planText, " Lbs")
            If Wfuel.Length = 4 Then
                fuel = Mid(planText, InStr(planText, " Lbs", CompareMethod.Binary) - 5, 5)
            ElseIf Wfuel.Length = 5 Then
                fuel = Microsoft.VisualBasic.Right(Wfuel(1), 5)
            End If
            If fuel.Chars(1) = ">" Then
                fuel = Microsoft.VisualBasic.Right(fuel, 3)
            End If
            WBFuelUsed = CInt(fuel)

            WebBrowser2.Navigate("http://aviationweather.gov/metar/data?ids=" + FPDep + "&format=raw&hours=0&taf=on&layout=off&date=0")
            WebBrowser3.Navigate("http://aviationweather.gov/metar/data?ids=" + FPDes + "&format=raw&hours=0&taf=on&layout=off&date=0")
            WebBrowser4.Navigate("http://aviationweather.gov/windtemp/data?level=low&fcst=06&region=all&layout=on&date=")
            WebBrowser5.Navigate("http://www.airnav.com/airport/" + FPDep)
            WebBrowser6.Navigate("http://www.airnav.com/airport/" + FPDes)

        End If
    End Sub



#End Region
End Class