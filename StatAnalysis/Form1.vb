Imports System.Deployment
Imports System.Linq
Imports System.IO
Imports Microsoft.office.interop.Excel

Public Class Form1
    Dim AR As New AccessRoutines
    Dim ExcelRoutn As New ExcelRoutines
    Dim TDHCoeff As New RegressionCoefficients

    Dim DSModelNumbers As DataSet
    Dim DSImpellerSizes As DataSet
    Dim DSSerialNumbers As DataSet

    Dim listImp As List(Of String) = New List(Of String)

    Dim SelectedModelNumber As String
    Dim SelectedImpellerSize As String

    'set up datasets for pipe diameters, vapor pressure and temperature correction
    Public PipeDiameters As DataSet
    Public VaporPressure As DataSet
    Public TemperatureCorrection As DataSet
    Public DischargeDiameter As DataSet
    Public SuctionDiameter As DataSet
    Public SuperMarketPumps As DataSet

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If System.Diagnostics.Debugger.IsAttached = False Then
            Me.Text = "Polar Statistical Analysis - Version: " &
            My.Application.Deployment.CurrentVersion.ToString
        Else
            Me.Text = "Polar Statistical Analysis - Version: " & My.Application.Info.Version.ToString
        End If

        'get the unc from f: and make the fully qualified database name
        Dim DriveLetter As String = "F"
        Dim UNC As String = AR.GetUncSourcePath(DriveLetter)
        If UNC = Nothing Then
            MsgBox("Cannot find the UNC path for " + DriveLetter, vbOKOnly, "Cannot find UNC")
            End
        End If

        Dim DatabaseName As String
        If Environment.MachineName = "MROSENBAUM-LT" Then
            DatabaseName = "C:\Databases\PolarData.mdb"
        Else
            DatabaseName = UNC & "\Groups\Shared\Databases\PolarData.mdb"
        End If


        'make the connection string for the Access Routines
        AR.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabaseName

        AR.SQL = "SELECT * FROM PipeDiameters;"
        PipeDiameters = AR.FillArrays("PipeDiameters")

        AR.SQL = "SELECT * FROM VaporPressure;"
        VaporPressure = AR.FillArrays("VaporPressure")

        AR.SQL = "SELECT * FROM TempCorrection;"
        TemperatureCorrection = AR.FillArrays("TempCorrection")

        AR.SQL = "SELECT * FROM DischargeDiameter;"
        DischargeDiameter = AR.FillArrays("DischargeDiameter")

        AR.SQL = "SELECT * FROM SuctionDiameter;"
        SuctionDiameter = AR.FillArrays("SuctionDiameter")

        AR.SQL = "SELECT * FROM SuperMarketPumpData"
        SuperMarketPumps = AR.FillArrays("Supermarket")

        btnGetModelNo_Click(sender, e)
        lblDesignFlow.Text = ""
        btnCalcRunCharts.Enabled = False

        ExcelRoutn.closeOpenExcelApps()

    End Sub

    Private Sub btnCalcRunCharts_Click(sender As Object, e As EventArgs) Handles btnCalcRunCharts.Click
        'get the serial numbers for the selected model number and impeller size

        Dim MultipleModels As Boolean = False

        If Not IsNumeric(txtDesignFlow.Text) Then
            MessageBox.Show("Please assure only numbers for Flow", "Only numbers for Flow", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        Label3.Visible = True
        Label3.Text = "Retrieving Serial Numbers . . ."
        Label3.Refresh()

        If lbSN.SelectedItems.Count = 1 Then
            AR.SQL = "SELECT SerialNumber FROM TempPumpData WHERE ModelNumber = '" & SelectedModelNumber & "' AND ImpellerDia = " & CDbl(Me.TextBox2.Text) & ";"
            DSSerialNumbers = AR.GetPolarQueryData("SerialNumbers")
            MultipleModels = False
        Else
            MultipleModels = True
            For j As Integer = 0 To lbSN.SelectedItems.Count - 1
                Dim SelMod As String = lbSN.SelectedItems(j)(0)
                AR.SQL = "SELECT SerialNumber FROM TempPumpData WHERE ModelNumber = '" & SelMod & "' AND ImpellerDia = " & CDbl(Me.TextBox2.Text) & ";"
                Dim dsTemp As DataSet = AR.GetPolarQueryData("SerialNumbers")
                dsTemp.Tables("SerialNumbers").Columns.Add("ModelNo", Type.GetType("System.String"))
                For Each row As DataRow In dsTemp.Tables(0).Rows
                    row("ModelNo") = SelMod
                Next
                If j = 0 Then
                    DSSerialNumbers = dsTemp
                Else
                    DSSerialNumbers.Merge(dsTemp, True, MissingSchemaAction.Add)
                    DSSerialNumbers.AcceptChanges()
                End If
            Next
        End If


        'add columns for the tdh readings
        DSSerialNumbers.Tables("SerialNumbers").Columns.Add("Design Head", Type.GetType("System.Single"))
        DSSerialNumbers.Tables("SerialNumbers").Columns.Add("WaterZero", Type.GetType("System.Single"))
        DSSerialNumbers.Tables("SerialNumbers").Columns.Add("WaterDesign", Type.GetType("System.Single"))

        If MultipleModels Then
            DSSerialNumbers.Tables("SerialNumbers").Columns("ModelNo").SetOrdinal(DSSerialNumbers.Tables("SerialNumbers").Columns.Count - 1)
        End If

        Dim HydTestValues(2) As Single

            '        Dim DVSN As DataView = DSSerialNumbers.Tables("SerialNumbers").DefaultView

            'sort the dataview
            'DVSN.Sort = "SerialNumber"

            'find location of serial numbers
            Dim FName As String = ""

            Dim i As Integer = DSSerialNumbers.Tables(0).Rows.Count

            For Each r As DataRow In DSSerialNumbers.Tables(0).Rows

                Label3.Text = "Retrieving data Database - " & i & " Serial Numbers remaining . . ."
                Label3.Refresh()

                i -= 1

                'get the data from temptestsetupdata required for velocity head and tdh calculations
                AR.SQL = "Select Date, SuctDiam, DischDiam, SuctionGageHeight, DischargeGageHeight, HDCor from TempTestSetupData WHERE TempTestSetupData.SerialNumber = '" & r("SerialNumber") & "' ORDER BY TempTestSetupData.Date;"
                Dim dsTestSetupData As DataSet = AR.GetPolarQueryData("TestSetupData")

                'the last row will be the latest date
                Dim DataOK As Boolean = False
                Dim lastRow As Integer = dsTestSetupData.Tables(0).Rows.Count - 1
                Dim TestDate As Date
                Dim dsTestData As DataSet

                If lastRow >= 0 Then
                    'make sure all flows are not 0.  if they are, get date before this date
                    While DataOK = False
                        TestDate = dsTestSetupData.Tables(0).Rows(lastRow)(0)

                        'get data from temptestdata
                        AR.SQL = "Select Flow, TemperatureSuction, SuctionPressure, DischargePressure, SuctionInHg from TempTestData WHERE SerialNumber = '" & r("SerialNumber") & "' AND DATE = #" & TestDate & "#;"
                        dsTestData = AR.GetPolarQueryData("TestData")

                        For Each rTest As DataRow In dsTestData.Tables(0).Rows
                            If rTest("Flow") <> 0 Then
                                DataOK = True
                            End If
                        Next

                        If lastRow = 0 Then
                            Exit While
                        Else
                            lastRow -= 1
                        End If
                    End While

                    If DataOK Then

                        Dim VelHead(7) As Double
                        Dim TDH(7) As Double

                        'clear coefficients and set degree to 3
                        TDHCoeff.Init()
                        TDHCoeff.Degree = 3

                        'calculate velocity head and TDH
                        Dim j As Integer = 0
                        For Each dr As DataRow In dsTestData.Tables(0).Rows
                            Dim disr As DataRow() = DischargeDiameter.Tables(0).Select("DischargeDiameter = " & dsTestSetupData.Tables(0)(lastRow)("DischDiam"))
                            Dim ActualDisc As Single = disr(0)("Description")
                            Dim sucr As DataRow() = DischargeDiameter.Tables(0).Select("DischargeDiameter = " & dsTestSetupData.Tables(0)(lastRow)("SuctDiam"))
                            Dim ActualSuct As Single = sucr(0)("Description")
                            VelHead(j) = AR.CalcVelHead(dr("Flow"), ActualDisc, ActualSuct, PipeDiameters)
                            TDH(j) = AR.CalcTDH(dr("DischargePressure"), dr("SuctionPressure"), dr("SuctionInHg"), VelHead(j), dr("TemperatureSuction"), dsTestSetupData.Tables(0)(lastRow)("HDCor") + ((dsTestSetupData.Tables(0)(lastRow)("DischargeGageHeight") / 12) - (dsTestSetupData.Tables(0)(lastRow)("SuctionGageHeight") / 12)), TemperatureCorrection)
                            TDHCoeff.XYAdd(dr("Flow"), TDH(j))
                            j += 1
                        Next

                        'get design flow for this pump
                        AR.SQL = "SELECT DesignFlow, DesignTDH FROM TempPumpData WHERE SerialNumber = '" & r("SerialNumber") & "'"
                        Dim dsDesign As DataSet = AR.GetPolarQueryData("Design")
                        '                Dim DesignFlow As Single = dsDesign.Tables(0).Rows(0)("DesignFlow")
                        Dim DesignFlow As Single = CSng(txtDesignFlow.Text)

                        Dim DesignHead As Single
                        If txtDesignHead.Text = "" Then
                            DesignHead = dsDesign.Tables(0).Rows(0)("DesignTDH")
                        Else
                            DesignHead = CSng(txtDesignHead.Text)
                        End If

                        'calculate 3rd order coefficients for TDH vs Flow
                        Dim TDHDesign As Single = TDHCoeff.Coeff(3) * DesignFlow ^ 3 + TDHCoeff.Coeff(2) * DesignFlow ^ 2 + TDHCoeff.Coeff(1) * DesignFlow + TDHCoeff.Coeff(0)

                        r("Design Head") = DesignHead
                        r("WaterZero") = TDHCoeff.Coeff(0)
                        r("WaterDesign") = TDHDesign
                    Else    'not data ok, make readings 0
                        r("Design Head") = 0
                        r("WaterZero") = 0
                        r("WaterDesign") = CSng(txtDesignHead.Text)
                    End If

                End If


            Next
        'End If

        'open template and save array
        Dim TemplateFile As Workbook
        TemplateFile = ExcelRoutn.CopyRunChartTemplate()

        Dim DV As DataView = DSSerialNumbers.Tables(0).DefaultView

        ExcelRoutn.WriteDataToExcel(ExcelRoutn.MetricDataViewToArray(DV), "A4", MultipleModels)

        'plot charts

        Label3.Text = "Plotting Charts . . ."
        Label3.Refresh()

        If Not MultipleModels Then
            ExcelRoutn.PlotCharts(DSSerialNumbers.Tables(0).Rows.Count, txtDesignHead.Text <> "", txtDesignFlow.Text, txtDesignHead.Text, SelectedModelNumber, SelectedImpellerSize, MultipleModels)
        Else
            ExcelRoutn.PlotCharts(DSSerialNumbers.Tables(0).Rows.Count, txtDesignHead.Text <> "", txtDesignFlow.Text, txtDesignHead.Text, "Multiple Models", SelectedImpellerSize, MultipleModels)
        End If
        ExcelRoutn.CloseExcel()



    End Sub

    Private Sub btnGetModelNo_Click(sender As Object, e As EventArgs) Handles btnGetModelNo.Click
        '        Dim AR As New AccessRoutines

        'set SQL string for query
        AR.SQL = "Select ModelNumber, count(ModelNumber) As CountOf from TempPumpData group by ModelNumber;"

        'get the dataset from the query
        DSModelNumbers = AR.GetPolarQueryData("ModelNumbers")
        DSModelNumbers.Tables("ModelNumbers").Columns.Add("ModelAndCount", Type.GetType("System.String"))
        For Each r As DataRow In DSModelNumbers.Tables("ModelNumbers").Rows
            Dim r1 As DataRow() = SuperMarketPumps.Tables(0).Select("Model = '" & r("ModelNumber") & "'")
            If r1.Length = 0 Then
                r("ModelAndCount") = r("ModelNumber") & "  (" & r("CountOf") & ")"
            Else
                r("ModelAndCount") = r("ModelNumber") & "  (" & r("CountOf") & ") *"
            End If
        Next

        'make a dataview, so we can easily sort
        Dim DV As DataView = DSModelNumbers.Tables("ModelNumbers").DefaultView

        'sort the dataview
        DV.Sort = "ModelNumber"


        lbSN.DisplayMember = "ModelAndCount"
        lbSN.ValueMember = "ModelNumber"
        lbSN.DataSource = DV
        lbSN.SelectedIndex = -1

        AddHandler lbSN.SelectedIndexChanged, AddressOf lbSN_SelectedIndexChanged

        cmbImpellers.Text = ""

        'show it
        'With DataGridView1
        '    .AutoGenerateColumns = True
        '    .DataSource = DV
        'End With

        '        Dim ans As Integer = MsgBox("Write Data To Excel?", vbYesNo, "Write Data To Excel")
        '       If ans = MsgBoxResult.No Then Exit Sub

    End Sub

    Private Sub lbSN_SelectedIndexChanged(sender As Object, e As EventArgs) ' Handles cmbSN.SelectedIndexChanged

        'see if there are multiple selections
        Dim ModelNo As String = ""
        Dim lInt As List(Of String) = New List(Of String)

        If lbSN.SelectedItems.Count < 2 Then

            ModelNo = lbSN.GetItemText(lbSN.SelectedItem)
            SelectedModelNumber = Trim(ModelNo.Substring(0, ModelNo.IndexOf("(") - 1))
            TextBox1.Text = SelectedModelNumber
            btnGetImpellers_Click(sender, e)

            'if we only have one entry selected, clear the intersection list
            If lbSN.SelectedItems.Count < 2 Then
                listImp.Clear()
            End If

            lblDesignFlow.Visible = True
            If ModelNo.Contains("*") Then   'supermarket pump, enter design flow in box
                Dim r1 As DataRow() = SuperMarketPumps.Tables(0).Select("Model = '" & SelectedModelNumber & "'")
                txtDesignFlow.Text = r1(0)("DesignFlow").ToString
                txtDesignHead.Text = r1(0)("DesignTDH").ToString
                txtDesignHead.Visible = True
                lblDesignHead.Visible = True
                lblDesignFlow.Text = "Flow from Supermarket Table"
                btnCalcRunCharts.Enabled = True
            Else
                btnCalcRunCharts.Enabled = False
                lblDesignFlow.Text = "Please Enter Desired Design Flow"
                txtDesignFlow.Text = ""
                txtDesignHead.Text = ""
                txtDesignHead.Visible = False
                lblDesignHead.Visible = False


            End If
        Else    'if multiple items selected

            For Each item In lbSN.SelectedItems

                ModelNo = lbSN.GetItemText(item)
                SelectedModelNumber = Trim(ModelNo.Substring(0, ModelNo.IndexOf("(") - 1))
                TextBox1.Text = "Multiple Selections"
                txtDesignFlow.Text = ""
                txtDesignHead.Text = ""
                lblDesignFlow.Text = ""
                btnGetImpellers_Click(sender, e)

                Dim l As List(Of String) = New List(Of String)


                'fill l with the impeller diameters
                For Each r As DataRow In DSImpellerSizes.Tables(0).AsEnumerable
                    l.Add(r("ImpellerDia"))
                Next

                Try
                    'If this Is the first selection (or only one), fill the intersection list with the impeller list
                    If listImp.Count = 0 Then
                        listImp = l
                        lInt = l
                    Else
                        'else find the intersection between the intersection list with the new impeller list
                        lInt = listImp.Intersect(l).ToList
                    End If
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try

                If lInt.Count = 0 And lbSN.SelectedItems.Count > 1 Then
                    MsgBox("There are no impeller sizes common to your selections.")
                End If


            Next

            cmbImpellers.DataSource = lInt

        End If 'end of only one item selected

    End Sub

    Private Sub btnGetImpellers_Click(sender As Object, e As EventArgs) Handles btnGetImpellers.Click

        'set SQL string for query
        AR.SQL = "Select ImpellerDia, count(ImpellerDia) As CountOf from TempPumpData WHERE ModelNumber = '" & SelectedModelNumber & "' group by ImpellerDia;" ' WHERE ModelNumber = '" & SelectedModelNumber & "';"

        'get the dataset from the query
        DSImpellerSizes = AR.GetPolarQueryData("ImpellerSizes")
        DSImpellerSizes.Tables("ImpellerSizes").Columns.Add("ImpSizeAndCount", Type.GetType("System.String"))
        For Each r As DataRow In DSImpellerSizes.Tables("ImpellerSizes").Rows
            r("ImpSizeAndCount") = r("ImpellerDia") & "  (" & r("CountOf") & ")"
        Next

        'make a dataview, so we can easily sort
        Dim DVImpellerSize As DataView = DSImpellerSizes.Tables("ImpellerSizes").DefaultView

        'sort the dataview
        DVImpellerSize.Sort = "ImpellerDia"

        cmbImpellers.DataSource = DVImpellerSize
        cmbImpellers.ValueMember = "ImpellerDia"
        cmbImpellers.DisplayMember = "ImpSizeAndCount"

        AddHandler cmbImpellers.SelectedIndexChanged, AddressOf cmbImpellers_SelectedIndexChanged

        cmbImpellers.SelectedIndex = 0
        TextBox2.Text = cmbImpellers.SelectedValue

    End Sub

    Private Sub cmbImpellers_SelectedIndexChanged(sender As Object, e As EventArgs) ' Handles cmbImpellers.SelectedIndexChanged

        Dim ImpDia As String = cmbImpellers.GetItemText(cmbImpellers.SelectedItem)
        If ImpDia.Contains("(") Then
            SelectedImpellerSize = ImpDia.Substring(0, ImpDia.IndexOf("(") - 1)
        Else
            SelectedImpellerSize = ImpDia
        End If
        TextBox2.Text = SelectedImpellerSize

    End Sub

    Private Sub cmbSN_MouseClick(sender As Object, e As MouseEventArgs) Handles cmbSN.MouseClick
        If Not DSImpellerSizes Is Nothing Then
            DSImpellerSizes.Tables(0).Rows.Clear()
        End If
    End Sub


    Private Sub txtDesignFlow_TextChanged(sender As Object, e As EventArgs) Handles txtDesignFlow.TextChanged
        If IsNumeric(txtDesignFlow.Text) Then
            btnCalcRunCharts.Enabled = True
        End If
    End Sub


End Class
