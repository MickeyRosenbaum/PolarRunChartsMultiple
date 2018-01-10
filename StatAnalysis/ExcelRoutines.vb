Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO
Public Class ExcelRoutines

    Public xlApp As Excel.Application
    Public xlWorkbook As Excel.Workbook
    Public xlWorkSheet As Excel.Worksheet
    Public xlChart As Excel.Chart

    Public SaveFileName As String
    Public WorksheetName As String

    Private mWB As Excel.Workbook
    Public Property WB() As Excel.Workbook
        Get
            Return mWB
        End Get
        Set(ByVal value As Excel.Workbook)
            mWB = value
        End Set
    End Property
    Private mWS As Excel.Worksheet
    Public Property WS() As Excel.Worksheet
        Get
            Return mWS
        End Get
        Set(ByVal value As Excel.Worksheet)
            mWS = value
        End Set
    End Property

    Public Sub closeOpenExcelApps()
        Dim Proc() As Process = Process.GetProcessesByName("Excel")
        If Proc.Count > 0 Then
            Dim ans As Integer = MsgBox("Excel is running. Is it ok if I close it?", MsgBoxStyle.YesNo, "Excel is already running")
            If ans = MsgBoxResult.No Then
                MsgBox("Exiting...", MsgBoxStyle.OkOnly, "Exiting Program")
                End
            End If
            For Each p As Process In Proc
                Try
                    p.Kill()
                Catch
                End Try
            Next
        End If
    End Sub
    Public Function CopyRunChartTemplate() As Excel.Workbook
        xlApp = New Excel.Application
        Dim sfd As New SaveFileDialog
        sfd.Title = "Save Run Chart"
        sfd.Filter = "Excel Files |*.xlsx"

        'If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\PolarQuery", "CellMetricsDirectory", Nothing) = "" Then
        '    sfd.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
        'Else
        '    sfd.InitialDirectory = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\PolarQuery", "CellMetricsDirectory", Nothing)
        'End If
        'sfd.CheckFileExists = False
        'sfd.OverwritePrompt = True

        If sfd.ShowDialog = DialogResult.OK Then
            SaveFileName = sfd.FileName
        End If
        sfd.FileName = ""
        sfd.Dispose()


        Dim xlTemplateName As String = "\\tei-main-01\F\Groups\SHARED\Databases\RunChartTemplate.xlsx"

        If File.Exists(xlTemplateName) Then
            System.IO.File.Copy(xlTemplateName, SaveFileName, True)
        Else
            Console.WriteLine("{0} does not exist", xlTemplateName)
        End If

        xlApp.Workbooks.Open(SaveFileName)

        xlApp.ActiveWorkbook.Save()

        Return xlApp.ActiveWorkbook

    End Function
    Public Function OpenHydraulicTest(FileName As String) As Array

        closeOpenExcelApps()

        Dim Values(4) As Single 'design, shutoff water, design water, shutoff fluid, design fluid

        Dim xlApp1 As Excel.Application
        Dim xlWorkbook1 As Excel.Workbook
        Dim xlWorkSheet1 As Excel.Worksheet

        xlApp1 = New Excel.Application
        'xlApp1.Visible = True

        xlWorkbook1 = xlApp1.Workbooks.Open(FileName)

        Dim i As Integer = xlWorkbook1.Worksheets.Count
        xlWorkSheet1 = xlWorkbook1.Worksheets(i)

        Dim currentFind As Excel.Range = Nothing
        Dim firstFind As Excel.Range = Nothing

        'Water Performance Trendlines								
        'Customer's Fluid Performance Trendlines								
        '@ Design Flow					


        currentFind = xlWorkSheet1.Cells.Find("Design Head", , Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, False)
        If Val(xlWorkSheet1.Cells(currentFind.Row, currentFind.Column + 2).Value) = 0 Then
            Values(0) = xlWorkSheet1.Cells(currentFind.Row, currentFind.Column + 3).Value
        Else
            Values(0) = xlWorkSheet1.Cells(currentFind.Row, currentFind.Column + 2).Value
        End If

        currentFind = xlWorkSheet1.Cells.Find("Water Performance Trendlines", , Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, False)
        Values(1) = xlWorkSheet1.Cells(currentFind.MergeArea.Row + 3, currentFind.MergeArea.Column + currentFind.MergeArea.Columns.Count - 1).Value

        currentFind = xlWorkSheet1.Cells.Find("Customer's Fluid Performance Trendlines", , Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, False)
        Values(2) = xlWorkSheet1.Cells(currentFind.MergeArea.Row + 3, currentFind.MergeArea.Column + currentFind.MergeArea.Columns.Count - 1).Value

        currentFind = xlWorkSheet1.Cells.Find("@ Design Flow", , Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, False)
        Values(3) = xlWorkSheet1.Cells(currentFind.Row + 2, currentFind.Column + 1).Value

        currentFind = xlWorkSheet1.Cells.Find("@ Design Flow", , Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, False)
        Values(4) = xlWorkSheet1.Cells(currentFind.Row + 3, currentFind.Column + 1).Value


        xlWorkbook1.Close(False)

        xlApp1.Quit()

        releaseObject(xlWorkbook1)
        releaseObject(xlApp1)

        xlApp1 = Nothing
        xlWorkbook1 = Nothing
        xlWorkSheet1 = Nothing

        Return Values
    End Function
    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
            Try
                'Dim MSExcelControl() As Process
                Dim iID As Integer
                Dim lastOpen As DateTime
                Dim obj1(10) As Process
                obj1 = Process.GetProcessesByName("EXCEL")
                lastOpen = obj1(0).StartTime
                For Each p As Process In obj1
                    p.Kill()
                Next

            Catch ex As Exception

            End Try

        End Try
    End Sub

    Public Sub OpenExcelWorkbook(NoTabs As Boolean)
        xlApp = New Excel.Application
        Dim ofd As New OpenFileDialog
        ofd.Title = "Open Excel Files"
        ofd.Filter = "Excel Files |*.xls;*.xlsx"

        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\PolarStatisticalAnalysis", "Directory", Nothing) = "" Then
            ofd.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
        Else
            ofd.InitialDirectory = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\PolarStatisticalAnalysis", "Directory", Nothing)
        End If
        ofd.CheckFileExists = False

        If ofd.ShowDialog() = DialogResult.OK Then
            SaveFileName = ofd.FileName
        End If

        If Dir(ofd.FileName) = "" Then
            SaveFileName = ofd.FileName
            If Not xlApp.Workbooks Is Nothing Then
                'close any open workbooks
                xlApp.Workbooks.Close()
            End If

            'create the workbook
            xlWorkbook = xlApp.Workbooks.Add
            WorksheetName = NewWorkBook()
            xlApp.ActiveWorkbook.SaveAs(SaveFileName, Excel.XlFileFormat.xlOpenXMLWorkbook)

        Else    'if the file name already exists
            SaveFileName = ofd.FileName
            If Not xlApp.Workbooks Is Nothing Then
                'close any open workbooks
                xlApp.Workbooks.Close()
            End If
            If SaveFileName = "" Then
                MsgBox("File not selected.  Exiting . . .", vbOKOnly, "File not selected.")
                End
            End If
            xlApp.Workbooks.Open(SaveFileName)

            If Not NoTabs Then
                If GetWorksheetTabs(SaveFileName, WorksheetName) = vbNo Then    'ask the user if he/she wants a new tab.
                    MsgBox("File not overwritten. Exiting...", vbOKOnly, "File not Opened")
                    End
                End If
            Else
                WorksheetName = "Summary"
            End If

        End If

        xlApp.ActiveWorkbook.Save()
        xlWorkSheet = xlApp.ActiveWorkbook.Worksheets(WorksheetName)

        If SaveFileName <> "" Then
            If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\PolarStatisticalAnalysis", "Directory", Nothing) Is Nothing Then
                My.Computer.Registry.CurrentUser.CreateSubKey("PolarStatisticalAnalysis")
            End If
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\PolarStatisticalAnalysis", "Directory", Path.GetDirectoryName(SaveFileName))
        End If
    End Sub
    Public Function SelectWS(WSName As String) As Excel.Worksheet
        xlWorkSheet = xlApp.ActiveWorkbook.Worksheets(WSName)
        Return xlWorkSheet
    End Function
    Public Function SelectChart(ChartName As String) As Excel.Chart
        xlChart = xlApp.ActiveWorkbook.Charts(ChartName)
        Return xlChart
    End Function
    Function NewWorkBook() As String
        Dim WorkSheetName As String

        'we've just added a new workbook, delete sheet1, sheet2, etc
        xlApp.DisplayAlerts = False
        While xlApp.Worksheets.Count > 1
            xlApp.Worksheets(1).Delete          'delete the sheet
        End While
        xlApp.DisplayAlerts = True

        WorkSheetName = InputBox("Enter the Worksheet Name for this run.")    'get the desired name
        xlApp.Worksheets(1).Name = WorkSheetName    'and name the sheet

        NewWorkBook = WorkSheetName
    End Function

    Function GetWorksheetTabs(filename As String, ByRef WorkSheetName As String)

        'see what worksheet tabs alread exist in the excel worksheet

        Dim intSheets As Integer    'number of sheets in the workbook
        Dim I As Integer
        Dim S As String
        Dim ans
        Dim NameOK As Boolean

        intSheets = xlApp.Worksheets.Count      'how many sheets are there?

        'define a crlf string
        S = vbCrLf

        For I = 1 To intSheets
            S = S & xlApp.Worksheets(I).Name & vbCrLf   'add in the worksheet name
        Next I

        'tell the user the names so far and ask if he/she wants to add another
        ans = MsgBox("You have the following Worksheet Names in " & filename & ": " & S & "Do you want to add another sheet to this file?", vbYesNo, "Sheets in Excel File")

        'get the answer
        If ans = vbNo Then
            GetWorksheetTabs = vbNo     'set up flag for when we return to the calling subroutine
            Exit Function
        End If

        'get worksheet name from user and check to see that it's not already used

        NameOK = False  'start assuming that the name is bad

        While Not NameOK    'as long as it's bad, stay in this loop
            WorkSheetName = InputBox("Enter Worksheet Name for this run.")  'ask for name

            If WorkSheetName = "" Then      'if we get a nul return or user presses cancel
                GetWorksheetTabs = vbNo
                Exit Function
            End If

            For I = 1 To xlApp.Worksheets.Count     'go through all of the existing sheets
                If WorkSheetName = xlApp.Worksheets(I).Name Then        'if the names are the same
                    MsgBox("The name " & WorkSheetName & " already exists for a Worksheet.  Please try again.", vbOKOnly, "Worksheet name already exists")  'tell the user
                    NameOK = False
                    Exit For
                End If
                NameOK = True       'if we make it thru say the name is ok
            Next I
        End While

        xlApp.Worksheets.Add(After:=xlApp.Worksheets(xlApp.Worksheets.Count))      'add a worksheer
        xlApp.Worksheets(xlApp.Worksheets.Count).Name = WorkSheetName       'give it the desired name
        GetWorksheetTabs = vbYes                                            'say that the results were ok

    End Function
    'Public Function DataViewToArray(DV As DataView) As Array
    '    Dim DT As DataTable = New DataTable
    '    DT = DV.ToTable

    '    Dim arr(DT.Rows.Count - 1, DT.Columns.Count - 1)
    '    'For c As Integer = 0 To DT.Columns.Count - 1
    '    '    arr(0, c) = DT.Columns(c).ColumnName
    '    'Next
    '    For r As Integer = 1 To DT.Rows.Count
    '        Dim dr As DataRow = DT.Rows(r - 1)
    '        For c As Integer = 0 To DT.Columns.Count - 1
    '            arr(r, c) = dr(c)
    '        Next
    '    Next
    '    Return arr
    'End Function
    Public Function MetricDataViewToArray(DV As DataView) As Array
        Dim DT As DataTable = New DataTable
        DV.RowFilter = ""
        DT = DV.ToTable

        Dim arr(DT.Rows.Count - 1, DT.Columns.Count - 1)
        For r As Integer = 0 To DT.Rows.Count - 1
            Dim dr As DataRow = DT.Rows(r)
            For c As Integer = 0 To DT.Columns.Count - 1
                arr(r, c) = dr(c)
            Next
        Next
        Return arr
    End Function
    Public Sub WriteDataToExcel(arr As Array, UpperLeft As String, MultipleModels As Boolean)
        'arr has arr.getlength(0) rows, and arr.getlength(1) columns 

        Dim endrange As Integer = 4 + arr.GetLength(0) - 1

        With xlApp.Range(UpperLeft)
            .ClearContents()
            xlApp.Range(UpperLeft).Resize(arr.GetLength(0), arr.GetLength(1)).Name = "MyNamedRange"
        End With

        xlApp.Range("MyNamedRange").Value2 = arr
        ' xlApp.Range("A1:A1").Value2 = "Query run " + Now

        If MultipleModels Then
            Dim sheet = xlApp.Worksheets("Summary")
            Dim destRange As Excel.Range
            Dim sourceRange As Excel.Range = xlApp.Range("E4:E1000")
            destRange = xlApp.Range("M4")
            sourceRange.Copy(destRange)
        End If

        xlApp.Columns("C:C").Select
        xlApp.Selection.NumberFormat = "m/d/yyyy h:mm"

        xlApp.Range("E4").Formula = "=Average($D$4:$D$" & 4 + arr.GetLength(0) - 1 & ")"
        xlApp.Range("F4").Formula = "=StDev($D$4:$D$" & 4 + arr.GetLength(0) - 1 & ")"
        '        xlApp.Range("O4").Formula = "=Average($F$4:$F$" & 4 + arr.GetLength(0) - 1 & ")"
        '        xlApp.Range("P4").Formula = "=StDev($F$4:$F$" & 4 + arr.GetLength(0) - 1 & ")"

        xlApp.Range("E4:L4").Select()
        xlApp.Selection.Copy

        xlApp.Range("E5:E" & endrange).Select()
        xlApp.ActiveSheet.Paste

        xlApp.Columns("C:L").Select
        xlApp.Selection.NumberFormat = "0.00"

    End Sub


    Public Sub PlotCharts(NumberOrRows As Integer, SuperMarketPump As Boolean, DesignFlow As String, DesignHead As String, ModelNo As String, Impeller As String, MultipleModels As Boolean)
        'plot a chart on a predetermined chart page
        'enter with the data sheet (and chart sheet) name, upperleft cell for data, and axis title

        With xlApp
            .Sheets("Run Chart - Water").Select
            .ActiveChart.ChartArea.Select()
            .ActiveChart.PlotArea.Select()
            .ActiveChart.SeriesCollection.NewSeries
            .ActiveChart.SeriesCollection(1).Name = "=""Water Perf"""
            .ActiveChart.SeriesCollection(1).Values = "=Summary!$D$4:$D$" & NumberOrRows + 4
            .ActiveChart.SeriesCollection(1).XValues = "=Summary!$A$4:$A$" & NumberOrRows + 4

            .ActiveChart.ChartArea.Select()
            .ActiveChart.PlotArea.Select()
            .ActiveChart.SeriesCollection.NewSeries
            .ActiveChart.SeriesCollection(2).Name = "=""+1 StdDev"""
            .ActiveChart.SeriesCollection(2).Values = "=Summary!$I$4:$I$" & NumberOrRows + 4
            .ActiveChart.SeriesCollection(2).XValues = "=Summary!$A$4:$A$" & NumberOrRows + 4

            .ActiveChart.ChartArea.Select()
            .ActiveChart.SeriesCollection(2).Select
            .Selection.MarkerStyle = -4142

            .ActiveChart.ChartArea.Select()
            .ActiveChart.PlotArea.Select()
            .ActiveChart.SeriesCollection.NewSeries
            .ActiveChart.SeriesCollection(3).Name = "=""-1 StdDev"""
            .ActiveChart.SeriesCollection(3).Values = "=Summary!$L$4:$L$" & NumberOrRows + 4
            .ActiveChart.SeriesCollection(3).XValues = "=Summary!$A$4:$A$" & NumberOrRows + 4

            .ActiveChart.ChartArea.Select()
            .ActiveChart.SeriesCollection(2).Select
            .Selection.MarkerStyle = -4142
            .ActiveChart.SeriesCollection(3).Select
            .Selection.MarkerStyle = -4142

            If SuperMarketPump Then
                .ActiveChart.ChartArea.Select()
                .ActiveChart.SeriesCollection.NewSeries
                .ActiveChart.SeriesCollection(4).Name = "=""Design"""
                .ActiveChart.SeriesCollection(4).Values = "=Summary!$B$4:$B$" & NumberOrRows + 4
                .ActiveChart.SeriesCollection(4).XValues = "=Summary!$A$4:$A$" & NumberOrRows + 4
                .ActiveChart.ChartArea.Select()
                .ActiveChart.SeriesCollection(4).Select
                .Selection.MarkerStyle = -4142
            End If


            If Val(DesignFlow) <> 0 And Not SuperMarketPump Then
                xlApp.Worksheets("Summary").Name = "Summary - Flow = " & DesignFlow
            End If


            Dim sChartTitle As String = "Model No. = " & ModelNo & " -- Impeller = " & Impeller & vbCrLf & "Flow = " & DesignFlow.ToString
            If DesignHead <> "" Then
                sChartTitle += " -- Head = " & DesignHead
            End If


            .ActiveChart.SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementChartTitleAboveChart)
            .ActiveChart.ChartTitle.Characters.Text = sChartTitle
            .ActiveChart.SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
            .ActiveChart.Axes(1).axistitle.caption = "Pump Serial Number"
            .ActiveChart.SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementPrimaryValueAxisTitleRotated)
            .ActiveChart.Axes(2).axistitle.caption = "TDH(ft)"

            .ActiveChart.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextDirection.msoTextDirectionLeftToRight, 500, 464, 150, 20).Select()
            .Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "Date Generated: " & Today().Date.ToShortDateString
            .Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 13).ParagraphFormat.FirstLineIndent = 0
            With .Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 13).Font
                .NameComplexScript = "+mn-cs"
                .NameFarEast = "+mn-ea"
                .Size = 11
                .Name = "+mn-lt"
            End With
            'With Selection.ShapeRange.Line
            '    .Visible = msoTrue
            '    .ForeColor.ObjectThemeColor = msoThemeColorAccent1
            '    .ForeColor.TintAndShade = 0
            '    .ForeColor.Brightness = 0
            'End With
            With .Selection.ShapeRange.Line
                .Visible = True
                '.ForeColor.ObjectThemeColor = msoThemeColorText1
                .ForeColor.TintAndShade = 0
                .ForeColor.Brightness = 0
                .Transparency = 0
            End With

            '.ActiveChart.SeriesCollection(1).Points(2).MarkerForegroundColor = RGB(0, 255, 0)
            If MultipleModels Then
                'label first datapoint
                Dim model As String = ""
                For dl As Integer = 1 To .ActiveChart.SeriesCollection(1).points.count
                    .ActiveChart.SeriesCollection(1).applydatalabels
                    If model <> xlApp.Worksheets(1).Range("M" & 3 + dl).Value Then
                        model = xlApp.Worksheets(1).Range("M" & 3 + dl).Value
                        .ActiveChart.SeriesCollection(1).datalabels(dl).text = xlApp.Worksheets(1).Range("M" & 3 + dl).Value
                    Else
                        .ActiveChart.SeriesCollection(1).datalabels(dl).text = ""
                    End If
                    .ActiveChart.SeriesCollection(1).datalabels(dl).position = Microsoft.Office.Interop.Excel.XlDataLabelPosition.xlLabelPositionAbove
                Next dl
            End If


            '           .ActiveChart.SeriesCollection(1).datalabels(1).hasdatalabel = True

            '.Sheets("Run Chart - Fluid").Select
            '.ActiveChart.ChartArea.Select()
            '.ActiveChart.SeriesCollection.NewSeries
            '.ActiveChart.SeriesCollection(1).Name = "=""Design"""
            '.ActiveChart.SeriesCollection(1).Values = "=Summary!$B$4:$B$" & NumberOrRows + 4
            '.ActiveChart.SeriesCollection(1).XValues = "=Summary!$A$4:$A$" & NumberOrRows + 4

            '.ActiveChart.ChartArea.Select()
            '.ActiveChart.PlotArea.Select()
            '.ActiveChart.SeriesCollection.NewSeries
            '.ActiveChart.SeriesCollection(2).Name = "=""Fluid Perf"""
            '.ActiveChart.SeriesCollection(2).Values = "=Summary!$F$4:$F$" & NumberOrRows + 4
            '.ActiveChart.SeriesCollection(2).XValues = "=Summary!$A$4:$A$" & NumberOrRows + 4

            '.ActiveChart.ChartArea.Select()
            '.ActiveChart.PlotArea.Select()
            '.ActiveChart.SeriesCollection.NewSeries
            '.ActiveChart.SeriesCollection(3).Name = "=""+1 StdDev"""
            '.ActiveChart.SeriesCollection(3).Values = "=Summary!$Q$4:$Q$" & NumberOrRows + 4
            '.ActiveChart.SeriesCollection(3).XValues = "=Summary!$A$4:$A$" & NumberOrRows + 4

            '.ActiveChart.ChartArea.Select()
            '.ActiveChart.PlotArea.Select()
            '.ActiveChart.SeriesCollection.NewSeries
            '.ActiveChart.SeriesCollection(4).Name = "=""-1 StdDev"""
            '.ActiveChart.SeriesCollection(4).Values = "=Summary!$T$4:$T$" & NumberOrRows + 4
            '.ActiveChart.SeriesCollection(4).XValues = "=Summary!$A$4:$A$" & NumberOrRows + 4
        End With


        ''select data sheet
        'WS = SelectWS(WSName)
        'With WS
        '    .Activate()
        '    .Range("G4").Activate()
        '    Dim lastrow As Excel.Range = xlApp.Rows.End(Excel.XlDirection.xlDown)
        '    Dim r As Excel.Range = xlApp.Range("G4" & lastrow.Row)
        '    r.Copy()
        'End With

        'Dim Cht As Excel.Chart = SelectChart("Run Chart - Water")
        'With Cht
        '    .Activate()
        '    .ChartArea.Select()
        '    .Paste()
        '    .ChartArea.Select()
        '    .SeriesCollection(1).Select
        '    .SeriesCollection(1).ChartType = Excel.XlChartType.xlLineMarkers
        '    .HasLegend = False
        '    .HasTitle = True
        '    .ChartTitle.Characters.Text = WSName
        '    .Axes(Excel.XlAxisType.xlCategory).HasTitle = True
        '    .Axes(Excel.XlAxisType.xlCategory).axistitle.Text = "Serial Number"

        '    ''if there are two series, make 2nd one on 2nd axis and label
        '    'Dim series As Excel.Series
        '    'Dim seriescollection As Excel.SeriesCollection = CType(Cht.SeriesCollection, Excel.SeriesCollection)
        '    'Dim i As Long = seriescollection.Count
        '    'If i = 2 Then
        '    '    series = seriescollection.Item(2)
        '    '    series.AxisGroup = 2
        '    '    With .Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary)
        '    '        .hastitle = True
        '    '        .axistitle.Text = "Pump Count"
        '    '        .hastitle = True
        '    '        .axistitle.Text = "Ave No of Tests"
        '    '    End With
        '    'End If
        'End With

    End Sub
    Public Sub CloseExcel()
        xlApp.ActiveWorkbook.Save()
        xlApp.Workbooks.Close()
        MsgBox("Query Data Saved in " + SaveFileName + " Exiting...", vbOKOnly, "Query Data Saved")
        xlApp = Nothing
        Dim Proc() As Process = Process.GetProcessesByName("Excel")
        For Each p As Process In Proc
            Try
                p.Kill()
            Catch
            End Try
        Next
        End

    End Sub
End Class
