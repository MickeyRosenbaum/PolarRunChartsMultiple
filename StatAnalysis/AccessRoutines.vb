Imports System.Data.Common
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data.OleDb


Public Class AccessRoutines
    Dim FACTORY As DbProviderFactory
    Dim CONN As DbConnection
    Dim CMD As DbCommand
    Dim DA As DbDataAdapter
    Dim DS As DataSet

    'Column number constants
    Public Const IDColNo As Integer = 0
    Public Const NominalColNo As Integer = 1
    Public Const ActualColNo As Integer = 2
    Public Const TempColNo As Integer = 1
    Public Const VaporPressureColNo As Integer = 2
    Public Const SpecificVolumeColNo As Integer = 3
    Public Const TDHColNo As Integer = 3

    'set connection string for the database
    Private mConnectionString As String
    Public Property ConnectionString() As String
        Get
            Return mConnectionString
        End Get
        Set(ByVal value As String)
            mConnectionString = value
        End Set
    End Property

    'set the SQL string for query
    Private mSQL As String
    Public Property SQL() As String
        Get
            Return mSQL
        End Get
        Set(ByVal value As String)
            mSQL = value
        End Set
    End Property

    'connect to the database
    Public Function DbConnect() As DbConnection
        'if we have a connection, then if it is closed, open it
        'then, return connection
        If CONN IsNot Nothing Then
            If CONN.State = ConnectionState.Closed Then
                CONN.Open()
            End If
            Return CONN
            Exit Function
        End If

        'if we have no connection, make one
        FACTORY = DbProviderFactories.GetFactory("System.Data.OleDb")
        CONN = FACTORY.CreateConnection()

        If ConnectionString = "" Then
            MsgBox("No database specified.", vbOKOnly, "No database specified")
            End
        End If

        If IsNothing(CONN) OrElse CONN.State = ConnectionState.Closed Then
            CONN.ConnectionString = ConnectionString
            CONN.Open()
        End If
        Return CONN
    End Function

    Public Function GetPolarQueryData(DataSetName As String) As DataSet
        'get the query from the database and return the dataset
        CONN = DbConnect()
        CMD = CONN.CreateCommand
        CMD.CommandType = CommandType.Text
        CMD.CommandText = SQL

        DA = FACTORY.CreateDataAdapter
        DA.SelectCommand = CMD

        DS = New DataSet
        DA.Fill(DS, DataSetName)
        CONN.Close()
        Return DS
    End Function
    Public Function GetUncSourcePath(ByVal driveLetter As Char) As String
        'find the UNC path from the mapped drive
        '  invoke net use and get remote name
        If String.IsNullOrEmpty(driveLetter) Then Return ""
        If (driveLetter < "a"c OrElse driveLetter > "z") AndAlso (driveLetter < "A"c OrElse driveLetter > "Z") Then Return ""
        Dim P As New Process()
        With P.StartInfo
            .FileName = "net"
            .Arguments = String.Format("use {0}:", driveLetter)
            .UseShellExecute = False
            .RedirectStandardOutput = True
            .CreateNoWindow = True
        End With
        P.Start()
        Dim T = P.StandardOutput.ReadToEnd()
        P.WaitForExit()
        For Each Line In Split(T, vbNewLine)
            If Line.StartsWith("Remote name") Then Return Line.Replace("Remote name", "").Trim()
        Next
        Return Nothing
    End Function

    Function CalcVelHead(Flow As Single, DischDiam As Single, SuctDiam As Single, PipeDiameters As DataSet) As Single
        CalcVelHead = 0
        If Not (DischDiam = 0 Or SuctDiam = 0) Then
            'lookup actual diameters in table
            Dim r As DataRow() = PipeDiameters.Tables(0).Select("NominalDia = " & DischDiam)
            Dim ActualDisc As Single = r(0)("ActualDia")
            r = PipeDiameters.Tables(0).Select("NominalDia = " & SuctDiam)
            Dim ActualSuct As Single = r(0)("ActualDia")

            If Not ((SuctDiam = -1 Or DischDiam = -1)) Then
                CalcVelHead = (0.00259 * Flow ^ 2) / (ActualDisc ^ 4) - (0.00259 * Flow ^ 2) / (ActualSuct ^ 4)
            End If
        End If
    End Function

    Function CalcTDH(DischargePressure As Single, SuctionPressure As Single, SuctionInHg As Single, VelHead As Single, SuctTemp As Single, HDCorr As Single, TempCorr As DataSet)
        If SuctTemp < 40 Or Val(SuctTemp) = 0 Then
            CalcTDH = 0
            Exit Function
        End If
        Dim r As DataRow() = TempCorr.Tables(0).Select("Temp = " & CInt(SuctTemp))
        Dim TempCorrect As Single = r(0)("TDHCorr")

        '    CalcTDH = (DischargePressure - CalculateSuctionPressure(SuctionPressure, SuctionInHg)) * 144 * DLookup("TDHCorr", "TempCorrection", "Temp = " & Int(SuctTemp)) + VelHead + HDCorr
        CalcTDH = (DischargePressure - CalculateSuctionPressure(SuctionPressure, SuctionInHg)) * 144 * TempCorrect + VelHead + HDCorr

    End Function

    Function CalculateSuctionPressure(SuctPress, SuctInHg)
        Dim sp As Single

        If (Not IsNumeric(SuctPress)) Then
            sp = 0
        Else
            sp = SuctPress
        End If

        CalculateSuctionPressure = sp - 0.4893 * SuctInHg
    End Function
    Function FillArrays(DataSetName As String) As DataSet
        CONN = DbConnect()
        CMD = CONN.CreateCommand
        CMD.CommandType = CommandType.Text
        CMD.CommandText = SQL

        DA = FACTORY.CreateDataAdapter
        DA.SelectCommand = CMD

        DS = New DataSet
        DA.Fill(DS, DataSetName)
        CONN.Close()
        Return DS

    End Function

End Class



