Public Class RegressionCoefficients
    Private Const MaximumOrder As Long = 25
    Private GlobalOrder As Long '"Ordnung" = degree of the polynom expected
    Private Finished As Boolean

    Private SumX(0 To 2 * MaximumOrder) As Double
    Private SumYX(0 To MaximumOrder) As Double
    Private M(0 To MaximumOrder, 0 To MaximumOrder + 1) As Double
    Private C(0 To MaximumOrder) As Double   'coefficients in: Y = C(0)*X^0 + C(1)*X^1 + C(2)*X^2 + ...

    Private Sub GaussSolve(O As Long)
        'gauss algorithm implementation,
        'following R.Sedgewick's "Algorithms in C", Addison-Wesley, with minor modifications
        'see http://en.wikipedia.org/wiki/Gauss%E2%80%93Jordan_elimination
        ' and http://www.arachnoid.com/sage/polynomial.html
        Dim i As Long, j As Long, k As Long, iMax As Long, T As Double, O1 As Double
        O1 = O + 1
        'first triangulize the matrix
        For i = 0 To O
            iMax = i : T = Math.Abs(M(iMax, i))
            For j = i + 1 To O 'find the line with the largest absvalue in this row
                If T < Math.Abs(M(j, i)) Then iMax = j : T = Math.Abs(M(iMax, i))
            Next j
            If i < iMax Then 'exchange the two lines
                For k = i To O1
                    T = M(i, k)
                    M(i, k) = M(iMax, k)
                    M(iMax, k) = T
                Next k
            End If
            For j = i + 1 To O 'scale all following lines to have a leading zero
                T = M(j, i) / M(i, i)
                M(j, i) = 0.0#
                For k = i + 1 To O1
                    M(j, k) = M(j, k) - M(i, k) * T
                Next k
            Next j
        Next i
        'then substitute the coefficients
        For j = O To 0 Step -1
            T = M(j, O1)
            For k = j + 1 To O
                T = T - M(j, k) * C(k)
            Next k
            C(j) = T / M(j, j)
        Next j
        Finished = True
    End Sub

    Private Sub BuildMatrix(Order As Long)
        'add a column for the Ys
        Dim Order1 As Long = Order + 1
        'copy the SumX to M
        For i As Integer = 0 To Order
            For k As Integer = 0 To Order
                M(i, k) = SumX(i + k)
            Next k
            'Add SumYX column
            M(i, Order1) = SumYX(i)
        Next i
    End Sub

    Private Sub FinalizeMatrix(Order As Long)
        Dim Order1 As Long = Order + 1
        For i As Integer = 0 To Order
            M(i, Order1) = SumYX(i)
        Next i
    End Sub

    Private Sub Solve()
        Dim Order As Long = GlobalOrder
        'adjust order if necessary for number of pairs so far
        If XYCount <= Order Then Order = XYCount - 1
        'trap on order of 0
        If Order < 0 Then Exit Sub

        BuildMatrix(Order)
        On Error Resume Next
        GaussSolve(Order)
        While (Err.Number <> 0) And (1 < Order)
            Err.Clear()
            C(0) = 0.0#
            Order = Order - 1
            FinalizeMatrix(Order)
        End While
        On Error GoTo 0
    End Sub

    Public Sub New()
        Init()
        GlobalOrder = 3 'third order as default
    End Sub

    Public Sub Init()
        Finished = False
        '0 out the matrices
        For i As Integer = 0 To MaximumOrder
            SumX(i) = 0.0#
            SumX(i + MaximumOrder) = 0.0#
            SumYX(i) = 0.0#
            C(i) = 0.0#
        Next i
    End Sub

    Private mCoeff As Double
    Public ReadOnly Property Coeff(Exponent As Long) As Double
        'retrieve the coefficient for the desired exponent
        Get
            Dim Ex As Long = Math.Abs(Exponent)
            Dim Order As Long = GlobalOrder
            If Not Finished Then Solve()

            'if we have less xy pairs than the desired order so far, change the order to one less than the number of pairs
            If XYCount <= Order Then Order = XYCount - 1
            'if the order is less than the exponent, return 0, else return the coefficient of the desired exponent
            If Order < Ex Then mCoeff = 0.0# Else mCoeff = C(Ex)
            Return mCoeff
        End Get
    End Property

    Public Property Degree() As Long
        'get and set desired polynomial order
        Get
            Return GlobalOrder
        End Get
        Set(ByVal value As Long)
            'trap order to be positive and less than the max order
            If value < 0 Or MaximumOrder < value Then
                Err.Raise(6000, "RegressionObject", value & " is an invalid property value! Use 0<= Degree <= " & MaximumOrder)
                Exit Property
            End If
            Init()  '0 out the matrices
            GlobalOrder = value     'set the desired order
        End Set
    End Property

    Private mXYCount As Long
    Public ReadOnly Property XYCount() As Long
        Get
            mXYCount = CLng(SumX(0))    'sumX(0) contains count of xy pairs so far
            Return mXYCount
        End Get
    End Property

    Public Sub XYAdd(ByVal NewX As Double, ByVal NewY As Double)
        'add new xy pair of data
        Dim TX As Double = 1.0#     'newX to a power
        Dim Max2O As Long = 2 * GlobalOrder  'twice the desired order
        Finished = False

        'sumX(0) contains count of number of xy pairs added so far
        SumX(0) = SumX(0) + 1

        'sumYx contains only the New Y
        SumYX(0) = SumYX(0) + NewY

        'calculate up to the desired order
        For i As Integer = 1 To GlobalOrder
            'TX = NewX to a power
            TX = TX * NewX
            'SumX = Summation of NewX to the power
            SumX(i) = SumX(i) + TX
            'sumYX = Summation of NewY * NewX to the power
            SumYX(i) = SumYX(i) + NewY * TX
        Next i

        For i As Integer = GlobalOrder + 1 To Max2O
            TX = TX * NewX
            SumX(i) = SumX(i) + TX
        Next i
    End Sub

    Public Function RegVal(X As Double) As Double

        Dim Order As Long = GlobalOrder
        If Not Finished Then Solve()
        RegVal = 0.0#

        If XYCount <= Order Then Order = XYCount - 1
        For i As Integer = 0 To Order
            RegVal = RegVal + C(i) * X ^ i
        Next i
    End Function

End Class
