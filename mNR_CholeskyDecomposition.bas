Attribute VB_Name = "mNR_CholeskyDecomposition"
Option Explicit
Option Private Module

Private Const A_IS_NOT_PD_MATRIX = vbObjectError + 30001     '  Matrix a, with rounding errors, is not positive definite

Function NRCholeskyDecomposition(io_aMatrix() As Double, n As Long, o_pVector2D() As Double) As HRESULT
    Dim i As Long, j As Long, k As Long, sum As Double
    For i = 1 To n
        For j = i To n
            sum = io_aMatrix(i, j)
            For k = i - 1 To 1 Step -1
                sum = sum - io_aMatrix(i, k) * io_aMatrix(j, k)
            Next
            If i = j Then
                If sum <= 0# Then NRCholeskyDecomposition.HRESULT = A_IS_NOT_PD_MATRIX: Exit Function
                o_pVector2D(i, 1) = Sqr(sum)
            Else
                io_aMatrix(j, i) = sum / o_pVector2D(i, 1)
            End If
        Next
    Next
End Function

Sub NRCholeskySolve(aMatrix() As Double, n As Long, pVector2D() As Double, bVector2D() As Double, oXVector2D() As Double)
    Dim i As Long, k As Long, sum As Double
    For i = 1 To n
        ' Solve L.y = b, storing y in x.
        sum = bVector2D(i, 1)
        For k = i - 1 To 1 Step -1
            sum = sum - aMatrix(i, k) * oXVector2D(k, 1)
        Next
        oXVector2D(i, 1) = sum / pVector2D(i, 1)
    Next
    For i = n To 1 Step -1
        ' Solve L^T.x = y
        sum = oXVector2D(i, 1)
        For k = i + 1 To n
            sum = sum - aMatrix(k, i) * oXVector2D(k, 1)
        Next
        oXVector2D(i, 1) = sum / pVector2D(i, 1)
    Next
End Sub
