Attribute VB_Name = "mMath_Moments"
'*************************************************************************************************************************************************************************************************************************************************
'
' Copyright (c) David Briant 2009-2011 - All rights reserved
'
'*************************************************************************************************************************************************************************************************************************************************

Option Explicit
Option Private Module
 
Function DBMoments1D(anArray2D() As Double) As Double()
    Dim compensation As Double, y As Double, t As Double, dummy As Double, sumXMinusXBar As Double, p As Double, oneOverCount As Double, xMinusXBar As Double
    Dim count As Long, sum As Double, mean As Double, var As Double, skew As Double, kurt As Double, sd As Double
    Dim i As Long, i1 As Long, i2 As Long, j As Long, j1 As Long, j2 As Long, answer1D() As Double
    
    DBGetArrayBounds anArray2D, 1, i1, i2
    DBGetArrayBounds anArray2D, 2, j1, j2
    
    ' high accuracy sum - see http://en.wikipedia.org/wiki/Compensated_summation
    For i = i1 To i2
        For j = j1 To j2
            count = count + 1#
            If count = 1# Then
                sum = anArray2D(i, j)
            Else
                y = anArray2D(i, j) - compensation
                t = sum + y
                compensation = t - sum
                compensation = compensation - y
                sum = t
            End If
        Next
    Next
    
    DBCreateNewArrayOfDoubles answer1D, 1, 6
    
    ' see p613 Numerical Receipes in C
    If count < 1 Then Exit Function
    oneOverCount = 1 / count
    mean = sum * oneOverCount
    If count < 2 Then
        answer1D(1) = count
        answer1D(2) = sum
        answer1D(3) = mean
        DBMoments1D = answer1D
        Exit Function
    End If
    For i = i1 To i2
        For j = j1 To j2
            xMinusXBar = anArray2D(i, j) - mean
            sumXMinusXBar = sumXMinusXBar + xMinusXBar
            p = xMinusXBar * xMinusXBar
            var = var + p
            p = p * xMinusXBar
            skew = skew + p
            p = p * xMinusXBar
            kurt = kurt + p
        Next
    Next
    var = (var - sumXMinusXBar * sumXMinusXBar * oneOverCount) / (count - 1)
    If var Then
        If var > 0 Then
            sd = Sqr(var)
            skew = skew / (count * var * sd)
            kurt = kurt / (count * var * var) - 3#
        End If
    End If
    answer1D(1) = count
    answer1D(2) = sum
    answer1D(3) = mean
    answer1D(4) = sd
    answer1D(5) = skew
    answer1D(6) = kurt
    DBMoments1D = answer1D
End Function

