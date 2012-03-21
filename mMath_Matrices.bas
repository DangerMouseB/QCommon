Attribute VB_Name = "mMath_Matrices"
'*************************************************************************************************************************************************************************************************************************************************
'
' Copyright (c) David Briant 2009-2010 - All rights reserved
'
'*************************************************************************************************************************************************************************************************************************************************

Option Explicit


'*************************************************************************************************************************************************************************************************************************************************
' Matrix operations - Double, return as Double() so can inline the calcs
'*************************************************************************************************************************************************************************************************************************************************

Function MSubMatrix(matrix() As Double, i1 As Long, i2 As Long, j1 As Long, j2 As Long) As Double()
    Dim iSize As Long, jSize As Long, i As Long, j As Long, answerMatrix() As Double, SA As SAFEARRAY, vType As Integer
    
    DBGetSafeArrayDetails matrix, SA, vType
    If SA.cbElements = 0 Then Exit Function
    
    iSize = i2 - i1 + 1
    jSize = j2 - j1 + 1
    DBCreateNewArrayOfDoubles answerMatrix, 1, iSize, 1, jSize
    For i = 1 To iSize
        For j = 1 To jSize
            answerMatrix(i, j) = matrix(i + i1 - 1, j + j1 - 1)
        Next
    Next
    MSubMatrix = answerMatrix
End Function

Function MScalerMult(scaler As Double, matrix() As Double) As Double()
    Dim i As Long, i1 As Long, i2 As Long, j As Long, j1 As Long, j2 As Long, answerMatrix() As Double, SA As SAFEARRAY, vType As Integer
    
    DBGetSafeArrayDetails matrix, SA, vType
    If SA.cbElements = 0 Then Exit Function
    
    DBGetArrayBounds matrix, 1, i1, i2
    DBGetArrayBounds matrix, 2, j1, j2
    DBCreateNewArrayOfDoubles answerMatrix, i1, i2, j1, j2
    For i = i1 To i2
        For j = j1 To j2
            answerMatrix(i, j) = scaler * matrix(i, j)
        Next
    Next
    MScalerMult = answerMatrix
End Function

Function MMult(matrixA() As Double, matrixB() As Double) As Double()
    Dim ai As Long, ai1 As Long, ai2 As Long, aj As Long, aj1 As Long, aj2 As Long, bi As Long, bi1 As Long, bi2 As Long, bj As Long, bj1 As Long, bj2 As Long, x As Long
    Dim sum As Double, i As Long, j As Long, answerMatrix() As Double, SA As SAFEARRAY, vType As Integer
    
    DBGetSafeArrayDetails matrixA, SA, vType
    If SA.cbElements = 0 Then Exit Function
    DBGetSafeArrayDetails matrixB, SA, vType
    If SA.cbElements = 0 Then Exit Function
    
    DBGetArrayBounds matrixA, 1, ai1, ai2
    DBGetArrayBounds matrixB, 2, bj1, bj2
    DBGetArrayBounds matrixA, 2, aj1, aj2
    DBGetArrayBounds matrixB, 1, bi1, bi2
    If ((bi2 - bi1) <> (aj2 - aj1)) Then Exit Function
    DBCreateNewArrayOfDoubles answerMatrix, ai1, ai2, bj1, bj2
    For ai = ai1 To ai2
        For bj = bj1 To bj2
            sum = 0
            For x = aj1 To aj2
                sum = sum + matrixA(ai, x) * matrixB(x, bj)
            Next
            answerMatrix(ai, bj) = sum
        Next
    Next
    MMult = answerMatrix
End Function

Function MTranspose(matrix() As Double) As Double()
    Dim i As Long, i1 As Long, i2 As Long, j As Long, j1 As Long, j2 As Long, answerMatrix() As Double, SA As SAFEARRAY, vType As Integer
    
    DBGetSafeArrayDetails matrix, SA, vType
    If SA.cbElements = 0 Then Exit Function
    
    DBGetArrayBounds matrix, 1, i1, i2
    DBGetArrayBounds matrix, 2, j1, j2
    DBCreateNewArrayOfDoubles answerMatrix, j1, j2, i1, i2
    For i = i1 To i2
        For j = j1 To j2
            answerMatrix(j, i) = matrix(i, j)
        Next
    Next
    MTranspose = answerMatrix
End Function

Function MAdd(matrixA() As Double, matrixB() As Double) As Double()
    Dim ai1 As Long, ai2 As Long, aj1 As Long, aj2 As Long, bi1 As Long, bi2 As Long, bj1 As Long, bj2 As Long, i As Long, j As Long, answerMatrix() As Double, SA As SAFEARRAY, vType As Integer
    
    DBGetSafeArrayDetails matrixA, SA, vType
    If SA.cbElements = 0 Then Exit Function
    DBGetSafeArrayDetails matrixB, SA, vType
    If SA.cbElements = 0 Then Exit Function
    
    DBGetArrayBounds matrixA, 1, ai1, ai2
    DBGetArrayBounds matrixA, 2, aj1, aj2
    DBGetArrayBounds matrixB, 1, bi1, bi2
    DBGetArrayBounds matrixB, 2, bj1, bj2
    If ai1 <> bi1 Or ai2 <> bi2 Or aj1 <> bj1 Or aj2 <> bj2 Then Exit Function
    DBCreateNewArrayOfDoubles answerMatrix, ai1, ai2, bj1, bj2
    For i = ai1 To ai2
        For j = aj1 To aj2
            answerMatrix(i, j) = matrixA(i, j) + matrixB(i, j)
        Next
    Next
    MAdd = answerMatrix
End Function
