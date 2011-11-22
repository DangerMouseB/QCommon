Attribute VB_Name = "mMath_Roots"
'*************************************************************************************************************************************************************************************************************************************************
'
' Copyright (c) David Briant 2009-2011 - All rights reserved
'
'*************************************************************************************************************************************************************************************************************************************************

Option Explicit
Option Private Module


' error reporting
Private Const MODULE_NAME As String = "mMath_Roots"
Private Const MODULE_VERSION As String = "0.0.0.1"

' mathematical constants
Private Const Pi As Double = 3.14159265358979
Private Const twoPi As Double = 6.28318530717958
Private Const twoThirdsPi As Double = 2.09439510239319
Private Const rootTwoPi As Double = 2.506628274631
Private Const root3Over2 As Double = 0.866025403784439
Private Const oneOverRootTwoPi As Double = 0.398942280401433
Private Const oneOver2 As Double = 1 / 2
Private Const oneOver3 As Double = 1 / 3
Private Const oneOver9 As Double = 1 / 9
Private Const oneOver54 As Double = 1 / 54

Private Const minusOneOver2 As Double = -1 / 2


'*************************************************************************************************************************************************************************************************************************************************
' roots
'*************************************************************************************************************************************************************************************************************************************************

Function DBRootsOfQuadratic(a As Double, b As Double, C As Double) As Double()
    Dim q As Double, b2Minus4ac As Double, answer() As Double
    
    ' p 184 Numerical Recipes
    ' i havn't handled the case of imaginary roots !!!!!
    
    b2Minus4ac = b * b - 4 * a * C
    DBCreateNewArrayOfDoubles answer, 1, 2, 1, 2
    If b2Minus4ac >= 0 Then
        q = (b + Sgn(b) * Sqr(b2Minus4ac)) * minusOneOver2
        answer(1, 1) = q / a
        answer(1, 2) = 0
        answer(2, 1) = C / q
        answer(2, 2) = 0
    Else
        DBErrors_raiseNotYetImplemented ModuleSummary(), "DMRootsOfQuadratic", "ImaginaryRootsOfQuadratic"
    End If
    DBRootsOfQuadratic = answer
End Function

Function DBRootsOfCubic(a As Double, b As Double, C As Double) As Variant
    Dim a2 As Double, q As Double, q3 As Double, r As Double, r2 As Double, aOver3 As Double, thetaOver3 As Double, minus2rootQ As Double, bigA As Double, bigB As Double, answer() As Double

    ' p 184 - 5 Numerical Recipes
    a2 = a * a
    q = (a2 - 3 * b) * oneOver9
    r = (2 * a2 * a - 9 * a * b + 27 * C) * oneOver54
    r2 = r * r
    q3 = q * q * q
    aOver3 = a * oneOver3
    DBCreateNewArrayOfDoubles answer, 1, 3, 1, 2
    If r2 < q3 Then
        thetaOver3 = DBInvCos(r / Sqr(q3)) * oneOver3
        minus2rootQ = -2 * Sqr(q)
        answer(1, 1) = minus2rootQ * Cos(thetaOver3) - aOver3
        answer(1, 2) = 0
        answer(2, 1) = minus2rootQ * Cos(thetaOver3 + twoThirdsPi) - aOver3
        answer(2, 2) = 0
        answer(3, 1) = minus2rootQ * Cos(thetaOver3 - twoThirdsPi) - aOver3
        answer(3, 2) = 0
    Else
        bigA = -Sgn(r) * ((Abs(r) + Sqr(r2 - q3)) ^ oneOver3)
        If bigA <> 0 Then
            bigB = q / bigA
        Else
            bigB = 0
        End If
        answer(1, 1) = bigA + bigB - aOver3
        answer(1, 2) = 0
        answer(2, 1) = minusOneOver2 * (bigA + bigB) - aOver3
        answer(2, 2) = root3Over2 * (bigA - bigB)
        answer(3, 1) = minusOneOver2 * (bigA + bigB) - aOver3
        answer(3, 2) = -root3Over2 * (bigA - bigB)
    End If
    DBRootsOfCubic = answer
End Function

Function DBRootsOfCubicWithD(a As Double, b As Double, C As Double, d As Double, Optional raiseErrors As Boolean = False) As Double()
    DBRootsOfCubicWithD = DBRootsOfCubic(b / a, C / a, d / a)
End Function


'*************************************************************************************************************************************************************************************************************************************************
' error reporting utilities
'*************************************************************************************************************************************************************************************************************************************************

Private Function ModuleSummary() As Variant()
    ModuleSummary = Array(1, GLOBAL_PROJECT_NAME, MODULE_NAME, MODULE_VERSION)
End Function



