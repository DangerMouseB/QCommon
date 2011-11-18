Attribute VB_Name = "mMath_Common"
Option Explicit
Option Private Module

Private Const oneOver2 As Double = 1# / 2#


Private myAtn1 As Double


'*************************************************************************************************************************************************************************************************************************************************
' rounding
'*************************************************************************************************************************************************************************************************************************************************

Function DBRoundToSF(x As Double, sf As Long) As Double
    Dim numberOfDigitsBeforePoint As Long, scaler As Double
    If x = 0 Then
        numberOfDigitsBeforePoint = 0
    Else
        numberOfDigitsBeforePoint = DBLogN(Abs(x), 10#)
    End If
    scaler = 10# ^ (sf - numberOfDigitsBeforePoint - 1)
    DBRoundToSF = DBRoundToDP(x * scaler, 0) / scaler
End Function

Function DBRoundToDP(x As Double, dp As Long) As Double
    Dim scaler As Double
    scaler = 10# ^ dp
    DBRoundToDP = CLng(x * scaler) / scaler              ' a little dodgy... :-|   => do properly sometime so that we won't get overflow errors
End Function


'*************************************************************************************************************************************************************************************************************************************************
' interpolation
'*************************************************************************************************************************************************************************************************************************************************

Function DBLinearInterpolate(x As Double, x1 As Double, x2 As Double, y1 As Double, y2 As Double) As Double
    DBLinearInterpolate = y1 + (x - x1) / (x2 - x1) * (y2 - y1)
End Function


'*************************************************************************************************************************************************************************************************************************************************
' min / max
'*************************************************************************************************************************************************************************************************************************************************

Function DBMinB(a As Byte, b As Byte) As Byte
    If a < b Then DBMinB = a Else DBMinB = b
End Function

Function DBMaxB(a As Byte, b As Byte) As Byte
    If a > b Then DBMaxB = a Else DBMaxB = b
End Function

Function DBMinI(a As Integer, b As Integer) As Integer
    If a < b Then DBMinI = a Else DBMinI = b
End Function

Function DBMaxI(a As Integer, b As Integer) As Integer
    If a > b Then DBMaxI = a Else DBMaxI = b
End Function

Function DBMinL(a As Long, b As Long) As Long
    If a < b Then DBMinL = a Else DBMinL = b
End Function

Function DBMaxL(a As Long, b As Long) As Long
    If a > b Then DBMaxL = a Else DBMaxL = b
End Function

Function DBMinS(a As Single, b As Single) As Single
    If a < b Then DBMinS = a Else DBMinS = b
End Function

Function DBMaxS(a As Single, b As Single) As Single
    If a > b Then DBMaxS = a Else DBMaxS = b
End Function

Function DBMinD(a As Double, b As Double) As Double
    If a < b Then DBMinD = a Else DBMinD = b
End Function

Function DBMaxD(a As Double, b As Double) As Double
    If a > b Then DBMaxD = a Else DBMaxD = b
End Function

Function DBMinDate(a As Date, b As Date) As Date
    If a < b Then DBMinDate = a Else DBMinDate = b
End Function

Function DBMaxDate(a As Date, b As Date) As Date
    If a > b Then DBMaxDate = a Else DBMaxDate = b
End Function

Function DBMinC(a As Currency, b As Currency) As Currency
    If a < b Then DBMinC = a Else DBMinC = b
End Function

Function DBMaxC(a As Currency, b As Currency) As Currency
    If a > b Then DBMaxC = a Else DBMaxC = b
End Function


'*************************************************************************************************************************************************************************************************************************************************
' trigonometrical functions - from Visual Basic Reference - Derived Math Functions
'*************************************************************************************************************************************************************************************************************************************************

Function DBSec(x As Double) As Double
    DBSec = 1 / Cos(x)
End Function

Function DBCosec(x As Double) As Double
    DBCosec = 1 / Sin(x)
End Function

Function DBCotan(x As Double) As Double
    DBCotan = 1 / Tan(x)
End Function

Function DBInvSin(x As Double) As Double
    DBInvSin = Atn(x / Sqr(-x * x + 1))
End Function

Function DBInvCos(x As Double) As Double
    If myAtn1 = 0 Then myAtn1 = Atn(1)
    DBInvCos = Atn(-x / Sqr(-x * x + 1)) + 2 * myAtn1
End Function

Function DBInvSec(x As Double) As Double
    If myAtn1 = 0 Then myAtn1 = Atn(1)
    DBInvSec = Atn(x / Sqr(x * x - 1)) + Sgn((x) - 1) * (2 * myAtn1)
End Function

Function DBInvCosec(x As Double) As Double
    If myAtn1 = 0 Then myAtn1 = Atn(1)
    DBInvCosec = Atn(x / Sqr(x * x - 1)) + (Sgn(x) - 1) * (2 * myAtn1)
End Function

Function DBInvCotan(x As Double) As Double
    If myAtn1 = 0 Then myAtn1 = Atn(1)
    DBInvCotan = Atn(x) + 2 * myAtn1
End Function

Function DBHSin(x As Double) As Double
    DBHSin = (Exp(x) - Exp(-x)) * oneOver2
End Function

Function DBHCos(x As Double) As Double
    DBHCos = (Exp(x) + Exp(-x)) * oneOver2
End Function

Function DBHTan(x As Double) As Double
    DBHTan = (Exp(x) - Exp(-x)) / (Exp(x) + Exp(-x))
End Function

Function DBHSec(x As Double) As Double
    DBHSec = 2 / (Exp(x) + Exp(-x))
End Function

Function DBHCosec(x As Double) As Double
    DBHCosec = 2 / (Exp(x) - Exp(-x))
End Function

Function DBHCotan(x As Double) As Double
    DBHCotan = (Exp(x) + Exp(-x)) / (Exp(x) - Exp(-x))
End Function

Function DBInvHSin(x As Double) As Double
    DBInvHSin = Log(x + Sqr(x * x + 1))
End Function

Function DBInvHCos(x As Double) As Double
    DBInvHCos = Log(x + Sqr(x * x - 1))
End Function

Function DBInvHTan(x As Double) As Double
    DBInvHTan = Log((1 + x) / (1 - x)) * oneOver2
End Function

Function DBInvHSec(x As Double) As Double
    DBInvHSec = Log((Sqr(-x * x + 1) + 1) / x)
End Function

Function DBInvHCosec(x As Double) As Double
    DBInvHCosec = Log((Sgn(x) * Sqr(x * x + 1) + 1) / x)
End Function

Function DBInvHCotan(x As Double) As Double
    DBInvHCotan = Log((x + 1) / (x - 1)) * oneOver2
End Function

Function DBLogN(x As Double, n As Double)
    DBLogN = Log(x) / Log(n)
End Function


