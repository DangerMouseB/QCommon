Attribute VB_Name = "mNR_BrentRootFinder"
Option Explicit
Option Private Module

' error reporting
Private Const MODULE_NAME = "mNR_BrentRootFinder"
Private Const MODULE_VERSION As String = "0.0.0.1"


'*************************************************************************************************************************************************************************************************************************************************
' based on the Brent root finding method as detailed in Numerical Recipes In C, section 9.3, p 359
' the control loop has been inverted to allow the client to remain in control and remove the necessity of a callback function
'*************************************************************************************************************************************************************************************************************************************************

Public Const ERROR_EXCEEDED_ITERATIONS As Long = 31001
Public Const ERROR_ROOT_NOT_BRACKETED As Long = 31002

' accessing - properties
Private Const cTolerance As Long = 1
Private Const cMaxIterations As Long = 2

' accessing - root finding state
Private Const cA As Long = 3
Private Const cB As Long = 4
Private Const cC As Long = 5
Private Const CD As Long = 6
Private Const cE As Long = 7
Private Const cFa As Long = 8
Private Const cFb As Long = 9
Private Const cFc As Long = 10
Private Const cIterationNumber As Long = 11
Private Const cIsWithinTolerance As Long = 12

' algorithm constants
Private Const EPS As Double = 0.000000000000003             ' machine floating point precision, compare with 3e-16


'*************************************************************************************************************************************************************************************************************************************************
' instantiation
'*************************************************************************************************************************************************************************************************************************************************

Function NRBRF_newToken(tolerance As Double, lower As Double, fLower As Double, UPPER As Double, fUpper As Double) As Double()
    Dim BRFToken() As Double
    DBCreateNewArrayOfDoubles BRFToken, 1, 12
    If (fLower > 0# And fUpper > 0#) Or (fLower < 0# And fUpper < 0#) Then DBErrors_raiseError ERROR_ROOT_NOT_BRACKETED, ModuleSummary(), "NRBRF_newToken", "#Root not bracketed!"
    BRFToken(cMaxIterations) = 100
    BRFToken(cTolerance) = tolerance
    BRFToken(cA) = lower
    BRFToken(cB) = UPPER
    BRFToken(cC) = UPPER
    BRFToken(cFa) = fLower
    BRFToken(cFb) = fUpper
    BRFToken(cFc) = fUpper
    BRFToken(cIsWithinTolerance) = False
    BRFToken(cIterationNumber) = 0
    NRBRF_reEstimateX BRFToken
    NRBRF_newToken = BRFToken
End Function


'*************************************************************************************************************************************************************************************************************************************************
' algorithm
'*************************************************************************************************************************************************************************************************************************************************

Sub NRBRF_reEstimateX(BRFToken() As Double)
    Dim min1 As Double, min2 As Double, min1or2 As Double, p As Double, q As Double, r As Double, S As Double, tol1 As Double, xm As Double
    
    BRFToken(cIterationNumber) = BRFToken(cIterationNumber) + 1
    If BRFToken(cIterationNumber) > BRFToken(cMaxIterations) Then DBErrors_raiseError ERROR_EXCEEDED_ITERATIONS, ModuleSummary(), "NRBRF_reEstimateX", "exceeded the maximun number of allowed iterations"
    If ((BRFToken(cFb) > 0# And BRFToken(cFc) > 0#) Or (BRFToken(cFb) < 0# And BRFToken(cFc) < 0#)) Then
        BRFToken(cC) = BRFToken(cA)
        BRFToken(cFc) = BRFToken(cFa)
        BRFToken(CD) = BRFToken(cB) - BRFToken(cA)
        BRFToken(cE) = BRFToken(CD)
    End If
    If Abs(BRFToken(cFc)) < Abs(BRFToken(cFb)) Then
        BRFToken(cA) = BRFToken(cB)                                                        ' rename a, b, c and adjust bounding interval d
        BRFToken(cB) = BRFToken(cC)
        BRFToken(cC) = BRFToken(cA)
        BRFToken(cFa) = BRFToken(cFb)
        BRFToken(cFb) = BRFToken(cFc)
        BRFToken(cFc) = BRFToken(cFa)
    End If
    tol1 = 2# * EPS * Abs(BRFToken(cB)) + 0.5 * BRFToken(cTolerance)       ' convergence check
    xm = 0.5 * (BRFToken(cC) - BRFToken(cB))
    If Abs(xm) <= tol1 Or BRFToken(cFb) = 0# Then BRFToken(cIsWithinTolerance) = True: Exit Sub
    If Abs(BRFToken(cE)) >= tol1 And Abs(BRFToken(cFa)) > Abs(BRFToken(cFb)) Then
        S = BRFToken(cFb) / BRFToken(cFa)                                                 ' attempt inverse quadratic interpolation
        If BRFToken(cA) = BRFToken(cC) Then
            p = 2# * xm * S
            q = 1# - S
        Else
            q = BRFToken(cFa) / BRFToken(cFc)
            r = BRFToken(cFb) / BRFToken(cFc)
            p = S * (2# * xm * q * (q - r) - (BRFToken(cB) - BRFToken(cA)) * (r - 1#))
            q = (q - 1#) * (r - 1#) * (S - 1#)
        End If
        If p > 0# Then q = -q                                             ' check whether in bounds
        p = Abs(p)
        min1 = 3# * xm * q - Abs(tol1 * q)
        min2 = Abs(BRFToken(cE) * q)
        If min1 < min2 Then min1or2 = min1 Else min1or2 = min2
        If 2# * p < min1or2 Then
            BRFToken(cE) = BRFToken(CD)                                                    ' accept interpolation
            BRFToken(CD) = p / q
        Else
            BRFToken(CD) = xm                                                      ' interpolation failed use bisection
            BRFToken(cE) = BRFToken(CD)
        End If
    Else                                                                       ' bounds decreasing too slowly, use bisection
        BRFToken(CD) = xm
        BRFToken(cE) = BRFToken(CD)
    End If
    BRFToken(cA) = BRFToken(cB)                                                             ' move last best guess to a
    BRFToken(cFa) = BRFToken(cFb)
    If Abs(BRFToken(CD)) > tol1 Then                                            ' evaluate new trial root
        BRFToken(cB) = BRFToken(cB) + BRFToken(CD)
    Else
        BRFToken(cB) = BRFToken(cB) + IIf(xm > 0, Abs(tol1), -Abs(tol1))
    End If
End Sub


'*************************************************************************************************************************************************************************************************************************************************
' Accessing
'*************************************************************************************************************************************************************************************************************************************************

Property Let NRBRF_fx(BRFToken() As Double, fx As Double)
    BRFToken(cFb) = fx
End Property

'*************************************************************************************************************************************************************************************************************************************************
Property Get NRBRF_x(BRFToken() As Double) As Double
    NRBRF_x = BRFToken(cB)
End Property

'*************************************************************************************************************************************************************************************************************************************************
Property Let NRBRF_maxIterations(BRFToken() As Double, maxIterations As Double)
    BRFToken(cMaxIterations) = maxIterations
End Property

'*************************************************************************************************************************************************************************************************************************************************
Property Get NRBRF_isWithinTolerance(BRFToken() As Double) As Boolean
    NRBRF_isWithinTolerance = BRFToken(cIsWithinTolerance)
End Property


'*************************************************************************************************************************************************************************************************************************************************
' module summary
'*************************************************************************************************************************************************************************************************************************************************

Private Function ModuleSummary() As Variant()
    ModuleSummary = Array(1, GLOBAL_PROJECT_NAME, MODULE_NAME, MODULE_VERSION)
End Function




