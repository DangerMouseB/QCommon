Attribute VB_Name = "mNR_BrentMinimizer"
Option Explicit
Option Private Module

' error reporting
Private Const MODULE_NAME = "mNR_BrentMinimizer"
Private Const MODULE_VERSION As String = "0.0.0.1"


'*************************************************************************************************************************************************************************************************************************************************
' based on the Brent minimization method as detailed in Numerical Recipes In C, section 10.2, p 402
' the control loop has been inverted to allow the client to remain in control and remove the necessity of a callback function
'*************************************************************************************************************************************************************************************************************************************************

Public Const ERROR_BM_EXCEEDED_ITERATIONS As Long = 31003

' accessing - properties
Private Const cTolerance As Long = 1
Private Const cMaxIterations As Long = 2
Private Const cIterationNumber As Long = 3
Private Const cIsWithinTolerance As Long = 4

' accessing - root finding state
Private Const c_a As Long = 5
Private Const c_b As Long = 6
Private Const c_d As Long = 7
Private Const c_e As Long = 8
Private Const c_u As Long = 9
Private Const c_v As Long = 10
Private Const c_w As Long = 11
Private Const c_x As Long = 12
Private Const c_fu As Long = 13
Private Const c_fv As Long = 14
Private Const c_fw As Long = 15
Private Const c_fx As Long = 16

' algorithm constants
Private Const CGOLD As Double = 0.381966
Private Const ZEPS As Double = 0.0000000001

'Private Const EPS As Double = 0.000000000000003             ' machine floating point precision, compare with 3e-16


'*************************************************************************************************************************************************************************************************************************************************
' instantiation
'*************************************************************************************************************************************************************************************************************************************************

Function NRBM_newToken(tolerance As Double, lower As Double, middle As Double, fMiddle As Double, UPPER As Double) As Double()
    Dim BMToken() As Double
    DBCreateNewArrayOfDoubles BMToken, 1, 16
    If lower > middle Or middle > UPPER Then DBErrors_raiseGeneralError ModuleSummary(), "NRBM_newToken", "lower, middle, and upper not in order"

    BMToken(cMaxIterations) = 100
    BMToken(cTolerance) = tolerance
    BMToken(c_a) = lower
    BMToken(c_b) = UPPER
    BMToken(c_e) = 0#                                                                  ' This will be the distance moved on the step before last.
    BMToken(c_x) = middle
    BMToken(c_w) = middle
    BMToken(c_v) = middle
    BMToken(c_fx) = fMiddle
    BMToken(c_fv) = fMiddle
    BMToken(c_fw) = fMiddle
    BMToken(cIsWithinTolerance) = False
    BMToken(cIterationNumber) = 0
    NRBM_reEstimateX BMToken
    NRBM_newToken = BMToken
End Function


'*************************************************************************************************************************************************************************************************************************************************
' algorithm
'*************************************************************************************************************************************************************************************************************************************************

Sub NRBM_reEstimateX(BMToken() As Double)
    Dim a As Double, b As Double, d As Double, e As Double, u As Double, v As Double, w As Double, x As Double, fu As Double, fv As Double, fw As Double, fx As Double

    Dim xm As Double, tol1 As Double, tol2 As Double, p As Double, q As Double, r As Double, eTemp As Double
    
    BMToken(cIterationNumber) = BMToken(cIterationNumber) + 1
    If BMToken(cIterationNumber) > BMToken(cMaxIterations) Then DBErrors_raiseError ERROR_BM_EXCEEDED_ITERATIONS, ModuleSummary(), "NRBM_reEstimateX", "exceeded the maximun number of allowed iterations"
    
    a = BMToken(c_a)
    b = BMToken(c_b)
    d = BMToken(c_d)
    e = BMToken(c_e)
    
    u = BMToken(c_u)
    v = BMToken(c_v)
    w = BMToken(c_w)
    x = BMToken(c_x)
    
    fu = BMToken(c_fu)
    fv = BMToken(c_fv)
    fw = BMToken(c_fw)
    fx = BMToken(c_fx)
    
    If BMToken(cIterationNumber) > 1 Then
        If fu <= fx Then                                                ' Now decide what to do with our function evalutation
            If u >= x Then a = x Else b = x
            SHFT v, w, x, u                                             ' Housekeeping follows
            SHFT fv, fw, fx, fu
        Else
            If u < x Then a = u Else b = u
            If fu <= fw Or w = x Then
                v = w
                w = u
                fv = fw
                fw = fu
            Else
                If fu <= fv Or v = x Or v = w Then
                    v = u
                    fv = fu
                End If
            End If
        End If
    End If
    
    xm = 0.5 * (a + b)
    tol1 = BMToken(cTolerance) * Abs(x) + ZEPS
    tol2 = 2# * tol1
    If Abs(x - xm) <= (tol2 - 0.5 * (b - a)) Then       ' test for done here
        BMToken(cIsWithinTolerance) = True
        ' put the results into the appropiate variables
        u = x
        fu = fx
        Exit Sub
    End If
    
    If Abs(e) > tol1 Then                                         ' Construct a trial parabolic fit.
        r = (x - w) * (fx - fv)
        q = (x - v) * (fx - fw)
        p = (x - v) * q - (x - w) * r
        q = 2# * (q - r)
        If q > 0# Then p = -q
        q = Abs(q)
        eTemp = e
        e = d
        If Abs(p) >= Abs(0.5 * q * eTemp) Or p <= q * (a - x) Or p >= q * (b - x) Then
            ' the above conditions determine the acceptability of the parabolic fit. Here we
            ' take the golden section step into the larger of the two segments.
            e = IIf(x >= xm, a - x, b - x)
            d = CGOLD * e
        Else
            d = p / q
            u = x + d
            If (u - a) < tol2 Or (b - u) < tol2 Then d = SIGN(tol1, d)
        End If
    Else
        e = IIf(x >= xm, a - x, b - x)
        d = CGOLD * e
    End If
    u = IIf(Abs(d) >= tol1, x + d, x + SIGN(tol1, d))

    BMToken(c_a) = a
    BMToken(c_b) = b
    BMToken(c_d) = d
    BMToken(c_e) = e
    
    BMToken(c_u) = u
    BMToken(c_v) = v
    BMToken(c_w) = w
    BMToken(c_x) = x
    
    BMToken(c_fu) = fu
    BMToken(c_fv) = fv
    BMToken(c_fw) = fw
    BMToken(c_fx) = fx

End Sub

Private Sub SHFT(a As Double, b As Double, c As Double, d As Double)
    a = b
    b = c
    c = d
End Sub

Private Function SIGN(a As Double, b As Double) As Double
    If b >= 0 Then
        SIGN = Abs(a)
    Else
        SIGN = -Abs(a)
    End If
End Function


'*************************************************************************************************************************************************************************************************************************************************
' Accessing
'*************************************************************************************************************************************************************************************************************************************************

Property Get NRBM_fx(BMToken() As Double) As Double
    NRBM_fx = BMToken(c_fu)
End Property

Property Let NRBM_fx(BMToken() As Double, fx As Double)
    BMToken(c_fu) = fx
End Property

'*************************************************************************************************************************************************************************************************************************************************
Property Get NRBM_x(BMToken() As Double) As Double
    NRBM_x = BMToken(c_u)
End Property

'*************************************************************************************************************************************************************************************************************************************************
Property Let NRBM_maxIterations(BMToken() As Double, maxIterations As Double)
    BMToken(cMaxIterations) = maxIterations
End Property

'*************************************************************************************************************************************************************************************************************************************************
Property Get NRBM_isWithinTolerance(BMToken() As Double) As Boolean
    NRBM_isWithinTolerance = BMToken(cIsWithinTolerance)
End Property

'*************************************************************************************************************************************************************************************************************************************************
Property Get NRBM_numberOfIterations(BMToken() As Double) As Long
    NRBM_numberOfIterations = BMToken(cIterationNumber)
End Property


'*************************************************************************************************************************************************************************************************************************************************
' module summary
'*************************************************************************************************************************************************************************************************************************************************

Private Function ModuleSummary() As Variant()
    ModuleSummary = Array(1, GLOBAL_PROJECT_NAME, MODULE_NAME, MODULE_VERSION)
End Function





