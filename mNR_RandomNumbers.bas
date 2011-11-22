Attribute VB_Name = "mNR_RandomNumbers"
Option Explicit
Option Private Module

' Numerical Recipes - p279 for m=&H7FFFFFFF also a=48271 (q=44488 & r=3399) and a=69621 (q=30845 & r=23902)
 
Private Const IM1 As Long = 2147483563
Private Const IA1 As Long = 40014
Private Const IQ1 As Long = 53668
Private Const IR1 As Long = 12211

Private Const IM2 As Long = 2147483399
Private Const IA2 As Long = 40692
Private Const IQ2 As Long = 52774
Private Const IR2 As Long = 3791

Private Const AM1 As Double = 1# / IM1
Private Const IMM1 As Long = IM1 - 1

Private Const NTAB As Long = 32
'Private Const NDIV0 As Double = (1 + (IM0 - 1) / NTAB)
'Private Const NDIV1 As Double = (1 + IMM1 / NTAB)
Private Const EPS As Double = 0.000000000000003                  '0.00000012       ' 1.2e-7
Private Const RNMX As Double = 1 - EPS

Function ran0(idum As Long) As Double
    ' p 279 - Numerical Recipes
    Dim k As Long
    
    Const IM As Long = 2147483647           ' 2 ^ 31 - 1           &H7FFFFFFF
    Const IA As Long = 16807
    Const IQ As Long = 127773
    Const IR As Long = 2836
    Const AM As Double = 1# / IM
    Const MASK As Long = 123459876
    
    idum = idum Xor MASK
    If idum = 0 Then Stop
    k = idum / IQ
    idum = IA * (idum - k * IQ) - IR * k
    If idum < 0 Then idum = idum + IM
    ran0 = AM * idum
    idum = idum Xor MASK
End Function

Function ran1(idum As Long) As Double
    Dim j As Long, k As Long, temp As Double
    Static iy As Long, iv(0 To NTAB - 1) As Long
    
    Const IM As Long = 2147483647           ' 2 ^ 31 - 1           &H7FFFFFFF
    Const IA As Long = 16807
    Const IQ As Long = 127773
    Const IR As Long = 2836
    Const AM As Double = 1# / IM
    Const NDIV As Double = (1 + (IM - 1) / NTAB)

    If (idum <= 0 Or iy = 0) Then
        If -idum < 1 Then idum = 1 Else idum = -idum
        For j = NTAB + 7 To 0 Step -1
            k = idum / IQ
            idum = IA * (idum - k * IQ) - IR * k
            If idum < 0 Then idum = idum + IM
            If j < NTAB Then iv(j) = idum
        Next
        iy = iv(0)
    End If
    k = idum / IQ
    idum = IA * (idum - k * IQ) - IR * k
    If idum < 0 Then idum = idum + IM
    j = Int(iy / NDIV)
    iy = iv(j)
    iv(j) = idum
    temp = AM * iy
    If temp > RNMX Then
        ran1 = RNMX
    Else
        ran1 = temp
    End If
End Function

Function gaussianRan1(idum As Long) As Double
    Static haveSpare As Boolean, spare As Double
    Dim v1 As Double, v2 As Double, rsq As Double, fac As Double
    If haveSpare Then
        haveSpare = False
        gaussianRan1 = spare
    Else
        Do
            v1 = 2# * ran1(idum) - 1#
            v2 = 2# * ran1(idum) - 1#
            rsq = v1 * v1 + v2 * v2
        Loop While rsq >= 1# Or rsq = 0
        fac = Sqr(-2# * Log(rsq) / rsq)
        spare = v1 * fac
        haveSpare = True
        gaussianRan1 = v2 * fac
    End If
End Function

Function logNormalFromNormal(zeroOneGaussian As Double, mean As Double, sd As Double) As Double
    Dim EX As Double, EX2 As Double, EX_2 As Double, a As Double, b As Double, P1 As Double, P2 As Double
    EX = mean
    EX_2 = EX * EX
    If EX_2 < 1E-299 Then logNormalFromNormal = mean: Exit Function
    EX2 = mean * mean + sd * sd
    If EX2 < 1E-299 Then logNormalFromNormal = mean: Exit Function
    a = EX_2 / Sqr(EX2)
    P1 = Log(EX2)
    P2 = Log(EX_2)
    b = Sqr(P1 - P2)
    logNormalFromNormal = a * Exp(b * zeroOneGaussian)
End Function



