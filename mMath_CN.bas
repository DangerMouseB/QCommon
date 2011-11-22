Attribute VB_Name = "mMath_CN"
'*************************************************************************************************************************************************************************************************************************************************
'
' Copyright (c) David Briant 2009-2011 - All rights reserved
'
'*************************************************************************************************************************************************************************************************************************************************

Option Explicit
Option Private Module

Private Const Pi As Double = 3.14159265358979
Private Const twoPi As Double = 6.28318530717958
Private Const twoThirdsPi As Double = 2.09439510239319
Private Const rootTwoPi As Double = 2.506628274631
Private Const root3Over2 As Double = 0.866025403784439
Private Const oneOverRootTwoPi As Double = 0.398942280401433
Private Const oneOver2 As Double = 1# / 2#
Private Const oneOver3 As Double = 1# / 3#
Private Const oneOver9 As Double = 1# / 9#
Private Const oneOver54 As Double = 1# / 54#

Private Const oneOverRoot2 As Double = 0.707106781186547     '  0.707106781186547514

Private Const minusOneOver2 As Double = -1 / 2


Function DBCN_Hart(x As Double) As Double
    ' from BETTER APPROXIMATIONS TO CUMULATIVE NORMAL FUNCTIONS By GRAEME WEST
    Dim xabs As Double, exponential As Double, build As Double
    xabs = Abs(x)
    If xabs > 37# Then
        DBCN_Hart = 0#
    Else
        exponential = Exp(-xabs * xabs * oneOver2)
        If xabs < 7.07106781186547 Then
            build = 3.52624965998911E-02 * xabs + 0.700383064443688
            build = build * xabs + 6.37396220353165
            build = build * xabs + 33.912866078383
            build = build * xabs + 112.079291497871
            build = build * xabs + 221.213596169931
            build = build * xabs + 220.206867912376
            DBCN_Hart = exponential * build
            build = 8.83883476483184E-02 * xabs + 1.75566716318264
            build = build * xabs + 16.064177579207
            build = build * xabs + 86.7807322029461
            build = build * xabs + 296.564248779674
            build = build * xabs + 637.333633378831
            build = build * xabs + 793.826512519948
            build = build * xabs + 440.413735824752
            DBCN_Hart = DBCN_Hart / build
        Else
            build = xabs + 0.65
            build = xabs + 4 / build
            build = xabs + 3 / build
            build = xabs + 2 / build
            build = xabs + 1 / build
            DBCN_Hart = exponential / build / 2.506628274631
        End If
    End If
    If x > 0# Then DBCN_Hart = 1 - DBCN_Hart
End Function

Function DBCN_Unknown(x As Double) As Double
    ' probably from hull but not sure
    
    Const cn1 As Double = 0.2316419
    Const cn2 As Double = 0.31938153
    Const cn3 As Double = -0.356563782
    Const cn4 As Double = 1.781477937
    Const cn5 As Double = -1.821255978
    Const cn6 As Double = 1.330274429

    Dim z As Double, z2 As Double, z3 As Double, z4 As Double, z5 As Double, answer As Double
    If x = 0# Then DBCN_Unknown = 0.5: Exit Function
    z = 1# / (1# + cn1 * Abs(x))
    z2 = z * z
    z3 = z2 * z
    z4 = z3 * z
    z5 = z4 * z
    answer = Exp(-(x * x) * oneOver2) * oneOverRootTwoPi * (cn2 * z + cn3 * z2 + cn4 * z3 + cn5 * z4 + cn6 * z5)
    If x > 0 Then
        DBCN_Unknown = 1# - answer
    Else
        DBCN_Unknown = answer
    End If
End Function

Function DBInvCN_Acklam(p As Double, Optional polishWithNewton As Boolean = True) As Double

    ' http://home.online.no/~pjacklam/notes/invnorm/index.html
    ' http://home.online.no/~pjacklam/notes/invnorm/impl/herrero/inversecdf.txt
    
    '  Adapted for Microsoft Visual Basic from Peter Acklam's
    '  "An algorithm for computing the inverse normal cumulative distribution function"
    '  (http://home.online.no/~pjacklam/notes/invnorm/)
    '  by John Herrero (3-Jan-03)
    
    Const p_low = 0.02425
    Const p_high = 1 - p_low
    
    Dim q As Double, r As Double, x As Double, e As Double, fx As Double
    Static constantInitialised As Boolean, a1 As Double, a2 As Double, a3 As Double, a4 As Double, a5 As Double, a6 As Double
    Static b1 As Double, b2 As Double, b3 As Double, b4 As Double, b5 As Double
    Static c1 As Double, c2 As Double, c3 As Double, c4 As Double, c5 As Double, c6 As Double
    Static d1 As Double, d2 As Double, d3 As Double, d4 As Double

    If Not constantInitialised Then
        a1 = CDec(-39.6968302866538) + CDec(0.00000000000004)
        a2 = CDec(220.946098424521) - CDec(0.0000000000005)
        a3 = CDec(-275.928510446969) + CDec(0.0000000000003)
        a4 = CDec(138.357751867269)
        a5 = CDec(-30.6647980661472) + CDec(0.00000000000004)
        a6 = CDec(2.50662827745924) - CDec(0.000000000000001)
        
        b1 = CDec(-54.4760987982241) + CDec(0.00000000000004)
        b2 = CDec(161.585836858041) - CDec(0.0000000000001)
        b3 = CDec(-155.698979859887) + CDec(0.0000000000004)
        b4 = CDec(66.8013118877197) + CDec(0.00000000000002)
        b5 = CDec(-13.2806815528857) - CDec(0.00000000000002)
        
        c1 = CDec(-7.78489400243029E-03) - CDec(3E-18)
        c2 = CDec(-0.322396458041136) - CDec(5E-16)
        c3 = CDec(-2.40075827716184) + CDec(0.000000000000002)
        c4 = CDec(-2.54973253934373) - CDec(0.000000000000004)
        c5 = CDec(4.37466414146497) - CDec(0.000000000000002)
        c6 = CDec(2.93816398269878) + CDec(0.000000000000003)
        
        d1 = CDec(7.78469570904146E-03) + CDec(2E-18)
        d2 = CDec(0.32246712907004) - CDec(2E-16)
        d3 = CDec(2.445134137143) - CDec(0.000000000000004)
        d4 = CDec(3.75440866190742) - CDec(0.000000000000004)
        constantInitialised = True
    End If
    
    'If argument out of bounds, raise error
    If p <= 0# Or p >= 1# Then Err.Raise 5
    
    If p < p_low Then
      'Rational approximation for lower region
      q = Sqr(-2# * Log(p))
      x = (((((c1 * q + c2) * q + c3) * q + c4) * q + c5) * q + c6) / _
        ((((d1 * q + d2) * q + d3) * q + d4) * q + 1#)
    ElseIf p <= p_high Then
      'Rational approximation for middle region
      q = p - 0.5
      r = q * q
      x = (((((a1 * r + a2) * r + a3) * r + a4) * r + a5) * r + a6) * q / _
        (((((b1 * r + b2) * r + b3) * r + b4) * r + b5) * r + 1#)
    ElseIf p < 1# Then
      'Rational approximation for upper region
      q = Sqr(-2 * Log(1# - p))
      x = -(((((c1 * q + c2) * q + c3) * q + c4) * q + c5) * q + c6) / _
        ((((d1 * q + d2) * q + d3) * q + d4) * q + 1#)
    End If

    If polishWithNewton Then
        ' refine to machine precision with a Newton step
        e = DBCN_Hart(x) - p
        fx = e * rootTwoPi * Exp(x * x * oneOver2)
        x = x - fx / (1 + x * fx * oneOver2)
    End If
    
    DBInvCN_Acklam = x
End Function


Function DBInvCN_Moro(p As Double, Optional polishWithNewton As Boolean = True) As Double
    ' http://www.mathfinance.com/FF/cpplib.php#normal.cpp

    ' returns the inverse of cumulative normal distribution function
    ' Reference> The Full Monte, by Boris Moro, Union Bank of Switzerland RISK 1995(2)

    Static constantInitialised As Boolean, c0 As Double, c1 As Double, c2 As Double, c3 As Double, c4 As Double, c5 As Double, c6 As Double, c7 As Double, c8 As Double
    Dim t As Double, x As Double, e As Double, fx As Double

    Const a0 As Double = 2.50662823884
    Const a1 As Double = -18.61500062529
    Const a2 As Double = 41.39119773534
    Const a3 As Double = -25.44106049637
    
    Const b0 As Double = -8.4735109309
    Const b1 As Double = 23.08336743743
    Const b2 As Double = -21.06224101826
    Const b3 As Double = 3.13082909833
    
    If Not constantInitialised Then
        c0 = CDec(0.337475482272614) + CDec(7E-16)          ' 0.3374754822726147
        c1 = CDec(0.976169019091718) + CDec(6E-16)          ' 0.9761690190917186
        c2 = CDec(0.16079797149182) + CDec(9E-16)            ' 0.1607979714918209
        c3 = CDec(2.76438810333863E-02)                            ' 0.0276438810333863
        c4 = CDec(3.8405729373609E-03)                              ' 0.0038405729373609
        c5 = CDec(3.951896511919E-04)                                ' 0.0003951896511919
        c6 = CDec(3.21767881768E-05)                                  ' 0.0000321767881768
        c7 = CDec(2.888167364E-07)                                     ' 0.0000002888167364
        c8 = CDec(3.960315187E-07)                                     ' 0.0000003960315187
        constantInitialised = True
    End If

    t = p - 0.5
    If (Abs(t) < 0.42) Then
        x = t * t
        x = t * (((a3 * x + a2) * x + a1) * x + a0) / _
        ((((b3 * x + b2) * x + b1) * x + b0) * x + 1#)
    Else
        If t > 0# Then
            x = 1# - p
        Else
            x = p
        End If
        x = Log(-Log(x))
        'x = c0 + x * (c1 + x * (c2 + x * (c3 + x * (c4 + x * (c5 + x * (c6 + x * (c7 + x * c8)))))))       ' stupid VB can't handle it this way round
        x = (((((((c8 * x + c7) * x + c6) * x + c5) * x + c4) * x + c3) * x + c2) * x + c1) * x + c0
        If t < 0# Then x = -x
    End If
    
    If polishWithNewton Then
        ' refine to machine precision with a Newton step
        e = DBCN_Hart(x) - p                  ' matches the hart CN
        fx = e * rootTwoPi * Exp(x * x * oneOver2)
        x = x - fx / (1 + x * fx * oneOver2)              ' x+ = xc - f(xc) / f'(xc)
    End If
    
    DBInvCN_Moro = x
End Function

