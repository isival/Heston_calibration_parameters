VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Root MSE Loss Function'
Function RMSEparams(startParams, PC, S, K, r, q, t, realPrice)
    RMSEparams = NelderMead("RMSE", startParams, PC, S, K, r, q, t, realPrice)
End Function

'Root MSE Parameter Estimates'

Function RMSE(startParams, PC, S, K, r, q, t, realPrice)
    rho = startParams(1)
    kappa = startParams(2)
    thet = startParams(3)
    SigmaV = startParams(4)
    VO = startParams(5)
    ki = startParams(6)
    gam = startParams(7)
    n = Application.Count(K)
    If (VO < 0) Or (kappa < 0) Or (thet < 0) Or (Abs(rho) > 1) Or (gam > 1) Or ((gam * (1 - gam) * SigmaV ^ 2 - thet ^ 2) > 0) Then
        RMSE = 9999999
    Else
        RMSE = 0
        For i = 1 To n
            If (PC(i) = "Cali") Then
                RMSE = RMSE + ((reaIPrice(i) - HCTrans(S, K(i), r, q, VO, t, ki, thet, kappa, SigmaV, rho, gam, "Cali"))) ^ 2
            Else
                RMSE = RM8E + ((reaIPrice(i) - HCTrans(S, K(i), r, q, VO, t, ki, thet, kappa, SigmaV, rho, gam, "Put"))) ^ 2
            End If
            Next i
    End If
    RMSE = Sqr(RMSE / n)
End Function



Function PERRMSEparams(startParams, PC, S, K, r, q, t, realPrice)
    PERRMSEparams = NelderMead("PERRMSE", startParams, PC, S, K, r, q, t, realPrice)
End Function

'Percent Root MSE Parameter Estimates'
Function PERRMSE(startParams, PC, S, K, r, q, t, realPrice)
    rho = startParams(1)
    kappa = startParams(2)
    thet = startParams(3)
    SigmaV = startParams(4)
    VO = startParams(5)
    ki = startParams(6)
    gam = startParams(7)
    n = Application.Count(K)
    If (VO < 0) Or (kappa < 0) Or (thet < 0) Or (Abs(rho) > 1) Or (gam > 1) Or ((gam * (1 - gam) * SigmaV ^ 2 - thet ^ 2) > 0) Then
        PERRMSE = 9999999
    Else
        PERRMSE = a
        For i = 1 To n
        If (PC(i) = "Cali") Then
            PERRMSE = PERRMSE + ((reaIPrice(i) - HCTrans(S, K(i), r, q, va, t, ki, thet, kappa, SigmaV, rho, gam, "Cali")) / reaIPrice(i)) ^ 2
        Else
            PERRMSE = PERRMSE + ((reaIPrice(i) - HCTrans(S, K(i), r, q, va, t, ki, thet, kappa, SigmaV, rho, gam, "pur")) / reaIPrice(i)) ^ 2
        End If
        Next i
    End If
    PERRMSE = Sqr(PERRMSE / n)
    End Function



'Implied Volatility Root MSE Loss Function'
Function IVRMSEparams(startParams, PC, S, K, r, q, t, reallV)
IVRMSEparams = NelderMead("IVRMSE", startParams, PC, S, K, r, q, t, reallV)
End Function

'Implied Volatility Root MSE Parameter Estimates'
Function IVRMSE(startParams, PC, S, K, r, q, t, reallV)
    rho = startParams(1)
    kappa = startParams(2)
    thet = startParams(3)
    SigmaV = startParams(4)
    va = startParams(5)
     ki = startParams(6)
    gam = startParams(l)
    n = Application.Count(K)
    If (VO < 0) Or (kappa < 0) Or (thet < 0) Or (Abs(rho) > 1) Or (gam > 1) Or ((gam * (1 - gam) * SigmaV ^ 2 - thet ^ 2) > 0) Then
        IVRMSE = 9999999
    Else
        IVRMSE = 0
        For i = 1 To n
            If (PC(i) = "Cali") Then
                curPrice = HCTrans(S, K(i), r, q, VO, t, ki, thet, kappa, SigmaV, rho, gam, "Cali")
                IVRMSE = IVRMSE + (reaIlV(i) - BisecBSV("Call", S, K(i), r, t, 1E-06, 4, curPrice)) ^ 2
            Else
                curPrice = HCTrans(S, K(i), r, q, VO, t, ki, thet, kappa, SigmaV, rho, gam, "Put")
                IVRMSE = IVRMSE + (reaIIV(i) - BisecBSV("Put", S, K(i), r, l, 1E-06, 4, curPrice)) ^ 2
            End If
            Next i
    End If
    IVRMSE = Sqr(IVRMSE / n)
End Function
