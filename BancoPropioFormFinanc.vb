Public Class BancoPropioFormFinanc

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBancoPropioFormulasFinancieras"

#End Region

#Region "Eventos BancoPropioFormFinanc"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDBancoPropio")) = 0 Then ApplicationService.GenerateError("El Campo ID Banco Propio es Obligatorio.")
    End Sub

#End Region

#Region "Funciones Públicas"

#Region "Funciones Calculo Intereses"

    Public Function CalculoInteresAplicadoBBVA(ByVal BlnBP As Boolean, ByVal DblInteresAplicado As Double) As Double
        If Not BlnBP Then
            Return DblInteresAplicado * 365 / 360
        Else
            Return DblInteresAplicado
        End If
    End Function

    Public Function CalculoInteresAplicadoCajaAstur(ByVal BlnBP As Boolean, ByVal DblInteresAplicado As Double) As Double
        If Not BlnBP Then
            Return DblInteresAplicado * 365 / 360
        Else
            Return DblInteresAplicado
        End If
    End Function

    Public Function CalculoInteresAplicadoBanesto(ByVal BlnBP As Boolean, ByVal DblInteresAplicado As Double) As Double
        If Not BlnBP Then
            Return DblInteresAplicado * 365 / 360
        Else
            Return DblInteresAplicado
        End If
    End Function

    Public Function CalculoInteresAplicadoCajaAsturias(ByVal BlnBP As Boolean, ByVal DblInteresAplicado As Double) As Double
        If Not BlnBP Then
            Return DblInteresAplicado * 365 / 360
        Else
            Return DblInteresAplicado
        End If
    End Function

    Public Function CalculoInteresAplicado(ByVal BlnBP As Boolean, ByVal DblInteresAplicado As Double) As Double
        Return DblInteresAplicado
    End Function

#End Region

#Region "Funciones Calculo Cuotas"

    Public Function CalculoCuotaBBVA(ByVal DblNTotalCuotas As Double, _
                                            ByVal DblNcuotasCarencia As Double, _
                                            ByVal DblTipoInteresAplicado As Double, _
                                            ByVal DblPagosAlAño As Double, _
                                            ByVal DblImpRecuperacionCoste As Double, _
                                            ByVal DblImpInteresesTotal As Double, _
                                            ByVal DblValorResidual As Double, _
                                            ByVal BlnValorResidualIgualCuota As Boolean, _
                                            ByVal DblCuota As Double, _
                                            ByVal DblRecuperacion As Double, _
                                            ByVal DblIntereses As Double, _
                                            ByVal StrDesglosePrimeraCuota As String, _
                                            ByVal LngNDecimalesImpA As Long) As DataTable
        Dim DtPagos As New DataTable
        Dim DtCalCuota As New DataTable
        Dim Dt As New DataTable
        Dim DblN, Dblk, DblBien, DblCuotaNew As Double
        Dim DblNumerador, DblNumerador2, DblExponente, DblExponenteMas As Double
        Dim DblDenominador1, DblDenominador2 As Double
        Dim DblRecuperacionNew, DblCapitalPte, DblInteresesNew As Double
        Dim Inti As Integer

        DblN = DblNTotalCuotas - DblNcuotasCarencia
        If DblPagosAlAño = 0 Then
            Exit Function
        Else
            If DblTipoInteresAplicado = 0 Then
                Exit Function
            Else
                Dblk = DblTipoInteresAplicado / (DblPagosAlAño * 100)
            End If
        End If
        DblBien = DblImpRecuperacionCoste
        DblNumerador = DblBien / (1 + Dblk)
        DblExponente = 1
        For i As Integer = 1 To DblN
            DblExponente *= (1 + Dblk)
        Next i
        DblExponenteMas = (1 + Dblk) * DblExponente
        DblDenominador1 = (DblExponente - 1) / (Dblk * DblExponente)
        If BlnValorResidualIgualCuota Then
            DblDenominador2 = 1 / DblExponenteMas
            DblCuotaNew = DblNumerador / (DblDenominador1 + DblDenominador2)
        Else
            DblNumerador2 = DblValorResidual / DblExponenteMas
            DblCuotaNew = (DblNumerador - DblNumerador2) / DblDenominador1
        End If
        Dt = CallByName(Me, StrDesglosePrimeraCuota, CallType.Method, Dblk, DblBien, DblCuotaNew, LngNDecimalesImpA)
        DblRecuperacionNew = Dt.Rows(0)("Recuperacion")
        DblInteresesNew = Dt.Rows(0)("Intereses")

        DtCalCuota.Columns.Add("Cuota", GetType(Double))
        DtCalCuota.Columns.Add("Recuperacion", GetType(Double))
        DtCalCuota.Columns.Add("Intereses", GetType(Double))

        Dim DrNew As DataRow = DtCalCuota.NewRow()
        DrNew("Cuota") = DblCuotaNew
        DrNew("Recuperacion") = DblRecuperacionNew
        DrNew("Intereses") = DblInteresesNew
        DtCalCuota.Rows.Add(DrNew)

        Return DtCalCuota
    End Function

    Public Function CalculoCuotaBanesto(ByVal DblNTotalCuotas As Double, _
                                              ByVal DblNcuotasCarencia As Double, _
                                              ByVal DblTipoInteresAplicado As Double, _
                                              ByVal DblPagosAlAño As Double, _
                                              ByVal DblImpRecuperacionCoste As Double, _
                                              ByVal DblImpInteresesTotal As Double, _
                                              ByVal DblValorResidual As Double, _
                                              ByVal BlnValorResidualIgualCuota As Boolean, _
                                              ByVal DblCuota As Double, _
                                              ByVal DblRecuperacion As Double, _
                                              ByVal DblIntereses As Double, _
                                              ByVal StrDesglosePrimeraCuota As String, _
                                              ByVal LngNDecimalesImpA As Long) As DataTable
        Dim DtPagos As New DataTable
        Dim DtCalCuota As New DataTable
        Dim Dt As New DataTable
        Dim DblN, Dblk, DblBien, DblCuotaNew As Double
        Dim DblNumerador, DblNumerador2, DblExponente, DblExponenteMas As Double
        Dim DblDenominador1, DblDenominador2 As Double
        Dim DblRecuperacionNew, DblCapitalPte, DblInteresesNew As Double
        Dim Inti As Integer

        DblN = DblNTotalCuotas - DblNcuotasCarencia
        If DblPagosAlAño = 0 Then
            Exit Function
        Else
            If DblTipoInteresAplicado = 0 Then
                Exit Function
            Else
                Dblk = DblTipoInteresAplicado / (DblPagosAlAño * 100)
            End If
        End If
        DblBien = DblImpRecuperacionCoste
        DblNumerador = DblBien / (1 + Dblk)
        DblExponente = 1
        For i As Integer = 1 To DblN
            DblExponente *= (1 + Dblk)
        Next i
        DblExponenteMas = (1 + Dblk) * DblExponente
        DblDenominador1 = (DblExponente - 1) / (Dblk * DblExponente)
        If BlnValorResidualIgualCuota Then
            DblDenominador2 = 1 / DblExponenteMas
            DblCuotaNew = DblNumerador / (DblDenominador1 + DblDenominador2)
        Else
            DblNumerador2 = DblValorResidual / DblExponenteMas
            DblCuotaNew = (DblNumerador - DblNumerador2) / DblDenominador1
        End If
        Dt = CallByName(Me, StrDesglosePrimeraCuota, CallType.Method, Dblk, DblBien, DblCuotaNew, LngNDecimalesImpA)
        DblRecuperacionNew = Dt.Rows(0)("Recuperacion")
        DblInteresesNew = Dt.Rows(0)("Intereses")

        DtCalCuota.Columns.Add("Cuota", GetType(Double))
        DtCalCuota.Columns.Add("Recuperacion", GetType(Double))
        DtCalCuota.Columns.Add("Intereses", GetType(Double))

        Dim DrNew As DataRow = DtCalCuota.NewRow()
        DrNew("Cuota") = DblCuotaNew
        DrNew("Recuperacion") = DblRecuperacionNew
        DrNew("Intereses") = DblInteresesNew
        DtCalCuota.Rows.Add(DrNew)

        Return DtCalCuota
    End Function

    Public Function CalculoCuotaCajaAstur(ByVal DblNTotalCuotas As Double, _
                                              ByVal DblNcuotasCarencia As Double, _
                                              ByVal DblTipoInteresAplicado As Double, _
                                              ByVal DblPagosAlAño As Double, _
                                              ByVal DblImpRecuperacionCoste As Double, _
                                              ByVal DblImpInteresesTotal As Double, _
                                              ByVal DblValorResidual As Double, _
                                              ByVal BlnValorResidualIgualCuota As Boolean, _
                                              ByVal DblCuota As Double, _
                                              ByVal DblRecuperacion As Double, _
                                              ByVal DblIntereses As Double, _
                                              ByVal StrDesglosePrimeraCuota As String, _
                                              ByVal LngNDecimalesImpA As Long) As DataTable
        Dim DtPagos As New DataTable
        Dim DtCalCuota As New DataTable
        Dim Dt As New DataTable
        Dim DblN, Dblk, DblBien, DblCuotaNew As Double
        Dim DblNumerador, DblNumerador2, DblExponente, DblExponenteMas As Double
        Dim DblDenominador1, DblDenominador2 As Double
        Dim DblRecuperacionNew, DblCapitalPte, DblInteresesNew As Double
        Dim Inti As Integer

        DblN = DblNTotalCuotas - DblNcuotasCarencia
        If DblPagosAlAño = 0 Then
            Exit Function
        Else
            If DblTipoInteresAplicado = 0 Then
                Exit Function
            Else
                Dblk = DblTipoInteresAplicado / (DblPagosAlAño * 100)
            End If
        End If
        DblBien = DblImpRecuperacionCoste
        DblNumerador = DblBien / (1 + Dblk)
        DblExponente = 1
        For i As Integer = 1 To DblN
            DblExponente *= (1 + Dblk)
        Next i
        DblExponenteMas = (1 + Dblk) * DblExponente
        DblDenominador1 = (DblExponente - 1) / (Dblk * DblExponente)
        If BlnValorResidualIgualCuota Then
            DblDenominador2 = 1 / DblExponenteMas
            DblCuotaNew = DblNumerador / (DblDenominador1 + DblDenominador2)
        Else
            DblNumerador2 = DblValorResidual / DblExponenteMas
            DblCuotaNew = (DblNumerador - DblNumerador2) / DblDenominador1
        End If
        Dt = CallByName(Me, StrDesglosePrimeraCuota, CallType.Method, Dblk, DblBien, DblCuotaNew, LngNDecimalesImpA)
        DblRecuperacionNew = Dt.Rows(0)("Recuperacion")
        DblInteresesNew = Dt.Rows(0)("Intereses")

        DtCalCuota.Columns.Add("Cuota", GetType(Double))
        DtCalCuota.Columns.Add("Recuperacion", GetType(Double))
        DtCalCuota.Columns.Add("Intereses", GetType(Double))

        Dim DrNew As DataRow = DtCalCuota.NewRow()
        DrNew("Cuota") = DblCuotaNew
        DrNew("Recuperacion") = DblRecuperacionNew
        DrNew("Intereses") = DblInteresesNew
        DtCalCuota.Rows.Add(DrNew)

        Return DtCalCuota
    End Function

    Public Function CalculoCuotaCajaAsturias(ByVal DblNTotalCuotas As Double, _
                                              ByVal DblNcuotasCarencia As Double, _
                                              ByVal DblTipoInteresAplicado As Double, _
                                              ByVal DblPagosAlAño As Double, _
                                              ByVal DblImpRecuperacionCoste As Double, _
                                              ByVal DblImpInteresesTotal As Double, _
                                              ByVal DblValorResidual As Double, _
                                              ByVal BlnValorResidualIgualCuota As Boolean, _
                                              ByVal DblCuota As Double, _
                                              ByVal DblRecuperacion As Double, _
                                              ByVal DblIntereses As Double, _
                                              ByVal StrDesglosePrimeraCuota As String, _
                                              ByVal LngNDecimalesImpA As Long) As DataTable
        Dim DtPagos As New DataTable
        Dim DtCalCuota As New DataTable
        Dim Dt As New DataTable
        Dim DblN, Dblk, DblBien, DblCuotaNew As Double
        Dim DblNumerador, DblNumerador2, DblExponente, DblExponenteMas As Double
        Dim DblDenominador1, DblDenominador2 As Double
        Dim DblRecuperacionNew, DblCapitalPte, DblInteresesNew As Double
        Dim Inti As Integer

        DblN = DblNTotalCuotas - DblNcuotasCarencia
        If DblPagosAlAño = 0 Then
            Exit Function
        Else
            If DblTipoInteresAplicado = 0 Then
                Exit Function
            Else
                Dblk = DblTipoInteresAplicado / (DblPagosAlAño * 100)
            End If
        End If
        DblBien = DblImpRecuperacionCoste
        DblNumerador = DblBien / (1 + Dblk)
        DblExponente = 1
        For i As Integer = 1 To DblN
            DblExponente *= (1 + Dblk)
        Next i
        DblExponenteMas = (1 + Dblk) * DblExponente
        DblDenominador1 = (DblExponente - 1) / (Dblk * DblExponente)
        If BlnValorResidualIgualCuota Then
            DblDenominador2 = 1 / DblExponenteMas
            DblCuotaNew = DblNumerador / (DblDenominador1 + DblDenominador2)
        Else
            DblNumerador2 = DblValorResidual / DblExponenteMas
            DblCuotaNew = (DblNumerador - DblNumerador2) / DblDenominador1
        End If
        Dt = CallByName(Me, StrDesglosePrimeraCuota, CallType.Method, Dblk, DblBien, DblCuotaNew, LngNDecimalesImpA)
        DblRecuperacionNew = Dt.Rows(0)("Recuperacion")
        DblInteresesNew = Dt.Rows(0)("Intereses")

        DtCalCuota.Columns.Add("Cuota", GetType(Double))
        DtCalCuota.Columns.Add("Recuperacion", GetType(Double))
        DtCalCuota.Columns.Add("Intereses", GetType(Double))

        Dim DrNew As DataRow = DtCalCuota.NewRow()
        DrNew("Cuota") = DblCuotaNew
        DrNew("Recuperacion") = DblRecuperacionNew
        DrNew("Intereses") = DblInteresesNew
        DtCalCuota.Rows.Add(DrNew)

        Return DtCalCuota
    End Function

    Public Function CalculoCuotaGenerica(ByVal DblNTotalCuotas As Double, _
                                         ByVal DblNcuotasCarencia As Double, _
                                         ByVal DblTipoInteresAplicado As Double, _
                                         ByVal DblPagosAlAño As Double, _
                                         ByVal DblImpRecuperacionCoste As Double, _
                                         ByVal DblImpInteresesTotal As Double, _
                                         ByVal DblValorResidual As Double, _
                                         ByVal BlnValorResidualIgualCuota As Boolean, _
                                         ByVal DblCuota As Double, _
                                         ByVal DblRecuperacion As Double, _
                                         ByVal DblIntereses As Double, _
                                         ByVal StrDesglosePrimeraCuota As String, _
                                         ByVal LngNDecimalesImpA As Long) As DataTable
        Dim DtCalCuota As New DataTable
        DtCalCuota.Columns.Add("Cuota", GetType(Double))
        DtCalCuota.Columns.Add("Recuperacion", GetType(Double))
        DtCalCuota.Columns.Add("Intereses", GetType(Double))

        Dim DrNew As DataRow = DtCalCuota.NewRow()
        DrNew("Cuota") = DblCuota
        DrNew("Recuperacion") = DblRecuperacion
        DrNew("Intereses") = DblIntereses
        DtCalCuota.Rows.Add(DrNew)

        Return DtCalCuota
    End Function

#End Region

#Region "Funciones Desglose Primera Cuota"

    Public Function DesglosePrimeraCuotaBBVA(ByVal Dblk As Double, _
                                                    ByVal DblBien As Double, _
                                                    ByVal DblCuota As Double, _
                                                    ByVal LngNDecimalesImp As Long) As DataTable
        Dim DblRecuperacion, DblCapitalPte, DblIntereses As Double
        Dim DtDesglose As New DataTable

        DblRecuperacion = (xRound(DblCuota, LngNDecimalesImp) * (1 + Dblk)) - (DblBien * Dblk)
        DblCapitalPte = DblBien - DblRecuperacion
        DblIntereses = DblCapitalPte * Dblk / (1 + Dblk)

        DtDesglose.Columns.Add("Recuperacion", GetType(Double))
        DtDesglose.Columns.Add("Intereses", GetType(Double))

        Dim DrDesglose As DataRow = DtDesglose.NewRow()
        DrDesglose("Recuperacion") = DblRecuperacion
        DrDesglose("Intereses") = DblIntereses
        DtDesglose.Rows.Add(DrDesglose)

        Return DtDesglose
    End Function

    Public Function DesglosePrimeraCuotaBanesto(ByVal Dblk As Double, _
                                                    ByVal DblBien As Double, _
                                                    ByVal DblCuota As Double, _
                                                    ByVal LngNDecimalesImp As Long) As DataTable
        Dim DblRecuperacion, DblCapitalPte, DblIntereses As Double
        Dim DtDesglose As New DataTable

        DblRecuperacion = (xRound(DblCuota, LngNDecimalesImp) * (1 + Dblk)) - (DblBien * Dblk)
        DblCapitalPte = DblBien - DblRecuperacion
        DblIntereses = DblCapitalPte * Dblk / (1 + Dblk)

        DtDesglose.Columns.Add("Recuperacion", GetType(Double))
        DtDesglose.Columns.Add("Intereses", GetType(Double))

        Dim DrDesglose As DataRow = DtDesglose.NewRow()
        DrDesglose("Recuperacion") = DblRecuperacion
        DrDesglose("Intereses") = DblIntereses
        DtDesglose.Rows.Add(DrDesglose)

        Return DtDesglose
    End Function

    Public Function DesglosePrimeraCuotaCajaAstur(ByVal Dblk As Double, _
                                                         ByVal DblBien As Double, _
                                                         ByVal DblCuota As Double, _
                                                         ByVal LngNDecimalesImp As Long) As DataTable
        Dim DblRecuperacion, DblCapitalPte, DblIntereses As Double
        Dim DtDesglose As New DataTable

        DblCapitalPte = DblBien - DblCuota
        DblIntereses = DblCapitalPte * Dblk
        DblRecuperacion = DblCuota - DblIntereses

        DtDesglose.Columns.Add("Recuperacion", GetType(Double))
        DtDesglose.Columns.Add("Intereses", GetType(Double))

        Dim DrNew As DataRow = DtDesglose.NewRow()
        DrNew("Recuperacion") = DblRecuperacion
        DrNew("Intereses") = DblIntereses
        DtDesglose.Rows.Add(DrNew)

        Return DtDesglose
    End Function

    Public Function DesglosePrimeraCuotaCajaAsturias(ByVal Dblk As Double, _
                                                       ByVal DblBien As Double, _
                                                       ByVal DblCuota As Double, _
                                                       ByVal LngNDecimalesImp As Long) As DataTable
        Dim DblRecuperacion, DblCapitalPte, DblIntereses As Double
        Dim DtDesglose As New DataTable

        DblCapitalPte = DblBien - DblCuota
        DblIntereses = DblCapitalPte * Dblk
        DblRecuperacion = DblCuota - DblIntereses

        DtDesglose.Columns.Add("Recuperacion", GetType(Double))
        DtDesglose.Columns.Add("Intereses", GetType(Double))

        Dim DrNew As DataRow = DtDesglose.NewRow()
        DrNew("Recuperacion") = DblRecuperacion
        DrNew("Intereses") = DblIntereses
        DtDesglose.Rows.Add(DrNew)

        Return DtDesglose
    End Function

#End Region

#Region "Funciones Desglose Sucesivas Cuotas"

    Public Function DesgloseSucesivasCuotasBBVA(ByVal Dblk As Double, _
                                                       ByVal DblBien As Double, _
                                                       ByVal DblAmortizacion As Double, _
                                                       ByVal DblCuota As Double, _
                                                       ByVal LngPeriodo As Long, _
                                                       ByVal BlnRedondeoFinal As Boolean, _
                                                       ByVal DblInteresAntCuota As Double, _
                                                       ByVal LngDecimalesImpA As Long, _
                                                       ByVal DblTotalIntereses As Double, _
                                                       ByVal DblTotalCuota As Double, _
                                                       ByVal DblTotalFinanciar As Double, _
                                                       ByVal DblValorResidual As Double, _
                                                       ByVal BlnCuotaVR As Boolean) As DataTable
        Dim DblRecuperacion, DblCapitalPte, DblIntereses As Double
        Dim DtDesglose As New DataTable

        If BlnCuotaVR Then
            DblIntereses = 0
            DblRecuperacion = DblValorResidual
            DblCapitalPte = DblBien - DblRecuperacion
        Else
            If BlnRedondeoFinal Then
                ' Se suma dos veces porque faltan dos cuotas que serían la que estamos analizando más la de VR.
                DblTotalCuota = DblTotalCuota + DblCuota + DblValorResidual
                DblIntereses = xRound(DblTotalCuota - (DblTotalFinanciar + DblTotalIntereses), LngDecimalesImpA)
                DblRecuperacion = DblCuota - xRound(DblIntereses, LngDecimalesImpA)
                DblCapitalPte = DblBien - DblRecuperacion
            Else
                DblRecuperacion = (DblAmortizacion * (1 + Dblk))
                DblCapitalPte = DblBien - DblRecuperacion
                DblIntereses = DblCuota - xRound(DblRecuperacion, LngDecimalesImpA)
            End If
        End If

        DtDesglose.Columns.Add("Recuperacion", GetType(Double))
        DtDesglose.Columns.Add("Intereses", GetType(Double))
        DtDesglose.Columns.Add("CapitalPte", GetType(Double))

        Dim DrNew As DataRow = DtDesglose.NewRow()
        DrNew("Recuperacion") = DblRecuperacion
        DrNew("Intereses") = DblIntereses
        DrNew("CapitalPte") = DblCapitalPte
        DtDesglose.Rows.Add(DrNew)

        Return DtDesglose
    End Function

    Public Function DesgloseSucesivasCuotasBanesto(ByVal Dblk As Double, _
                                                           ByVal DblBien As Double, _
                                                           ByVal DblAmortizacion As Double, _
                                                           ByVal DblCuota As Double, _
                                                           ByVal LngPeriodo As Long, _
                                                           ByVal BlnRedondeoFinal As Boolean, _
                                                           ByVal DblInteresAntCuota As Double, _
                                                           ByVal LngDecimalesImpA As Long, _
                                                           ByVal DblTotalIntereses As Double, _
                                                           ByVal DblTotalCuota As Double, _
                                                           ByVal DblTotalFinanciar As Double, _
                                                           ByVal DblValorResidual As Double, _
                                                           ByVal BlnCuotaVR As Boolean) As DataTable
        Dim DblRecuperacion, DblCapitalPte, DblIntereses As Double
        Dim DtDesglose As New DataTable

        If BlnCuotaVR Then
            DblIntereses = 0
            DblRecuperacion = DblValorResidual
            DblCapitalPte = DblBien - DblRecuperacion
        Else
            If BlnRedondeoFinal Then
                ' Se suma dos veces porque faltan dos cuotas que serían la que estamos analizando más la de VR.
                DblTotalCuota = DblTotalCuota + DblCuota + DblValorResidual
                DblIntereses = xRound(DblTotalCuota - (DblTotalFinanciar + DblTotalIntereses), LngDecimalesImpA)
                DblRecuperacion = DblCuota - xRound(DblIntereses, LngDecimalesImpA)
                DblCapitalPte = DblBien - DblRecuperacion
            Else
                DblRecuperacion = (DblAmortizacion * (1 + Dblk))
                DblCapitalPte = DblBien - DblRecuperacion
                DblIntereses = (xRound(DblBien - DblRecuperacion, LngDecimalesImpA)) * Dblk / (1 + Dblk)
                DblRecuperacion = DblCuota - xRound(DblIntereses, LngDecimalesImpA)
                DblCapitalPte = DblBien - DblRecuperacion
            End If
        End If

        DtDesglose.Columns.Add("Recuperacion", GetType(Double))
        DtDesglose.Columns.Add("Intereses", GetType(Double))
        DtDesglose.Columns.Add("CapitalPte", GetType(Double))

        Dim DrNew As DataRow = DtDesglose.NewRow
        DrNew("Recuperacion") = DblRecuperacion
        DrNew("Intereses") = DblIntereses
        DrNew("CapitalPte") = DblCapitalPte
        DtDesglose.Rows.Add(DrNew)

        Return DtDesglose
    End Function

    Public Function DesgloseSucesivasCuotasCajaAstur(ByVal Dblk As Double, _
                                                            ByVal DblBien As Double, _
                                                            ByVal DblAmortizacion As Double, _
                                                            ByVal DblCuota As Double, _
                                                            ByVal LngPeriodo As Long, _
                                                            ByVal BlnRedondeoFinal As Boolean, _
                                                            ByVal DblInteresAntCuota As Double, _
                                                            ByVal LngDecimalesImpA As Long, _
                                                            ByVal DblTotalIntereses As Double, _
                                                            ByVal DblTotalCuota As Double, _
                                                            ByVal DblTotalFinanciar As Double, _
                                                            ByVal DblValorResidual As Double, _
                                                            ByVal BlnCuotaVR As Boolean) As DataTable
        Dim DblRecuperacion, DblCapitalPte, DblIntereses, DblMargenInteres As Double
        Dim DtDesglose As New DataTable

        If LngPeriodo = 1 Then DblBien = xRound(DblBien, LngDecimalesImpA) - DblCuota

        If BlnRedondeoFinal Then
            DblTotalCuota = DblTotalCuota + DblCuota + DblValorResidual
            DblIntereses = xRound(DblTotalCuota - (DblTotalFinanciar + DblTotalIntereses), LngDecimalesImpA)
            DblRecuperacion = DblCuota - xRound(DblIntereses, LngDecimalesImpA)
            DblCapitalPte = DblBien - DblRecuperacion
        Else
            DblIntereses = xRound(DblBien, LngDecimalesImpA) * Dblk
            DblRecuperacion = DblCuota - DblIntereses
            DblCapitalPte = xRound(DblBien, LngDecimalesImpA) - DblRecuperacion
        End If

        DtDesglose.Columns.Add("Recuperacion", GetType(Double))
        DtDesglose.Columns.Add("Intereses", GetType(Double))
        DtDesglose.Columns.Add("CapitalPte", GetType(Double))

        Dim DrNew As DataRow = DtDesglose.NewRow
        DrNew("Recuperacion") = DblRecuperacion
        DrNew("Intereses") = DblIntereses
        DrNew("CapitalPte") = DblCapitalPte
        DtDesglose.Rows.Add(DrNew)

        Return DtDesglose
    End Function

    Public Function DesgloseSucesivasCuotasCajaAsturias(ByVal Dblk As Double, _
                                                            ByVal DblBien As Double, _
                                                            ByVal DblAmortizacion As Double, _
                                                            ByVal DblCuota As Double, _
                                                            ByVal LngPeriodo As Long, _
                                                            ByVal BlnRedondeoFinal As Boolean, _
                                                            ByVal DblInteresAntCuota As Double, _
                                                            ByVal LngDecimalesImpA As Long, _
                                                            ByVal DblTotalIntereses As Double, _
                                                            ByVal DblTotalCuota As Double, _
                                                            ByVal DblTotalFinanciar As Double, _
                                                            ByVal DblValorResidual As Double, _
                                                            ByVal BlnCuotaVR As Boolean) As DataTable
        Dim DblRecuperacion, DblCapitalPte, DblIntereses, DblMargenInteres As Double
        Dim DtDesglose As New DataTable

        If LngPeriodo = 1 Then DblBien = xRound(DblBien, LngDecimalesImpA) - DblCuota

        If BlnRedondeoFinal Then
            DblTotalCuota = DblTotalCuota + DblCuota + DblValorResidual
            DblIntereses = xRound(DblTotalCuota - (DblTotalFinanciar + DblTotalIntereses), LngDecimalesImpA)
            DblRecuperacion = DblCuota - xRound(DblIntereses, LngDecimalesImpA)
            DblCapitalPte = DblBien - DblRecuperacion
        Else
            DblIntereses = xRound(DblBien, LngDecimalesImpA) * Dblk
            DblRecuperacion = DblCuota - DblIntereses
            DblCapitalPte = xRound(DblBien, LngDecimalesImpA) - DblRecuperacion
        End If

        DtDesglose.Columns.Add("Recuperacion", GetType(Double))
        DtDesglose.Columns.Add("Intereses", GetType(Double))
        DtDesglose.Columns.Add("CapitalPte", GetType(Double))

        Dim DrNew As DataRow = DtDesglose.NewRow
        DrNew("Recuperacion") = DblRecuperacion
        DrNew("Intereses") = DblIntereses
        DrNew("CapitalPte") = DblCapitalPte
        DtDesglose.Rows.Add(DrNew)

        Return DtDesglose
    End Function

#End Region

#Region "Funciones Actualizacion Pagos"

    Public Function ActualizacionPagosBBVA(ByVal DteFecha As Date, _
                                          ByVal DblTipoInteresAplicado As Double, _
                                          ByVal DblNPagosAño As Double, _
                                          ByVal LngID As Long, _
                                          ByVal DblRecuperacionCoste As Double, _
                                          ByVal StrIdMoneda As String, _
                                          ByVal BlnPrepagable As Boolean, _
                                          ByVal DblTipoInteres As Double, _
                                          ByVal IntUnidad As Integer, _
                                          ByVal BlnInicial As Boolean, _
                                          ByRef DtCuota As DataTable) As Long
        Dim services As New ServiceProvider
        Dim ClsPagoPer As New PagoPeriodico
        Dim ClsTipoIva As New TipoIva
        Dim ClsMoneda As New Moneda
        Dim ClsPago As New Pago
        Dim Dt As New DataTable
        Dim DtTipoIva As New DataTable
        Dim DtPagoPer As New DataTable
        Dim DtMoneda As New DataTable
        Dim DtMonedaA As New DataTable
        Dim DtMonedaB As New DataTable
        Dim DblPagosRecorridos, Dblk, DblIntereses1, DblIntereses2, DblCuota, DblInteresesFin As Double
        Dim IntDif As Integer
        Dim DteFechaNext, DteFechaPrev As Date
        Dim BlnValorResidual As Boolean

        BlnValorResidual = False
        Dblk = DblTipoInteresAplicado / (DblNPagosAño * 100)
        DtMoneda = ClsMoneda.Filter(New FilterItem("IdMoneda", FilterOperator.Equal, StrIdMoneda, FilterType.String))

        DtMonedaA = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf Moneda.ObtenerMonedaA, Nothing, services)
        DtMonedaB = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf Moneda.ObtenerMonedaB, Nothing, services)

        DtPagoPer = ClsPagoPer.SelOnPrimaryKey(LngID)

        If Length(DtPagoPer.Rows(0)("IDTipoIva") & String.Empty) > 0 Then
            DtTipoIva = ClsTipoIva.SelOnPrimaryKey(DtPagoPer.Rows(0)("IDTipoIva"))
        End If

        Dt = ClsPago.Filter(New FilterItem("IDPagoPeriodo", FilterOperator.Equal, LngID), "FechaVencimiento ASC")
        If Not Dt Is Nothing AndAlso Dt.Rows.Count > 0 Then
            For i As Integer = 0 To Dt.Rows.Count - 1
                DblPagosRecorridos += Dt.Rows(i)("ImpRecuperacionCoste")
                If Dt.Rows(i)("FechaVencimiento") >= DteFecha Then
                    DblIntereses1 = (DblRecuperacionCoste - DblPagosRecorridos) * DblTipoInteres / 36000
                    DteFechaPrev = Dt.Rows(i)("FechaVencimiento")
                    'NEXT
                    If Not (i + 1) = Dt.Rows.Count Then
                        If Length(Dt.Rows(i + 1)("FechaVencimiento") & String.Empty) > 0 Then
                            DteFechaNext = Dt.Rows(i + 1)("FechaVencimiento")
                        Else
                            Select Case IntUnidad
                                Case 3
                                    DteFechaNext = DateAdd("yyyy", 1, DteFechaPrev)
                                Case 2
                                    DteFechaNext = DateAdd("m", 1, DteFechaPrev)
                                Case 1
                                    DteFechaNext = DateAdd("ww", 1, DteFechaPrev)
                                Case 0
                                    DteFechaNext = DateAdd("d", 1, DteFechaPrev)
                            End Select
                        End If
                    Else
                        BlnValorResidual = True
                    End If
                    'PREVIOUS
                    If BlnValorResidual = False Then
                        DblIntereses2 = (DblRecuperacionCoste - DblPagosRecorridos) * (DateDiff("d", Dt.Rows(i)("FechaVencimiento"), DteFechaNext) * DblTipoInteres / 36000)
                        If BlnPrepagable = False Then
                            DblInteresesFin = IIf(DblIntereses2 = 0, 0, DblIntereses2 / (1 + Dblk))
                        Else
                            DblInteresesFin = IIf(DblIntereses2 = 0, 0, DblIntereses2)
                        End If
                        Dt.Rows(i)("ImpIntereses") = xRound(DblInteresesFin, DtMoneda.Rows(0)("NDecimalesImp"))
                        Dt.Rows(i)("ImpInteresesA") = xRound(DblInteresesFin * DtMoneda.Rows(0)("CambioA"), DtMonedaA.Rows(0)("NDecimalesImp"))
                        Dt.Rows(i)("ImpInteresesB") = xRound(DblInteresesFin * DtMoneda.Rows(0)("CambioB"), DtMonedaB.Rows(0)("NDecimalesImp"))
                        DblCuota = Dt.Rows(i)("ImpRecuperacionCoste") + DblInteresesFin
                        Dt.Rows(i)("ImpCuota") = xRound(DblCuota, DtMoneda.Rows(0)("NDecimalesImp"))
                        Dt.Rows(i)("ImpCuotaA") = xRound(DblCuota * DtMoneda.Rows(0)("CambioA"), DtMonedaA.Rows(0)("NDecimalesImp"))
                        Dt.Rows(i)("ImpCuotaB") = xRound(DblCuota * DtMoneda.Rows(0)("CambioB"), DtMonedaB.Rows(0)("NDecimalesImp"))
                        Dt.Rows(i)("ImpVencimiento") = xRound(Dt.Rows(i)("ImpCuota") * (1 + (DtTipoIva.Rows(0)("Factor")) / 100), DtMoneda.Rows(0)("NDecimalesImp"))
                        Dt.Rows(i)("ImpVencimientoA") = xRound(Dt.Rows(i)("ImpCuotaA") * (1 + (DtTipoIva.Rows(0)("Factor")) / 100) * DtMoneda.Rows(0)("CambioA"), DtMonedaA.Rows(0)("NDecimalesImp"))
                        Dt.Rows(i)("ImpVencimientoB") = xRound(Dt.Rows(i)("ImpCuotaA") * (1 + (DtTipoIva.Rows(0)("Factor")) / 100) * DtMoneda.Rows(0)("CambioB"), DtMonedaB.Rows(0)("NDecimalesImp"))
                    End If
                End If
            Next
            If Me.Update(Dt) Is Nothing Then
                Return 0
            Else
                Return -1
            End If
        End If
    End Function

    Public Function ActualizacionPagosBanesto(ByVal DteFecha As Date, _
                                              ByVal DblTipoInteresAplicado As Double, _
                                              ByVal DblNPagosAño As Double, _
                                              ByVal LngID As Long, _
                                              ByVal DblRecuperacionCoste As Double, _
                                              ByVal StrIdMoneda As String, _
                                              ByVal BlnPrepagable As Boolean, _
                                              ByVal DblTipoInteres As Double, _
                                              ByVal IntUnidad As Integer, _
                                              ByVal BlnInicial As Boolean, _
                                              ByRef DtCuota As DataTable) As Long
        Dim services As New ServiceProvider
        Dim ClsPagoPeriodico As New PagoPeriodico
        Dim ClsPago As New Pago
        Dim ClsBPFF As New BancoPropioFormFinanc
        Dim ClsMoneda As New Moneda
        Dim ClsTipoIva As New TipoIva
        Dim DtTipoIva As New DataTable
        Dim Dt As New DataTable
        Dim DtBPFF As New DataTable
        Dim DtPagoPeriodico As New DataTable
        Dim DtPagoVR As New DataTable
        Dim Dt1 As New DataTable
        Dim BlnValorResidual, BlnPrimera, BlnCuotaVR, BlnRedondeoFinal, BlnCarencia As Boolean
        Dim LngNumCuotas, Lng1 As Long
        Dim StrFuncionCuotasSucesivas As String
        Dim Dblk, DblBien, DblAmortizacion, DblRecuperacion, DblIntereses As Double
        Dim DblRecuperacionCosteFinal, DblBienTotal, DblRecCosteUltima As Double
        Dim DblTotalIntereses, DblTotalCuota, DblTotalFinanciar, DblVR, DblImpCuota As Double

        'ActualizacionPagosBanesto = fwmActionError
        If BlnInicial = True Then
            DtCuota = CuotaPrevioActualizacion(LngID, DteFecha, Nz(DblTipoInteresAplicado))
            If Not DtCuota Is Nothing AndAlso DtCuota.Rows.Count > 0 Then
                Return -1
            End If
        Else
            Lng1 = CuotaFinalActualizacion(LngID, DtCuota)
            BlnPrimera = False
            BlnValorResidual = False
            BlnCarencia = False
            DblTotalIntereses = 0
            DblTotalCuota = 0
            DblTotalFinanciar = 0
            Dblk = DblTipoInteresAplicado / (DblNPagosAño * 100)
            Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
            Dim MonInfo As MonedaInfo = Monedas.GetMoneda(StrIdMoneda)
            Dim MonInfoA As MonedaInfo = Monedas.MonedaA
            Dim MonInfoB As MonedaInfo = Monedas.MonedaB
           
            DtPagoPeriodico = ClsPagoPeriodico.SelOnPrimaryKey(LngID)
            If Length(DtPagoPeriodico.Rows(0)("IDTipoIva") & String.Empty) > 0 Then
                DtTipoIva = ClsTipoIva.SelOnPrimaryKey(DtPagoPeriodico.Rows(0)("IDTipoIva"))
            End If
            DtBPFF = ClsBPFF.Filter("fDesgloseSucesivasCuotas", "IDBancoPropio = '" & DtPagoPeriodico.Rows(0)("IDBancoPropio") & "'")
            If Not DtBPFF Is Nothing AndAlso DtBPFF.Rows.Count > 0 Then
                StrFuncionCuotasSucesivas = DtBPFF.Rows(0)("fDesgloseSucesivasCuotas") & String.Empty
            End If
            DblImpCuota = xRound(DtPagoPeriodico.Rows(0)("ImpCuotaPeriodo"), MonInfo.NDecimalesImporte)
            DblBien = DtPagoPeriodico.Rows(0)("ImpRecuperacionCoste") '- rcsPagoPeriodico!AportacionInicial
            DblBienTotal = DtPagoPeriodico.Rows(0)("ImpRecuperacionCoste") '- rcsPagoPeriodico!AportacionInicial
            DblTotalFinanciar = DtPagoPeriodico.Rows(0)("ImpRecuperacionCoste")

            If DtPagoPeriodico.Rows(0)("ValorResidualIgualCota") = False Then BlnValorResidual = True

            'Buscar el valor residual que está en la última cuota del leasing
            DtPagoVR = ClsPago.Filter(New FilterItem("IDPagoPeriodo", FilterOperator.Equal, LngID), "FechaVencimiento DESC")
            If Not DtPagoVR Is Nothing AndAlso DtPagoVR.Rows.Count > 0 Then
                DblVR = Nz(DtPagoVR.Rows(0)("ImpRecuperacionCoste"), 0)
            End If

            Dt = ClsPago.Filter(New FilterItem("IDPagoPeriodo", FilterOperator.Equal, LngID), "FechaVencimiento ASC")
            If Not Dt Is Nothing AndAlso Dt.Rows.Count > 0 Then
                For i As Integer = 0 To Dt.Rows.Count - 1
                    If Dt.Rows(i + 1) Is Nothing AndAlso BlnValorResidual = True Then Exit For
                    If Dt.Rows(i)("FechaVencimiento") >= DteFecha Then
                        While Dt.Rows(i)("RecuperacionCoste") = 0
                            BlnCarencia = True
                            Dt.Rows(i)("ImpIntereses") = xRound(DtPagoPeriodico.Rows(0)("ImpRecuperacionCoste") * Dblk, MonInfo.NDecimalesImporte)
                            Dt.Rows(i)("ImpInteresesA") = xRound(DtPagoPeriodico.Rows(0)("ImpRecuperacionCoste") * Dblk * MonInfo.CambioA, MonInfoA.NDecimalesImporte)
                            Dt.Rows(i)("ImpInteresesB") = xRound(DtPagoPeriodico.Rows(0)("ImpRecuperacionCoste") * Dblk * MonInfo.CambioB, MonInfoB.NDecimalesImporte)
                            Dt.Rows(i)("ImpCuota") = Dt.Rows(i)("ImpIntereses")
                            Dt.Rows(i)("ImpCuotaA") = xRound(Dt.Rows(i)("ImpInteresesA"), MonInfoA.NDecimalesImporte)
                            Dt.Rows(i)("ImpCuotaB") = xRound(Dt.Rows(i)("ImpInteresesB"), MonInfoB.NDecimalesImporte)
                            Dt.Rows(i)("ImpVencimiento") = xRound(Dt.Rows(i)("ImpIntereses") * (1 + (DtTipoIva.Rows(0)("Factor") / 100)), MonInfo.NDecimalesImporte)
                            Dt.Rows(i)("ImpVencimientoA") = xRound(Dt.Rows(i)("ImpInteresesA") * (1 + (DtTipoIva.Rows(0)("Factor") / 100)), MonInfoA.NDecimalesImporte)
                            Dt.Rows(i)("ImpVencimientoB") = xRound(Dt.Rows(i)("ImpInteresesB") * (1 + (DtTipoIva.Rows(0)("Factor") / 100)), MonInfoB.NDecimalesImporte)
                            i += 1
                        End While
                        If BlnCarencia = True Then
                            If Not DtPagoPeriodico.Rows(0)("PagoIntereses") Then
                                Dt.Rows(i)("ImpIntereses") = xRound(DtPagoPeriodico.Rows(0)("ImpRecuperacionCoste") * Dblk, MonInfo.NDecimalesImporte)
                                Dt.Rows(i)("ImpInteresesA") = xRound(DtPagoPeriodico.Rows(0)("ImpRecuperacionCoste") * Dblk * MonInfo.CambioA, MonInfoA.NDecimalesImporte)
                                Dt.Rows(i)("ImpInteresesB") = xRound(DtPagoPeriodico.Rows(0)("ImpRecuperacionCoste") * Dblk * MonInfo.CambioB, MonInfoB.NDecimalesImporte)
                                Dt.Rows(i)("ImpRecuperacionCoste") = xRound(DtPagoPeriodico.Rows(0)("ImpCuotaPeriodo"), MonInfo.NDecimalesImporte)
                                DblBien -= DblAmortizacion
                                Dt.Rows(i)("ImpRecuperacionCosteA") = xRound(DtPagoPeriodico.Rows(0)("ImpCuotaPeriodo") * MonInfo.CambioA, MonInfoA.NDecimalesImporte)
                                Dt.Rows(i)("ImpRecuperacionCosteB") = xRound(DtPagoPeriodico.Rows(0)("ImpCuotaPeriodo") * MonInfo.CambioB, MonInfoB.NDecimalesImporte)
                                Dt.Rows(i)("ImpCuota") = DtPagoPeriodico.Rows(0)("ImpCuotaPeriodo") + Dt.Rows(i)("ImpIntereses")
                                Dt.Rows(i)("ImpCuotaA") = xRound(DtPagoPeriodico.Rows(0)("ImpCuotaPeriodo") + Dt.Rows(i)("ImpIntereses") * MonInfo.CambioA, MonInfoA.NDecimalesImporte)
                                Dt.Rows(i)("ImpCuotaB") = xRound(DtPagoPeriodico.Rows(0)("ImpCuotaPeriodo") + Dt.Rows(i)("ImpIntereses") * MonInfo.CambioB, MonInfoB.NDecimalesImporte)
                                Dt.Rows(i)("ImpVencimiento") = xRound(Dt.Rows(i)("ImpCuota") * (1 + DtTipoIva.Rows(0)("Factor") / 100), MonInfo.NDecimalesImporte)
                                Dt.Rows(i)("ImpVencimientoA") = xRound(Dt.Rows(i)("ImpCuotaA") * (1 + DtTipoIva.Rows(0)("Factor") / 100) * MonInfo.CambioA, MonInfoA.NDecimalesImporte)
                                Dt.Rows(i)("ImpVencimientoB") = xRound(Dt.Rows(i)("ImpCuotaA") * (DtTipoIva.Rows(0)("Factor") / 100) * MonInfo.CambioB, MonInfoB.NDecimalesImporte)
                                DblTotalIntereses += Dt.Rows(i)("ImpIntereses")
                                DblTotalCuota += Dt.Rows(i)("ImpCuota")
                                i += 1
                                BlnCarencia = False
                            End If
                        End If
                        If BlnPrimera = False Then
                            If Not DtPagoPeriodico.Rows(0)("PagoIntereses") Then
                                'En el caso de que el leasing sea PostPagable, los intereses de la primera cuota a modificar,
                                'no varían, por lo tanto los intereses son los mismos que el del cuadro de amortización
                                'inicial, y los demás valores se ajustan.
                                Dt.Rows(i)("ImpRecuperacionCoste") = xRound(DtPagoPeriodico.Rows(0)("ImpCuotaPeriodo") - Dt.Rows(i)("ImpIntereses"), MonInfo.NDecimalesImporte)
                                DblAmortizacion = xRound(DtPagoPeriodico.Rows(0)("ImpCuotaPeriodo") - Dt.Rows(i)("ImpIntereses"), MonInfo.NDecimalesImporte)
                                DblBien -= DblAmortizacion
                                Dt.Rows(i)("ImpRecuperacionCosteA") = xRound((DtPagoPeriodico.Rows(0)("ImpCuotaPeriodo") - Dt.Rows(i)("ImpIntereses")) * MonInfo.CambioA, MonInfoA.NDecimalesImporte)
                                Dt.Rows(i)("ImpRecuperacionCosteB") = xRound((DtPagoPeriodico.Rows(0)("ImpCuotaPeriodo") - Dt.Rows(i)("ImpIntereses")) * MonInfo.CambioB, MonInfoB.NDecimalesImporte)
                                Dt.Rows(i)("ImpCuota") = DtPagoPeriodico.Rows(0)("ImpCuotaPeriodo")
                                Dt.Rows(i)("ImpCuotaA") = xRound(DtPagoPeriodico.Rows(0)("ImpCuotaPeriodo") * MonInfo.CambioA, MonInfoA.NDecimalesImporte)
                                Dt.Rows(i)("ImpCuotaB") = xRound(DtPagoPeriodico.Rows(0)("ImpCuotaPeriodo") * MonInfo.CambioB, MonInfoB.NDecimalesImporte)
                                Dt.Rows(i)("ImpVencimiento") = xRound(DtPagoPeriodico.Rows(0)("ImpCuota") * (1 + DtTipoIva.Rows(0)("Factor") / 100), MonInfo.NDecimalesImporte)
                                Dt.Rows(i)("ImpVencimientoA") = xRound(DtPagoPeriodico.Rows(0)("ImpCuotaA") * (1 + DtTipoIva.Rows(0)("Factor") / 100) * MonInfo.CambioA, MonInfoA.NDecimalesImporte)
                                Dt.Rows(i)("ImpVencimientoB") = xRound(DtPagoPeriodico.Rows(0)("ImpCuotaA") * (1 + DtTipoIva.Rows(0)("Factor") / 100) * MonInfo.CambioB, MonInfoB.NDecimalesImporte)
                                DblTotalIntereses += Dt.Rows(i)("ImpIntereses")
                                DblTotalCuota += Dt.Rows(i)("ImpCuota")
                                LngNumCuotas += 1
                                i += 1
                            End If
                            Dt.Rows(i)("ImpRecuperacionCoste") = xRound(DtPagoPeriodico.Rows(0)("ImpRecuperacionCostePeriodo"), MonInfo.NDecimalesImporte)
                            DblAmortizacion = DtPagoPeriodico.Rows(0)("ImpRecuperacionCostePeriodo")
                            DblBien -= DblAmortizacion
                            Dt.Rows(i)("ImpRecuperacionCosteA") = xRound(DtPagoPeriodico.Rows(0)("ImpRecuperacionCostePeriodo") * MonInfo.CambioA, MonInfoA.NDecimalesImporte)
                            Dt.Rows(i)("ImpRecuperacionCosteB") = xRound(DtPagoPeriodico.Rows(0)("ImpRecuperacionCostePeriodo") * MonInfo.CambioB, MonInfoB.NDecimalesImporte)
                            Dt.Rows(i)("ImpIntereses") = xRound(DtPagoPeriodico.Rows(0)("ImpInteresPeriodo"), MonInfo.NDecimalesImporte)
                            Dt.Rows(i)("ImpInteresesA") = xRound(DtPagoPeriodico.Rows(0)("ImpInteresPeriodo") * MonInfo.CambioA, MonInfoA.NDecimalesImporte)
                            Dt.Rows(i)("ImpInteresesB") = xRound(DtPagoPeriodico.Rows(0)("ImpInteresPeriodo") * MonInfo.CambioB, MonInfoB.NDecimalesImporte)
                            BlnPrimera = True
                        Else
                            If Length(StrFuncionCuotasSucesivas & String.Empty) > 0 Then
                                If Nz(DtPagoPeriodico.Rows(0)("ValorResidualIgualCuota"), False) AndAlso LngNumCuotas = (Nz(DtPagoPeriodico.Rows(0)("NTotalCuotas"), 0) - Nz(DtPagoPeriodico.Rows(0)("NCuotasCarencia"), 0)) - 1 Then
                                    BlnRedondeoFinal = True
                                Else
                                    BlnRedondeoFinal = False
                                End If
                                If Nz(DtPagoPeriodico.Rows(0)("ValorResidualCuota"), False) AndAlso LngNumCuotas = (Nz(DtPagoPeriodico.Rows(0)("NTotalCuotas"), 0) - Nz(DtPagoPeriodico.Rows(0)("NCuotasCarencia"), 0)) Then
                                    BlnCuotaVR = True
                                    Exit For
                                Else
                                    BlnCuotaVR = False
                                End If
                                Dt1 = CallByName(Me, StrFuncionCuotasSucesivas, CallType.Method, Dblk, DblBien, _
                                                 DblAmortizacion, DblImpCuota, 0, BlnRedondeoFinal, 0, _
                                                 MonInfoA.NDecimalesImporte, DblTotalIntereses, _
                                                 DblTotalCuota, DblTotalFinanciar, DblVR, BlnCuotaVR)
                                DblBien = Dt1.Rows(0)("CapitalPte")
                                DblRecuperacion = Dt1.Rows(0)("Recuperacion")
                                DblIntereses = Dt1.Rows(0)("Intereses")
                                Dt.Rows(i)("ImpRecuperacionCoste") = xRound(DblRecuperacion, MonInfo.NDecimalesImporte)
                                DblAmortizacion = Dt.Rows(i)("ImpRecuperacionCoste")
                                Dt.Rows(i)("ImpRecuperacionCosteA") = xRound(DblRecuperacion * MonInfo.CambioA, MonInfoA.NDecimalesImporte)
                                Dt.Rows(i)("ImpRecuperacionCosteB") = xRound(DblRecuperacion * MonInfo.CambioB, MonInfoB.NDecimalesImporte)
                                Dt.Rows(i)("ImpIntereses") = xRound(DblIntereses, MonInfo.NDecimalesImporte)
                                Dt.Rows(i)("ImpInteresesA") = xRound(DblIntereses * MonInfo.CambioA, MonInfoA.NDecimalesImporte)
                                Dt.Rows(i)("ImpInteresesB") = xRound(DblIntereses * MonInfo.CambioB, MonInfoB.NDecimalesImporte)
                            Else
                                Dt.Rows(i)("ImpRecuperacionCoste") = xRound(DtPagoPeriodico.Rows(0)("ImpRecuperacionCostePeriodo"), MonInfo.NDecimalesImporte)
                                Dt.Rows(i)("ImpRecuperacionCosteA") = xRound(DtPagoPeriodico.Rows(0)("ImpRecuperacionCostePeriodo") * MonInfo.CambioA, MonInfoA.NDecimalesImporte)
                                Dt.Rows(i)("ImpRecuperacionCosteB") = xRound(DtPagoPeriodico.Rows(0)("ImpRecuperacionCostePeriodo") * MonInfo.CambioB, MonInfoB.NDecimalesImporte)
                                Dt.Rows(i)("ImpIntereses") = xRound(DtPagoPeriodico.Rows(0)("ImpInteresPeriodo"), MonInfo.NDecimalesImporte)
                                Dt.Rows(i)("ImpInteresesA") = xRound(DtPagoPeriodico.Rows(0)("ImpInteresPeriodo") * MonInfo.CambioA, MonInfoA.NDecimalesImporte)
                                Dt.Rows(i)("ImpInteresesB") = xRound(DtPagoPeriodico.Rows(0)("ImpInteresPeriodo") * MonInfo.CambioB, MonInfoB.NDecimalesImporte)
                            End If
                        End If
                        DblRecuperacionCosteFinal += Dt.Rows(i)("ImpRecuperacionCoste")
                        i += 1
                        If Dt.Rows(i) Is Nothing Then
                            i -= 1
                            If DblBienTotal <> DblRecuperacionCosteFinal Then
                                DblRecCosteUltima = Dt.Rows(i)("ImpRecuperacionCoste") + (DblBienTotal - DblRecuperacionCosteFinal)
                                Dt.Rows(i)("ImpRecuperacionCoste") = xRound(DblRecCosteUltima, MonInfo.NDecimalesImporte)
                                Dt.Rows(i)("ImpRecuperacionCosteA") = xRound(DblRecCosteUltima * MonInfo.CambioA, MonInfoA.NDecimalesImporte)
                                Dt.Rows(i)("ImpRecuperacionCosteB") = xRound(DblRecCosteUltima * MonInfo.CambioB, MonInfoB.NDecimalesImporte)
                            End If
                        Else : i -= 1
                        End If
                        Dt.Rows(i)("ImpCuota") = DtPagoPeriodico.Rows(0)("ImpCuotaPeriodo")
                        Dt.Rows(i)("ImpCuotaA") = xRound(DtPagoPeriodico.Rows(0)("ImpCuotaPeriodo") * MonInfo.CambioA, MonInfoA.NDecimalesImporte)
                        Dt.Rows(i)("ImpCuotaB") = xRound(DtPagoPeriodico.Rows(0)("ImpCuotaPeriodo") * MonInfo.CambioB, MonInfoB.NDecimalesImporte)
                        Dt.Rows(i)("ImpVencimiento") = xRound(Dt.Rows(i)("ImpCuota") * (1 + (DtTipoIva.Rows(0)("Factor") / 100)), MonInfo.NDecimalesImporte)
                        Dt.Rows(i)("ImpVencimientoA") = xRound(Dt.Rows(i)("ImpCuotaA") * (1 + (DtTipoIva.Rows(0)("Factor") / 100)) * MonInfo.CambioA, MonInfoA.NDecimalesImporte)
                        Dt.Rows(i)("ImpVencimientoB") = xRound(Dt.Rows(i)("ImpCuotaA") * (1 + (DtTipoIva.Rows(0)("Factor") / 100)) * MonInfo.CambioB, MonInfoB.NDecimalesImporte)
                    Else
                        DblBien -= Dt.Rows(i)("ImpRecuperacionCoste")
                        DblRecuperacionCosteFinal += Dt.Rows(i)("ImpRecuperacionCoste")
                    End If
                    DblTotalIntereses += Dt.Rows(i)("ImpIntereses")
                    DblTotalCuota += Dt.Rows(i)("ImpIntereses")
                    LngNumCuotas += 1
                Next
                If ClsPago.Update(Dt) Is Nothing Then
                    Return 0
                Else
                    Return -1
                End If
            End If
        End If
    End Function

    Public Function ActualizacionPagosCajastur(ByVal DteFecha As Date, _
                                               ByVal DblTipoInteresAplicado As Double, _
                                               ByVal DblNPagosAño As Double, _
                                               ByVal LngID As Long, _
                                               ByVal DblRecuperacionCoste As Double, _
                                               ByVal StrIdMoneda As String, _
                                               ByVal BlnPrepagable As Boolean, _
                                               ByVal DblTipoInteres As Double, _
                                               ByVal IntUnidad As Integer, _
                                               ByVal BlnInicial As Boolean, _
                                               ByRef DtCuota As DataTable) As Long
        Dim services As New ServiceProvider
        Dim ClsMoneda As New Moneda
        Dim ClsBPFF As New BancoPropioFormFinanc
        Dim ClsPagoPeriodico As New PagoPeriodico
        Dim ClsPago As New Pago
        Dim ClsTipoIva As New TipoIva
        Dim DtPagoVR As New DataTable
        Dim DtBPFF As New DataTable
        Dim DtPagoPeriodico As New DataTable
        Dim Dt As New DataTable
        Dim Dt1 As New DataTable
        Dim DtTipoIva As New DataTable
        Dim Dblk, DblBien, DblAmortizacion, DblRecuperacion, _
        DblIntereses, DblRecuperacionCosteFinal, DblBienTotal, _
        DblRecCosteUltima, DblInteresAnterCuota, DblImpCuota, _
        DblTotalIntereses, DblTotalCuota, DblTotalFinanciar, DblVR As Double
        Dim StrFuncionCuotasSucesivas As String
        Dim BlnPrimera, BlnCuotaVR, BlnCarencia, BlnValorResidual, BlnRedondeoFinal As Boolean
        Dim LngNumCuotas, Lng1 As Long
        ' Los parámetros opcionales, únicamente se pasarán cuando sean actualizaciones

        'return  fwmActionError

        If BlnInicial = True Then
            DtCuota = CuotaPrevioActualizacion(LngID, DteFecha, Nz(DblTipoInteresAplicado))
            If Not DtCuota Is Nothing AndAlso DtCuota.Rows.Count > 0 Then
                Return -1
            End If
        Else
            Lng1 = CuotaFinalActualizacion(LngID, DtCuota)
            BlnPrimera = False
            BlnValorResidual = False
            BlnCarencia = False
            DblTotalIntereses = 0
            DblTotalCuota = 0
            DblTotalFinanciar = 0
            Dblk = DblTipoInteresAplicado / (DblNPagosAño * 100)
       
            Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
            Dim MonInfo As MonedaInfo = Monedas.GetMoneda(StrIdMoneda)
            Dim MonInfoA As MonedaInfo = Monedas.MonedaA
            Dim MonInfoB As MonedaInfo = Monedas.MonedaB

            DtPagoPeriodico = ClsPagoPeriodico.SelOnPrimaryKey(LngID)
            If Length(DtPagoPeriodico.Rows(0)("IDTipoIva") & String.Empty) > 0 Then
                DtTipoIva = ClsTipoIva.SelOnPrimaryKey(DtPagoPeriodico.Rows(0)("IDTipoIva"))
            End If
            DtBPFF = ClsBPFF.Filter(New FilterItem("IDBancoPropio", FilterOperator.Equal, DtPagoPeriodico.Rows(0)("IDBancoPropio")))
            If Not DtBPFF Is Nothing Then
                If DtBPFF.Rows.Count > 0 Then
                    StrFuncionCuotasSucesivas = DtBPFF.Rows(0)("fDesgloseSucesivasCuotas") & String.Empty
                End If
            End If
            DblImpCuota = xRound(DtPagoPeriodico.Rows(0)("ImpCuotaPeriodo"), MonInfo.NDecimalesImporte)
            DblBien = DtPagoPeriodico.Rows(0)("ImpRecuperacionCoste") - DblImpCuota '+ rcsPagoPeriodico!ImpInteresesTotal
            DblBienTotal = DtPagoPeriodico.Rows(0)("ImpRecuperacionCoste") '+ rcsPagoPeriodico!ImpInteresesTotal
            DblTotalFinanciar = DtPagoPeriodico.Rows(0)("ImpRecuperacionCoste")

            DblInteresAnterCuota = 0
            If DtPagoPeriodico.Rows(0)("ValorResidualIgualCota") = False Then BlnValorResidual = True
            'Buscar el valor residual que está en la última cuota del leasing
            DtPagoVR = ClsPago.Filter("*", "IDPagoPeriodo = '" & LngID & "'", "FechaVencimiento DESC")
            If Not DtPagoVR Is Nothing AndAlso DtPagoVR.Rows.Count > 0 Then
                DblVR = Nz(DtPagoVR.Rows(0)("ImpRecuperacionCoste"), 0)
            End If
            Dt = ClsPago.Filter(New FilterItem("IDPagoPeriodo", FilterOperator.Equal, LngID), "FechaVencimiento ASC")
            If Not Dt Is Nothing AndAlso Dt.Rows.Count > 0 Then
                For i As Integer = 0 To Dt.Rows.Count - 1
                    If Dt.Rows(i + 1) Is Nothing And BlnValorResidual = True Then Exit For
                    If Dt.Rows(i)("FechaVencimiento") >= DteFecha Then
                        Do While Dt.Rows(i)("ImpRecuperacionCoste") = 0
                            BlnCarencia = True
                            Dt.Rows(i)("ImpIntereses") = xRound(DtPagoPeriodico.Rows(0)("ImpRecuperacionCoste") * Dblk, MonInfo.NDecimalesImporte)
                            Dt.Rows(i)("ImpInteresesA") = xRound(DtPagoPeriodico.Rows(0)("ImpRecuperacionCoste") * Dblk * MonInfo.CambioA, MonInfoA.NDecimalesImporte)
                            Dt.Rows(i)("ImpInteresesB") = xRound(DtPagoPeriodico.Rows(0)("ImpRecuperacionCoste") * Dblk * MonInfo.CambioB, MonInfoB.NDecimalesImporte)
                            Dt.Rows(i)("ImpCuota") = Dt.Rows(i)("ImpIntereses")
                            Dt.Rows(i)("ImpCuotaA") = xRound(Dt.Rows(i)("ImpInteresesA"), MonInfoA.NDecimalesImporte)
                            Dt.Rows(i)("ImpCuotaB") = xRound(Dt.Rows(i)("ImpInteresesB"), MonInfoB.NDecimalesImporte)
                            Dt.Rows(i)("ImpVencimiento") = xRound(Dt.Rows(i)("ImpIntereses") * (1 + DtTipoIva.Rows(0)("Factor") / 100), MonInfo.NDecimalesImporte)
                            Dt.Rows(i)("ImpVencimientoA") = xRound(Dt.Rows(i)("ImpInteresesA") * (1 + DtTipoIva.Rows(0)("Factor") / 100) * MonInfo.CambioA, MonInfoA.NDecimalesImporte)
                            Dt.Rows(i)("ImpVencimientoB") = xRound(Dt.Rows(i)("ImpInteresesB") * (1 + DtTipoIva.Rows(0)("Factor") / 100) * MonInfo.CambioB, MonInfoB.NDecimalesImporte)
                            DblTotalIntereses += Dt.Rows(i)("ImpIntereses")
                            DblTotalCuota += Dt.Rows(i)("ImpCuota")
                            i += 1
                        Loop
                        If Nz(DtPagoPeriodico.Rows(0)("ValorResidualIgualCuota"), False) AndAlso LngNumCuotas = (Nz(DtPagoPeriodico.Rows(0)("NTotalCuotas"), 0) - Nz(DtPagoPeriodico.Rows(0)("NCuotasCarencia"), 0) - 1) Then
                            BlnRedondeoFinal = True
                        Else
                            BlnRedondeoFinal = False
                        End If
                        If Nz(DtPagoPeriodico.Rows(0)("ValorResidualIgualCuota"), False) AndAlso LngNumCuotas = (Nz(DtPagoPeriodico.Rows(0)("NTotalCuotas"), 0) - Nz(DtPagoPeriodico.Rows(0)("NCuotasCarencia"), 0)) Then
                            BlnCuotaVR = True
                            Exit For
                        Else
                            BlnCuotaVR = False
                        End If
                        If BlnCarencia = True Then
                            If Not DtPagoPeriodico.Rows(0)("PagoIntereses") Then
                                Dt.Rows(i)("ImpIntereses") = xRound(DtPagoPeriodico.Rows(0)("ImpRecuperacionCoste") * Dblk, MonInfo.NDecimalesImporte)
                                Dt.Rows(i)("ImpInteresesA") = xRound(DtPagoPeriodico.Rows(0)("ImpRecuperacionCoste") * Dblk * MonInfo.CambioA, MonInfoA.NDecimalesImporte)
                                Dt.Rows(i)("ImpInteresesB") = xRound(DtPagoPeriodico.Rows(0)("ImpRecuperacionCoste") * Dblk * MonInfo.CambioB, MonInfoB.NDecimalesImporte)
                                Dt.Rows(i)("ImpRecuperacionCoste") = xRound(DtPagoPeriodico.Rows(0)("ImpCuotaPeriodo"), MonInfo.NDecimalesImporte)
                                DblAmortizacion = xRound(DtPagoPeriodico.Rows(0)("ImpCuotaPeriodo") - Dt.Rows(i)("ImpIntereses"), MonInfo.NDecimalesImporte)
                                DblBien -= DblAmortizacion
                                Dt.Rows(i)("ImpRecuperacionCosteA") = xRound((DtPagoPeriodico.Rows(0)("ImpCuotaPeriodo") * MonInfo.CambioA), MonInfoA.NDecimalesImporte)
                                Dt.Rows(i)("ImpRecuperacionCosteB") = xRound((DtPagoPeriodico.Rows(0)("ImpCuotaPeriodo") * MonInfo.CambioB), MonInfoB.NDecimalesImporte)
                                Dt.Rows(i)("ImpCuota") = DtPagoPeriodico.Rows(0)("ImpCuotaPeriodo") + Dt.Rows(i)("ImpIntereses")
                                Dt.Rows(i)("ImpCuotaA") = xRound((DtPagoPeriodico.Rows(0)("ImpCuotaPeriodo") + Dt.Rows(i)("ImpIntereses") * MonInfo.CambioA), MonInfoA.NDecimalesImporte)
                                Dt.Rows(i)("ImpCuotaB") = xRound((DtPagoPeriodico.Rows(0)("ImpCuotaPeriodo") + Dt.Rows(i)("ImpIntereses") + MonInfo.CambioB), MonInfoB.NDecimalesImporte)
                                Dt.Rows(i)("ImpVencimiento") = xRound(Dt.Rows(i)("ImpCuota") * (1 + (DtTipoIva.Rows(0)("Factor") / 100)), MonInfo.NDecimalesImporte)
                                Dt.Rows(i)("ImpVencimientoA") = xRound(Dt.Rows(i)("ImpCuotaA") * (1 + (DtTipoIva.Rows(0)("Factor") / 100)) * MonInfo.CambioA, MonInfoA.NDecimalesImporte)
                                Dt.Rows(i)("ImpVencimientoB") = xRound(Dt.Rows(i)("ImpCuotaA") * (1 + (DtTipoIva.Rows(0)("Factor") / 100)) * MonInfo.CambioB, MonInfoB.NDecimalesImporte)
                                DblTotalIntereses += Dt.Rows(i)("ImpIntereses")
                                DblTotalCuota += Dt.Rows(i)("ImpCuota")
                                i += 1
                                BlnCarencia = False
                            End If
                        End If
                        If Length(StrFuncionCuotasSucesivas & String.Empty) > 0 Then
                            Dt1 = CallByName(Me, StrFuncionCuotasSucesivas, CallType.Method, Dblk, DblBien, _
                                             DblAmortizacion, DblImpCuota, LngNumCuotas, BlnRedondeoFinal, _
                                             DblInteresAnterCuota, MonInfo.NDecimalesImporte, DblTotalIntereses, _
                                             DblTotalCuota, DblTotalFinanciar, DblVR, BlnCuotaVR)
                            DblBien = Dt1.Rows(0)("CapitalPte")
                            DblRecuperacion = Dt1.Rows(0)("Recuperacion")
                            DblIntereses = Dt1.Rows(0)("Intereses")
                            DblInteresAnterCuota = Dt1.Rows(0)("Intereses")
                            Dt.Rows(i)("ImpRecuperacionCoste") = xRound(DblRecuperacion, MonInfo.NDecimalesImporte)
                            DblAmortizacion = Dt.Rows(0)("ImpRecuperacionCoste")
                            Dt.Rows(i)("ImpRecuperacionCosteA") = xRound(DblRecuperacion * MonInfo.CambioA, MonInfoA.NDecimalesImporte)
                            Dt.Rows(i)("ImpRecuperacionCosteB") = xRound(DblRecuperacion * MonInfo.CambioB, MonInfoB.NDecimalesImporte)
                            Dt.Rows(i)("ImpIntereses") = xRound(DblIntereses, MonInfo.NDecimalesImporte)
                            Dt.Rows(i)("ImpInteresesA") = xRound(DblIntereses * MonInfo.CambioA, MonInfoA.NDecimalesImporte)
                            Dt.Rows(i)("ImpInteresesB") = xRound(DblIntereses * MonInfo.CambioB, MonInfoB.NDecimalesImporte)
                        Else
                            Dt.Rows(i)("ImpRecuperacionCoste") = xRound(DtPagoPeriodico.Rows(0)("ImpRecuperacionCostePeriodo"), MonInfo.NDecimalesImporte)
                            Dt.Rows(i)("ImpRecuperacionCosteA") = xRound(DtPagoPeriodico.Rows(0)("ImpRecuperacionCostePeriodo") * MonInfo.CambioA, MonInfoA.NDecimalesImporte)
                            Dt.Rows(i)("ImpRecuperacionCosteB") = xRound(DtPagoPeriodico.Rows(0)("ImpRecuperacionCostePeriodo") * MonInfo.CambioB, MonInfoB.NDecimalesImporte)
                            Dt.Rows(i)("ImpIntereses") = xRound(DtPagoPeriodico.Rows(0)("ImpInteresPeriodo"), MonInfo.NDecimalesImporte)
                            Dt.Rows(i)("ImpInteresesA") = xRound(DtPagoPeriodico.Rows(0)("ImpInteresPeriodo") * MonInfo.CambioA, MonInfoA.NDecimalesImporte)
                            Dt.Rows(i)("ImpInteresesB") = xRound(DtPagoPeriodico.Rows(0)("ImpInteresPeriodo") * MonInfo.CambioB, MonInfoB.NDecimalesImporte)
                        End If
                        DblRecuperacionCosteFinal += Dt.Rows(i)("ImpRecuperacionCoste")
                        i += 1
                        If Dt.Rows(i) Is Nothing Then
                            i -= 1
                            If DblBienTotal <> DblRecuperacionCosteFinal Then
                                DblRecCosteUltima = Dt.Rows(i)("ImpRecuperacionCoste") + (DblBienTotal - DblRecuperacionCosteFinal)
                                Dt.Rows(i)("ImpRecuperacionCoste") = xRound(DblRecCosteUltima, MonInfo.NDecimalesImporte)
                                Dt.Rows(i)("ImpRecuperacionCosteA") = xRound(DblRecCosteUltima * MonInfo.CambioA, MonInfoA.NDecimalesImporte)
                                Dt.Rows(i)("ImpRecuperacionCosteB") = xRound(DblRecCosteUltima * MonInfo.CambioB, MonInfoB.NDecimalesImporte)
                            End If
                        Else : i -= 1
                        End If
                        Dt.Rows(i)("ImpCuota") = DtPagoPeriodico.Rows(0)("ImpCuotaPeriodo")
                        Dt.Rows(i)("ImpCuotaA") = xRound(DtPagoPeriodico.Rows(0)("ImpCuotaPeriodo") * MonInfo.CambioA, MonInfoA.NDecimalesImporte)
                        Dt.Rows(i)("imPCuotaB") = xRound(DtPagoPeriodico.Rows(0)("ImpCuotaPeriodo") * MonInfo.CambioB, MonInfoB.NDecimalesImporte)
                        Dt.Rows(i)("ImpVencimiento") = xRound(Dt.Rows(i)("ImpCuota") * (1 + (DtTipoIva.Rows(0)("Factor") / 100)), MonInfo.NDecimalesImporte)
                        Dt.Rows(i)("ImpVencimientoA") = xRound(Dt.Rows(i)("ImpCuotaA") * (1 + (DtTipoIva.Rows(0)("Factor") / 100)) * MonInfo.CambioA, MonInfoA.NDecimalesImporte)
                        Dt.Rows(i)("ImpVencimientoB") = xRound(Dt.Rows(i)("ImpCuotaA") * (1 + (DtTipoIva.Rows(0)("Factor") / 100)) * MonInfo.CambioB, MonInfoB.NDecimalesImporte)
                    Else
                        DblBien -= Dt.Rows(i)("ImpRecuperacionCoste")
                        DblRecuperacionCosteFinal += Dt.Rows(i)("ImpRecuperacionCoste")
                    End If
                    DblTotalIntereses += Dt.Rows(i)("ImpIntereses")
                    DblTotalCuota += Dt.Rows(i)("ImpCuota")
                    LngNumCuotas += 1
                Next
                If ClsPago.Update(Dt) Is Nothing Then
                    Return 0
                Else
                    Return -1
                End If
            End If
        End If
    End Function

    Public Function ActualizacionPagosCajaAsturias(ByVal DteFecha As Date, _
                                              ByVal DblTipoInteresAplicado As Double, _
                                              ByVal DblNPagosAño As Double, _
                                              ByVal LngID As Long, _
                                              ByVal DblRecuperacionCoste As Double, _
                                              ByVal StrIdMoneda As String, _
                                              ByVal BlnPrepagable As Boolean, _
                                              ByVal DblTipoInteres As Double, _
                                              ByVal IntUnidad As Integer, _
                                              ByVal BlnInicial As Boolean, _
                                              ByRef DtCuota As DataTable) As Long
        Dim services As New ServiceProvider

        Dim ClsBPFF As New BancoPropioFormFinanc
        Dim ClsPagoPeriodico As New PagoPeriodico
        Dim ClsPago As New Pago
        Dim ClsTipoIva As New TipoIva
        Dim DtPagoVR As New DataTable
        Dim DtBPFF As New DataTable
        Dim DtPagoPeriodico As New DataTable
        Dim Dt As New DataTable
        Dim Dt1 As New DataTable
        Dim DtTipoIva As New DataTable
        Dim Dblk, DblBien, DblAmortizacion, DblRecuperacion, _
        DblIntereses, DblRecuperacionCosteFinal, DblBienTotal, _
        DblRecCosteUltima, DblInteresAnterCuota, DblImpCuota, _
        DblTotalIntereses, DblTotalCuota, DblTotalFinanciar, DblVR As Double
        Dim StrFuncionCuotasSucesivas As String
        Dim BlnPrimera, BlnCuotaVR, BlnCarencia, BlnValorResidual, BlnRedondeoFinal As Boolean
        Dim LngNumCuotas, Lng1 As Long
        ' Los parámetros opcionales, únicamente se pasarán cuando sean actualizaciones

        'return  fwmActionError

        If BlnInicial = True Then
            DtCuota = CuotaPrevioActualizacion(LngID, DteFecha, Nz(DblTipoInteresAplicado))
            If Not DtCuota Is Nothing AndAlso DtCuota.Rows.Count > 0 Then
                Return -1
            End If
        Else
            Lng1 = CuotaFinalActualizacion(LngID, DtCuota)
            BlnPrimera = False
            BlnValorResidual = False
            BlnCarencia = False
            DblTotalIntereses = 0
            DblTotalCuota = 0
            DblTotalFinanciar = 0
            Dblk = DblTipoInteresAplicado / (DblNPagosAño * 100)
            Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
            Dim MonInfo As MonedaInfo = Monedas.GetMoneda(StrIdMoneda)
            Dim MonInfoA As MonedaInfo = Monedas.MonedaA
            Dim MonInfoB As MonedaInfo = Monedas.MonedaB
            DtPagoPeriodico = ClsPagoPeriodico.SelOnPrimaryKey(LngID)
            If Length(DtPagoPeriodico.Rows(0)("IDTipoIva") & String.Empty) > 0 Then
                DtTipoIva = ClsTipoIva.SelOnPrimaryKey(DtPagoPeriodico.Rows(0)("IDTipoIva"))
            End If
            DtBPFF = ClsBPFF.Filter(New FilterItem("IDBancoPropio", FilterOperator.Equal, DtPagoPeriodico.Rows(0)("IDBancoPropio")))
            If Not DtBPFF Is Nothing Then
                If DtBPFF.Rows.Count > 0 Then
                    StrFuncionCuotasSucesivas = DtBPFF.Rows(0)("fDesgloseSucesivasCuotas") & String.Empty
                End If
            End If
            DblImpCuota = xRound(DtPagoPeriodico.Rows(0)("ImpCuotaPeriodo"), MonInfo.NDecimalesImporte)
            DblBien = DtPagoPeriodico.Rows(0)("ImpRecuperacionCoste") - DblImpCuota '+ rcsPagoPeriodico!ImpInteresesTotal
            DblBienTotal = DtPagoPeriodico.Rows(0)("ImpRecuperacionCoste") '+ rcsPagoPeriodico!ImpInteresesTotal
            DblTotalFinanciar = DtPagoPeriodico.Rows(0)("ImpRecuperacionCoste")

            DblInteresAnterCuota = 0
            If DtPagoPeriodico.Rows(0)("ValorResidualIgualCota") = False Then BlnValorResidual = True
            'Buscar el valor residual que está en la última cuota del leasing
            DtPagoVR = ClsPago.Filter("*", "IDPagoPeriodo = '" & LngID & "'", "FechaVencimiento DESC")
            If Not DtPagoVR Is Nothing AndAlso DtPagoVR.Rows.Count > 0 Then
                DblVR = Nz(DtPagoVR.Rows(0)("ImpRecuperacionCoste"), 0)
            End If
            Dt = ClsPago.Filter(New FilterItem("IDPagoPeriodo", FilterOperator.Equal, LngID), "FechaVencimiento ASC")
            If Not Dt Is Nothing AndAlso Dt.Rows.Count > 0 Then
                For i As Integer = 0 To Dt.Rows.Count - 1
                    If Dt.Rows(i + 1) Is Nothing And BlnValorResidual = True Then Exit For
                    If Dt.Rows(i)("FechaVencimiento") >= DteFecha Then
                        Do While Dt.Rows(i)("ImpRecuperacionCoste") = 0
                            BlnCarencia = True
                            Dt.Rows(i)("ImpIntereses") = xRound(DtPagoPeriodico.Rows(0)("ImpRecuperacionCoste") * Dblk, MonInfo.NDecimalesImporte)
                            Dt.Rows(i)("ImpInteresesA") = xRound(DtPagoPeriodico.Rows(0)("ImpRecuperacionCoste") * Dblk * MonInfo.CambioA, MonInfoA.NDecimalesImporte)
                            Dt.Rows(i)("ImpInteresesB") = xRound(DtPagoPeriodico.Rows(0)("ImpRecuperacionCoste") * Dblk * MonInfo.CambioB, MonInfoB.NDecimalesImporte)
                            Dt.Rows(i)("ImpCuota") = Dt.Rows(i)("ImpIntereses")
                            Dt.Rows(i)("ImpCuotaA") = xRound(Dt.Rows(i)("ImpInteresesA"), MonInfoA.NDecimalesImporte)
                            Dt.Rows(i)("ImpCuotaB") = xRound(Dt.Rows(i)("ImpInteresesB"), MonInfoB.NDecimalesImporte)
                            Dt.Rows(i)("ImpVencimiento") = xRound(Dt.Rows(i)("ImpIntereses") * (1 + DtTipoIva.Rows(0)("Factor") / 100), MonInfo.NDecimalesImporte)
                            Dt.Rows(i)("ImpVencimientoA") = xRound(Dt.Rows(i)("ImpInteresesA") * (1 + DtTipoIva.Rows(0)("Factor") / 100) * MonInfo.CambioA, MonInfoA.NDecimalesImporte)
                            Dt.Rows(i)("ImpVencimientoB") = xRound(Dt.Rows(i)("ImpInteresesB") * (1 + DtTipoIva.Rows(0)("Factor") / 100) * MonInfo.CambioB, MonInfoB.NDecimalesImporte)
                            DblTotalIntereses += Dt.Rows(i)("ImpIntereses")
                            DblTotalCuota += Dt.Rows(i)("ImpCuota")
                            i += 1
                        Loop
                        If Nz(DtPagoPeriodico.Rows(0)("ValorResidualIgualCuota"), False) AndAlso LngNumCuotas = (Nz(DtPagoPeriodico.Rows(0)("NTotalCuotas"), 0) - Nz(DtPagoPeriodico.Rows(0)("NCuotasCarencia"), 0) - 1) Then
                            BlnRedondeoFinal = True
                        Else
                            BlnRedondeoFinal = False
                        End If
                        If Nz(DtPagoPeriodico.Rows(0)("ValorResidualIgualCuota"), False) AndAlso LngNumCuotas = (Nz(DtPagoPeriodico.Rows(0)("NTotalCuotas"), 0) - Nz(DtPagoPeriodico.Rows(0)("NCuotasCarencia"), 0)) Then
                            BlnCuotaVR = True
                            Exit For
                        Else
                            BlnCuotaVR = False
                        End If
                        If BlnCarencia = True Then
                            If Not DtPagoPeriodico.Rows(0)("PagoIntereses") Then
                                Dt.Rows(i)("ImpIntereses") = xRound(DtPagoPeriodico.Rows(0)("ImpRecuperacionCoste") * Dblk, MonInfo.NDecimalesImporte)
                                Dt.Rows(i)("ImpInteresesA") = xRound(DtPagoPeriodico.Rows(0)("ImpRecuperacionCoste") * Dblk * MonInfo.CambioA, MonInfoA.NDecimalesImporte)
                                Dt.Rows(i)("ImpInteresesB") = xRound(DtPagoPeriodico.Rows(0)("ImpRecuperacionCoste") * Dblk * MonInfo.CambioB, MonInfoB.NDecimalesImporte)
                                Dt.Rows(i)("ImpRecuperacionCoste") = xRound(DtPagoPeriodico.Rows(0)("ImpCuotaPeriodo"), MonInfo.NDecimalesImporte)
                                DblAmortizacion = xRound(DtPagoPeriodico.Rows(0)("ImpCuotaPeriodo") - Dt.Rows(i)("ImpIntereses"), MonInfo.NDecimalesImporte)
                                DblBien -= DblAmortizacion
                                Dt.Rows(i)("ImpRecuperacionCosteA") = xRound((DtPagoPeriodico.Rows(0)("ImpCuotaPeriodo") * MonInfo.CambioA), MonInfoA.NDecimalesImporte)
                                Dt.Rows(i)("ImpRecuperacionCosteB") = xRound((DtPagoPeriodico.Rows(0)("ImpCuotaPeriodo") * MonInfo.CambioB), MonInfoB.NDecimalesImporte)
                                Dt.Rows(i)("ImpCuota") = DtPagoPeriodico.Rows(0)("ImpCuotaPeriodo") + Dt.Rows(i)("ImpIntereses")
                                Dt.Rows(i)("ImpCuotaA") = xRound((DtPagoPeriodico.Rows(0)("ImpCuotaPeriodo") + Dt.Rows(i)("ImpIntereses") * MonInfo.CambioA), MonInfoA.NDecimalesImporte)
                                Dt.Rows(i)("ImpCuotaB") = xRound((DtPagoPeriodico.Rows(0)("ImpCuotaPeriodo") + Dt.Rows(i)("ImpIntereses") + MonInfo.CambioB), MonInfoB.NDecimalesImporte)
                                Dt.Rows(i)("ImpVencimiento") = xRound(Dt.Rows(i)("ImpCuota") * (1 + (DtTipoIva.Rows(0)("Factor") / 100)), MonInfo.NDecimalesImporte)
                                Dt.Rows(i)("ImpVencimientoA") = xRound(Dt.Rows(i)("ImpCuotaA") * (1 + (DtTipoIva.Rows(0)("Factor") / 100)) * MonInfo.CambioA, MonInfoA.NDecimalesImporte)
                                Dt.Rows(i)("ImpVencimientoB") = xRound(Dt.Rows(i)("ImpCuotaA") * (1 + (DtTipoIva.Rows(0)("Factor") / 100)) * MonInfo.CambioB, MonInfoB.NDecimalesImporte)
                                DblTotalIntereses += Dt.Rows(i)("ImpIntereses")
                                DblTotalCuota += Dt.Rows(i)("ImpCuota")
                                i += 1
                                BlnCarencia = False
                            End If
                        End If
                        If Length(StrFuncionCuotasSucesivas & String.Empty) > 0 Then
                            Dt1 = CallByName(Me, StrFuncionCuotasSucesivas, CallType.Method, Dblk, DblBien, _
                                             DblAmortizacion, DblImpCuota, LngNumCuotas, BlnRedondeoFinal, _
                                             DblInteresAnterCuota, MonInfo.NDecimalesImporte, DblTotalIntereses, _
                                             DblTotalCuota, DblTotalFinanciar, DblVR, BlnCuotaVR)
                            DblBien = Dt1.Rows(0)("CapitalPte")
                            DblRecuperacion = Dt1.Rows(0)("Recuperacion")
                            DblIntereses = Dt1.Rows(0)("Intereses")
                            DblInteresAnterCuota = Dt1.Rows(0)("Intereses")
                            Dt.Rows(i)("ImpRecuperacionCoste") = xRound(DblRecuperacion, MonInfo.NDecimalesImporte)
                            DblAmortizacion = Dt.Rows(0)("ImpRecuperacionCoste")
                            Dt.Rows(i)("ImpRecuperacionCosteA") = xRound(DblRecuperacion * MonInfo.CambioA, MonInfoA.NDecimalesImporte)
                            Dt.Rows(i)("ImpRecuperacionCosteB") = xRound(DblRecuperacion * MonInfo.CambioB, MonInfoB.NDecimalesImporte)
                            Dt.Rows(i)("ImpIntereses") = xRound(DblIntereses, MonInfo.NDecimalesImporte)
                            Dt.Rows(i)("ImpInteresesA") = xRound(DblIntereses * MonInfo.CambioA, MonInfoA.NDecimalesImporte)
                            Dt.Rows(i)("ImpInteresesB") = xRound(DblIntereses * MonInfo.CambioB, MonInfoB.NDecimalesImporte)
                        Else
                            Dt.Rows(i)("ImpRecuperacionCoste") = xRound(DtPagoPeriodico.Rows(0)("ImpRecuperacionCostePeriodo"), MonInfo.NDecimalesImporte)
                            Dt.Rows(i)("ImpRecuperacionCosteA") = xRound(DtPagoPeriodico.Rows(0)("ImpRecuperacionCostePeriodo") * MonInfo.CambioA, MonInfoA.NDecimalesImporte)
                            Dt.Rows(i)("ImpRecuperacionCosteB") = xRound(DtPagoPeriodico.Rows(0)("ImpRecuperacionCostePeriodo") * MonInfo.CambioB, MonInfoB.NDecimalesImporte)
                            Dt.Rows(i)("ImpIntereses") = xRound(DtPagoPeriodico.Rows(0)("ImpInteresPeriodo"), MonInfo.NDecimalesImporte)
                            Dt.Rows(i)("ImpInteresesA") = xRound(DtPagoPeriodico.Rows(0)("ImpInteresPeriodo") * MonInfo.CambioA, MonInfoA.NDecimalesImporte)
                            Dt.Rows(i)("ImpInteresesB") = xRound(DtPagoPeriodico.Rows(0)("ImpInteresPeriodo") * MonInfo.CambioB, MonInfoB.NDecimalesImporte)
                        End If
                        DblRecuperacionCosteFinal += Dt.Rows(i)("ImpRecuperacionCoste")
                        i += 1
                        If Dt.Rows(i) Is Nothing Then
                            i -= 1
                            If DblBienTotal <> DblRecuperacionCosteFinal Then
                                DblRecCosteUltima = Dt.Rows(i)("ImpRecuperacionCoste") + (DblBienTotal - DblRecuperacionCosteFinal)
                                Dt.Rows(i)("ImpRecuperacionCoste") = xRound(DblRecCosteUltima, MonInfo.NDecimalesImporte)
                                Dt.Rows(i)("ImpRecuperacionCosteA") = xRound(DblRecCosteUltima * MonInfo.CambioA, MonInfoA.NDecimalesImporte)
                                Dt.Rows(i)("ImpRecuperacionCosteB") = xRound(DblRecCosteUltima * MonInfo.CambioB, MonInfoB.NDecimalesImporte)
                            End If
                        Else : i -= 1
                        End If
                        Dt.Rows(i)("ImpCuota") = DtPagoPeriodico.Rows(0)("ImpCuotaPeriodo")
                        Dt.Rows(i)("ImpCuotaA") = xRound(DtPagoPeriodico.Rows(0)("ImpCuotaPeriodo") * MonInfo.CambioA, MonInfoA.NDecimalesImporte)
                        Dt.Rows(i)("imPCuotaB") = xRound(DtPagoPeriodico.Rows(0)("ImpCuotaPeriodo") * MonInfo.CambioB, MonInfoB.NDecimalesImporte)
                        Dt.Rows(i)("ImpVencimiento") = xRound(Dt.Rows(i)("ImpCuota") * (1 + (DtTipoIva.Rows(0)("Factor") / 100)), MonInfo.NDecimalesImporte)
                        Dt.Rows(i)("ImpVencimientoA") = xRound(Dt.Rows(i)("ImpCuotaA") * (1 + (DtTipoIva.Rows(0)("Factor") / 100)) * MonInfo.CambioA, MonInfoA.NDecimalesImporte)
                        Dt.Rows(i)("ImpVencimientoB") = xRound(Dt.Rows(i)("ImpCuotaA") * (1 + (DtTipoIva.Rows(0)("Factor") / 100)) * MonInfo.CambioB, MonInfoB.NDecimalesImporte)
                    Else
                        DblBien -= Dt.Rows(i)("ImpRecuperacionCoste")
                        DblRecuperacionCosteFinal += Dt.Rows(i)("ImpRecuperacionCoste")
                    End If
                    DblTotalIntereses += Dt.Rows(i)("ImpIntereses")
                    DblTotalCuota += Dt.Rows(i)("ImpCuota")
                    LngNumCuotas += 1
                Next
                If ClsPago.Update(Dt) Is Nothing Then
                    Return 0
                Else
                    Return -1
                End If
            End If
        End If
    End Function

#End Region

#Region "Funciones Cuotas"

    Private Function CuotaPrevioActualizacion(ByVal LngID As Long, _
                                          Optional ByVal DtmFechaAct As Date = cnMinDate, _
                                          Optional ByVal DblIntAplicado As Double = 0) As DataTable
        Dim services As New ServiceProvider
        Dim ClsBPFF As New BancoPropioFormFinanc
        Dim ClsPago As New Pago
        Dim ClsPagoPeriodico As New PagoPeriodico
        Dim DtBPFF As New DataTable
        Dim DtPagos As New DataTable
        Dim Dt As New DataTable
        Dim DtApp As New DataTable
        Dim DtPago As New DataTable
        Dim DtPagoPeriodico As New DataTable
        Dim DblBienTotal, DblPagos, DblNumCuotas, DblNumCoutasCarencia, DblInteresAplicado, _
        DblPagosAlAño, DblValorResidual, DblCuotaPeriodo, DblRecCostePeriodo, _
        DblInteresPeriodo, DblImpInteresesTotal As Double
        Dim BlnValorResidualIgualCuota As Boolean
        Dim FCalculoCuota, FDesglosePrimCuota As String
        ' Los parámetros opcionales, únicamente se pasarán cuando sean actualizaciones

        CuotaPrevioActualizacion = Nothing
        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
        Dim MonInfoA As MonedaInfo = Monedas.MonedaA
        DtPagoPeriodico = ClsPagoPeriodico.SelOnPrimaryKey(LngID)
        DblPagosAlAño = Nz(DtPagoPeriodico.Rows(0)("NPagosAlAño"))
        DblImpInteresesTotal = Nz(DtPagoPeriodico.Rows(0)("ImpInteresesTotal"))
        DblValorResidual = Nz(DtPagoPeriodico.Rows(0)("ValorResidual"))
        BlnValorResidualIgualCuota = CBool(DtPagoPeriodico.Rows(0)("ValorResidualIgualCuota"))
        DblCuotaPeriodo = Nz(DtPagoPeriodico.Rows(0)("ImpCuotaPeriodo"))
        DblRecCostePeriodo = Nz(DtPagoPeriodico.Rows(0)("ImpRecuperacionCostePeriodo"))
        DblInteresPeriodo = Nz(DtPagoPeriodico.Rows(0)("ImpInteresPeriodo"))

        DtApp = AdminData.Filter("vLeasingTotalInmovilizado", "*", "IDInmovilizado = '" & DtPagoPeriodico.Rows(0)("IDInmovilizado") & "'")
        If Not DtApp Is Nothing AndAlso DtApp.Rows.Count > 0 Then
            DblBienTotal = Nz(DtApp.Rows(0)("TotalRevalorizadoA"))
            If DtPagoPeriodico.Rows(0)("AportacionInicial") > 0 Then DblBienTotal -= DtPagoPeriodico.Rows(0)("AportacionInicial")
        End If

        DblPagos = 0
        'Recalcular el Bien Total sobre el que parte ahora después de la actualización
        Dim FilPago As New Filter
        FilPago.Add("IDPagoPeriodo", FilterOperator.Equal, LngID, FilterType.Numeric)
        FilPago.Add("FechaVencimiento", FilterOperator.LessThan, DtmFechaAct, FilterType.DateTime)
        DtPago = ClsPago.Filter(FilPago)
        If Not DtPago Is Nothing AndAlso DtPago.Rows.Count > 0 Then
            For Each Dr As DataRow In DtPago.Select
                DblPagos += DtPago.Rows(0)("ImpRecuperacionCoste")
            Next
            DblBienTotal -= DblPagos
            'Recalcular el Número de cuotas
            DblNumCuotas = CDbl(DtPagoPeriodico.Rows(0)("NTotalCuotas")) - DtPago.Rows.Count
            'Recalcular el número de cuotas de carencia
            If DtPago.Rows.Count >= CDbl(DtPagoPeriodico.Rows(0)("NCuotasCarencia")) Then
                DblNumCoutasCarencia = 0
            Else
                DblNumCoutasCarencia = CDbl(DtPagoPeriodico.Rows(0)("NCuotasCarencia")) - DtPago.Rows.Count
            End If
        End If
        DblInteresAplicado = DblIntAplicado

        DtBPFF = ClsBPFF.Filter(New FilterItem("IDBancoPropio", FilterOperator.Equal, DtPagoPeriodico.Rows(0)("IDBancoPropio")))
        If Not DtBPFF Is Nothing AndAlso DtBPFF.Rows.Count > 0 Then
            If Length(DtBPFF.Rows(0)("fCalculoCuota") & String.Empty) > 0 Then FCalculoCuota = DtBPFF.Rows(0)("fCalculoCuota")
            If Length(DtBPFF.Rows(0)("fDesglosePrimeraCuota") & String.Empty) > 0 Then FDesglosePrimCuota = DtBPFF.Rows(0)("fDesglosePrimeraCuota")

            DtPagos = ClsPago.Filter(New FilterItem("IdPagoPeriodo", FilterOperator.Equal, LngID, FilterType.Numeric))
            If Not DtPagos Is Nothing Then
                If DtPagos.Rows.Count > 0 Then
                    If Length(FCalculoCuota & String.Empty) > 0 AndAlso Length(FDesglosePrimCuota & String.Empty) > 0 Then
                        Dt = CallByName(Me, FCalculoCuota, CallType.Method, DblNumCuotas, _
                            DblNumCoutasCarencia, DblInteresAplicado, DblPagosAlAño, _
                            DblBienTotal, DblImpInteresesTotal, DblValorResidual, BlnValorResidualIgualCuota, _
                            DblCuotaPeriodo, DblRecCostePeriodo, DblInteresPeriodo, FDesglosePrimCuota, MonInfoA.NDecimalesImporte)
                        If Not Dt Is Nothing AndAlso Dt.Rows.Count > 0 Then
                            Return Dt
                        End If
                    End If
                Else
                    If Length(FCalculoCuota & String.Empty) > 0 AndAlso Length(FDesglosePrimCuota & String.Empty) > 0 Then
                        Dt = CallByName(Me, FCalculoCuota, CallType.Method, DblNumCuotas, _
                            DblNumCoutasCarencia, DblInteresAplicado, DblPagosAlAño, _
                            DblBienTotal, DblImpInteresesTotal, DblValorResidual, BlnValorResidualIgualCuota, _
                            DblCuotaPeriodo, DblRecCostePeriodo, DblInteresPeriodo, FDesglosePrimCuota, MonInfoA.NDecimalesImporte)
                        If Not Dt Is Nothing AndAlso Dt.Rows.Count > 0 Then
                            Return Dt
                        End If
                    End If
                End If
            Else
                If Length(FCalculoCuota & String.Empty) > 0 AndAlso Length(FDesglosePrimCuota & String.Empty) > 0 Then
                    Dt = CallByName(Me, FCalculoCuota, CallType.Method, DblNumCuotas, _
                        DblNumCoutasCarencia, DblInteresAplicado, DblPagosAlAño, _
                        DblBienTotal, DblImpInteresesTotal, DblValorResidual, BlnValorResidualIgualCuota, _
                        DblCuotaPeriodo, DblRecCostePeriodo, DblInteresPeriodo, FDesglosePrimCuota, MonInfoA.NDecimalesImporte)
                    If Not Dt Is Nothing AndAlso Dt.Rows.Count > 0 Then
                        Return Dt
                    End If
                End If
            End If
        End If
    End Function

    Private Function CuotaFinalActualizacion(ByVal LngID As Long, ByVal Dt As DataTable) As Long
        Dim ClsMoneda As New Moneda
        Dim ClsTipoIVa As New TipoIva
        Dim ClsPagoPeriodico As New PagoPeriodico
        Dim DtPagoPeriodico As New DataTable
        Dim DtTipoIva As New DataTable
        Dim DblFactorIva As Double
        ' Los parámetros opcionales, únicamente se pasarán cuando sean actualizaciones

        Dim services As New ServiceProvider
        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
        Dim MonInfoA As MonedaInfo = Monedas.MonedaA
        CuotaFinalActualizacion = 1
        DtPagoPeriodico = ClsPagoPeriodico.SelOnPrimaryKey(LngID)
        DtTipoIva = ClsTipoIVa.Filter(New FilterItem("IDTipoIva", FilterOperator.Equal, DtPagoPeriodico.Rows(0)("IDTipoIva")))
        DblFactorIva = DtTipoIva.Rows(0)("Factor")

        If Not Dt Is Nothing AndAlso Dt.Rows.Count > 0 Then
            DtPagoPeriodico.Rows(0)("CuotaAutomatica") = True
            DtPagoPeriodico.Rows(0)("ImpCuotaPeriodo") = xRound(Dt.Rows(0)("Cuota"), MonInfoA.NDecimalesImporte)
            DtPagoPeriodico.Rows(0)("ImpRecuperacionCostePeriodo") = Dt.Rows(0)("Recuperacion")
            DtPagoPeriodico.Rows(0)("ImpInteresPeriodo") = xRound(Dt.Rows(0)("Intereses"), MonInfoA.NDecimalesImporte)
            DtPagoPeriodico.Rows(0)("Importe") = xRound(DtPagoPeriodico.Rows(0)("ImpCuotaPeriodo") * (1 + (CDbl(DblFactorIva) / 100)), MonInfoA.NDecimalesImporte)
            ClsPagoPeriodico.Update(DtPagoPeriodico)
        End If
    End Function

#End Region

#End Region

End Class