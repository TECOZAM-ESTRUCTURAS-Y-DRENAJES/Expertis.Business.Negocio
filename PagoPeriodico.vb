Public Class PagoPeriodicoInfo
    Inherits ClassEntityInfo

    Public ID As Integer
    Public IDInmovilizado As String

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal data As DataRow)
        MyBase.New(data)
    End Sub

    Public Overloads Overrides Sub Fill(ByVal ParamArray PrimaryKey() As Object)
        Dim dt As DataTable
        If Not IsNothing(PrimaryKey) AndAlso PrimaryKey.Length > 0 AndAlso Length(PrimaryKey(0)) > 0 Then
            dt = New BE.DataEngine().Filter("tbPagoPeriodico", New StringFilterItem("ID", PrimaryKey(0)))
        End If

        If IsNothing(dt) OrElse dt.Rows.Count = 0 Then
            ApplicationService.GenerateError("El Pago Periódico | no existe.", Quoted(PrimaryKey(0)))
        Else
            Me.Fill(dt.Rows(0))
        End If
    End Sub

End Class

Public Class PagoPeriodico
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

#Region " Constructor "

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbPagoPeriodico"

#End Region

#Region " RegisterAddNewTasks "

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        data("NoCalcularImpuesto") = False
        data("Contabilizado") = False
        data("CuotaFija") = False
        data("RecuperacionFija") = False
        data("PagoIntereses") = False
        data("CarenciaConIntereses") = False
        data("ValorResidualIgualCuota") = False
        data("ImpRecuperacionCoste") = 0
        data("ImpInteresesTotal") = 0
        data("ImpIVAOperacion") = 0
        data("OpcionCompra") = 0
        data("ImpCortoPlazoDeuda") = 0
        data("ImpCuotaPeriodo") = 0
        data("ImpRecuperacionCostePeriodo") = 0
        data("ImpInteresPeriodo") = 0
        data("Importe") = 0
        data("NPagosAlAño") = 12
        data("ValorResidual") = 0
    End Sub

#End Region

#Region " Validate "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)

        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarFechaInicioObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarFechaFinObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarUnidadObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarPeriodoObligatorio)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarIDCContableObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarTipoPagoObligatorio)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarMonedaObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarImporteObligatorio)
    End Sub

#End Region

#Region " Update "

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarIdentificador)
    End Sub

    <Task()> Public Shared Sub AsignarIdentificador(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("ID")) = 0 Then data("ID") = AdminData.GetAutoNumeric
        End If
    End Sub

#End Region

#Region " Business Rules "

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBRL As New BusinessRules
        oBRL.Add("CContable", AddressOf NegocioGeneral.FormatoCuentaContable)
        oBRL.Add("IdMoneda", AddressOf ProcesoComunes.CambioMoneda)
        oBRL.Add("IDProveedor", AddressOf CambioProveedor)
        oBRL.Add("FechaInicio", AddressOf CambioFechasUnidadPeriodo)
        oBRL.Add("FechaFin", AddressOf CambioFechasUnidadPeriodo)
        oBRL.Add("Unidad", AddressOf CambioFechasUnidadPeriodo)
        oBRL.Add("Periodo", AddressOf CambioFechasUnidadPeriodo)
        oBRL.Add("ValorResidualIgualCuota", AddressOf CambioValorResidualIgualCuota)
        oBRL.Add("NoCalcularImpuesto", AddressOf CambioNoCalcularImpuesto)
        oBRL.Add("IDBancoPropio", AddressOf CambioBancoPropio)
        oBRL.Add("BaseCalculo", AddressOf CambioBaseCalculoTipoInteres)
        oBRL.Add("TipoInteres", AddressOf CambioBaseCalculoTipoInteres)
        oBRL.Add("NTotalCuotas", AddressOf CambioNTotalCuotas)
        oBRL.Add("ImpIntereses", AddressOf CambioImpIntereses)
        oBRL.Add("ImpCuota", AddressOf CambioImpCuota)
        oBRL.Add("ImpRecuperacionCoste", AddressOf CambioImpRecuperacionCoste)
        Return oBRL
    End Function

    <Task()> Public Shared Sub CambioProveedor(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDProveedor")) > 0 Then
            Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
            Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(data.Current("IDProveedor"))
            If Not ProvInfo Is Nothing AndAlso Length(ProvInfo.IDProveedor) > 0 Then
                data.Current("IdFormaPago") = ProvInfo.IDFormaPago
                data.Current("IdBancoPropio") = ProvInfo.IDBancoPropio
                data.Current("IdMoneda") = ProvInfo.IDMoneda
                data.Current("IDTipoIva") = ProvInfo.IDTipoIVA
            End If

            ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CambioBancoPropio, data, services)
        End If
    End Sub

    <Task()> Public Shared Sub CambioBancoPropio(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IdBancoPropio")) > 0 Then
            Dim BancosPropios As EntityInfoCache(Of BancoPropioInfo) = services.GetService(Of EntityInfoCache(Of BancoPropioInfo))()
            Dim BPInfo As BancoPropioInfo = BancosPropios.GetEntity(data.Current("IdBancoPropio"))
            If Not BPInfo Is Nothing AndAlso Length(BPInfo.IDBancoPropio) > 0 Then
                data.Current("PagoIntereses") = BPInfo.PagoIntereses
                data.Current("BaseCalculo") = BPInfo.BaseCalculo
                data.Current("CarenciaConIntereses") = BPInfo.CarenciaConIntereses
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioFechasUnidadPeriodo(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        Dim Dbli, Dblj, DblK As Double
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("FechaFin")) > 0 AndAlso Length(data.Current("FechaInicio")) > 0 Then
            If Length(data.Current("Periodo")) > 0 AndAlso Length(data.Current("Unidad")) > 0 Then
                Select Case CInt(Nz(data.Current("Unidad"), -1))
                    Case 3
                        Dbli = DateDiff("d", data.Current("FechaInicio"), data.Current("FechaFin"))
                        DblK = Dbli / 365
                        Dblj = Math.Round(DblK / data.Current("Periodo"))
                        If Dblj < DblK / data.Current("Periodo") Then Dblj += 1

                    Case 2
                        Dbli = DateDiff("m", data.Current("FechaInicio"), data.Current("FechaFin"))
                        Dblj = Math.Round(Dbli / data.Current("Periodo"))

                    Case 1
                        Dbli = DateDiff("d", data.Current("FechaInicio"), data.Current("FechaFin"))
                        DblK = Dbli / 7
                        Dblj = Math.Round(DblK / data.Current("Periodo"))
                        If Dblj < (DblK / data.Current("Periodo")) Then Dblj += 1

                    Case 0
                        Dbli = DateDiff("d", data.Current("FechaInicio"), data.Current("FechaFin"))
                        Dblj = Math.Round(Dbli / data.Current("Periodo"))
                End Select
                data.Current("NTotalCuotas") = Dblj
            End If
        End If

    End Sub

    <Task()> Public Shared Sub CambioValorResidualIgualCuota(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Nz(data.Current("ValorResidualIgualCuota"), False) Then
            data.Current("ValorResidual") = 0
            data.Current("ValorResidualA") = 0
            data.Current("ValorResidualB") = 0
        End If
    End Sub

    <Task()> Public Shared Sub CambioNoCalcularImpuesto(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Nz(data.Current("NoCalcularImpuesto"), False) Then
            ProcessServer.ExecuteTask(Of Integer)(AddressOf BorrarPagoPeriodicoImpuesto, data.Current("ID"), services)
        End If
    End Sub

    <Task()> Public Shared Sub CambioBaseCalculoTipoInteres(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        Dim ClsPago As New Pago
        Dim ClsBPFF As New BancoPropioFormFinanc
        Dim DblInteresAplicado As Double
        Dim DtBPFF As DataTable = ClsBPFF.Filter(New FilterItem("IDBancoPropio", FilterOperator.Equal, data.Current("IDBancoPropio")))
        If Not DtBPFF Is Nothing AndAlso DtBPFF.Rows.Count > 0 Then
            If Length(DtBPFF.Rows(0)("fCalculoInteresAplicado")) > 0 Then
                If Length(data.Current("BaseCalculo")) > 0 AndAlso Length(data.Current("TipoInteres")) > 0 Then
                    DblInteresAplicado = CallByName(ClsBPFF, DtBPFF.Rows(0)("fCalculoInteresAplicado"), CallType.Method, data.Current("BaseCalculo"), data.Current("TipoInteres"))
                Else
                    DblInteresAplicado = 0
                End If
                data.Current("TipoInteresAplicado") = DblInteresAplicado
            Else
                ApplicationService.GenerateError("El Cálculo de Interés Aplicado no ha producido ningún resultado.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioNTotalCuotas(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        Dim intDia As Integer
        Dim intMes As Integer
        Dim intAño As Integer
        Dim dbli As Double
        Dim intAux As Integer
        If Length(data.Current("FechaInicio")) > 0 AndAlso Length(data.Current("Unidad")) > 0 Then
            If Length(data.Current("Periodo")) > 0 Then
                Select Case CInt(Nz(data.Current("Unidad"), -1))
                    Case 3
                        dbli = data.Current("NTotalCuotas") * 365 * data.Current("Periodo")
                        data.Current("FechaFin") = data.Current("FechaInicio") + dbli
                    Case 2
                        intDia = CDate(data.Current("FechaInicio")).Day
                        intMes = CDate(data.Current("FechaInicio")).Month
                        intAño = CDate(data.Current("FechaInicio")).Year
                        intMes = CDate(data.Current("FechaInicio")).Month + (data.Current("NTotalCuotas") * data.Current("Periodo"))
                        intAux = intMes
                        If intMes / 12 > 1 Then
                            If intMes Mod 12 = 0 Then
                                intMes -= (12 * (Int(intMes / 12) - 1))
                                intAño += (1 * (Int(intAux / 12) - 1))
                            Else
                                intMes -= (12 * Int(intMes / 12))
                                intAño += (1 * Int(intAux / 12))
                            End If
                        End If
                        data.Current("FechaFin") = New Date(intAño, intMes, intDia)
                        'current("FechaFin") = CDate(lngDia & "/" & lngMes & "/" & lngAño)

                    Case 1
                        dbli = data.Current("NTotalCuotas") * 7 * data.Current("Periodo")
                        data.Current("FechaFin") = data.Current("FechaInicio") + dbli

                    Case 0
                        dbli = data.Current("NTotalCuotas") * data.Current("Periodo")
                        data.Current("FechaFin") = data.Current("FechaInicio") + dbli
                End Select
            End If

        End If
    End Sub

    <Task()> Public Shared Sub CambioImpIntereses(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        'Comprobamos que no tenga cadena vacia y que el valor sea numerico
        If Length(data.Current("ImpIntereses")) > 0 Then
            'Actualizamos ImpInteresesA y ImpInteresesB y las bases imponibles
            If Not data.Current("Contabilizado") Then
                Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
                Dim MonInfo As MonedaInfo = Monedas.GetMoneda(data.Current("IDMoneda"))

                Dim ValAyB As New ValoresAyB(CDbl(Nz(data.Current("ImpIntereses"), 0)), data.Current("IDMoneda"), MonInfo.CambioA, MonInfo.CambioB)
                Dim fImp As fImporte = ProcessServer.ExecuteTask(Of ValoresAyB, fImporte)(AddressOf NegocioGeneral.MantenimientoValoresImporteAyB, ValAyB, services)
                data.Current("ImpIntereses") = fImp.Importe
                data.Current("ImpInteresesA") = fImp.ImporteA
                data.Current("ImpInteresesB") = fImp.ImporteB

                ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CambioImpCuota, data, services)
            Else
                'No se puede modificar porque ya esta contabilizado
                ApplicationService.GenerateError("No se puede modificar el importe ni la moneda porque ya está contabilizado.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioImpCuota(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
        Dim MonInfo As MonedaInfo = Monedas.GetMoneda(data.Current("IDMoneda"))
        If data.Current("RecuperacionFija") OrElse (Not data.Current("CuotaFija") AndAlso Not data.Current("RecuperacionFija")) Then
            data.Current("ImpCuota") = Nz(data.Current("ImpIntereses"), 0) + Nz(data.Current("ImpRecuperacionCoste"), 0)
            Dim ValAyB As ValoresAyB = New ValoresAyB(CDbl(Nz(data.Current("ImpCuota"), 0)), data.Current("IDMoneda"), MonInfo.CambioA, MonInfo.CambioB)
            Dim fImp As fImporte = ProcessServer.ExecuteTask(Of ValoresAyB, fImporte)(AddressOf NegocioGeneral.MantenimientoValoresImporteAyB, ValAyB, services)
            data.Current("ImpCuota") = fImp.Importe
            data.Current("ImpCuotaA") = fImp.ImporteA
            data.Current("ImpCuotaB") = fImp.ImporteB
        End If
        If Length(data.Current("IDTipoIVA")) > 0 Then
            Dim dtTipoIva As DataTable = New TipoIva().SelOnPrimaryKey(data.Current("IDTipoIVA"))
            If Not dtTipoIva Is Nothing AndAlso dtTipoIva.Rows.Count > 0 Then
                If dtTipoIva.Rows(0)("Factor") <> 0 Then
                    data.Current("ImpVencimiento") = data.Current("ImpCuota") + data.Current("ImpCuota") * dtTipoIva.Rows(0)("Factor") / 100
                Else
                    data.Current("ImpVencimiento") = data.Current("ImpCuota")
                End If
                Dim ValAyB As ValoresAyB = New ValoresAyB(CDbl(Nz(data.Current("ImpVencimiento"), 0)), data.Current("IDMoneda"), MonInfo.CambioA, MonInfo.CambioB)
                Dim fImp As fImporte = ProcessServer.ExecuteTask(Of ValoresAyB, fImporte)(AddressOf NegocioGeneral.MantenimientoValoresImporteAyB, ValAyB, services)
                data.Current("ImpVencimiento") = fImp.Importe
                data.Current("ImpVencimientoA") = fImp.ImporteA
                data.Current("ImpVencimientoB") = fImp.ImporteB
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioImpRecuperacionCoste(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("ImpRecuperacionCoste")) > 0 Then
            If Not data.Current("Contabilizado") Then
                Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
                Dim MonInfo As MonedaInfo = Monedas.GetMoneda(data.Current("IDMoneda"))
                Dim ValAyB As New ValoresAyB(CDbl(Nz(data.Current("ImpRecuperacionCoste"), 0)), data.Current("IDMoneda"), MonInfo.CambioA, MonInfo.CambioB)
                Dim fImp As fImporte = ProcessServer.ExecuteTask(Of ValoresAyB, fImporte)(AddressOf NegocioGeneral.MantenimientoValoresImporteAyB, ValAyB, services)
                data.Current("ImpRecuperacionCoste") = fImp.Importe
                data.Current("ImpRecuperacionCosteA") = fImp.ImporteA
                data.Current("ImpRecuperacionCosteB") = fImp.ImporteB

                ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CambioImpIntereses, data, services)
            Else
                'No se puede modificar porque ya esta contabilizado
                ApplicationService.GenerateError("No se puede modificar el importe ni la moneda porque ya está contabilizado.")
            End If
        End If
    End Sub

#End Region

#Region "Funciones / Procesos Públicos / Privados"

    Public Function InsertPagoPeriodicoCuota(ByVal Dt As DataTable, _
                                            ByVal DteFechaFinal As Date, _
                                            ByVal BlnSimulacion As Boolean) As DataTable
        Dim services As New ServiceProvider
        'Prepara el datatable para actualizar los Pagos que se han modificado manualmente
        Dim ClsMoneda As New Moneda
        Dim ClsPago As New Pago
        Dim ClsTipoPago As New TipoPago
        Dim ClsTipoIva As New TipoIva
        Dim ClsBPFF As New BancoPropioFormFinanc
        Dim DtMoneda, DtMonedaA, DtPago, DtTipoPago, Dt1, _
        DtApp, DtTipoIva, DtBPFF As DataTable 'DtMonedaB,
        Dim DtPagoPer As New DataTable
        Dim DteFechaCom, DteFechaCom2, DteFechaTope, DteFechaComCar As Date
        Dim LngPeriodo, LngDecImpA As Long ', LngDecImpB As Long
        Dim LngNumCuotas As Integer
        Dim StrUnidad, StrAgrup, StrIDPagoPer, StrFuncCuotasSuces As String
        Dim BlnRedondeoFinal, BlnCuotaVR, BlnCarConInt, BlnPrimCuota As Boolean
        Dim DblRecupCosteFinal, DblBienTotal, DblRecCosteUltima, DblImpCuota, DblIntAntCuota, _
        DblAmortini, DblTotalInt, DblTotalCuota, DblTotalFinanc, DblVR, DblIntCar, DblAmort, _
         DblK, DblBien, DblRecup, DblIntereses, DblImpIntTemp, DblImpCuotaTemp As Double

        DtMoneda = ClsMoneda.Filter()
        DtMonedaA = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf Moneda.ObtenerMonedaA, Nothing, services)
        'DtMonedaB = ClsMoneda.ObtenerMonedaB
        LngDecImpA = DtMonedaA.Rows(0)("NDecimalesImp")
        'LngDecImpB = DtMonedaB.Rows(0)("NDecimalesImp")
        BlnCarConInt = False
        BlnPrimCuota = False
        DblRecupCosteFinal = 0
        DblAmortini = 0
        DblTotalInt = 0
        DblTotalCuota = 0
        DblVR = 0
        BlnCuotaVR = False
        If Dt.Rows.Count = 1 Then
            DblBien = Dt.Rows(0)("ImpRecuperacionCoste") + Dt.Rows(0)("ImpInteresesTotal")
            DblBienTotal = Dt.Rows(0)("ImpRecuperacionCoste") + Dt.Rows(0)("ImpInteresesTotal")
            DblTotalFinanc = Dt.Rows(0)("ImpRecuperacionCoste") + Dt.Rows(0)("ImpInteresesTotal")
            DtApp = New BE.DataEngine().Filter("vLeasingTotalInmovilizado", New StringFilterItem("IDInmovilizado", Dt.Rows(0)("IDInmovilizado")))
            If Not DtApp Is Nothing AndAlso DtApp.Rows.Count > 0 Then
                DblBien = Nz(DtApp.Rows(0)("TotalRevalorizadoA") - Dt.Rows(0)("AportacionInicial"), Dt.Rows(0)("ImpRecuperacionCoste") + Dt.Rows(0)("ImpInteresesTotal"))
                DblBienTotal = Nz(DtApp.Rows(0)("TotalRevalorizadoA") - Dt.Rows(0)("AportacionInicial"), Dt.Rows(0)("ImpRecuperacionCoste") + Dt.Rows(0)("ImpInteresesTotal"))
                DblTotalFinanc = Nz(DtApp.Rows(0)("TotalRevalorizadoA") - Dt.Rows(0)("AportacionInicial"), Dt.Rows(0)("ImpRecuperacionCoste") + Dt.Rows(0)("ImpInteresesTotal"))
            End If
            LngNumCuotas = 0
            DblK = Dt.Rows(0)("TipoInteresAplicado") / (Dt.Rows(0)("NPagosAlAño") * 100)
            DblImpCuota = Dt.Rows(0)("ImpCuotaPeriodo")
            BlnRedondeoFinal = False
            DblIntAntCuota = 0

            DtBPFF = ClsBPFF.Filter(New FilterItem("IDBancoPropio", FilterOperator.Equal, Dt.Rows(0)("IDBancoPropio"), FilterType.String))
            If Not DtBPFF Is Nothing AndAlso DtBPFF.Rows.Count > 0 Then
                StrFuncCuotasSuces = DtBPFF.Rows(0)("fDesgloseSucesivasCuotas") & String.Empty
            End If
            StrIDPagoPer = Dt.Rows(0)("ID")
            If Length(StrIDPagoPer) > 0 Then
                DtPagoPer = Me.Filter(New NumberFilterItem("ID", StrIDPagoPer))
                DtPago = ClsPago.AddNew()
                Dim DrDatos() As DataRow = DtPagoPer.Select("ID=" & Dt.Rows(0)("ID"))
                If Length(DrDatos.Length) > 0 AndAlso Length(DrDatos(0)("IDTipoIva")) > 0 Then
                    DtTipoIva = ClsTipoIva.SelOnPrimaryKey(DrDatos(0)("IDTipoIva"))
                End If
                With DtPagoPer.Rows(0)
                    Dim Drmoneda() As DataRow = DtMoneda.Select("IdMoneda = '" & .Item("IdMoneda") & "'")
                    If Nz(.Item("FechaUltimaActualizacion"), cnMinDate) < .Item("FechaFin") Then
                        StrUnidad = New NegocioGeneral().GetPeriodString(.Item("Unidad"))
                        If Length(.Item("FechaUltimaActualizacion") & String.Empty) = 0 Then
                            DteFechaCom = .Item("FechaInicio")
                            If .Item("NCuotasCarencia") > 0 Then
                                DteFechaCom = DateAdd(StrUnidad, .Item("NCuotasCarencia"), .Item("FechaInicio"))
                                If .Item("CarenciaConIntereses") = True Then
                                    DteFechaComCar = .Item("FechaInicio")
                                    BlnCarConInt = True
                                    DblIntCar = Dt.Rows(0)("ImpRecuperacionCoste") * DblK
                                End If
                            End If
                        Else
                            DteFechaCom = DateAdd(StrUnidad, .Item("Periodo"), .Item("FechaUltimaActualizacion"))
                            If .Item("NCuotasCarencia") > 0 Then
                                DteFechaCom2 = DateAdd(StrUnidad, .Item("NCuotasCarencia"), .Item("FechaInicio"))
                                If DteFechaCom2 > DteFechaCom Then
                                    DteFechaCom = DteFechaCom2
                                End If
                            End If
                        End If
                        DteFechaTope = IIf(.Item("FechaFin") < DteFechaFinal, .Item("FechaFin"), DteFechaFinal)
                        If Not ClsTipoPago Is Nothing Then
                            DtTipoPago = ClsTipoPago.Filter(New FilterItem("IdTipoPago", FilterOperator.Equal, DtPagoPer.Rows(0)("IdTipoPago"), FilterType.String))
                            If Not DtTipoPago Is Nothing AndAlso DtTipoPago.Rows.Count > 0 Then
                                StrAgrup = DtTipoPago.Rows(0)("IDAgrupacion") & String.Empty
                            End If
                        End If
                        LngPeriodo = 0
                        DblVR = IIf(.Item("ValorResidualIgualCuota"), DblImpCuota, .Item("ValorResidual"))
                        ''''''''''''''''''''''''''''''''''''
                        '''Caso de Carencia con Intereses'''
                        ''''''''''''''''''''''''''''''''''''
                        If BlnCarConInt = True Then
                            If .Item("PagoIntereses") = True Then
                                For IntI As Integer = 0 To .Item("NCuotasCarencia") - 1
                                    Dim DrNew As DataRow = DtPago.NewRow()
                                    If Not BlnSimulacion Then DrNew("IDPago") = AdminData.GetAutoNumeric
                                    DrNew("Titulo") = .Item("DescPago")
                                    DrNew("FechaVencimiento") = DateAdd(StrUnidad, IntI, DteFechaComCar)
                                    DrNew("CContable") = .Item("IDCContable")
                                    DrNew("IDProveedor") = .Item("IDProveedor")
                                    DrNew("IDTipoPago") = .Item("IDTipoPago")
                                    DrNew("IDFormaPago") = .Item("IDFormaPago")
                                    DrNew("IDBancoPropio") = .Item("IDBancoPropio")
                                    DrNew("IDMoneda") = .Item("IDMoneda")
                                    DrNew("CambioA") = Drmoneda(0)("CambioA")
                                    DrNew("CambioB") = Drmoneda(0)("CambioB")
                                    DrNew("Situacion") = enumPagoSituacion.NoPagado
                                    DrNew("Contabilizado") = enumPagoContabilizado.PagoNoContabilizado
                                    DrNew("ImpRecuperacionCoste") = 0
                                    DrNew("ImpRecuperacionCosteA") = 0
                                    'DrNew("ImpRecuperacionCosteB") = 0
                                    DrNew("ImpIntereses") = xRound(DblIntCar, Drmoneda(0)("NDecimalesImp"))
                                    DrNew("ImpInteresesA") = xRound(DblIntCar * Drmoneda(0)("CambioA"), DtMonedaA.Rows(0)("NDecimalesImp"))
                                    'DrNew("ImpInteresesB") = xRound(DblIntCar * DtMoneda.Rows(0)("CambioB"), DtMoneda.Rows(0)("NDecimalesImp"))
                                    DrNew("ImpCuota") = DrNew("ImpIntereses")
                                    DrNew("ImpCuotaA") = DrNew("ImpInteresesA")
                                    'DrNew("impCuotaB") = DrNew("ImpInteresesB")
                                    DrNew("ImpVencimiento") = xRound(DrNew("ImpCuota") * (1 + (DtTipoIva.Rows(0)("Factor")) / 100), Drmoneda(0)("NDecimalesimp"))
                                    DrNew("ImpVencimientoA") = xRound(DrNew("ImpCuotaA") * (1 + (DtTipoIva.Rows(0)("Factor")) / 100) * Drmoneda(0)("CambioA"), DtMonedaA.Rows(0)("NDecimalesImp"))
                                    'DrNew("ImpVencimientoB") = xRound(DrNew("ImpCuotaA") * (1 + (DtTipoIva.Rows(0)("Factor")) / 100) * DtMoneda.Rows(0)("CambioB"), DtMonedaB.Rows(0)("NDecimalesImp"))
                                    If StrAgrup & String.Empty <> String.Empty Then
                                        DrNew("IDAgrupacion") = StrAgrup
                                    End If
                                    DrNew("IdPagoPeriodo") = Dt.Rows(0)("ID")
                                    DtPago.Rows.Add(DrNew)
                                Next IntI
                            ElseIf .Item("PagoIntereses") = False Then
                                For IntI As Integer = 1 To .Item("NCuotasCarencia") - 1
                                    Dim DrNew As DataRow = DtPago.NewRow()
                                    If Not BlnSimulacion Then DrNew("IDPago") = AdminData.GetAutoNumeric
                                    DrNew("Titulo") = .Item("DescPago")
                                    DrNew("FechaVencimiento") = DateAdd(StrUnidad, IntI, DteFechaComCar)
                                    DrNew("CContable") = .Item("IDCContable")
                                    DrNew("IDProveedor") = .Item("IDProveedor")
                                    DrNew("IDTipoPago") = .Item("IDTipoPago")
                                    DrNew("IDFormaPago") = .Item("IDTipoPago")
                                    DrNew("IDBancoPropio") = .Item("IDBancoPropio")
                                    DrNew("IDMoneda") = .Item("IDMoneda")
                                    DrNew("CambioA") = Drmoneda(0)("CambioA")
                                    DrNew("CambioB") = Drmoneda(0)("CambioB")
                                    DrNew("Situacion") = enumPagoSituacion.NoPagado
                                    DrNew("Contabilizado") = enumPagoContabilizado.PagoNoContabilizado
                                    DrNew("ImpRecuperacionCoste") = 0
                                    DrNew("ImpRecuperacionCosteA") = 0
                                    'drnew("ImpRecuperacionCosteB") = 0
                                    DrNew("ImpIntereses") = xRound(DblIntCar, Drmoneda(0)("NDecimalesIMP"))
                                    DrNew("ImpInteresesA") = xRound(DblIntCar * Drmoneda(0)("CambioA"), DtMonedaA.Rows(0)("NDecimalesImp"))
                                    'drnew("ImpInteresesB") = xRound(DblIntCar * DtMoneda.Rows(0)("CambioB"), DtMonedaB.Rows(0)("NDecimalesImp"))
                                    DrNew("ImpCuota") = DtPago.Rows(0)("ImpIntereses")
                                    DrNew("ImpCuotaA") = DtPago.Rows(0)("ImpInteresesA")
                                    'drnew("ImpCuotaB") = DtPago.Rows(0)("ImpInteresesB")
                                    DrNew("ImpVencimiento") = xRound(DtPago.Rows(0)("ImpCuota") * (1 + (DtTipoIva.Rows(0)("Factor")) / 100), Drmoneda(0)("NDecimalesIMP"))
                                    DrNew("ImpVencimientoA") = xRound(DtPago.Rows(0)("ImpCuotaA") * (1 + (DtTipoIva.Rows(0)("Factor")) / 100) * Drmoneda(0)("CambioA"), DtMonedaA.Rows(0)("NDecimalesImp"))
                                    'drnew("ImpVencimientoB") = xRound(DtPago.Rows(0)("ImpCuotaB") * (1 + (DtTipoIva.Rows(0)("Factor")) / 100) * DtMoneda.Rows(0)("CambioB"), DtMonedaB.Rows(0)("NDecimalesImp"))
                                    If StrAgrup & String.Empty <> String.Empty Then
                                        DtPago.Rows(0)("IDAgrupacion") = StrAgrup
                                    End If
                                    DtPago.Rows(0)("IDPagoPeriodo") = Dt.Rows(0)("ID")
                                    DtPago.Rows.Add(DrNew)
                                Next IntI
                            End If
                        End If
                        While DateAdd(StrUnidad, LngPeriodo * .Item("Periodo"), DteFechaCom) <= DteFechaTope
                            If Not Nz(DtPagoPer.Rows(0)("ValorResidualIgualCuota"), False) AndAlso LngNumCuotas >= Nz(DtPagoPer.Rows(0)("NTotalCuotas"), 0) - Nz(DtPagoPer.Rows(0)("NCuotasCarencia"), 0) Then
                                Exit While
                            Else
                                If Nz(DtPagoPer.Rows(0)("ValorResidualIgualCuota"), False) AndAlso LngNumCuotas = Nz(DtPagoPer.Rows(0)("NTotalCuotas"), 0) - Nz(DtPagoPer.Rows(0)("NCuotasCarencia"), 0) - 1 Then
                                    BlnRedondeoFinal = True
                                Else
                                    BlnRedondeoFinal = False
                                End If
                                If Nz(DtPagoPer.Rows(0)("ValorResidualIgualCuota"), False) AndAlso LngNumCuotas = Nz(DtPagoPer.Rows(0)("NTotalCuotas"), 0) - Nz(DtPagoPer.Rows(0)("NCuotasCarencia"), 0) Then
                                    BlnCuotaVR = True
                                Else
                                    BlnCuotaVR = False
                                End If
                                If Dt.Rows(0)("PagoIntereses") = False AndAlso BlnPrimCuota = False Then
                                    Dim DrNew As DataRow = DtPago.NewRow()
                                    If Not BlnSimulacion Then DrNew("IDPago") = AdminData.GetAutoNumeric()
                                    DrNew("Titulo") = .Item("DescPago")
                                    If LngPeriodo = 0 Then
                                        DrNew("FechaVencimiento") = DteFechaCom
                                    Else
                                        DrNew("FechaVencimiento") = DateAdd(StrUnidad, LngPeriodo * .Item("Periodo"), DteFechaCom)
                                    End If
                                    .Item("FechaUltimaActualizacion") = DrNew("FechaVencimiento")
                                    DrNew("CContable") = .Item("IDCContable")
                                    DrNew("IDProveedor") = .Item("IDProveedor")
                                    DrNew("IdTipoPago") = .Item("IDTipoPago")
                                    DrNew("IDFormaPago") = .Item("IDFormaPago")
                                    DrNew("IDBancoPropio") = .Item("IDBancoPropio")
                                    DrNew("IDMoneda") = .Item("IDMoneda")
                                    DrNew("CambioA") = Drmoneda(0)("CambioA")
                                    DrNew("CambioB") = Drmoneda(0)("CambioB")
                                    DrNew("Situacion") = enumPagoSituacion.NoPagado
                                    DrNew("Contabilizado") = enumPagoContabilizado.PagoNoContabilizado
                                    DrNew("ImpRecuperacionCoste") = xRound(.Item("ImpCuotaPeriodo"), Drmoneda(0)("NDecimalesIMP"))
                                    DrNew("ImpRecuperacionCosteA") = xRound(.Item("ImpCuotaPeriodo") * Drmoneda(0)("CambioA"), DtMonedaA.Rows(0)("NDecimalesImp"))
                                    'drnew("ImpRecuperacionCosteB") = xRound(.Item("ImpCuotaPeriodo") * DtMoneda.Rows(0)("CambioB"), DtMonedaB.Rows(0)("NDecimalesImp"))
                                    If BlnCarConInt = False Then
                                        DrNew("ImpIntereses") = 0
                                        DrNew("ImpInteresesA") = 0
                                        'drnew("ImpInteresesB") = 0
                                    Else
                                        DrNew("ImpIntereses") = xRound(DblIntCar, Drmoneda(0)("NDecimalesIMP"))
                                        DrNew("ImpInteresesA") = xRound(DblIntCar * Drmoneda(0)("CambioA"), DtMonedaA.Rows(0)("NDecimalesImp"))
                                        'drnew("ImpInteresesB") = xRound(DblIntCar * DtMoneda.Rows(0)("CambioB"), DtMonedaB.Rows(0)("NDecimalesImp"))
                                    End If
                                    If BlnCarConInt = False Then
                                        DrNew("ImpVencimiento") = xRound(.Item("Importe"), Drmoneda(0)("NDecimalesIMP"))
                                        DrNew("ImpVencimientoA") = xRound(.Item("Importe") * Drmoneda(0)("CambioA"), DtMonedaA.Rows(0)("NDecimalesImp"))
                                        'drnew("ImpVencimientoB") = xRound(.Item("Importe") * DtMoneda.Rows(0)("CambioB"), DtMonedaB.Rows(0)("NDecimalesImp"))
                                    Else
                                        DrNew("ImpVencimiento") = xRound((.Item("Importe") + (DrNew("ImpIntereses") * (1 + (DtTipoIva.Rows(0)("Factor")) / 100))), Drmoneda(0)("NDecimalesIMP"))
                                        DrNew("ImpVencimientoA") = xRound((.Item("Importe") + (DrNew("ImpInteresesA") * (1 + (DtTipoIva.Rows(0)("Factor")) / 100))) * Drmoneda(0)("CambioA"), DtMonedaA.Rows(0)("NDecimalesImp"))
                                        'drnew("ImpVencimientoB") = xRound((.Item("Importe") + (drnew("ImpInteresesB") * (1 + (DtTipoIva.Rows(0)("Factor")) / 100))) * DtMoneda.Rows(0)("CambioB"), DtMonedaB.Rows(0)("NDecimalesImp"))
                                    End If
                                    If BlnCarConInt = False Then
                                        DrNew("ImpCuota") = Dt.Rows(0)("ImpCuotaPeriodo")
                                        DrNew("ImpCuotaA") = xRound(Dt.Rows(0)("ImpCuotaPeriodo") * Drmoneda(0)("CambioA"), DtMonedaA.Rows(0)("NDecimalesImp"))
                                        'drnew("ImpCuotaB") = xRound(Dt.Rows(0)("ImpCuotaPeriodo") * DtMoneda.Rows(0)("CambioB"), DtMonedaB.Rows(0)("NDecimalesImp"))
                                    Else
                                        DrNew("ImpCuota") = Dt.Rows(0)("ImpCuotaPeriodo") + DrNew("ImpIntereses")
                                        DrNew("ImpCuotaA") = xRound((Dt.Rows(0)("ImpCuotaPeriodo") + DrNew("ImpIntereses")) * Drmoneda(0)("CambioA"), DtMonedaA.Rows(0)("NDecimalesImp"))
                                        'drnew("ImpCuotaB") = xRound((Dt.Rows(0)("ImpCuotaPeriodo") + drnew("ImpIntereses")) * DtMoneda.Rows(0)("CambioB"), DtMonedaB.Rows(0)("NDecimalesImp"))
                                    End If
                                    If StrAgrup & String.Empty <> String.Empty Then
                                        DrNew("IDAgrupacion") = StrAgrup
                                    End If
                                    DrNew("IdPagoPeriodo") = Dt.Rows(0)("ID")
                                    LngPeriodo += 1
                                    LngNumCuotas += 1
                                    DblTotalInt += DrNew("ImpIntereses")
                                    DblTotalCuota += DrNew("ImpCuota")
                                    BlnPrimCuota = True
                                    DblRecupCosteFinal += DrNew("ImpRecuperacionCoste")
                                    DtPago.Rows.Add(DrNew)
                                End If
                                Dim DrNew2 As DataRow = DtPago.NewRow()
                                If Not BlnSimulacion Then DrNew2("IDPago") = AdminData.GetAutoNumeric
                                DrNew2("Titulo") = .Item("DescPago")
                                If LngPeriodo = 0 Then
                                    DrNew2("FechaVencimiento") = DteFechaCom
                                Else
                                    DrNew2("FechaVencimiento") = DateAdd(StrUnidad, LngPeriodo * .Item("Periodo"), DteFechaCom)
                                End If
                                .Item("FechaUltimaActualizacion") = DrNew2("FechaVencimiento")
                                DrNew2("CContable") = .Item("IDCContable")
                                DrNew2("IDProveedor") = .Item("IDProveedor")
                                DrNew2("IDTipoPago") = .Item("IDTipoPago")
                                DrNew2("IDFormaPago") = .Item("IDFormaPago")
                                DrNew2("IDBancoPropio") = .Item("IDBancoPropio")
                                DrNew2("IDMoneda") = .Item("IDMoneda")
                                DrNew2("CambioA") = Drmoneda(0)("CambioA")
                                DrNew2("CambioB") = Drmoneda(0)("CambioB")
                                DrNew2("Situacion") = enumPagoSituacion.NoPagado
                                DrNew2("Contabilizado") = enumPagoContabilizado.PagoNoContabilizado
                                DrNew2("ImpVencimiento") = xRound(.Item("Importe"), Drmoneda(0)("NDecimalesIMP"))
                                DrNew2("ImpVencimientoA") = xRound(.Item("Importe") * Drmoneda(0)("CambioA"), DtMonedaA.Rows(0)("NDecimalesImp"))
                                'DrNew2("ImpVencimientoB") = xRound(.Item("Importe") * DtMoneda.Rows(0)("CambioB"), DtMonedaB.Rows(0)("NDecimalesImp"))
                                If LngPeriodo <> 0 AndAlso (LngPeriodo <> 1 OrElse BlnPrimCuota = False) Then
                                    If Length(StrFuncCuotasSuces & String.Empty) > 0 Then
                                        Dt1 = CallByName(ClsBPFF, StrFuncCuotasSuces, CallType.Method, DblK, DblBien, _
                                        DblAmort, DblImpCuota, LngNumCuotas, BlnRedondeoFinal, DblIntAntCuota, _
                                        LngDecImpA, DblTotalInt, DblTotalCuota, DblTotalFinanc, DblVR, BlnCuotaVR)
                                        DblBien = Dt1.Rows(0)("CapitalPte")
                                        DblRecup = Dt1.Rows(0)("Recuperacion")
                                        DblIntereses = Dt1.Rows(0)("Intereses")
                                        DblIntAntCuota = Dt1.Rows(0)("Intereses")
                                        DrNew2("ImpRecuperacionCoste") = xRound(DblRecup, Drmoneda(0)("NDecimalesIMP"))
                                        DblAmort = DrNew2("ImpRecuperacionCoste")

                                        DrNew2("ImpRecuperacionCosteA") = xRound(DblRecup * Drmoneda(0)("CambioA"), DtMonedaA.Rows(0)("NDecimalesImp"))
                                        'DrNew2("ImpRecuperacionCosteB") = xRound(DblRecup * DtMoneda.Rows(0)("CambioB"), DtMonedaB.Rows(0)("NDecimalesImp"))
                                        DrNew2("ImpIntereses") = xRound(DblIntereses, Drmoneda(0)("NDecimalesIMP"))
                                        DrNew2("ImpInteresesA") = xRound(DblIntereses * Drmoneda(0)("CambioA"), DtMonedaA.Rows(0)("NDecimalesImp"))
                                        'DrNew2("ImpInteresesB") = xRound(DblIntereses * DtMoneda.Rows(0)("CambioB"), DtMoneda.Rows(0)("NDecimalesImp"))
                                    Else
                                        DrNew2("ImpRecuperacionCoste") = xRound(.Item("ImpRecuperacionCostePeriodo"), Drmoneda(0)("NDecimalesIMP"))
                                        DrNew2("ImpRecuperacionCosteA") = xRound(.Item("ImpRecuperacionCostePeriodo") * Drmoneda(0)("CambioA"), DtMonedaA.Rows(0)("NDecimalesImp"))
                                        'DrNew2("ImpRecuperacionCosteB") = xRound(.Item("ImpRecuperacionCostePeriodo") * DtMoneda.Rows(0)("CambioB"), DtMoneda.Rows(0)("NDecimalesImp"))
                                        DrNew2("ImpIntereses") = xRound(.Item("ImpInteresPeriodo"), Drmoneda(0)("NDecimalesIMP"))
                                        DrNew2("ImpInteresesA") = xRound(.Item("ImpInteresPeriodo") * Drmoneda(0)("CambioA"), DtMonedaA.Rows(0)("NDecimalesImp"))
                                        'DrNew2("ImpInteresesB") = xRound(.Item("ImpInteresPeriodo") * DtMoneda.Rows(0)("CambioB"), DtMoneda.Rows(0)("NDecimalesImp"))
                                    End If
                                Else
                                    DrNew2("ImpRecuperacionCoste") = xRound(.Item("ImpRecuperacionCostePeriodo"), Drmoneda(0)("NDecimalesIMP"))
                                    DblAmort = .Item("ImpRecuperacionCostePeriodo")
                                    DblBien -= DblAmort
                                    DrNew2("ImpRecuperacionCosteA") = xRound(.Item("ImpRecuperacionCostePeriodo") * Drmoneda(0)("CambioA"), DtMonedaA.Rows(0)("NDecimalesImp"))
                                    'DrNew2("ImpRecuperacionCosteB") = xRound(.Item("ImpRecuperacionCostePeriodo") * DtMoneda.Rows(0)("CambioB"), DtMonedaB.Rows(0)("NDecimalesImp"))
                                    DrNew2("ImpIntereses") = xRound(.Item("ImpInteresPeriodo"), Drmoneda(0)("NDecimalesIMP"))
                                    DrNew2("ImpInteresesA") = xRound(!ImpInteresPeriodo * Drmoneda(0)("CambioA"), DtMonedaA.Rows(0)("NDecimalesImp"))
                                    'DrNew2("ImpInteresesB") = xRound(!ImpInteresPeriodo * DtMoneda.Rows(0)("CambioB"), DtMonedaB.Rows(0)("NDecimalesImp"))
                                End If
                                DblRecupCosteFinal += DrNew2("ImpRecuperacionCoste")
                                If DateAdd(StrUnidad, LngPeriodo * .Item("Periodo"), DteFechaCom) >= DteFechaTope Then
                                    If DblBienTotal <> DblRecupCosteFinal Then
                                        DblRecCosteUltima = DrNew2("ImpRecuperacionCoste") + (DblBienTotal - DblRecupCosteFinal)
                                        DrNew2("ImpRecuperacionCoste") = xRound(DblRecCosteUltima, Drmoneda(0)("NDecimalesIMP"))
                                        DrNew2("ImpRecuperacionCosteA") = xRound(DblRecCosteUltima * Drmoneda(0)("CambioA"), DtMonedaA.Rows(0)("NDecimalesImp"))
                                        'DrNew2("ImpRecuperacionCosteB") = xRound(DblRecCosteUltima * DtMoneda.Rows(0)("CambioB"), DtMoneda.Rows(0)("NDecimalesImp"))
                                    End If
                                End If
                                DrNew2("ImpCuota") = Dt.Rows(0)("ImpCuotaPeriodo")
                                DrNew2("ImpCuotaA") = xRound(Dt.Rows(0)("ImpCuotaPeriodo") * Drmoneda(0)("CambioA"), DtMonedaA.Rows(0)("NDecimalesImp"))
                                'DrNew2("ImpCuotaB") = xRound(Dt.Rows(0)("ImpCuotaPeriodo") * DtMoneda.Rows(0)("CambioB"), DtMoneda.Rows(0)("NDecimalesImp"))
                                If StrAgrup & String.Empty <> String.Empty Then DrNew2("IDAgrupacion") = StrAgrup
                                DrNew2("IDPagoPeriodo") = Dt.Rows(0)("ID")
                                DblImpIntTemp = DrNew2("ImpIntereses")
                                DblImpCuotaTemp = DrNew2("ImpCuota")
                                DtPago.Rows.Add(DrNew2)
                            End If
                            DblTotalInt += DblImpIntTemp
                            DblTotalCuota += DblImpCuotaTemp
                            LngPeriodo += 1
                            LngNumCuotas += 1
                        End While
                        If Dt.Rows(0)("ValorResidualIgualCuota") = False Then
                            Dim DrNew As DataRow = DtPago.NewRow()
                            DrNew("ImpCuota") = Dt.Rows(0)("ValorResidual")
                            DrNew("ImpCuotaA") = Dt.Rows(0)("ValorResidualA")
                            'drnew("ImpCuotaB") = Dt.Rows(0)("ValorResidualB")
                            If Not BlnSimulacion Then DrNew("IDPago") = AdminData.GetAutoNumeric
                            DrNew("Titulo") = Dt.Rows(0)("DescPago")
                            If LngPeriodo = 0 Then
                                DrNew("FechaVencimiento") = DteFechaCom
                            Else
                                DrNew("FechaVencimiento") = DateAdd(StrUnidad, LngPeriodo * Dt.Rows(0)("Periodo"), DteFechaCom)
                            End If
                            .Item("FechaUltimaActualizacion") = DrNew("FechaVencimiento")
                            DrNew("CContable") = .Item("IDCContable")
                            DrNew("IDProveedor") = .Item("IDProveedor")
                            DrNew("IDTipoPago") = .Item("IDTipoPago")
                            DrNew("IDFormaPago") = .Item("IDFormaPago")
                            DrNew("IDBancoPropio") = .Item("IDBancoPropio")
                            DrNew("IDMoneda") = .Item("IDMoneda")
                            DrNew("CambioA") = Drmoneda(0)("CambioA")
                            DrNew("CambioB") = Drmoneda(0)("CambioB")
                            DrNew("Situacion") = enumPagoSituacion.NoPagado
                            DrNew("Contabilizado") = enumPagoContabilizado.PagoNoContabilizado
                            DrNew("ImpVencimiento") = xRound(DrNew("ImpCuota") * (1 + DtTipoIva.Rows(0)("Factor") / 100), Drmoneda(0)("NDecimalesIMP"))
                            DrNew("ImpVencimientoA") = xRound(DrNew("ImpCuotaA") * (1 + DtTipoIva.Rows(0)("Factor") / 100) * Drmoneda(0)("CambioA"), DtMonedaA.Rows(0)("NDecimalesImp"))
                            'drnew("ImpVencimientoB") = xRound(drnew("ImpCuotaA") * (1 + DtTipoIva.Rows(0)("Factor") / 100) * DtMoneda.Rows(0)("CambioB"), DtMonedaB.Rows(0)("NDecimalesImp"))
                            DrNew("ImpRecuperacionCoste") = 0
                            DrNew("ImpRecuperacionCosteA") = 0
                            'drnew("ImpRecuperacionCosteB") = 0
                            DrNew("ImpIntereses") = 0
                            DrNew("ImpInteresesA") = 0
                            'drnew("ImpInteresesB") = 0
                            If StrAgrup & String.Empty <> String.Empty Then DrNew("IDAgrupacion") = StrAgrup
                            DrNew("IDPagoPeriodo") = Dt.Rows(0)("ID")
                            DtPago.Rows.Add(DrNew)
                        End If
                    End If
                    BusinessHelper.UpdateTable(DtPagoPer)
                End With
                If BlnSimulacion Then
                    Return DtPago
                Else
                    Return ClsPago.Update(DtPago)
                End If
            End If
        End If
    End Function

    Public Function InsertPagoPeriodico(ByVal Dt As DataTable, _
                                        ByVal DteFechaFinal As Date, _
                                        ByVal BlnSimulacion As Boolean) As DataTable
        Dim services As New ServiceProvider
        'Prepara el datatable para actualizar los Pagos que se han modificado manualmente
        Dim ClsMoneda As New Moneda
        Dim ClsPago As New Pago
        Dim ClsPagoPer As New PagoPeriodico
        Dim ClsTipoPago As New TipoPago
        Dim ClsTipoIva As New TipoIva
        Dim DtMoneda, DtMonedaA As DataTable ', DtMonedaB
        Dim DtPagoPer, DtPago, DtTipoPago, DtTipoIva As DataTable
        Dim DteFechaComienzo, DteFechaTope As Date
        Dim LngPeriodo As Long
        Dim StrUnidad, StrAgrup As String

        'DtMoneda = ClsMoneda.Filter()
        DtMonedaA = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf Moneda.ObtenerMonedaA, Nothing, services)
        'DtMonedaB = ClsMoneda.ObtenerMonedaB

        If Not Dt Is Nothing AndAlso Dt.Rows.Count <> 0 Then
            'Forma una lista de pagos periodicos
            Dim f As New Filter(FilterUnionOperator.Or)
            For Each Dr As DataRow In Dt.Select
                f.Add(New NumberFilterItem("ID", Dr("ID")))
            Next

            DtPagoPer = ClsPagoPer.Filter(f)
            ' DtPagoPer = ClsPagoPer.Filter(, " ID IN (" & StrIDPagoPer & ")")
            DtPago = ClsPago.Filter(New FilterItem("IDPago", FilterOperator.Equal, "-1"))
            Dim g As New NegocioGeneral
            For Each Dr As DataRow In Dt.Select
                Dim DrDatos() As DataRow = DtPagoPer.Select("Id=" & Dr("ID"))
                If Length(DrDatos.Length) > 0 Then
                    If Length(DrDatos(0)("IDTipoIva") & String.Empty) > 0 Then
                        DtTipoIva = ClsTipoIva.SelOnPrimaryKey(DrDatos(0)("IDTipoIva"))
                    End If
                    With DtPagoPer
                        'DrMoneda = DtMoneda.Select("IdMoneda = '" & .Rows(0)("IDMoneda") & "'")
                        DtMoneda = ClsMoneda.SelOnPrimaryKey(.Rows(0)("IDMoneda"))
                        If Nz(.Rows(0)("FechaUltimaActualizacion"), cnMinDate) < .Rows(0)("FechaFin") Then
                            StrUnidad = g.GetPeriodString(.Rows(0)("Unidad"))
                            If Length(.Rows(0)("FechaUltimaActualizacion") & String.Empty) = 0 Then
                                DteFechaComienzo = .Rows(0)("FechaInicio")
                            Else
                                DteFechaComienzo = DateAdd(StrUnidad, .Rows(0)("Periodo"), .Rows(0)("FechaUltimaActualizacion"))
                            End If
                            DteFechaTope = IIf(.Rows(0)("FechaFin") < DteFechaFinal, .Rows(0)("FechaFin"), DteFechaFinal)

                            DtTipoPago = ClsTipoPago.Filter(New FilterItem("IdTipoPago", FilterOperator.Equal, DtPagoPer.Rows(0)("IdTipoPago")))
                            If DtTipoPago.Rows.Count > 0 Then
                                StrAgrup = DtTipoPago.Rows(0)("IDAgrupacion") & String.Empty
                            End If
                            LngPeriodo = 0
                            While DateAdd(StrUnidad, LngPeriodo * .Rows(0)("Periodo"), DteFechaComienzo) <= DteFechaTope
                                Dim DrNew As DataRow = DtPago.NewRow()
                                If Not BlnSimulacion Then DrNew("IDPago") = AdminData.GetAutoNumeric
                                DrNew("Titulo") = .Rows(0)("DescPago")
                                If LngPeriodo = 0 Then
                                    DrNew("FechaVencimiento") = DteFechaComienzo
                                Else
                                    DrNew("FechaVencimiento") = DateAdd(StrUnidad, LngPeriodo * .Rows(0)("Periodo"), DteFechaComienzo)
                                End If
                                .Rows(0)("FechaUltimaActualizacion") = DrNew("FechaVencimiento")
                                DrNew("CContable") = .Rows(0)("IDCContable")
                                DrNew("IDProveedor") = .Rows(0)("IDProveedor")
                                DrNew("IdTipoPago") = .Rows(0)("IdTipoPago")
                                DrNew("IDFormaPago") = .Rows(0)("IDFormaPago")
                                DrNew("IDBancoPropio") = .Rows(0)("IDBancoPropio")
                                DrNew("IDMoneda") = .Rows(0)("IDMoneda")
                                DrNew("CambioA") = DtMoneda.Rows(0)("CambioA")
                                DrNew("CambioB") = DtMoneda.Rows(0)("CambioB")
                                DrNew("Situacion") = enumPagoSituacion.NoPagado
                                DrNew("Contabilizado") = enumPagoContabilizado.PagoNoContabilizado
                                DrNew("ImpVencimiento") = xRound(.Rows(0)("Importe"), DtMoneda.Rows(0)("NDecimalesImp"))
                                DrNew("ImpVencimientoA") = xRound(.Rows(0)("Importe") * DtMoneda.Rows(0)("CambioA"), DtMonedaA.Rows(0)("NDecimalesImp"))
                                'DrNew("ImpVencimientoB") = xRound(.Rows(0)("Importe") * DtMoneda.Rows(0)("CambioB"), DtMonedaB.Rows(0)("NDecimalesImp"))
                                DrNew("ImpRecuperacionCoste") = xRound(.Rows(0)("ImpRecuperacionCostePeriodo"), DtMoneda.Rows(0)("NDecimalesImp"))
                                DrNew("ImpRecuperacionCosteA") = xRound(.Rows(0)("ImpRecuperacionCostePeriodo") * DtMoneda.Rows(0)("CambioA"), DtMonedaA.Rows(0)("NDecimalesImp"))
                                'DrNew("ImpRecuperacionCosteB") = xRound(.Rows(0)("ImpRecuperacionCostePeriodo") * DtMoneda.Rows(0)("CambioB"), DtMonedaB.Rows(0)("NDecimalesImp"))
                                DrNew("ImpIntereses") = xRound(.Rows(0)("ImpInteresPeriodo"), DtMoneda.Rows(0)("NDecimalesImp"))
                                DrNew("ImpInteresesA") = xRound(.Rows(0)("ImpInteresPeriodo") * DtMoneda.Rows(0)("CambioA"), DtMonedaA.Rows(0)("NDecimalesImp"))
                                'DrNew("ImpInteresesB") = xRound(.Rows(0)("ImpInteresPeriodo") * DtMoneda.Rows(0)("CambioB"), DtMonedaB.Rows(0)("NDecimalesImp)"))
                                DrNew("ImpCuota") = DrNew("ImpIntereses") + DrNew("ImpRecuperacionCoste")
                                DrNew("ImpCuotaA") = DrNew("ImpInteresesA") + DrNew("ImpRecuperacionCosteA")
                                'DrNew("ImpCuotaB") = DrNew("ImpInteresesB") + DrNew("ImpRecuperacionCosteB")
                                If StrAgrup & String.Empty <> String.Empty Then
                                    DrNew("IDAgrupacion") = StrAgrup
                                End If
                                DrNew("IdPagoPeriodo") = Dt.Rows(0)("ID")
                                DtPago.Rows.Add(DrNew)
                                LngPeriodo += 1
                            End While
                        End If
                    End With
                End If
            Next
            If Not BlnSimulacion Then
                ClsPago.Update(DtPago)
                ClsPagoPer.Update(DtPagoPer)
            End If
            Return DtPago

        End If
    End Function

    Public Function CargarAnticipadoDiferido(ByVal LngIDPagoPer As Integer) As DataTable
        Dim services As New ServiceProvider
        Dim ClsInmov As New Inmovilizado
        Dim DtPagoPer, DtGFT, DtAños As DataTable
        Dim DteFechaIni, DteFechaFin As Date
        Dim StrIDInmov As String
        Dim LngAñoIni, LngAñoFin, LngAño As Integer
        Dim DblGastoCont, DblTotalFiscal, DblTotalGastoTeor, DblTotalDif, DblTotalAnti As Double

        If LngIDPagoPer > 0 Then
            DtPagoPer = Me.SelOnPrimaryKey(LngIDPagoPer)
            If Not DtPagoPer Is Nothing AndAlso DtPagoPer.Rows.Count > 0 Then
                DteFechaIni = CDate(Format(DtPagoPer.Rows(0)("FechaContrato"), "dd/mm/yyyy"))
                DteFechaFin = CDate(Format(DtPagoPer.Rows(0)("FechaFin"), "dd/mm/yyyy"))
                StrIDInmov = DtPagoPer.Rows(0)("IDInmovilizado") & String.Empty
            End If
            Dim DtAntDif As New DataTable
            DtAntDif.Columns.Add("ID", Type.GetType("System.Int32"))
            DtAntDif.Columns.Add("IDInmovilizado", Type.GetType("System.String"))
            DtAntDif.Columns.Add("AñoContabilizado", Type.GetType("System.Int32"))
            DtAntDif.Columns.Add("GastoContable", Type.GetType("System.Double"))
            DtAntDif.Columns.Add("GastoFiscalTeorico", Type.GetType("System.Int32"))
            DtAntDif.Columns.Add("LimiteAmortizacion", Type.GetType("System.Double"))
            DtAntDif.Columns.Add("GastoFiscal", Type.GetType("System.Double"))
            DtAntDif.Columns.Add("Anticipado", Type.GetType("System.Double"))
            DtAntDif.Columns.Add("Diferido", Type.GetType("System.Double"))
            DtAntDif.Columns.Add("Impuesto", Type.GetType("System.Double"))

            If Length(StrIDInmov) > 0 Then
                Dim BEDataEngine As New BE.DataEngine
                Dim f As New Filter
                f.Add(New StringFilterItem("IDInmovilizado", StrIDInmov))
                DtAños = BEDataEngine.Filter("vNegCargarAnticipadoDiferidoAñoFinAmortizacion", f)
                f.Clear()
                f.Add(New NumberFilterItem("ID", LngIDPagoPer))
                DtGFT = BEDataEngine.Filter("vNegCargarAnticipadoDiferidoImpRecuperacionCoste", f)
                If Not DtAños Is Nothing AndAlso DtAños.Rows.Count > 0 Then
                    LngAñoFin = DtAños.Rows(0)("AñoFin")
                    LngAñoIni = DtAños.Rows(0)("AñoInicio")
                    LngAño = LngAñoIni
                    If LngAñoFin < Year(DteFechaFin) Then
                        LngAñoFin = Year(DteFechaFin)
                    End If
                    While LngAño <= LngAñoFin
                        Dim DrNew As DataRow = DtAntDif.NewRow
                        DrNew("ID") = LngIDPagoPer
                        DrNew("IDInmovilizado") = StrIDInmov
                        DrNew("AñoContabilizado") = LngAño
                        Dim StDatos As New Inmovilizado.DatosAmortContAño
                        StDatos.IDInmovilizado = StrIDInmov
                        StDatos.Año = LngAño
                        StDatos.BlnAño = False
                        DblGastoCont = ProcessServer.ExecuteTask(Of Inmovilizado.DatosAmortContAño, Double)(AddressOf Inmovilizado.ObtenerAmortContAño, StDatos, services)
                        If DblGastoCont > 0 Then DrNew("GastoContable") = DblGastoCont
                        If Length(DrNew("GastoCont")) > 0 Then
                            DrNew("LimiteAmortizacion") = DrNew("GastoContable") * 2
                        End If
                        Dim DrDatos() As DataRow = DtGFT.Select("Año=" & LngAño)
                        If DrDatos.Length > 0 Then
                            DrNew("GastoFiscalTeorico") = DrDatos(0)("ImpRecuperacionCoste")
                            DblTotalGastoTeor += Nz(DrNew("GastoFiscalTeorico"), 0)
                        End If
                        If DrNew("GastoContable") > Nz(DrNew("GastoFiscalTeorico"), 0) Then
                            DrNew("GastoFiscal") = DblTotalGastoTeor - DblTotalFiscal
                        Else
                            If DblTotalFiscal + Nz(DrNew("LimiteAmortizacion"), 0) > DtPagoPer.Rows(0)("ImpRecuperacionCoste") Then
                                DrNew("GastoFiscal") = DrNew("ImpRecuperacionCoste") - DblTotalFiscal
                            Else
                                DrNew("GastoFiscal") = DrNew("LimiteAmortizacion")
                            End If
                        End If
                        DblTotalFiscal += Nz(DrNew("GastoFiscal"), 0)
                        DrNew("Impuesto") = DrNew("GastoFiscal") - DrNew("GastoContable") * 0.35
                        If DrNew("GastoFiscal") - DrNew("GastoContable") > 0 Then
                            DblTotalAnti += Nz(DrNew("Impuesto"), 0)
                        Else
                            DblTotalDif += Nz(DrNew("Impuesto"), 0)
                            If Math.Abs(DblTotalDif) > DblTotalAnti Then
                                DrNew("Impuesto") = DblTotalAnti + DblTotalDif + Nz(DrNew("Impuesto"), 0)
                            End If
                        End If
                        LngAño += 1
                        DtAntDif.Rows.Add(DrNew)
                    End While
                Else
                    LngAñoFin = Year(DteFechaFin)
                    LngAñoIni = Year(DteFechaIni)
                End If
                Return DtAntDif
            End If
        End If
    End Function

    Public Sub RevisarCondiciones(ByRef Dt As DataTable)
        Dim services As New ServiceProvider
        Dim ClsPago As New Pago
        Dim ClsTipoIVA As New TipoIva
        Dim ClsMoneda As New Moneda
        Dim DtPago, DtTipoIVA, DtMoneda As DataTable
        Dim DblFactorIVA, DblCuotaFija As Double
        Dim LngIDPagoPerAux, LngDecA, LngDecB As Integer

        If Not Dt Is Nothing AndAlso Dt.Rows.Count > 0 Then
            'For Each Dr As DataRow In Dt.Select
            '    StrListPagosPer = Dt.Rows(0)("IDPagoPeriodico")
            'Next
            Dim Values(-1) As Object
            For Each r As DataRow In Dt.Select
                ReDim Preserve Values(UBound(Values) + 1)
                Values(UBound(Values)) = Dt.Rows(0)("IDPagoPeriodico")
            Next
            'If Length(StrListPagosPer) > 0 Then
            If Values.Length > 0 Then
                Dim FilPagosPer As New Filter
                FilPagosPer.Add(New InListFilterItem("IDPagoPeriodo", Values, FilterType.Numeric))
                'FilPagosPer.Add(New InListFilterItem("IDPagoPeriodo", StrListPagosPer, FilterType.String))
                FilPagosPer.Add(New NumberFilterItem("Situacion", FilterOperator.Equal, enumPagoSituacion.Pagado))
                'FilPagosPer.Add("Situacion", FilterOperator.Equal, enumPagoSituacion.Pagado)
                DtPago = ClsPago.Filter(FilPagosPer, "IDPagoPeriodo")
                If Not DtPago Is Nothing AndAlso DtPago.Rows.Count > 0 Then
                    DtMoneda = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf Moneda.ObtenerMonedaA, Nothing, services)
                    If Not DtMoneda Is Nothing AndAlso DtMoneda.Rows.Count > 0 Then
                        LngDecA = DtMoneda.Rows(0)("NDecimalesImp")
                    End If
                    DtMoneda.Rows.Clear()
                    DtMoneda = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf Moneda.ObtenerMonedaB, Nothing, services)
                    If Not DtMoneda Is Nothing AndAlso DtMoneda.Rows.Count > 0 Then
                        LngDecB = DtMoneda.Rows(0)("NDecimalesImp")
                    End If
                    For Each Dr As DataRow In DtPago.Select()
                        If LngIDPagoPerAux <> Dr("IdPagoPeriodo") Then
                            LngIDPagoPerAux = Dt.Rows(0)("IDPagoPeriodico")
                            DblCuotaFija = Nz(Dt.Rows(0)("CuotaFija"))
                            Dim FiltroTipoIva As New Filter
                            FiltroTipoIva.Add(New NumberFilterItem("ID", FilterOperator.Equal, Dr("IDPagoPeriodo")))
                            DtTipoIVA = Filter(FiltroTipoIva)
                            If Not DtTipoIVA Is Nothing AndAlso DtTipoIVA.Rows.Count > 0 Then
                                DblFactorIVA = 1
                            Else
                                DtTipoIVA = ClsTipoIVA.Filter(New FilterItem("IDTipoIva=", FilterOperator.Equal, DtTipoIVA.Rows(0)("IDTipoIva")))
                                If Not DtTipoIVA Is Nothing AndAlso DtTipoIVA.Rows.Count > 0 Then
                                    DblFactorIVA = 1
                                Else
                                    DblFactorIVA = (DtTipoIVA.Rows(0)("Factor") / 100) + 1
                                End If
                            End If
                        End If
                        DtMoneda = ClsMoneda.SelOnPrimaryKey(Dr("IDMoneda"))
                        Dr("ImpCuota") = DblCuotaFija
                        If Not DtMoneda Is Nothing AndAlso DtMoneda.Rows.Count > 0 Then
                            Dr("ImpCuota") = xRound(Dr("ImpCuota"), DtMoneda.Rows(0)("NDecimalesImp"))
                            Dr("ImpCuotaA") = xRound(Dr("ImpCuota") * DtMoneda.Rows(0)("CambioA"), LngDecA)
                            Dr("ImpCuotaB") = xRound(Dr("ImpCuota") * DtMoneda.Rows(0)("CambioB"), LngDecB)
                        End If
                        Dr("ImpIntereses") = Dr("ImpCuota") - Dr("ImpRecuperacionCoste")
                        Dr("ImpVencimiento") = Dr("ImpCuota") * DblFactorIVA
                        If Not DtMoneda Is Nothing AndAlso DtMoneda.Rows.Count > 0 Then
                            Dr("ImpInteresesA") = xRound(Dr("ImpIntereses") * DtMoneda.Rows(0)("CambioA"), LngDecA)
                            Dr("ImpInteresesB") = xRound(Dr("ImpIntereses") * DtMoneda.Rows(0)("CambioB"), LngDecB)
                            Dr("ImpVencimiento") = xRound(Dr("ImpVencimiento"), DtMoneda.Rows(0)("NDecimalesImp"))
                            Dr("ImpVencimientoA") = xRound(Dr("ImpVencimiento") * DtMoneda.Rows(0)("CambioA"), LngDecA)
                            Dr("ImpVencimientoB") = xRound(Dr("ImpVencimiento") * DtMoneda.Rows(0)("CambioB"), LngDecB)
                        End If
                    Next
                    BusinessHelper.UpdateTable(DtPago)
                End If
            End If
        End If

    End Sub

    Public Sub AnularRevision(ByVal LngIDPagoPer As Integer)
        Dim LngNOper As Integer
        Dim ClsDiarioCont As BusinessHelper = BusinessHelper.CreateBusinessObject("DiarioContable")
        Dim DtPagoPer As DataTable = Me.Filter(New FilterItem("ID", FilterOperator.Equal, LngIDPagoPer, FilterType.Numeric))
        If Not DtPagoPer Is Nothing AndAlso DtPagoPer.Rows.Count > 0 Then
            LngNOper = DtPagoPer.Rows(0)("NOperacion")
            DtPagoPer = ClsDiarioCont.Filter(New FilterItem("IDApunte", FilterOperator.Equal, LngNOper, FilterType.Numeric))
            If Not DtPagoPer Is Nothing AndAlso DtPagoPer.Rows.Count > 0 Then
                NegocioGeneral.DeleteWhere(DtPagoPer.Rows(0)("IDEjercicio"), "NAsiento=" & DtPagoPer.Rows(0)("NAsiento"))
            End If
        End If

        Exit Sub
    End Sub

    <Task()> Public Shared Sub BorrarPagoPeriodicoImpuesto(ByVal IDPagoPeriodico As Integer, ByVal services As ServiceProvider)
        Dim Clsp As New PagoPeriodicoImpuesto
        Dim dtPagoPerImpuesto As DataTable = Clsp.SelOnPrimaryKey(IDPagoPeriodico)
        Clsp.Delete(dtPagoPerImpuesto)
    End Sub

    Public Function DtLeasingPorBanco(ByVal pStrClausulaWhere As String) As DataTable
        Dim vSQL As String
        vSQL = "SELECT tbPago.IdTipoPago, tbPago.IDBancoPropio, tbMaestroBancoPropio.DescBancoPropio," _
                & "tbMaestroBancoPropio.IDBanco, SUM(CASE MONTH(FechaVencimiento) WHEN 1 THEN tbPago.ImpCuotaA ELSE 0 END) AS SEnero," _
                & "SUM(CASE MONTH(FechaVencimiento) WHEN 2 THEN tbPago.ImpCuotaA ELSE 0 END) AS SFebrero," _
                & "SUM(CASE MONTH(FechaVencimiento) WHEN 3 THEN tbPago.ImpCuotaA ELSE 0 END) AS SMarzo," _
                & "SUM(CASE MONTH(FechaVencimiento) WHEN 4 THEN tbPago.ImpCuotaA ELSE 0 END) AS SAbril," _
                & "SUM(CASE MONTH(FechaVencimiento) WHEN 5 THEN tbPago.ImpCuotaA ELSE 0 END) AS SMayo," _
                & "SUM(CASE MONTH(FechaVencimiento) WHEN 6 THEN tbPago.ImpCuotaA ELSE 0 END) AS SJunio," _
                & "SUM(CASE MONTH(FechaVencimiento) WHEN 7 THEN tbPago.ImpCuotaA ELSE 0 END) AS SJulio," _
                & "SUM(CASE MONTH(FechaVencimiento) WHEN 8 THEN tbPago.ImpCuotaA ELSE 0 END) AS SAgosto," _
                & "SUM(CASE MONTH(FechaVencimiento) WHEN 9 THEN tbPago.ImpCuotaA ELSE 0 END) AS SSeptiembre," _
                & "SUM(CASE MONTH(FechaVencimiento) WHEN 10 THEN tbPago.ImpCuotaA ELSE 0 END) AS SOctubre," _
                & "SUM(CASE MONTH(FechaVencimiento) WHEN 11 THEN tbPago.ImpCuotaA ELSE 0 END) AS SNoviembre," _
                & "SUM(CASE MONTH(FechaVencimiento) WHEN 12 THEN tbPago.ImpCuotaA ELSE 0 END) AS SDiciembre," _
                & "SUM(tbPago.ImpCuotaA) AS STotalLinea " _
                & "FROM tbPago INNER JOIN tbMaestroBancoPropio ON tbPago.IDBancoPropio = tbMaestroBancoPropio.IDBancoPropio INNER JOIN  tbMaestroTipoPago ON tbPago.IdTipoPago = tbMaestroTipoPago.IdTipoPago"
        'vSQL &= " WHERE (tbPago.IdTipoPago = 8) "
        If Length(pStrClausulaWhere) > 0 Then vSQL &= " WHERE " & pStrClausulaWhere
        vSQL &= "GROUP BY tbPago.IdTipoPago, tbPago.IDBancoPropio, tbMaestroBancoPropio.DescBancoPropio, tbMaestroBancoPropio.IDBanco, tbMaestroTipoPago.LeasingSN "
        vSQL &= " HAVING (tbMaestroTipoPago.LeasingSN = 1)"
        vSQL &= "ORDER BY tbPago.IDBancoPropio"
        Dim CmdInversion As Common.DbCommand = AdminData.GetCommand
        CmdInversion.CommandType = CommandType.Text
        CmdInversion.CommandText = vSQL
        Return AdminData.Execute(CmdInversion, ExecuteCommand.ExecuteReader)
    End Function

    Public Function DtLeasingOpcionCompra(ByVal pStrClausulaWhere As String) As DataTable
        Dim vSQL As String
        vSQL = "SELECT tbPago.IdTipoPago, tbPago.IDBancoPropio, tbMaestroBancoPropio.DescBancoPropio," _
                & "tbMaestroBancoPropio.IDBanco, SUM(CASE MONTH(tbPago.FechaVencimiento) WHEN 1 THEN tbPago.ImpCuotaA ELSE 0 END) AS SEnero," _
                & "SUM(CASE MONTH(tbPago.FechaVencimiento) WHEN 2 THEN tbPago.ImpCuotaA ELSE 0 END) AS SFebrero," _
                & "SUM(CASE MONTH(tbPago.FechaVencimiento) WHEN 3 THEN tbPago.ImpCuotaA ELSE 0 END) AS SMarzo," _
                & "SUM(CASE MONTH(tbPago.FechaVencimiento) WHEN 4 THEN tbPago.ImpCuotaA ELSE 0 END) AS SAbril," _
                & "SUM(CASE MONTH(tbPago.FechaVencimiento) WHEN 5 THEN tbPago.ImpCuotaA ELSE 0 END) AS SMayo," _
                & "SUM(CASE MONTH(tbPago.FechaVencimiento) WHEN 6 THEN tbPago.ImpCuotaA ELSE 0 END) AS SJunio," _
                & "SUM(CASE MONTH(tbPago.FechaVencimiento) WHEN 7 THEN tbPago.ImpCuotaA ELSE 0 END) AS SJulio," _
                & "SUM(CASE MONTH(tbPago.FechaVencimiento) WHEN 8 THEN tbPago.ImpCuotaA ELSE 0 END) AS SAgosto," _
                & "SUM(CASE MONTH(tbPago.FechaVencimiento) WHEN 9 THEN tbPago.ImpCuotaA ELSE 0 END) AS SSeptiembre," _
                & "SUM(CASE MONTH(tbPago.FechaVencimiento) WHEN 10 THEN tbPago.ImpCuotaA ELSE 0 END) AS SOctubre," _
                & "SUM(CASE MONTH(tbPago.FechaVencimiento) WHEN 11 THEN tbPago.ImpCuotaA ELSE 0 END) AS SNoviembre," _
                & "SUM(CASE MONTH(tbPago.FechaVencimiento) WHEN 12 THEN tbPago.ImpCuotaA ELSE 0 END) AS SDiciembre," _
                & "SUM(tbPago.ImpCuotaA) AS STotalLinea " _
                & "FROM tbPago INNER JOIN tbMaestroBancoPropio ON tbPago.IDBancoPropio = tbMaestroBancoPropio.IDBancoPropio INNER JOIN " _
                & "vrptLeasingUltimaCuotaAux ON tbPago.FechaVencimiento = vrptLeasingUltimaCuotaAux.FechaVencimiento AND " _
                & "tbPago.IdPagoPeriodo = vrptLeasingUltimaCuotaAux.IdPagoPeriodo "
        'vSQL &= " WHERE (tbPago.IdTipoPago = 8) "
        If Length(pStrClausulaWhere) > 0 Then vSQL &= "WHERE " & pStrClausulaWhere
        vSQL &= "GROUP BY tbPago.IdTipoPago, tbPago.IDBancoPropio, tbMaestroBancoPropio.DescBancoPropio, tbMaestroBancoPropio.IDBanco, vrptLeasingUltimaCuotaAux.LeasingSN"
        vSQL &= " HAVING(vrptLeasingUltimaCuotaAux.LeasingSN = 1)"
        vSQL &= "ORDER BY tbPago.IDBancoPropio"

        Dim CmdInversion As Common.DbCommand = AdminData.GetCommand
        CmdInversion.CommandType = CommandType.Text
        CmdInversion.CommandText = vSQL
        Return AdminData.Execute(CmdInversion, ExecuteCommand.ExecuteReader)

    End Function

    Public Function DtLeasingUltimaCuota(ByVal pStrClausulaWhere As String) As DataTable
        Dim vSQL As String
        vSQL = "SELECT tbPago.IdTipoPago, tbPago.IDBancoPropio, tbMaestroBancoPropio.DescBancoPropio," _
                & "tbMaestroBancoPropio.IDBanco, SUM(CASE MONTH(tbPago.FechaVencimiento) WHEN 1 THEN tbPago.ImpCuotaA ELSE 0 END) AS SEnero," _
                & "SUM(CASE MONTH(tbPago.FechaVencimiento) WHEN 2 THEN tbPago.ImpCuotaA ELSE 0 END) AS SFebrero," _
                & "SUM(CASE MONTH(tbPago.FechaVencimiento) WHEN 3 THEN tbPago.ImpCuotaA ELSE 0 END) AS SMarzo," _
                & "SUM(CASE MONTH(tbPago.FechaVencimiento) WHEN 4 THEN tbPago.ImpCuotaA ELSE 0 END) AS SAbril," _
                & "SUM(CASE MONTH(tbPago.FechaVencimiento) WHEN 5 THEN tbPago.ImpCuotaA ELSE 0 END) AS SMayo," _
                & "SUM(CASE MONTH(tbPago.FechaVencimiento) WHEN 6 THEN tbPago.ImpCuotaA ELSE 0 END) AS SJunio," _
                & "SUM(CASE MONTH(tbPago.FechaVencimiento) WHEN 7 THEN tbPago.ImpCuotaA ELSE 0 END) AS SJulio," _
                & "SUM(CASE MONTH(tbPago.FechaVencimiento) WHEN 8 THEN tbPago.ImpCuotaA ELSE 0 END) AS SAgosto," _
                & "SUM(CASE MONTH(tbPago.FechaVencimiento) WHEN 9 THEN tbPago.ImpCuotaA ELSE 0 END) AS SSeptiembre," _
                & "SUM(CASE MONTH(tbPago.FechaVencimiento) WHEN 10 THEN tbPago.ImpCuotaA ELSE 0 END) AS SOctubre," _
                & "SUM(CASE MONTH(tbPago.FechaVencimiento) WHEN 11 THEN tbPago.ImpCuotaA ELSE 0 END) AS SNoviembre," _
                & "SUM(CASE MONTH(tbPago.FechaVencimiento) WHEN 12 THEN tbPago.ImpCuotaA ELSE 0 END) AS SDiciembre," _
                & "SUM(tbPago.ImpCuotaA) AS STotalLinea " _
                & "FROM tbPago INNER JOIN tbMaestroBancoPropio ON tbPago.IDBancoPropio = tbMaestroBancoPropio.IDBancoPropio INNER JOIN " _
                & "vrptLeasingPenultimaCuotaAux ON tbPago.FechaVencimiento = vrptLeasingPenultimaCuotaAux.FechaVencimiento AND " _
                & "tbPago.IdPagoPeriodo = vrptLeasingPenultimaCuotaAux.IdPagoPeriodo "

        'vSQL &= " WHERE (tbPago.IdTipoPago = 8) "
        If Length(pStrClausulaWhere) > 0 Then vSQL &= "WHERE " & pStrClausulaWhere
        vSQL &= "GROUP BY tbPago.IdTipoPago, tbPago.IDBancoPropio, tbMaestroBancoPropio.DescBancoPropio, tbMaestroBancoPropio.IDBanco, vrptLeasingPenultimaCuotaAux.LeasingSN"
        vSQL &= " HAVING(vrptLeasingPenultimaCuotaAux.LeasingSN = 1)"
        vSQL &= "ORDER BY tbPago.IDBancoPropio"
        Dim CmdInversion As Common.DbCommand = AdminData.GetCommand
        CmdInversion.CommandType = CommandType.Text
        CmdInversion.CommandText = vSQL
        Return AdminData.Execute(CmdInversion, ExecuteCommand.ExecuteReader)
    End Function

    Public Function TienePagosContabilizados(ByVal lngIDPagoPeriodico As Long) As Boolean

        Dim fwnPago As New Pago
        Dim Filtro As New Filter
        Filtro.Add("Contabilizado", FilterOperator.Equal, True, FilterType.Boolean)
        Filtro.Add("IdPagoPeriodo", FilterOperator.Equal, lngIDPagoPeriodico, FilterType.Numeric)
        Dim dtPago As DataTable = fwnPago.Filter(Filtro)
        If Not IsNothing(dtPago) AndAlso dtPago.Rows.Count > 0 Then
            TienePagosContabilizados = True
        Else
            TienePagosContabilizados = False
        End If

    End Function

    'Private Function GenerarFacturaCabecera(ByRef DtPagoPer As DataTable) As Integer
    '    Dim ClsFCC As New FacturaCompraCabecera
    '    Dim ClsFCL As New FacturaCompraLinea
    '    Dim ClsProv As New Proveedor
    '    Dim ClsBancoProv As New ProveedorBanco
    '    Dim ClsMoneda As New Moneda
    '    Dim DtFCC, DtFCL, DtProv, DtBancoProv, DtMoneda As DataTable

    '    DtFCC = ClsFCC.AddNewForm
    '    With DtFCC.Rows(0)
    '        .Item("FechaFactura") = Today.Date
    '        .Item("SuFechaFactura") = Today.Date
    '        .Item("IDProveedor") = DtPagoPer.Rows(0)("IDProveedor")
    '        DtProv = ClsProv.SelOnPrimaryKey(DtPagoPer.Rows(0)("IDProveedor"))
    '        If Not DtProv Is Nothing AndAlso DtProv.Rows.Count > 0 Then
    '            .Item("RazonSocial") = DtProv.Rows(0)("RazonSocial")
    '            .Item("CifProveedor") = DtProv.Rows(0)("CifProveedor")
    '            .Item("IDFormaPago") = DtProv.Rows(0)("IDFormaPago")
    '            .Item("IDCondicionPago") = DtProv.Rows(0)("IDCondicionPago")
    '            .Item("IDDiaPago") = DtProv.Rows(0)("IDDiaPago")
    '            .Item("IDMoneda") = DtPagoPer.Rows(0)("IDMoneda")
    '            DtMoneda = ClsMoneda.SelOnPrimaryKey(.Item("IDMoneda"))
    '            If Not DtMoneda Is Nothing AndAlso DtMoneda.Rows.Count > 0 Then
    '                .Item("CambioA") = DtMoneda.Rows(0)("CambioA")
    '                .Item("CambioB") = DtMoneda.Rows(0)("CambioB")
    '            End If
    '            .Item("IDCondicionPago") = DtProv.Rows(0)("IDCondicionPago")
    '            .Item("IDBancoPropio") = DtPagoPer.Rows(0)("IDBancoPropio")
    '            Dim FilProv As New Filter
    '            FilProv.Add("IDProveedor", FilterOperator.Equal, DtProv.Rows(0)("IDProveedor"), FilterType.String)
    '            FilProv.Add("Predeterminado", FilterOperator.Equal, 1, FilterType.Boolean)
    '            DtBancoProv = ClsBancoProv.Filter(FilProv)
    '            If Not DtBancoProv Is Nothing AndAlso DtBancoProv.Rows.Count > 0 Then
    '                .Item("IDProveedorBanco") = DtBancoProv.Rows(0)("IDProveedorBanco")
    '            End If
    '        End If
    '        .Item("Estado") = enumfccEstado.fccContabilizado
    '        .Item("FacturaPagoPeriodicoSN") = 1
    '    End With
    '    DtFCC = ClsFCC.Update(DtFCC)
    '    Return DtFCC.Rows(0)("IDFactura")
    'End Function

    'Private Sub GenerarFacturaLinea(ByRef DtFCC As DataTable, ByRef DtPago As DataTable, _
    'ByRef DtPagoPer As DataTable, ByVal LngTipoFact As Integer)
    '    Dim ClsFCL As New FacturaCompraLinea
    '    Dim ClsRef As New ArticuloProveedor
    '    Dim ClsCompra As New Compra
    '    Dim ClsUds As New ArticuloUnidadAB
    '    Dim ClsProv As New Proveedor
    '    Dim DtFCL, DtArt, DtRef, DtProv As DataTable

    '    DtFCL = ClsFCL.AddNewForm
    '    With DtFCL.Rows(0)
    '        .Item("IDFactura") = DtFCC.Rows(0)("IDFactura")
    '        .Item("NFactura") = DtFCC.Rows(0)("NFactura")
    '        .Item("IDCentroGestion") = DtFCC.Rows(0)("IDCentroGestion")
    '        .Item("Cantidad") = 1
    '        .Item("Factor") = 1
    '        .Item("QInterna") = 1
    '        .Item("UdValoracion") = 1
    '        If LngTipoFact = 0 Then
    '            .Item("Precio") = DtPagoPer.Rows(0)("ImporteNetoNominal")
    '            .Item("Importe") = DtPagoPer.Rows(0)("ImporteNetoNominal")
    '            .Item("CContable") = DtPagoPer.Rows(0)("CCNominal")
    '        ElseIf LngTipoFact = 1 Then
    '            .Item("Precio") = DtPago.Rows(0)("ImpBaseImponible")
    '            .Item("Importe") = DtPago.Rows(0)("ImpBaseImponible")
    '            If Length(DtPagoPer.Rows(0)("IDProveedor")) > 0 Then
    '                DtProv = ClsProv.SelOnPrimaryKey(DtPagoPer.Rows(0)("IDProveedor"))
    '                If Not DtProv Is Nothing AndAlso DtProv.Rows.Count > 0 Then
    '                    .Item("CContable") = DtProv.Rows(0)("CCInmovilizadoCortoPlazo")
    '                End If
    '            End If
    '        End If
    '        .Item("Dto1") = 0
    '        .Item("Dto2") = 0
    '        .Item("Dto3") = 0
    '        .Item("TipoGastoObra") = enumfclTipoGastoObra.enumfclMaterial
    '        DtArt = AdminData.Filter("vNegCaractArticulo", , "IDArticulo='" & DtPagoPer.Rows(0)("IDArticulo") & "' AND Compra=1 AND Activo=1")
    '        If Not DtArt Is Nothing AndAlso DtArt.Rows.Count > 0 Then
    '            .Item("IDArticulo") = DtPagoPer.Rows(0)("IDArticulo")
    '            DtRef = ClsRef.SelOnPrimaryKey(DtPagoPer.Rows(0)("IDProveedor"), DtPagoPer.Rows(0)("IDArticulo"))
    '            If Not DtRef Is Nothing AndAlso DtRef.Rows.Count > 0 Then
    '                .Item("RefProveedor") = DtRef.Rows(0)("RefProveedor") & String.Empty
    '                .Item("DescArticulo") = DtRef.Rows(0)("DescRefProveedor") & String.Empty
    '                .Item("IDUDMedida") = DtRef.Rows(0)("IdUdCompra") & String.Empty
    '                .Item("UdValoracion") = DtRef.Rows(0)("UdValoracion")
    '            End If
    '            If Length(.Item("DescArticulo")) = 0 Then .Item("DescArticulo") = DtArt.Rows(0)("DescArticulo")
    '            .Item("IDTipoIva") = ClsCompra.ObtenerIVA(DtPagoPer.Rows(0)("IDProveedor"), DtPagoPer.Rows(0)("IDArticulo"))
    '            If .Item("UdValoracion") = 0 Then .Item("UdValoracion") = IIf(DtArt.Rows(0)("UdValoracion") > 0, DtArt.Rows(0)("UdValoracion"), 1)
    '            .Item("Factor") = ClsUds.FactorDeConversion(.Item("IDArticulo"), .Item("IDUDMedida"), .Item("IDUDInterna"))
    '        End If
    '    End With
    '    ClsFCL.Update(DtFCL)
    'End Sub

    'Crea una factura a partir de un pago de un pago periódico
    'Public Function GenerarFactura(ByVal LngIDPago As Integer, ByVal LngTipoFact As Integer) As String
    '    Dim ClsFCC As New FacturaCompraCabecera
    '    Dim ClsPago As New Pago
    '    Dim DtFCC, DtPago, DtPagoPer As DataTable
    '    Dim LngIDFact As Integer

    '    DtPago = ClsPago.SelOnPrimaryKey(LngIDPago)
    '    If Not DtPago Is Nothing AndAlso DtPago.Rows.Count > 0 Then
    '        'obtiene el pago periodico a traves del pago
    '        DtPagoPer = Me.SelOnPrimaryKey(DtPago.Rows(0)("IdPagoPeriodo"))
    '        If Not DtPagoPer Is Nothing AndAlso DtPagoPer.Rows.Count > 0 Then
    '            LngIDFact = GenerarFacturaCabecera(DtPagoPer)
    '            DtFCC = ClsFCC.SelOnPrimaryKey(LngIDFact)
    '            If Not DtFCC Is Nothing AndAlso DtFCC.Rows.Count > 0 Then
    '                GenerarFactura = DtFCC.Rows(0)("NFactura")
    '                GenerarFacturaLinea(DtFCC, DtPago, DtPagoPer, LngTipoFact)
    '            End If
    '        End If
    '        DtPago.Rows(0)("IDFactura") = DtFCC.Rows(0)("IDFactura")
    '        DtPago.Rows(0)("NFactura") = DtFCC.Rows(0)("NFactura")
    '        ClsPago.Update(DtPago)
    '    End If
    'End Function

#End Region

End Class

