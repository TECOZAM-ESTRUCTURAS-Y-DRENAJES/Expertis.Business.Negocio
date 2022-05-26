Public Class BancoPropioInfo
    Inherits ClassEntityInfo

    Public IDBancoPropio As String
    Public DescBancoPropio As String
    Public CContable As String
    Public CEfectosDescontados As String
    Public IDCContableTalon As String
    Public IDBanco As String
    Public Sucursal As String
    Public DigitoControl As String
    Public NCuenta As String
    Public PagoIntereses As Boolean
    Public BaseCalculo As Integer
    Public CarenciaConIntereses As Boolean
    Public IDCContableAnticipo As String
    Public SufijoRemesas As String
    Public SWIFT As String
    Public CodigoIBAN As String

    Public Overrides Sub Fill(ByVal ParamArray PrimaryKey() As Object)
        Dim dt As DataTable = New BancoPropio().SelOnPrimaryKey(PrimaryKey)
        If IsNothing(dt) OrElse dt.Rows.Count = 0 Then
            ApplicationService.GenerateError("El Banco Propio {0} no existe.", Quoted(PrimaryKey(0)))
        Else
            Me.Fill(dt.Rows(0))
        End If
    End Sub

End Class

Public Class BancoPropio

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroBancoPropio"

#End Region

#Region "Eventos GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBrl As New BusinessRules
        oBrl.Add("IDBanco", AddressOf CambioBanco)
        Return oBrl
    End Function

    <Task()> Public Shared Sub CambioBanco(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) = 0 Then
            data.Current("Sucursal") = String.Empty
            data.Current("DigitoControl") = String.Empty
            data.Current("NCuenta") = String.Empty
        End If
    End Sub

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
        validateProcess.AddTask(Of DataRow)(AddressOf General.Comunes.ValidarCodigoIBAN)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("DescBancoPropio")) = 0 Then ApplicationService.GenerateError("Introduzca la descripción del banco propio")
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDBancoPropio")) = 0 Then ApplicationService.GenerateError("El Banco Propio es un dato obligatorio.")
            Dim dt As DataTable = New BancoPropio().SelOnPrimaryKey(data("IDBancoPropio"))
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then ApplicationService.GenerateError("Ya existe un Banco Propio con ese valor.")
        End If
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarDigitoControl)
    End Sub

    <Task()> Public Shared Sub AsignarDigitoControl(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDBanco")) > 0 AndAlso Length(data("Sucursal")) > 0 AndAlso Length(data("NCuenta")) > 0 Then
            Dim dataDC As New NegocioGeneral.dataCalculoDigitosControl(data("IDBanco"), data("Sucursal"), data("NCuenta"))
            data("DigitoControl") = ProcessServer.ExecuteTask(Of NegocioGeneral.dataCalculoDigitosControl, String)(AddressOf NegocioGeneral.CalculoDigitosControl, dataDC, services)
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Serializable()> _
    Public Class DatosCalcTesBanco
        Public DtTesoreria As DataTable
        Public Atrasados As Boolean
        Public FechaDesde As Date
        Public FechaHasta As Date

        Public Sub New(ByVal DtTesoreria As DataTable, ByVal Atrasados As Boolean, ByVal FechaDesde As Date, Optional ByVal FechaHasta As Date = cnMinDate)
            Me.DtTesoreria = DtTesoreria
            Me.Atrasados = Atrasados
            Me.FechaDesde = FechaDesde
            Me.FechaHasta = FechaHasta
        End Sub
    End Class

    <Task()> Public Shared Function CalcularTesoreriaPorBanco(ByVal data As DatosCalcTesBanco, ByVal services As ServiceProvider) As DataTable
        Dim drNewRowTesoreria As DataRow
        Dim strSelect As String
        Dim dtRemesa As DataTable
        Dim objNegBP As BancoPropio
        Dim dtBP As DataTable
        Dim rowBP As DataRow
        Dim dtSaldos As DataTable

        Dim dblPagoAt As Double
        Dim dblCobroAt As Double
        Dim dblPagoPe As Double
        Dim dblCobroPe As Double
        Dim dblCobroRemesa As Double
        Dim Importes As New DatosImportes

        Dim strIDEjercicio As String
        Dim aColumnsKey(1) As DataColumn

        '//Recuperar la información de la moneda
        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
        Dim MonInfoA As MonedaInfo = Monedas.MonedaA

        '//Cogemos las remesas en un DataTable
        strSelect = "IdRemesa, IdTipoNegociacion, FechaNegociacion"
        dtRemesa = AdminData.Filter("tbRemesa", strSelect)
        '//Establecer cual es la PrimaryKey de la tabla de Remesas, para poder buscar uan remesa concreta.
        aColumnsKey(0) = dtRemesa.Columns("IDRemesa")
        dtRemesa.PrimaryKey = aColumnsKey

        '//Cogemos el Ejercicio Predeterminado
        strIDEjercicio = ProcessServer.ExecuteTask(Of Date, String)(AddressOf NegocioGeneral.EjercicioPredeterminado, Today, services)

        '//Cogemos todos los  Bancos Propios
        dtBP = New BancoPropio().Filter("IdBancoPropio, DescBancoPropio, CContable")


        If Not IsNothing(dtBP) AndAlso dtBP.Rows.Count > 0 Then
            For Each rowBP In dtBP.Rows
                dblCobroAt = 0
                dblPagoAt = 0
                dblCobroPe = 0
                dblPagoPe = 0
                dblCobroRemesa = 0

                '//Calculamos los Importes Atrasados, del Periodo y de Remesa, de los cobros
                '//que pertenecen al Banco Propio actual
                Dim StDatos As New DatosObImpTesBanco(True, data.FechaDesde, data.FechaHasta, dtRemesa, rowBP("IDBancoPropio"))
                Importes = ProcessServer.ExecuteTask(Of DatosObImpTesBanco, DatosImportes)(AddressOf ObtenerImportesTesoreriaBanco, StDatos, services)
                If data.Atrasados Then dblCobroAt = CDbl(Importes.Atrasado)
                dblCobroPe = CDbl(Importes.Periodo)
                dblCobroRemesa = CDbl(Importes.Remesa)

                '//Calculamos los Importes Atrasados y del Periodo de los pagos que pertenecen
                '//al Banco Propio actual
                StDatos.FechaDesde = data.FechaDesde
                StDatos.FechaHasta = data.FechaHasta
                StDatos.DtRemesa = dtRemesa
                StDatos.Cobros = False
                StDatos.IDBancoPropio = rowBP("IDBancoPropio")
                Importes = ProcessServer.ExecuteTask(Of DatosObImpTesBanco, DatosImportes)(AddressOf ObtenerImportesTesoreriaBanco, StDatos, services)
                If data.Atrasados Then dblPagoAt = CDbl(Importes.Atrasado)
                dblPagoPe = CDbl(Importes.Periodo)

                '//Insertamos la linea en el DataTable
                With data.DtTesoreria
                    drNewRowTesoreria = .NewRow()
                    drNewRowTesoreria("IDBancoPropio") = rowBP("IDBancoPropio")
                    drNewRowTesoreria("DescBancoPropio") = rowBP("DescBancoPropio")
                    '//Calculamos el Saldo del BancoPropio
                    drNewRowTesoreria("SaldoA") = 0
                    If Not IsDBNull(rowBP("CContable")) Then
                        dtSaldos = NegocioGeneral.CuentaSaldo(strIDEjercicio, Nz(rowBP("CContable"), String.Empty))
                        If Not IsNothing(dtSaldos) AndAlso dtSaldos.Rows.Count > 0 Then
                            drNewRowTesoreria("SaldoA") = xRound(dtSaldos.Rows(0)("SaldoA"), MonInfoA.NDecimalesImporte)
                        End If
                    End If
                    If data.Atrasados Then
                        drNewRowTesoreria("PagosAtrasados") = xRound(dblPagoAt, MonInfoA.NDecimalesImporte)
                        drNewRowTesoreria("CobrosAtrasados") = xRound(dblCobroAt, MonInfoA.NDecimalesImporte)
                    End If
                    drNewRowTesoreria("PagosPeriodo") = xRound(dblPagoPe, MonInfoA.NDecimalesImporte)
                    drNewRowTesoreria("CobrosPeriodo") = xRound(dblCobroPe, MonInfoA.NDecimalesImporte)
                    drNewRowTesoreria("Remesas") = xRound(dblCobroRemesa, MonInfoA.NDecimalesImporte)
                    drNewRowTesoreria("PagosTotales") = xRound(dblPagoAt + dblPagoPe, MonInfoA.NDecimalesImporte)
                    drNewRowTesoreria("CobrosTotales") = xRound(dblCobroAt + dblCobroPe + dblCobroRemesa, MonInfoA.NDecimalesImporte)
                    drNewRowTesoreria("Total") = xRound(drNewRowTesoreria("SaldoA") - drNewRowTesoreria("PagosTotales") + drNewRowTesoreria("CobrosTotales"), MonInfoA.NDecimalesImporte)
                    .Rows.Add(drNewRowTesoreria)
                End With
            Next rowBP

            '//AÑADIMOS UNA LINEA PARA LOS COBROS Y PAGOS QUE NO TIENEN BANCOPROPIO

            '//Calculamos los Importes Atrasados, del Periodo y de Remesa, de los cobros
            '//que no tienen BancoPropio
            Dim StDatosF1 As New DatosObImpTesBanco(True, data.FechaDesde, data.FechaHasta, dtRemesa, String.Empty)
            Importes = ProcessServer.ExecuteTask(Of DatosObImpTesBanco, DatosImportes)(AddressOf ObtenerImportesTesoreriaBanco, StDatosF1, services)
            If data.Atrasados Then dblCobroAt = CDbl(Importes.Atrasado)
            dblCobroPe = CDbl(Importes.Periodo)
            dblCobroRemesa = CDbl(Importes.Remesa)

            '//Calculamos los Importes Atrasados y del Periodo de los pagos que no tienen BancoPropio
            Dim StDatosF2 As New DatosObImpTesBanco(False, data.FechaDesde, data.FechaHasta, dtRemesa, String.Empty)
            Importes = ProcessServer.ExecuteTask(Of DatosObImpTesBanco, DatosImportes)(AddressOf ObtenerImportesTesoreriaBanco, StDatosF2, services)
            If data.Atrasados Then dblPagoAt = CDbl(Importes.Atrasado)
            dblPagoPe = CDbl(Importes.Periodo)

            'Insertamos la linea en el DataTable
            With data.DtTesoreria
                drNewRowTesoreria = .NewRow()
                drNewRowTesoreria("DescBancoPropio") = "DESCONOCIDO"
                drNewRowTesoreria("SaldoA") = 0
                If data.Atrasados Then
                    drNewRowTesoreria("PagosAtrasados") = xRound(dblPagoAt, MonInfoA.NDecimalesImporte)
                    drNewRowTesoreria("CobrosAtrasados") = xRound(dblCobroAt, MonInfoA.NDecimalesImporte)
                End If
                drNewRowTesoreria("PagosPeriodo") = xRound(dblPagoPe, MonInfoA.NDecimalesImporte)
                drNewRowTesoreria("CobrosPeriodo") = xRound(dblCobroPe, MonInfoA.NDecimalesImporte)
                drNewRowTesoreria("Remesas") = xRound(dblCobroRemesa, MonInfoA.NDecimalesImporte)
                drNewRowTesoreria("PagosTotales") = xRound(dblPagoAt + dblPagoPe, MonInfoA.NDecimalesImporte)
                drNewRowTesoreria("CobrosTotales") = xRound(dblCobroAt + dblCobroPe + dblCobroRemesa, MonInfoA.NDecimalesImporte)
                drNewRowTesoreria("Total") = xRound(drNewRowTesoreria("SaldoA") - drNewRowTesoreria("PagosTotales") + drNewRowTesoreria("CobrosTotales"), MonInfoA.NDecimalesImporte)
                .Rows.Add(drNewRowTesoreria)
            End With
        End If

        Return data.DtTesoreria
    End Function

    <Serializable()> _
    Public Class DatosObImpTesBanco
        Public Cobros As Boolean
        Public FechaDesde As Date
        Public FechaHasta As Date
        Public DtRemesa As DataTable
        Public IDBancoPropio As String

        Public Sub New(ByVal Cobros As Boolean, ByVal FechaDesde As Date, ByVal FechaHasta As Date, ByVal DtRemesa As DataTable, ByVal IDBancoPropio As String)
            Me.Cobros = Cobros
            Me.FechaDesde = FechaDesde
            Me.FechaHasta = FechaHasta
            Me.DtRemesa = DtRemesa
            Me.IDBancoPropio = IDBancoPropio
        End Sub
    End Class

    <Serializable()> _
    Public Class DatosImportes
        Public Atrasado As Double
        Public Periodo As Double
        Public Remesa As Double

        Public Sub New()
        End Sub

        Public Sub New(ByVal Atrasado As Double, ByVal Periodo As Double, ByVal Remesa As Double)
            Me.Atrasado = Atrasado
            Me.Periodo = Periodo
            Me.Remesa = Remesa
        End Sub
    End Class

    <Task()> Public Shared Function ObtenerImportesTesoreriaBanco(ByVal data As DatosObImpTesBanco, ByVal services As ServiceProvider) As DatosImportes
        '//Calcula los Importes Atrasados, del Periodo y de Remesa, de los cobros o pagos
        '//que pertenecen al Banco Propio  que recibe como parametro
        Dim strWhere As String
        Dim strSelect As String
        Dim dt As DataTable
        Dim Importes As New DatosImportes

        If data.IDBancoPropio & String.Empty <> String.Empty Then
            strWhere = "IdBancoPropio='" & data.IDBancoPropio & "'"
        Else : strWhere = "IdBancoPropio IS NULL"
        End If

        If data.Cobros Then
            strSelect = "FechaVencimiento, ImpVencimientoA, IdRemesa, Contabilizado"
            dt = AdminData.Filter("vctlCITesoreriaPorBancoCobros", strSelect, strWhere)
        Else
            strSelect = "FechaVencimiento, ImpVencimientoA, Contabilizado"
            dt = AdminData.Filter("vctlCITesoreriaPorBancoPagos", strSelect, strWhere)
        End If

        If data.FechaHasta = DateTime.MinValue Then
            Dim StDatos As New DatosCalcImpTes(dt, data.Cobros, data.FechaDesde, Nothing, data.DtRemesa)
            Importes = ProcessServer.ExecuteTask(Of DatosCalcImpTes, DatosImportes)(AddressOf CalcularImportesTesoreria, StDatos, services)
        Else
            Dim StDatos As New DatosCalcImpTes(dt, data.Cobros, data.FechaDesde, data.FechaHasta, data.DtRemesa)
            Importes = ProcessServer.ExecuteTask(Of DatosCalcImpTes, DatosImportes)(AddressOf CalcularImportesTesoreria, StDatos, services)
        End If
        If Not IsNothing(dt) Then dt.Dispose()
        Return Importes
    End Function

    <Serializable()> _
    Public Class DatosCalcImpTes
        Public Dt As DataTable
        Public Cobro As Boolean
        Public FechaDesde As Date
        Public FechaHasta As Date
        Public DtRemesa As DataTable

        Public Sub New(ByVal Dt As DataTable, ByVal Cobro As Boolean, ByVal FechaDesde As Date, ByVal FechaHasta As Date, ByVal DtRemesa As DataTable)
            Me.Dt = Dt
            Me.Cobro = Cobro
            Me.FechaDesde = FechaDesde
            Me.FechaHasta = FechaHasta
            Me.DtRemesa = DtRemesa
        End Sub
    End Class

    <Task()> Public Shared Function CalcularImportesTesoreria(ByVal data As DatosCalcImpTes, ByVal services As ServiceProvider) As DatosImportes
        CalcularImportesTesoreria = New DatosImportes
        Dim dblPeriodo As Double
        Dim dblAtrasado As Double
        Dim dblRemesa As Double
        Dim row As DataRow
        Dim rowRemesa As DataRow

        If Not IsNothing(data.Dt) AndAlso data.Dt.Rows.Count > 0 Then
            For Each row In data.Dt.Rows
                '//Comprueba la fecha de Vencimiento
                If Not row("Contabilizado") Then
                    If row("FechaVencimiento") < data.FechaDesde Then
                        '//Si es anterior a la fecha desde, este es un Pago/Cobro atrasado, o no se
                        '//ha introducido FechaHasta
                        dblAtrasado = dblAtrasado + row("ImpVencimientoA")
                    ElseIf (data.FechaHasta = cnMinDate AndAlso row("FechaVencimiento") >= data.FechaDesde) OrElse _
                           (Not (data.FechaHasta = cnMinDate) AndAlso row("FechaVencimiento") >= data.FechaDesde AndAlso row("FechaVencimiento") <= data.FechaHasta) Then
                        '//Si está entre la fecha desde y la fecha hasta,  es un pago/cobro del periodo
                        If data.Cobro Then
                            'Si es un cobro, calcula tambien el Importe Remesa
                            If Not AreEquals(row("IdRemesa") & String.Empty, String.Empty) Then
                                'REMESA
                                With data.DtRemesa
                                    If Not IsNothing(data.DtRemesa) Then
                                        '//Cogemos los datos de la remesa del cobro actual
                                        rowRemesa = .Rows.Find(row("IdRemesa"))
                                        If Not IsNothing(rowRemesa) Then
                                            '//Si la remesa es alDescuento, se compara la FechaNegociacion de la remesa
                                            If AreEquals(rowRemesa("IdTipoNegociacion"), enumTipoRemesa.RemesaAlDescuento) Then
                                                If (rowRemesa("FechaNegociacion") >= data.FechaDesde And rowRemesa("FechaNegociacion") <= data.FechaHasta And Not (data.FechaHasta = cnMinDate)) Or ((data.FechaHasta = cnMinDate) And rowRemesa("FechaNegociacion") >= data.FechaDesde) Then
                                                    dblRemesa = dblRemesa + row("ImpVencimientoA")
                                                End If
                                            ElseIf AreEquals(rowRemesa("IdTipoNegociacion"), enumTipoRemesa.RemesaAlCobro) Then
                                                '//Si la remesa es AlCobro, se compara la FechaVencimiento del cobro
                                                If (row("FechaVencimiento") >= data.FechaDesde And row("FechaVencimiento") <= data.FechaHasta And Not IsNothing(data.FechaHasta)) Or (IsNothing(data.FechaHasta) And row("FechaVencimiento") >= data.FechaDesde) Then
                                                    dblRemesa = dblRemesa + row("ImpVencimientoA")
                                                End If
                                            End If
                                        End If
                                    End If
                                End With
                            Else
                                'COBRO
                                dblPeriodo = dblPeriodo + row("ImpVencimientoA")
                            End If
                        Else
                            'PAGO
                            dblPeriodo = dblPeriodo + row("ImpVencimientoA")
                        End If
                    End If

                End If

            Next row
        End If
        CalcularImportesTesoreria.Atrasado = dblAtrasado
        CalcularImportesTesoreria.Periodo = dblPeriodo
        CalcularImportesTesoreria.Remesa = dblRemesa
    End Function

    <Serializable()> _
    Public Class DatosGetDispenBancos
        Public IDBancoPropio As String
        Public CContable As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDBancoPropio As String, ByVal CContable As String)
            Me.IDBancoPropio = IDBancoPropio
            Me.CContable = CContable
        End Sub
    End Class

    <Task()> Public Shared Function GetDispenBancos(ByVal data As DatosGetDispenBancos, ByVal services As ServiceProvider) As DataTable
        Dim StrCadena As String
        Dim DblSaldoCtble As Double
        Dim dtSaldo As DataTable
        Dim StDatos As New DatosSaldoAFecha(data.IDBancoPropio, Today())
        dtSaldo = ProcessServer.ExecuteTask(Of DatosSaldoAFecha, DataTable)(AddressOf SaldoBancoPropioAFecha, StDatos, services)
        If Not IsNothing(dtSaldo) AndAlso dtSaldo.Rows.Count > 0 Then
            DblSaldoCtble = dtSaldo.Rows(0)("Saldo")
        Else
            DblSaldoCtble = 0
        End If

        Dim f As New Filter
        f.Add("FechaApertura", FilterOperator.LessThanOrEqual, Today(), FilterType.DateTime)
        f.Add("FechaFinalizacion", FilterOperator.GreaterThanOrEqual, Today(), FilterType.DateTime)
        f.Add("IDBancoPropio", FilterOperator.Equal, data.IDBancoPropio, FilterType.String)
        StrCadena = "SELECT SUM(LimitePoliza) AS LimiteCredito"
        StrCadena += "  FROM tbMaestroPolizaCredito"
        StrCadena += " WHERE " & AdminData.ComposeFilter(f) & ""
        StrCadena += " GROUP BY IDBancoPropio"
        Dim DblPolCred As Double = AdminData.Execute(StrCadena, ExecuteCommand.ExecuteScalar)

        Dim f1 As New Filter
        f1.Add("FechaVencimiento", FilterOperator.LessThanOrEqual, Today(), FilterType.DateTime)
        f1.Add("IDBancoPropio", FilterOperator.Equal, data.IDBancoPropio, FilterType.String)
        StrCadena = "SELECT SUM (frmPagos.ImpVencimientoA) AS CargoPendiente"
        StrCadena += " FROM frmPagos INNER JOIN tbMaestroEstadoPago ON frmPagos.Situacion = tbMaestroEstadoPago.IDEstado"
        StrCadena += " WHERE " & AdminData.ComposeFilter(f1) & ""
        StrCadena += " GROUP BY frmPagos.IDBancoPropio, tbMaestroEstadoPago.Riesgo"
        StrCadena += " HAVING(tbMaestroEstadoPago.Riesgo <> 0)"
        Dim DblCargoPend As Double = AdminData.Execute(StrCadena, ExecuteCommand.ExecuteScalar)

        Dim f2 As New Filter
        f2.Add("FechaVencimiento", FilterOperator.LessThanOrEqual, Today(), FilterType.DateTime)
        f2.Add("IDBancoPropio", FilterOperator.Equal, data.IDBancoPropio, FilterType.String)
        StrCadena = "SELECT SUM (frmCobros.ImpVencimientoA) AS CargoPendiente"
        StrCadena += " FROM frmCobros INNER JOIN tbMaestroEstadoCobro ON frmCobros.Situacion = tbMaestroEstadoCobro.IDEstado"
        StrCadena += " WHERE " & AdminData.ComposeFilter(f2) & ""
        StrCadena += " GROUP BY frmCobros.IDBancoPropio, tbMaestroEstadoCobro.Riesgo"
        StrCadena += " HAVING(tbMaestroEstadoCobro.Riesgo <> 0)"
        Dim DblAbonoPend As Double = AdminData.Execute(StrCadena, ExecuteCommand.ExecuteScalar)

        Dim DtBanco As New DataTable("tbDispBancos")
        DtBanco.Columns.Add(New DataColumn("SaldoContable", System.Type.GetType("System.Double")))
        DtBanco.Columns.Add(New DataColumn("LimiteCredito", System.Type.GetType("System.Double")))
        DtBanco.Columns.Add(New DataColumn("CargosPendientes", System.Type.GetType("System.Double")))
        DtBanco.Columns.Add(New DataColumn("AbonosPendientes", System.Type.GetType("System.Double")))
        DtBanco.Columns.Add(New DataColumn("DisponibleHoy", System.Type.GetType("System.Double")))
        Dim Drnew As DataRow = DtBanco.NewRow()
        Drnew("SaldoContable") = DblSaldoCtble
        Drnew("LimiteCredito") = DblPolCred
        Drnew("CargosPendientes") = DblCargoPend
        Drnew("AbonosPendientes") = DblAbonoPend
        Drnew("DisponibleHoy") = (DblSaldoCtble + DblPolCred + DblAbonoPend) - DblCargoPend
        DtBanco.Rows.Add(Drnew)
        Return DtBanco
    End Function

    <Task()> Public Shared Function SaldosBancosPropiosAFecha(ByVal data As Date, ByVal services As ServiceProvider) As DataTable
        Dim dtSaldosBancosPropios As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf ConstruirDTSaldosBancosPropiosAFecha, Nothing, services)
        Dim dtBancosPropios As DataTable = New BancoPropio().Filter()
        For Each drRowBanco As DataRow In dtBancosPropios.Rows
            With dtSaldosBancosPropios
                Dim drRowSaldoBanco As DataRow = .NewRow
                drRowSaldoBanco("IDBancoPropio") = drRowBanco("IDBancoPropio")
                drRowSaldoBanco("DescBancoPropio") = drRowBanco("DescBancoPropio")
                Dim dtSaldo As DataTable
                If New Parametro().Contabilidad Then
                    Dim StDatos As New DatosSaldoAFecha(drRowBanco("IDBancoPropio"), data)
                    dtSaldo = ProcessServer.ExecuteTask(Of DatosSaldoAFecha, DataTable)(AddressOf SaldoBancoPropioAFecha, StDatos, services)
                End If
                If Not IsNothing(dtSaldo) AndAlso dtSaldo.Rows.Count > 0 Then
                    drRowSaldoBanco("Fecha") = dtSaldo.Rows(0)("Fecha")
                    drRowSaldoBanco("Saldo") = dtSaldo.Rows(0)("Saldo")
                Else
                    drRowSaldoBanco("Fecha") = Today
                    drRowSaldoBanco("Saldo") = 0
                End If
                .Rows.Add(drRowSaldoBanco)
            End With
        Next

        Return dtSaldosBancosPropios
    End Function

    <Task()> Public Shared Function ConstruirDTSaldosBancosPropiosAFecha(ByVal data As Object, ByVal services As ServiceProvider) As DataTable
        Dim dtSaldos As New DataTable
        dtSaldos.Columns.Add("IDBancoPropio", GetType(String))
        dtSaldos.Columns.Add("DescBancoPropio", GetType(String))
        dtSaldos.Columns.Add("Fecha", GetType(Date))
        dtSaldos.Columns.Add("Saldo", GetType(Double))
        Return dtSaldos
    End Function

    <Serializable()> _
    Public Class DatosSaldoAFecha
        Public IDBancoPropio As String
        Public Fecha As Date

        Public Sub New(ByVal IDBancoPropio As String, ByVal Fecha As Date)
            Me.IDBancoPropio = IDBancoPropio
            Me.Fecha = Fecha
        End Sub
    End Class

    <Task()> Public Shared Function SaldoBancoPropioAFecha(ByVal data As DatosSaldoAFecha, ByVal services As ServiceProvider) As DataTable
        'Dim dblSaldoBanco As Double = 0
        Dim dtSaldo As New DataTable
        With dtSaldo
            .Columns.Clear()
            .Columns.Add("IDBancoPropio", GetType(String))
            .Columns.Add("Saldo", GetType(Double))
            .Columns.Add("Fecha", GetType(Date))
        End With

        '//Obtenemos la C.Contable del Banco indicado.
        Dim dtBanco As DataTable = New BancoPropio().SelOnPrimaryKey(data.IDBancoPropio)
        If Not IsNothing(dtBanco) AndAlso dtBanco.Rows.Count > 0 Then
            If Length(dtBanco.Rows(0)("CContable") & String.Empty) > 0 Then
                '//Recuperamos el Ejercicio predeterminado para la fecha dada.
                Dim strEjercicio As String = ProcessServer.ExecuteTask(Of Date, String)(AddressOf NegocioGeneral.EjercicioPredeterminado, data.Fecha, services)

                '//Recuperamos los decimales de la Moneda A.
                Dim intDecimalesA As Integer = ProcessServer.ExecuteTask(Of Date, MonedaInfo)(AddressOf Moneda.MonedaA, cnMinDate, services).NDecimalesImporte

                '//Recuperamos los Saldos para la C.Contable del Banco dado.
                Dim dtSaldoCuenta As DataTable = NegocioGeneral.ExtractoCuenta(strEjercicio, dtBanco.Rows(0)("CContable") & String.Empty)

                If Not IsNothing(dtSaldoCuenta) AndAlso dtSaldoCuenta.Rows.Count > 0 Then
                    Dim drRowSaldo As DataRow = dtSaldo.NewRow
                    drRowSaldo("IDBancoPropio") = data.IDBancoPropio

                    '//Nos quedamos con el Saldo correspondiente a la fecha dada, pero indicando cual es la Fecha del último apunte.
                    Dim fFilterFecha As New Filter
                    fFilterFecha.Add(New DateFilterItem("FechaApunte", FilterOperator.LessThanOrEqual, data.Fecha))
                    Dim WhereFechaApunte As String = fFilterFecha.Compose(New AdoFilterComposer)
                    Dim adrSaldoFecha() As DataRow = dtSaldoCuenta.Select(WhereFechaApunte, "FechaApunte DESC, NAsiento DESC, IDApunte DESC")
                    If Not IsNothing(adrSaldoFecha) AndAlso adrSaldoFecha.Length > 0 Then
                        drRowSaldo("Fecha") = adrSaldoFecha(0)("FechaApunte")
                        drRowSaldo("Saldo") = xRound(adrSaldoFecha(0)("SaldoA"), intDecimalesA)
                    Else
                        drRowSaldo("Fecha") = Today
                        drRowSaldo("Saldo") = xRound(0, intDecimalesA)
                    End If
                    dtSaldo.Rows.Add(drRowSaldo)

                    adrSaldoFecha = Nothing
                    dtSaldoCuenta.Rows.Clear()
                End If
            End If
            dtBanco.Rows.Clear()
        End If

        Return dtSaldo
    End Function

    <Serializable()> _
    Public Class BancoPropioFactoringInfo
        Public IDBanco As String
        Public Factoring As Boolean
        Public TipoFactoring As String
        Public IDClienteFactoring As String
        Public IDContadorFactoring As String
    End Class

    <Task()> Public Shared Function BancoFactoring(ByVal data As String, ByVal services As ServiceProvider) As Boolean
        BancoFactoring = True
        Dim dtBP As DataTable = New BancoPropio().SelOnPrimaryKey(data)
        If Not IsNothing(dtBP) AndAlso dtBP.Rows.Count > 0 Then
            If Not CBool(dtBP.Rows(0)("Factoring")) Then
                BancoFactoring = False
                dtBP.Rows.Clear()
            End If
        End If
    End Function

    <Task()> Public Shared Function BancoFactoringInfo(ByVal data As String) As BancoPropioFactoringInfo
        Dim objBPInfoFactoring As BancoPropioFactoringInfo
        Dim dtBP As DataTable = New BancoPropio().SelOnPrimaryKey(data)
        If Not IsNothing(dtBP) AndAlso dtBP.Rows.Count > 0 Then
            objBPInfoFactoring = New BancoPropioFactoringInfo
            objBPInfoFactoring.IDBanco = data
            objBPInfoFactoring.Factoring = dtBP.Rows(0)("Factoring")
            objBPInfoFactoring.TipoFactoring = dtBP.Rows(0)("TipoFactoring") & String.Empty
            objBPInfoFactoring.IDClienteFactoring = dtBP.Rows(0)("IDClienteFactoring") & String.Empty
            objBPInfoFactoring.IDContadorFactoring = dtBP.Rows(0)("IDContadorFactoring") & String.Empty
        End If
        Return objBPInfoFactoring
    End Function

#End Region

End Class