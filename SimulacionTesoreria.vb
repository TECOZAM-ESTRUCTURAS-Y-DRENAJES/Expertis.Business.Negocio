<Transactional()> _
Public Class SimulacionTesoreria
    Inherits ContextBoundObject

#Region " Proceso de Simulación "

    <Serializable()> _
    Public Class DataSimulTes
        Public FechaDesde As Date
        Public FechaHasta As Date
        Public FechaValidez As Date
        Public BlnBancoPropio As Boolean
        Public BlnFactura As Boolean
        Public BlnAlbaran As Boolean
        Public BlnPedido As Boolean
        Public BlnObra As Boolean
        Public BlnPromotoras As Boolean

        Public Sub New()
        End Sub
        Public Sub New(ByVal FechaDesde As Date, ByVal FechaHasta As Date, ByVal FechaValidez As Date, ByVal BlnBancoPropio As Boolean, _
                       ByVal BlnFactura As Boolean, ByVal BlnAlbaran As Boolean, ByVal BlnPedido As Boolean, ByVal BlnObra As Boolean, _
                       ByVal BlnPromotoras As Boolean)
            Me.FechaDesde = FechaDesde
            Me.FechaHasta = FechaHasta
            Me.FechaValidez = FechaValidez
            Me.BlnBancoPropio = BlnBancoPropio
            Me.BlnFactura = BlnFactura
            Me.BlnAlbaran = BlnAlbaran
            Me.BlnPedido = BlnPedido
            Me.BlnObra = BlnObra
            Me.BlnPromotoras = BlnPromotoras
        End Sub
    End Class

    <Task()> Public Shared Function SimulacionTesoreria(ByVal data As DataSimulTes, ByVal services As ServiceProvider) As DataTable
        '////////////////////////////PENDIENTE////////////////////////////
        'Conexion = New AdminData
        ''REVISAR(11/10/05)
        ''BDActual = Conexion.GetSessionInfo.Database
        'BDActual = GetPropertyValue(BDActual, "DataBase")
        'clsAdmin = New AdminData
        '////////////////////////////FIN PENDIENTE////////////////////////////


        '//Se guardan los Parámetros de la Simulación, para que al guardar la simulación en un fichero se pueda
        '//ver cuales han sido esos parámetros.
        Dim strParametrosSimulacion As String = "FechaDesde=" & data.FechaDesde & ";" & _
                                                "FechaHasta=" & data.FechaHasta & ";" & _
                                                "FechaValidez=" & data.FechaValidez & ";" & _
                                                "Facturas=" & data.BlnFactura & ";" & _
                                                "Albaranes=" & data.BlnAlbaran & ";" & _
                                                "Pedidos=" & data.BlnPedido & ";" & _
                                                "BancoPropios=" & data.BlnBancoPropio & ";" & _
                                                "Promotoras=" & data.BlnPromotoras & ";" & _
                                                "Obras=" & data.BlnObra & ";"

        Dim objFilter As New Filter
        objFilter.Clear()

        '//Contruimos la estructura a devolver por la simulación.
        Dim dtSimulacionTesoreria As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf ConstruirDTSimTesoreria, Nothing, services)

        '//Recuperamos la información necesaria de la Moneda A
        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
        Dim MonInfoA As MonedaInfo = Monedas.MonedaA

        '//Obtenemos el Ejecicio Predeterminado
        Dim strEjercicio As String = ProcessServer.ExecuteTask(Of Date, String)(AddressOf NegocioGeneral.EjercicioPredeterminado, Today, services)

        '//Obtenemos algunos parámetros de la aplicación.
        Dim objNegParametro As New Parametro
        Dim blnCalcularImpRE As Boolean = objNegParametro.EmpresaRecargoEquivPredeterminado
        Dim strAgrupacionGasto As String = objNegParametro.ParametroGasto
        Dim strAgrupacionIngreso As String = objNegParametro.ParametroIngreso
        Dim strBancoPropio As String = objNegParametro.ParametroBancoPropio
        Dim strCondPagoPred As String = objNegParametro.CondicionPago
        Dim strDiaPagoPred As String = objNegParametro.DiaPago
        Dim strIdProvRetencion As String = objNegParametro.ProveedorRetencion

        Dim strDescProvRetencion As String
        Dim objNegProveedor As New Proveedor
        objFilter.Clear()
        objFilter.Add(New StringFilterItem("IDProveedor", strIdProvRetencion))
        Dim dtProveedor As DataTable = objNegProveedor.Filter(objFilter)
        If Not IsNothing(dtProveedor) AndAlso dtProveedor.Rows.Count > 0 Then
            strDescProvRetencion = IIf(Length(dtProveedor.Rows(0)("DescProveedor")) > 0, dtProveedor.Rows(0)("DescProveedor") & String.Empty, String.Empty)
            dtProveedor.Rows.Clear()
        End If

        '// 1.- Se incorporan los COBROS a la Simulación.
        Dim StSimulTes As New DataSimulTesALL(dtSimulacionTesoreria, strParametrosSimulacion, enumSimulacionTesoreria.Cobro, MonInfoA.NDecimalesImporte, data.FechaDesde, data.FechaHasta, , strBancoPropio, strAgrupacionIngreso)
        StSimulTes.DtSimulacionTesoreria = ProcessServer.ExecuteTask(Of DataSimulTesALL, DataTable)(AddressOf AddingRecordSimulacionTesoreria, StSimulTes, services)

        '// 2.- Se incorporan los COBROS PERIODICOS a la Simulación.
        StSimulTes.TipoDocumento = enumSimulacionTesoreria.CobroPeriodico
        StSimulTes.DtSimulacionTesoreria = ProcessServer.ExecuteTask(Of DataSimulTesALL, DataTable)(AddressOf AddingRecordSimulacionTesoreria, StSimulTes, services)

        '// 3.- Se incorporan los PAGOS a la Simulación.
        StSimulTes.TipoDocumento = enumSimulacionTesoreria.Pago
        StSimulTes.AgrupacionGasto = strAgrupacionGasto
        StSimulTes.AgrupacionIngreso = String.Empty
        StSimulTes.DtSimulacionTesoreria = ProcessServer.ExecuteTask(Of DataSimulTesALL, DataTable)(AddressOf AddingRecordSimulacionTesoreria, StSimulTes, services)

        '// 4.- Se incorporan los PAGOS PERIODICOS a la Simulación.
        StSimulTes.TipoDocumento = enumSimulacionTesoreria.PagoPeriodico
        StSimulTes.DtSimulacionTesoreria = ProcessServer.ExecuteTask(Of DataSimulTesALL, DataTable)(AddressOf AddingRecordSimulacionTesoreria, StSimulTes, services)

        '// 5.- Se incorporan los COBROS/PAGOS de las FACTURAS a la Simulación.
        If data.BlnFactura Then
            '// 5. A.- Se incorporan los COBROS de las FACTURAS de VENTA a la Simulación.
            StSimulTes.TipoDocumento = enumSimulacionTesoreria.FVNoContabilizada
            StSimulTes.AgrupacionGasto = String.Empty
            StSimulTes.AgrupacionIngreso = strAgrupacionIngreso
            StSimulTes.DtSimulacionTesoreria = ProcessServer.ExecuteTask(Of DataSimulTesALL, DataTable)(AddressOf AddingRecordSimulacionTesoreria, StSimulTes, services)
            
            '// 5. B.- Se incorporan los PAGOS de las FACTURAS de COMPRA a la Simulación.
            StSimulTes.TipoDocumento = enumSimulacionTesoreria.FCNoContabilizada
            StSimulTes.AgrupacionGasto = strAgrupacionGasto
            StSimulTes.AgrupacionIngreso = String.Empty
            StSimulTes.DtSimulacionTesoreria = ProcessServer.ExecuteTask(Of DataSimulTesALL, DataTable)(AddressOf AddingRecordSimulacionTesoreria, StSimulTes, services)
        End If

        '// 6.- Se incorporan los COBROS/PAGOS de los ALBARANES a la Simulación.
        If data.BlnAlbaran Then
            '// 6. A.- Se incorporan los COBROS de los ALBARANES de VENTA a la Simulación.
            StSimulTes.TipoDocumento = enumSimulacionTesoreria.AVNoFacturado
            StSimulTes.FechaValidez = data.FechaValidez
            StSimulTes.AgrupacionGasto = String.Empty
            StSimulTes.AgrupacionIngreso = strAgrupacionIngreso
            StSimulTes.CondpagoPred = strCondPagoPred
            StSimulTes.DiaPagoPred = strDiaPagoPred
            StSimulTes.DtSimulacionTesoreria = ProcessServer.ExecuteTask(Of DataSimulTesALL, DataTable)(AddressOf AddingRecordSimulacionTesoreria, StSimulTes, services)

            '// 6. B.- Se incorporan los PAGOS de los ALBARANES de COMPRA a la Simulación.
            StSimulTes.TipoDocumento = enumSimulacionTesoreria.ACNoFacturado
            StSimulTes.FechaValidez = data.FechaValidez
            StSimulTes.AgrupacionGasto = strAgrupacionGasto
            StSimulTes.AgrupacionIngreso = String.Empty
            StSimulTes.CondpagoPred = strCondPagoPred
            StSimulTes.DiaPagoPred = strDiaPagoPred
            StSimulTes.IDProvRetencion = strIdProvRetencion
            StSimulTes.DescProvRetencion = strDescProvRetencion
            StSimulTes.CalcularImpRE = blnCalcularImpRE
            StSimulTes.DtSimulacionTesoreria = ProcessServer.ExecuteTask(Of DataSimulTesALL, DataTable)(AddressOf AddingRecordSimulacionTesoreria, StSimulTes, services)
        End If

        '// 7.- Se incorporan los COBROS/PAGOS de los PEDIDOS a la Simulación.
        If data.BlnPedido Then
            '// 7. A.- Se incorporan los COBROS de los PEDIDOS de VENTA a la Simulación.
            StSimulTes.TipoDocumento = enumSimulacionTesoreria.PVPendiente
            StSimulTes.FechaValidez = data.FechaValidez
            StSimulTes.AgrupacionGasto = String.Empty
            StSimulTes.AgrupacionIngreso = strAgrupacionIngreso
            StSimulTes.CondpagoPred = strCondPagoPred
            StSimulTes.DiaPagoPred = strDiaPagoPred
            StSimulTes.IDProvRetencion = String.Empty
            StSimulTes.DescProvRetencion = String.Empty
            StSimulTes.CalcularImpRE = False
            StSimulTes.DtSimulacionTesoreria = ProcessServer.ExecuteTask(Of DataSimulTesALL, DataTable)(AddressOf AddingRecordSimulacionTesoreria, StSimulTes, services)

            '// 7. B.- Se incorporan los PAGOS de los PEDIDOS de COMPRA a la Simulación.
            StSimulTes.TipoDocumento = enumSimulacionTesoreria.PCPendiente
            StSimulTes.FechaValidez = data.FechaValidez
            StSimulTes.AgrupacionGasto = strAgrupacionGasto
            StSimulTes.AgrupacionIngreso = String.Empty
            StSimulTes.CondpagoPred = strCondPagoPred
            StSimulTes.DiaPagoPred = strDiaPagoPred
            StSimulTes.IDProvRetencion = strIdProvRetencion
            StSimulTes.DescProvRetencion = strDescProvRetencion
            StSimulTes.CalcularImpRE = blnCalcularImpRE
            StSimulTes.DtSimulacionTesoreria = ProcessServer.ExecuteTask(Of DataSimulTesALL, DataTable)(AddressOf AddingRecordSimulacionTesoreria, StSimulTes, services)
        End If

        '// 8.- Se incorporan los Saldos de los Bancos Propios a la simulación
        If data.BlnBancoPropio Then
            StSimulTes.TipoDocumento = enumSimulacionTesoreria.BancoPropio
            StSimulTes.FechaDesde = cnMinDate
            StSimulTes.FechaHasta = cnMinDate
            StSimulTes.FechaValidez = cnMinDate
            StSimulTes.AgrupacionGasto = String.Empty
            StSimulTes.AgrupacionIngreso = String.Empty
            StSimulTes.CondpagoPred = String.Empty
            StSimulTes.DiaPagoPred = String.Empty
            StSimulTes.IDProvRetencion = String.Empty
            StSimulTes.DescProvRetencion = String.Empty
            StSimulTes.CalcularImpRE = False
            StSimulTes.DtSimulacionTesoreria = ProcessServer.ExecuteTask(Of DataSimulTesALL, DataTable)(AddressOf AddingRecordSimulacionTesoreria, StSimulTes, services)
        End If

        If data.BlnPromotoras Then
            StSimulTes.TipoDocumento = enumSimulacionTesoreria.Promotoras
            StSimulTes.FechaDesde = data.FechaDesde
            StSimulTes.FechaHasta = data.FechaHasta
            StSimulTes.FechaValidez = cnMinDate
            StSimulTes.AgrupacionGasto = String.Empty
            StSimulTes.AgrupacionIngreso = String.Empty
            StSimulTes.CondpagoPred = String.Empty
            StSimulTes.DiaPagoPred = String.Empty
            StSimulTes.IDProvRetencion = String.Empty
            StSimulTes.DescProvRetencion = String.Empty
            StSimulTes.CalcularImpRE = False
            StSimulTes.DtSimulacionTesoreria = ProcessServer.ExecuteTask(Of DataSimulTesALL, DataTable)(AddressOf AddingRecordSimulacionTesoreria, StSimulTes, services)
        End If

        '// 9.- Se incorporan los Cobros/Pagos de las Obras de la Simulación
        If data.BlnObra Then
            StSimulTes.TipoDocumento = enumSimulacionTesoreria.OVNoFacturada
            StSimulTes.FechaDesde = data.FechaDesde
            StSimulTes.FechaHasta = data.FechaHasta
            StSimulTes.FechaValidez = data.FechaValidez
            StSimulTes.AgrupacionGasto = String.Empty
            StSimulTes.AgrupacionIngreso = strAgrupacionIngreso
            StSimulTes.CondpagoPred = strCondPagoPred
            StSimulTes.DiaPagoPred = strDiaPagoPred
            StSimulTes.IDProvRetencion = String.Empty
            StSimulTes.DescProvRetencion = String.Empty
            StSimulTes.CalcularImpRE = False
            '// 6. A.- Se incorporan los COBROS de los ALBARANES de VENTA a la Simulación.
            StSimulTes.DtSimulacionTesoreria = ProcessServer.ExecuteTask(Of DataSimulTesALL, DataTable)(AddressOf AddingRecordSimulacionTesoreria, StSimulTes, services)
        End If

        '//Calculamos la columna Acumulado
        Dim dblAcumulado As Double = 0
        For Each drRowSimulacion As DataRow In dtSimulacionTesoreria.Select(Nothing, "Fecha")
            drRowSimulacion("Acumulado") = xRound(drRowSimulacion("Importe"), MonInfoA.NDecimalesImporte) + xRound(dblAcumulado, MonInfoA.NDecimalesImporte)
            dblAcumulado = xRound(drRowSimulacion("Acumulado"), MonInfoA.NDecimalesImporte)
        Next drRowSimulacion

        ''//Construimos el cubo de OLAP.
        'CrearSimulacionOLAP(dtSimulacionTesoreria)

        Return dtSimulacionTesoreria
    End Function

    <Task()> Public Shared Function ConstruirDTSimTesoreria(ByVal data As Object, ByVal services As ServiceProvider) As DataTable
        Dim dtSimulacion As New DataTable
        dtSimulacion.RemotingFormat = SerializationFormat.Binary
        dtSimulacion.Columns.Add("ParametrosSimulacion", GetType(String))
        dtSimulacion.Columns.Add("EnumTipoDocumento", GetType(Integer))
        dtSimulacion.Columns.Add("IdOrigen", GetType(Integer))
        dtSimulacion.Columns.Add("IdClienteProveedor", GetType(String))
        dtSimulacion.Columns.Add("DescClienteProveedor", GetType(String))
        dtSimulacion.Columns.Add("Importe", GetType(Double))
        dtSimulacion.Columns.Add("Fecha", GetType(Date))
        dtSimulacion.Columns.Add("IdBancoPropio", GetType(String))
        dtSimulacion.Columns.Add("IdTipoCobroPago", GetType(Integer))
        dtSimulacion.Columns.Add("Acumulado", GetType(Double))
        dtSimulacion.Columns.Add("IDAgrupacion", GetType(String))
        'dtSimulacion.Columns.Add("Empresa", GetType(String))
        dtSimulacion.Columns.Add("Modificado", GetType(Boolean))
        dtSimulacion.Columns.Add("Situacion", GetType(String))
        dtSimulacion.Columns.Add("FormaPago", GetType(String))
        dtSimulacion.Columns.Add("DescSituacion", GetType(String))
        Return dtSimulacion
    End Function

    <Serializable()> _
    Public Class DataSimulTesALL
        Public DtSimulacionTesoreria As DataTable
        Public ParametrosSimulacion As String
        Public TipoDocumento As enumSimulacionTesoreria
        Public DecimalesImpA As Integer
        Public FechaDesde As Date
        Public FechaHasta As Date
        Public FechaValidez As Date
        Public BancoPropio As String
        Public AgrupacionIngreso As String
        Public AgrupacionGasto As String
        Public CondpagoPred As String
        Public DiaPagoPred As String
        Public IDProvRetencion As String
        Public DescProvRetencion As String
        Public CalcularImpRE As Boolean

        Public Sub New()
        End Sub
        Public Sub New(ByVal DtSimulacionTesoreria As DataTable, ByVal ParametrosSimulacion As String, ByVal TipoDocumento As enumSimulacionTesoreria, ByVal DecimalesImpA As Integer, _
                       Optional ByVal FechaDesde As Date = cnMinDate, Optional ByVal FechaHasta As Date = cnMinDate, Optional ByVal FechaValidez As Date = cnMinDate, _
                       Optional ByVal BancoPropio As String = "", Optional ByVal AgrupacionIngreso As String = "", Optional ByVal AgrupacionGasto As String = "", Optional ByVal CondPagoPred As String = "", _
                       Optional ByVal DiaPagoPred As String = "", Optional ByVal IDProvRetencion As String = "", Optional ByVal DescProvRetencion As String = "", Optional ByVal CalcularImpRE As Boolean = False)
            Me.DtSimulacionTesoreria = DtSimulacionTesoreria
            Me.ParametrosSimulacion = ParametrosSimulacion
            Me.TipoDocumento = TipoDocumento
            Me.DecimalesImpA = DecimalesImpA
            Me.FechaDesde = FechaDesde
            Me.FechaHasta = FechaHasta
            Me.FechaValidez = FechaValidez
            Me.BancoPropio = BancoPropio
            Me.AgrupacionIngreso = AgrupacionIngreso
            Me.AgrupacionGasto = AgrupacionGasto
            Me.CondpagoPred = CondPagoPred
            Me.DiaPagoPred = DiaPagoPred
            Me.IDProvRetencion = IDProvRetencion
            Me.DescProvRetencion = DescProvRetencion
            Me.CalcularImpRE = CalcularImpRE
        End Sub
    End Class

    <Task()> Public Shared Function AddingRecordSimulacionTesoreria(ByVal data As DataSimulTesALL, ByVal services As ServiceProvider) As DataTable
        Select Case data.TipoDocumento
            Case enumSimulacionTesoreria.Cobro
                data.DtSimulacionTesoreria = ProcessServer.ExecuteTask(Of DataSimulTesALL, DataTable)(AddressOf SimulacionTesoreriaCobros, data, services)
            Case enumSimulacionTesoreria.CobroPeriodico
                data.DtSimulacionTesoreria = ProcessServer.ExecuteTask(Of DataSimulTesALL, DataTable)(AddressOf SimulacionTesoreriaCobrosPeriodicos, data, services)
            Case enumSimulacionTesoreria.Pago
                data.DtSimulacionTesoreria = ProcessServer.ExecuteTask(Of DataSimulTesALL, DataTable)(AddressOf SimulacionTesoreriaPagos, data, services)
            Case enumSimulacionTesoreria.PagoPeriodico
                data.DtSimulacionTesoreria = ProcessServer.ExecuteTask(Of DataSimulTesALL, DataTable)(AddressOf SimulacionTesoreriaPagosPeriodicos, data, services)
            Case enumSimulacionTesoreria.FCNoContabilizada
                data.DtSimulacionTesoreria = ProcessServer.ExecuteTask(Of DataSimulTesALL, DataTable)(AddressOf SimulacionTesoreriaFacturasCompras, data, services)
            Case enumSimulacionTesoreria.FVNoContabilizada
                data.DtSimulacionTesoreria = ProcessServer.ExecuteTask(Of DataSimulTesALL, DataTable)(AddressOf SimulacionTesoreriaFacturasVentas, data, services)
            Case enumSimulacionTesoreria.AVNoFacturado
                data.DtSimulacionTesoreria = ProcessServer.ExecuteTask(Of DataSimulTesALL, DataTable)(AddressOf SimulacionTesoreriaAlbaranesVentas, data, services)
            Case enumSimulacionTesoreria.ACNoFacturado
                data.DtSimulacionTesoreria = ProcessServer.ExecuteTask(Of DataSimulTesALL, DataTable)(AddressOf SimulacionTesoreriaAlbaranesCompras, data, services)
            Case enumSimulacionTesoreria.PVPendiente
                data.DtSimulacionTesoreria = ProcessServer.ExecuteTask(Of DataSimulTesALL, DataTable)(AddressOf SimulacionTesoreriaPedidosVentas, data, services)
            Case enumSimulacionTesoreria.PCPendiente
                data.DtSimulacionTesoreria = ProcessServer.ExecuteTask(Of DataSimulTesALL, DataTable)(AddressOf SimulacionTesoreriaPedidosCompras, data, services)
            Case enumSimulacionTesoreria.BancoPropio
                data.DtSimulacionTesoreria = ProcessServer.ExecuteTask(Of DataSimulTesALL, DataTable)(AddressOf SimulacionTesoreriaBancosPropios, data, services)
            Case enumSimulacionTesoreria.Promotoras
                data.DtSimulacionTesoreria = ProcessServer.ExecuteTask(Of DataSimulTesALL, DataTable)(AddressOf SimulacionTesoreriaPromotoras, data, services)
            Case enumSimulacionTesoreria.OVNoFacturada
                data.DtSimulacionTesoreria = ProcessServer.ExecuteTask(Of DataSimulTesALL, DataTable)(AddressOf SimulacionTesoreriaObrasHitos, data, services)
            Case enumSimulacionTesoreria.OCNoFacturada
                data.DtSimulacionTesoreria = ProcessServer.ExecuteTask(Of DataSimulTesALL, DataTable)(AddressOf SimulacionTesoreriaObrasCompras, data, services)
        End Select
        Return data.DtSimulacionTesoreria
    End Function

    <Task()> Public Shared Function SimulacionTesoreriaCobros(ByVal data As DataSimulTesALL, ByVal services As ServiceProvider) As DataTable
        Dim objFilter As New Filter
        objFilter.Clear()
        If data.FechaDesde <> cnMinDate AndAlso data.FechaHasta <> cnMinDate Then
            objFilter.Add(New BetweenFilterItem("FechaVencimiento", data.FechaDesde, data.FechaHasta, FilterType.DateTime))
        End If

        Dim dtCobros As DataTable = New BE.DataEngine().Filter("vNegSimulacionTesoreriaC", objFilter)

        For Each drRowCobro As DataRow In dtCobros.Select(Nothing, "FechaVencimiento")
            Dim dblDiasRetraso As Double = 0
            If Length(drRowCobro("IdCliente") & String.Empty) > 0 Then
                dblDiasRetraso = ProcessServer.ExecuteTask(Of String, Double)(AddressOf RecuperarDiasRetraso, drRowCobro("IdCliente") & String.Empty, services)
            End If

            With data.DtSimulacionTesoreria
                Dim drRowSim As DataRow = .NewRow
                drRowSim("ParametrosSimulacion") = data.ParametrosSimulacion
                drRowSim("EnumTipoDocumento") = enumSimulacionTesoreria.Cobro
                drRowSim("IdOrigen") = drRowCobro("IdCobro")
                drRowSim("IdClienteProveedor") = drRowCobro("IdCliente")
                drRowSim("DescClienteProveedor") = drRowCobro("Titulo")
                drRowSim("Importe") = drRowCobro("ImpVencimientoA")
                drRowSim("Fecha") = DateAdd(DateInterval.Day, dblDiasRetraso, drRowCobro("FechaVencimiento"))
                drRowSim("IdBancoPropio") = Nz(drRowCobro("IdBancoPropio"), data.BancoPropio)
                drRowSim("IdTipoCobroPago") = drRowCobro("IdTipoCobro")
                drRowSim("IDAgrupacion") = Nz(drRowCobro("IDAgrupacion"), data.AgrupacionIngreso)
                'drRowSim("Empresa") = rsCobro.Fields("SyncDB").Value & String.Empty
                drRowSim("Modificado") = False
                drRowSim("FormaPago") = drRowCobro("IDFormaPago")
                drRowSim("Situacion") = drRowCobro("Situacion")
                .Rows.Add(drRowSim)
            End With
        Next drRowCobro

        objFilter = Nothing
        Return data.DtSimulacionTesoreria
    End Function

    <Task()> Public Shared Function SimulacionTesoreriaPromotoras(ByVal data As DataSimulTesALL, ByVal services As ServiceProvider) As DataTable
        Dim objFilter As New Filter
        objFilter.Clear()
        If data.FechaDesde <> cnMinDate AndAlso data.FechaHasta <> cnMinDate Then
            objFilter.Add(New BetweenFilterItem("FechaVencimiento", data.FechaDesde, data.FechaHasta, FilterType.DateTime))
        End If

        Dim dtCobros As DataTable = New BE.DataEngine().Filter("vNegSimulacionTesoreriaPromo", objFilter)

        For Each drRowCobro As DataRow In dtCobros.Select(Nothing, "FechaVencimiento")
            Dim dblDiasRetraso As Double = 0
            If Length(drRowCobro("IdCliente") & String.Empty) > 0 Then
                dblDiasRetraso = ProcessServer.ExecuteTask(Of String, Double)(AddressOf RecuperarDiasRetraso, drRowCobro("IdCliente") & String.Empty, services)
            End If

            With data.DtSimulacionTesoreria
                Dim drRowSim As DataRow = .NewRow
                drRowSim("ParametrosSimulacion") = data.ParametrosSimulacion
                drRowSim("EnumTipoDocumento") = enumSimulacionTesoreria.Promotoras
                drRowSim("IdOrigen") = drRowCobro("IdLocalVencimiento")
                drRowSim("IdClienteProveedor") = drRowCobro("IdCliente")
                drRowSim("DescClienteProveedor") = drRowCobro("DescVencimiento")
                drRowSim("Importe") = drRowCobro("ImpTotalA")
                drRowSim("Fecha") = DateAdd(DateInterval.Day, dblDiasRetraso, drRowCobro("FechaVencimiento"))
                drRowSim("IdBancoPropio") = drRowCobro("IdBancoPropio")
                drRowSim("IdTipoCobroPago") = System.DBNull.Value ' poner el enumerado
                drRowSim("IDAgrupacion") = data.AgrupacionIngreso
                drRowSim("Modificado") = False
                drRowSim("FormaPago") = drRowCobro("IDFormaPago")
                .Rows.Add(drRowSim)
            End With
        Next drRowCobro

        Return data.DtSimulacionTesoreria
    End Function

    <Task()> Public Shared Function SimulacionTesoreriaObrasVentas(ByVal data As DataSimulTesALL, ByVal services As ServiceProvider) As DataTable
        Dim FactorIva, FactorRE, TotalTramo, TotalDia, TotalMes As Double
        Dim Meses, DiasTramo As Integer
        Dim i, pi, mi, mes, dias As Integer
        Dim fecha, fecha2 As Date
        Dim FilDatos As New Filter
        If data.FechaValidez <> cnMinDate AndAlso data.FechaHasta <> cnMinDate Then
            FilDatos.Add("Desde", FilterOperator.GreaterThanOrEqual, data.FechaValidez, FilterType.DateTime)
        End If
        Dim DtAV As DataTable = New BE.DataEngine().Filter("vNegSimulacionTesoreriaOV", FilDatos)
        FilDatos.Clear()
        If data.FechaDesde <> cnMinDate AndAlso data.FechaHasta <> cnMinDate Then
            FilDatos.Add("FechaVencimiento", FilterOperator.GreaterThanOrEqual, data.FechaDesde, FilterType.DateTime)
            FilDatos.Add("FechaVencimiento", FilterOperator.LessThanOrEqual, data.FechaHasta, FilterType.DateTime)
        End If
        For Each Dr As DataRow In DtAV.Select
            Dim FilObra As New Filter
            FilObra.Add("IDObra", FilterOperator.Equal, Dr("IDObra"), FilterType.Numeric)
            FilObra.Add(New StringFilterItem("IDTipoTrabajo", Dr("IDTipoTrabajo")))
            Dim ClsBE As New BE.DataEngine
            Dim DtTrabajo As DataTable = ClsBE.Filter("vFrmObraTrabajoTipo", FilObra)
            For Each DrTrab As DataRow In DtTrabajo.Select
                DiasTramo = DateDiff(DateInterval.Day, Dr("Desde"), Dr("Hasta"))
                If DiasTramo > 0 Then
                    TotalTramo = DrTrab("Venta") * Dr("Porcentaje") / 100
                    TotalDia = TotalTramo / DiasTramo
                End If
                'Meses por tramo
                fecha = Dr("Desde")
                mes = CDate(Dr("Desde")).Month
                dias = 0
                While fecha <= Dr("Hasta")
                    If mes <> fecha.Month Then
                        TotalMes = dias * TotalDia
                        ' Aqui tenemos el importe y calculamos los vencimientos                        
                        If Length(Dr("FactorIVA")) = 0 Then
                            FactorIva = IIf(Length(Dr("CliFactorIVA")), 0, Dr("CliFactorIVA"))
                        Else
                            FactorIva = Dr("FactorIVA")
                        End If
                        If Length(Dr("FactorRE")) = 0 Then
                            FactorRE = IIf(Length(Dr("CliFactorRE")), 0, Dr("CliFactorRE"))
                        Else
                            FactorRE = Dr("FactorRE")
                        End If
                        Dim StData As New DataVto(Dr("IDCliente"), String.Empty, TotalMes, DateSerial(Year(fecha), Month(fecha) - 1, 0), _
                                                  Nz(Dr("IDCondicionPago"), data.CondpagoPred), Nz(Dr("IDDiaPago"), data.DiaPagoPred), _
                                                  FactorIva, IIf(Dr("TieneRE"), FactorRE, 0), Dr("DtoComercial"), Nz(Dr("FactorDPP"), 0), Nz(Dr("FactorRecFinan"), 0), data.DecimalesImpA)
                        Dim DtVtoOV As DataTable = ProcessServer.ExecuteTask(Of DataVto, DataTable)(AddressOf VtoObra, StData, services)
                        With data.DtSimulacionTesoreria
                            '//Recorremos los Vtos. del las Obras q se encuentran en el intervalo indicado.
                            Dim WherePeriodo As String = FilDatos.Compose(New AdoFilterComposer)
                            For Each drRowVtoOV As DataRow In DtVtoOV.Select(WherePeriodo)
                                Dim dblDiasRetraso As Double = 0
                                If Length(Dr("IdCliente") & String.Empty) > 0 Then
                                    dblDiasRetraso = ProcessServer.ExecuteTask(Of String, Double)(AddressOf RecuperarDiasRetraso, Dr("IdCliente") & String.Empty, services)
                                End If

                                Dim drRowSim As DataRow = .NewRow
                                drRowSim("ParametrosSimulacion") = data.ParametrosSimulacion
                                drRowSim("EnumTipoDocumento") = enumSimulacionTesoreria.OVNoFacturada
                                drRowSim("IdOrigen") = Dr("IDObra")
                                drRowSim("IdClienteProveedor") = Dr("IDCliente")
                                drRowSim("DescClienteProveedor") = Dr("DescCliente")
                                drRowSim("Importe") = Math.Round(drRowVtoOV("ImpVencimientoA"), 2)
                                drRowSim("Fecha") = DateAdd(DateInterval.Day, dblDiasRetraso, drRowVtoOV("FechaVencimiento"))
                                drRowSim("IdBancoPropio") = Nz(Dr("IdBancoPropio"), data.BancoPropio)
                                drRowSim("IdTipoCobroPago") = System.DBNull.Value ' poner el enumerado
                                drRowSim("IDAgrupacion") = data.AgrupacionIngreso
                                'drRowSim("Empresa") = drRowOV("SyncDB") & String.Empty
                                drRowSim("Modificado") = False
                                .Rows.Add(drRowSim)
                            Next drRowVtoOV
                        End With
                        dias = 0
                        mes = fecha.Month
                    End If
                    fecha = fecha.AddDays(1)
                    dias += 1
                End While
                If dias > 1 Then
                    TotalMes = dias * TotalDia
                    If Length(Dr("FactorIVA")) = 0 Then
                        If Length(Dr("CliFactorIVA")) Then
                            FactorIva = 0
                        Else
                            FactorIva = Dr("CliFactorIVA")
                        End If
                    Else
                        FactorIva = Dr("FactorIVA")
                    End If

                    If Length(Dr("FactorRE")) = 0 Then
                        If Length(Dr("CliFactorRE")) Then
                            FactorRE = 0
                        Else
                            FactorRE = Dr("CliFactorRE")
                        End If
                    Else
                        FactorRE = Dr("FactorRE")
                    End If
                    Dim StData As New DataVto(Dr("IDCliente"), String.Empty, TotalMes, DateSerial(Year(fecha), Month(fecha) + 1, 0), _
                                              Nz(Dr("IDCondicionPago"), data.CondpagoPred), Nz(Dr("IDDiaPago"), data.DiaPagoPred), _
                                              FactorIva, IIf(Dr("TieneRE"), FactorRE, 0), Dr("DtoComercial"), Nz(Dr("FactorDPP"), 0), _
                                              Nz(Dr("FactorRecFinan"), 0), data.DecimalesImpA)
                    Dim dtVtoOV As DataTable = ProcessServer.ExecuteTask(Of DataVto, DataTable)(AddressOf VtoObra, StData, services)
                    With data.DtSimulacionTesoreria
                        '//Recorremos los Vtos. del las Obras q se encuentran en el intervalo indicado.
                        Dim WherePeriodo As String = FilDatos.Compose(New AdoFilterComposer)
                        For Each DrVto As DataRow In dtVtoOV.Select(WherePeriodo)
                            Dim dblDiasRetraso As Double = 0
                            If Length(Dr("IdCliente") & String.Empty) > 0 Then
                                dblDiasRetraso = ProcessServer.ExecuteTask(Of String, Double)(AddressOf RecuperarDiasRetraso, Dr("IdCliente") & String.Empty, services)
                            End If
                            Dim drRowSim As DataRow = .NewRow
                            drRowSim("ParametrosSimulacion") = data.ParametrosSimulacion
                            drRowSim("EnumTipoDocumento") = enumSimulacionTesoreria.OVNoFacturada
                            drRowSim("IdOrigen") = Dr("IDObra")
                            drRowSim("IdClienteProveedor") = Dr("IDCliente")
                            drRowSim("DescClienteProveedor") = Dr("DescCliente")
                            drRowSim("Importe") = Math.Round(DrVto("ImpVencimientoA"), 2)
                            drRowSim("Fecha") = DateAdd(DateInterval.Day, dblDiasRetraso, DrVto("FechaVencimiento"))
                            drRowSim("IdBancoPropio") = Nz(Dr("IdBancoPropio"), data.BancoPropio)
                            drRowSim("IdTipoCobroPago") = System.DBNull.Value ' poner el enumerado
                            drRowSim("IDAgrupacion") = data.AgrupacionIngreso
                            drRowSim("Modificado") = False
                            .Rows.Add(drRowSim)
                        Next
                    End With
                End If
            Next
        Next
        Return data.DtSimulacionTesoreria
    End Function

    <Task()> Public Shared Function SimulacionTesoreriaObrasCompras(ByVal data As DataSimulTesALL, ByVal services As ServiceProvider) As DataTable
        Dim ClsBE As New BE.DataEngine
        Dim ClsParam As New Parametro
        Dim ClsIva As New TipoIva
        Dim DtIva As DataTable
        Dim FactorIva, FactorRE, TotalTramo, TotalDia, TotalMes As Double
        Dim CondPago, DiaPago As String
        Dim Meses, DiasTramo As Integer
        Dim i, pi, mi, mes, dias As Integer
        Dim fecha, fecha2 As Date
        CondPago = ClsParam.CondicionPago
        DiaPago = ClsParam.DiaPago
        DtIva = ClsIva.SelOnPrimaryKey(ClsParam.TipoIva)
        If DtIva.Rows.Count > 0 Then
            FactorIva = DtIva.Rows(0)("Factor")
            FactorRE = DtIva.Rows(0)("Factor")
        Else
            FactorIva = 0
            FactorRE = 0
        End If
        Dim FilDatos As New Filter
        If data.FechaValidez <> cnMinDate AndAlso data.FechaHasta <> cnMinDate Then
            FilDatos.Add("Desde", FilterOperator.GreaterThanOrEqual, data.FechaValidez, FilterType.DateTime)
        End If
        Dim DtAV As DataTable = New BE.DataEngine().Filter("vNegSimulacionTesoreriaOV", FilDatos)
        FilDatos.Clear()
        If data.FechaDesde <> cnMinDate AndAlso data.FechaHasta <> cnMinDate Then
            FilDatos.Add("FechaVencimiento", FilterOperator.GreaterThanOrEqual, data.FechaDesde, FilterType.DateTime)
            FilDatos.Add("FechaVencimiento", FilterOperator.LessThanOrEqual, data.FechaHasta, FilterType.DateTime)
        End If
        For Each DrRowOV As DataRow In DtAV.Select
            Dim FilObra As New Filter
            FilObra.Add("IDObra", FilterOperator.Equal, DrRowOV("IDObra"), FilterType.Numeric)
            FilObra.Add("IDTipoTrabajo", FilterOperator.Equal, DrRowOV("IDTipoTrabajo"), FilterType.String)
            Dim DtTrabajo As DataTable = ClsBE.Filter("vFrmObraTrabajoTipo", FilObra)
            For Each DrTrabajos As DataRow In DtTrabajo.Select
                DiasTramo = DateDiff(DateInterval.Day, DrRowOV("Desde"), DrRowOV("Hasta"))
                If DiasTramo > 0 Then
                    TotalTramo = DrTrabajos("Total") * DrRowOV("Porcentaje") / 100
                    TotalDia = TotalTramo / DiasTramo
                End If
                'Meses por tramo
                fecha = DrRowOV("Desde")
                mes = CDate(DrRowOV("Desde")).Month
                dias = 0
                While fecha <= DrRowOV("Hasta")
                    If mes <> CDate(fecha).Month Then
                        TotalMes = dias * TotalDia
                        ' Aqui tenemos el importe y calculamos los vencimientos
                        Dim StData As New DataVto(String.Empty, String.Empty, TotalMes, DateSerial(Year(fecha), Month(fecha) - 1, 0), _
                                                  CondPago, DiaPago, FactorIva, 0, 0, 0, 0, data.DecimalesImpA)
                        Dim dtVtoOV As DataTable = ProcessServer.ExecuteTask(Of DataVto, DataTable)(AddressOf VtoObra, StData, services)
                        With data.DtSimulacionTesoreria
                            '//Recorremos los Vtos. del las Obras q se encuentran en el intervalo indicado.
                            Dim WherePeriodo As String = FilDatos.Compose(New AdoFilterComposer)
                            For Each DrRowVtoOV As DataRow In dtVtoOV.Select(WherePeriodo)
                                Dim DblDiasRetraso As Double = 0
                                If Length(DrRowOV("IdCliente") & String.Empty) > 0 Then
                                    DblDiasRetraso = ProcessServer.ExecuteTask(Of String, Double)(AddressOf RecuperarDiasRetraso, DrRowOV("IdCliente") & String.Empty, services)
                                End If
                                Dim DrRowSim As DataRow = .NewRow
                                DrRowSim("ParametrosSimulacion") = data.ParametrosSimulacion
                                DrRowSim("EnumTipoDocumento") = enumSimulacionTesoreria.OCNoFacturada
                                DrRowSim("IdOrigen") = DrRowOV("IDObra")
                                DrRowSim("IdClienteProveedor") = ""
                                DrRowSim("DescClienteProveedor") = ""
                                DrRowSim("Importe") = Math.Round(DrRowVtoOV("ImpVencimientoA"), 2) * (-1)
                                DrRowSim("Fecha") = CDate(DrRowVtoOV("FechaVencimiento")).AddDays(DblDiasRetraso)
                                DrRowSim("IdBancoPropio") = Nz(DrRowOV("IdBancoPropio"), data.BancoPropio)
                                DrRowSim("IdTipoCobroPago") = System.DBNull.Value ' poner el enumerado
                                DrRowSim("IDAgrupacion") = data.AgrupacionIngreso
                                DrRowSim("Modificado") = False
                                .Rows.Add(DrRowSim)
                            Next DrRowVtoOV
                        End With
                        dias = 0
                        mes = CDate(fecha).Month
                    End If
                    fecha = fecha.AddDays(1)
                    dias += 1
                End While
                If dias > 1 Then
                    TotalMes = dias * TotalDia
                    Dim StData As New DataVto(String.Empty, String.Empty, TotalMes, DateSerial(Year(fecha), Month(fecha) + 1, 0), _
                                              CondPago, DiaPago, FactorIva, 0, 0, 0, 0, data.DecimalesImpA)
                    Dim dtVtoOV As DataTable = ProcessServer.ExecuteTask(Of DataVto, DataTable)(AddressOf VtoObra, StData, services)
                    With data.DtSimulacionTesoreria
                        '//Recorremos los Vtos. del las Obras q se encuentran en el intervalo indicado.
                        Dim WherePeriodo As String = FilDatos.Compose(New AdoFilterComposer)
                        For Each DrRowVtoOV As DataRow In dtVtoOV.Select(WherePeriodo)
                            Dim DblDiasRetraso As Double = 0
                            If Length(DrRowOV("IdCliente") & String.Empty) > 0 Then
                                DblDiasRetraso = ProcessServer.ExecuteTask(Of String, Double)(AddressOf RecuperarDiasRetraso, DrRowOV("IdCliente") & String.Empty, services)
                            End If
                            Dim DrRowSim As DataRow = .NewRow
                            DrRowSim("ParametrosSimulacion") = data.ParametrosSimulacion
                            DrRowSim("EnumTipoDocumento") = enumSimulacionTesoreria.OCNoFacturada
                            DrRowSim("IdOrigen") = DrRowOV("IDObra")
                            DrRowSim("IdClienteProveedor") = ""
                            DrRowSim("DescClienteProveedor") = ""
                            DrRowSim("Importe") = Math.Round(DrRowVtoOV("ImpVencimientoA"), 2) * (-1)
                            DrRowSim("Fecha") = CDate(DrRowVtoOV("FechaVencimiento")).AddDays(DblDiasRetraso)
                            DrRowSim("IdBancoPropio") = Nz(DrRowOV("IdBancoPropio"), data.BancoPropio)
                            DrRowSim("IdTipoCobroPago") = System.DBNull.Value ' poner el enumerado
                            DrRowSim("IDAgrupacion") = data.AgrupacionIngreso
                            'drRowSim("Empresa") = drRowOV("SyncDB") & String.Empty
                            DrRowSim("Modificado") = False
                            .Rows.Add(DrRowSim)
                        Next DrRowVtoOV
                    End With
                End If
            Next
        Next
        Return data.DtSimulacionTesoreria
    End Function

    <Task()> Public Shared Function SimulacionTesoreriaCobrosPeriodicos(ByVal data As DataSimulTesALL, ByVal services As ServiceProvider) As DataTable
        '//Recuperamos todos los Cobros periódicos y los añadimos a Cobros.
        Dim dtCobrosPer As DataTable = New BE.DataEngine().Filter("vNegSimulacionTesoreriaCP", "*", "")
        If Not IsNothing(dtCobrosPer) AndAlso dtCobrosPer.Rows.Count > 0 Then
            Dim datosCobroPer As New Cobro.DataAddCobroPeriodico(dtCobrosPer, data.FechaHasta, True)
            Dim dtCobros As DataTable = ProcessServer.ExecuteTask(Of Cobro.DataAddCobroPeriodico, DataTable)(AddressOf Cobro.AddCobroPeriodico, datosCobroPer, services)
            If Not dtCobros Is Nothing AndAlso dtCobros.Rows.Count > 0 Then
                For Each drRowCobro As DataRow In dtCobros.Select(Nothing, "FechaVencimiento")
                    With data.DtSimulacionTesoreria
                        If drRowCobro("FechaVencimiento") >= Today Then
                            Dim drRowSim As DataRow = .NewRow
                            drRowSim("ParametrosSimulacion") = data.ParametrosSimulacion
                            drRowSim("EnumTipoDocumento") = enumSimulacionTesoreria.CobroPeriodico
                            drRowSim("IdOrigen") = drRowCobro("IdCobro")
                            If Length(drRowCobro("IdCliente")) > 0 Then
                                drRowSim("IdClienteProveedor") = drRowCobro("IdCliente")
                            End If
                            drRowSim("DescClienteProveedor") = drRowCobro("Titulo")
                            drRowSim("Importe") = drRowCobro("ImpVencimientoA")
                            drRowSim("Fecha") = drRowCobro("FechaVencimiento")
                            drRowSim("IdBancoPropio") = Nz(drRowCobro("IdBancoPropio"), data.BancoPropio)
                            drRowSim("IdTipoCobroPago") = drRowCobro("IdTipoCobro")
                            drRowSim("IDAgrupacion") = Nz(drRowCobro("IDAgrupacion"), data.AgrupacionIngreso)
                            'drRowSim("Empresa") = BDActual & String.Empty
                            drRowSim("Modificado") = False
                            drRowSim("FormaPago") = drRowCobro("IDFormaPago")
                            .Rows.Add(drRowSim)
                        End If
                    End With
                Next drRowCobro
            End If
        End If

        Return data.DtSimulacionTesoreria
    End Function

    <Task()> Public Shared Function SimulacionTesoreriaPagos(ByVal data As DataSimulTesALL, ByVal services As ServiceProvider) As DataTable
        Dim objFilter As New Filter
        objFilter.Clear()
        If data.FechaDesde <> cnMinDate AndAlso data.FechaHasta <> cnMinDate Then
            objFilter.Add(New BetweenFilterItem("FechaVencimiento", data.FechaDesde, data.FechaHasta, FilterType.DateTime))
        End If

        Dim dtPagos As DataTable = New BE.DataEngine().Filter("vNegSimulacionTesoreriaP", objFilter)

        For Each drRowPago As DataRow In dtPagos.Select(Nothing, "FechaVencimiento")
            With data.DtSimulacionTesoreria
                Dim drRowSim As DataRow = .NewRow
                drRowSim("ParametrosSimulacion") = data.ParametrosSimulacion
                drRowSim("EnumTipoDocumento") = enumSimulacionTesoreria.Pago
                drRowSim("IdOrigen") = drRowPago("IdPago")
                drRowSim("IdClienteProveedor") = drRowPago("IDProveedor")
                drRowSim("DescClienteProveedor") = drRowPago("Titulo")
                drRowSim("Importe") = drRowPago("ImpVencimientoA")
                drRowSim("Fecha") = drRowPago("FechaVencimiento")
                drRowSim("IdBancoPropio") = Nz(drRowPago("IdBancoPropio"), data.BancoPropio)
                drRowSim("IdTipoCobroPago") = drRowPago("IdTipoPago")
                drRowSim("IDAgrupacion") = Nz(drRowPago("IDAgrupacion"), data.AgrupacionGasto)
                'drRowSim("Empresa") = drRowPago("SyncDB") & String.Empty
                drRowSim("Modificado") = False
                drRowSim("FormaPago") = drRowPago("IDFormaPago")
                drRowSim("Situacion") = drRowPago("Situacion")
                .Rows.Add(drRowSim)
            End With
        Next drRowPago

        objFilter = Nothing
        Return data.DtSimulacionTesoreria
    End Function

    <Task()> Public Shared Function SimulacionTesoreriaPagosPeriodicos(ByVal data As DataSimulTesALL, ByVal services As ServiceProvider) As DataTable
        '//Recuperamos todos los Cobros periódicos y los añadimos a Cobros.
        Dim dtPagosPer As DataTable = New BE.DataEngine().Filter("vNegSimulacionTesoreriaPP", "*", "")

        If Not IsNothing(dtPagosPer) AndAlso dtPagosPer.Rows.Count > 0 Then
            Dim addPagoPer As New Pago.DataAddPagoPeriodico(dtPagosPer, data.FechaHasta, True)
            Dim dtPagos As DataTable = ProcessServer.ExecuteTask(Of Pago.DataAddPagoPeriodico, DataTable)(AddressOf Pago.AddPagoPeriodico, addPagoPer, services)
            If Not dtPagos Is Nothing AndAlso dtPagos.Rows.Count > 0 Then
                For Each drRowPago As DataRow In dtPagos.Select(Nothing, "FechaVencimiento")
                    With data.DtSimulacionTesoreria
                        If drRowPago("FechaVencimiento") >= Today Then
                            Dim drRowSim As DataRow = .NewRow
                            drRowSim("ParametrosSimulacion") = data.ParametrosSimulacion
                            drRowSim("EnumTipoDocumento") = enumSimulacionTesoreria.PagoPeriodico
                            drRowSim("IdOrigen") = drRowPago("IdPago")
                            If Length(drRowPago("IdProveedor")) > 0 Then
                                drRowSim("IdClienteProveedor") = drRowPago("IdProveedor")
                            End If
                            drRowSim("DescClienteProveedor") = drRowPago("Titulo")
                            drRowSim("Importe") = -1 * drRowPago("ImpVencimientoA")
                            drRowSim("Fecha") = drRowPago("FechaVencimiento")
                            drRowSim("IdBancoPropio") = Nz(drRowPago("IdBancoPropio"), data.BancoPropio)
                            drRowSim("IdTipoCobroPago") = drRowPago("IdTipoPago")
                            drRowSim("IDAgrupacion") = Nz(drRowPago("IDAgrupacion"), data.AgrupacionGasto)
                            'drRowSim("Empresa") = BDActual & String.Empty
                            drRowSim("Modificado") = False
                            drRowSim("FormaPago") = drRowPago("IDFormaPago")
                            .Rows.Add(drRowSim)
                        End If
                    End With
                Next drRowPago
                dtPagosPer.Rows.Clear()
            End If
        End If

        Return data.DtSimulacionTesoreria
    End Function

    <Task()> Public Shared Function SimulacionTesoreriaFacturasCompras(ByVal data As DataSimulTesALL, ByVal services As ServiceProvider) As DataTable
        Dim objFilter As New Filter
        objFilter.Clear()

        If data.FechaDesde <> cnMinDate AndAlso data.FechaHasta <> cnMinDate Then
            objFilter.Add(New BetweenFilterItem("FechaVencimiento", data.FechaDesde, data.FechaHasta, FilterType.DateTime))
        End If

        Dim dtFC As DataTable = New BE.DataEngine().Filter("vNegSimulacionTesoreriaFC", objFilter)
        If Not dtFC Is Nothing AndAlso dtFC.Rows.Count > 0 Then
            For Each drRowFC As DataRow In dtFC.Rows
                With data.DtSimulacionTesoreria
                    If drRowFC("FechaVencimiento") >= Today Then
                        Dim drRowSim As DataRow = .NewRow
                        drRowSim("ParametrosSimulacion") = data.ParametrosSimulacion
                        drRowSim("EnumTipoDocumento") = enumSimulacionTesoreria.FCNoContabilizada
                        drRowSim("IdOrigen") = drRowFC("IDFactura")
                        drRowSim("IdClienteProveedor") = drRowFC("IDProveedor")
                        drRowSim("DescClienteProveedor") = drRowFC("DescProveedor")
                        drRowSim("Importe") = drRowFC("ImpVencimientoA")
                        drRowSim("Fecha") = drRowFC("FechaVencimiento")
                        drRowSim("IdBancoPropio") = Nz(drRowFC("IdBancoPropio"), data.BancoPropio)
                        drRowSim("IdTipoCobroPago") = System.DBNull.Value '0 ' poner el enumerado
                        drRowSim("IDAgrupacion") = data.AgrupacionGasto
                        'drRowSim("Empresa") = drRowFC("SyncDB") & String.Empty
                        drRowSim("Modificado") = False
                        drRowSim("FormaPago") = drRowFC("IDFormaPago")
                        .Rows.Add(drRowSim)
                    End If
                End With
            Next drRowFC
        End If

        Return data.DtSimulacionTesoreria
    End Function

    <Task()> Public Shared Function SimulacionTesoreriaFacturasVentas(ByVal data As DataSimulTesALL, ByVal services As ServiceProvider) As DataTable
        Dim objFilter As New Filter
        objFilter.Clear()

        If data.FechaDesde <> cnMinDate AndAlso data.FechaHasta <> cnMinDate Then
            objFilter.Add(New BetweenFilterItem("FechaVencimiento", data.FechaDesde, data.FechaHasta, FilterType.DateTime))
        End If

        Dim dtFV As DataTable = New BE.DataEngine().Filter("vNegSimulacionTesoreriaFV", objFilter)
        If Not dtFV Is Nothing AndAlso dtFV.Rows.Count > 0 Then
            For Each drRowFV As DataRow In dtFV.Rows
                Dim dblDiasRetraso As Double = 0
                If Length(drRowFV("IdCliente") & String.Empty) > 0 Then
                    dblDiasRetraso = ProcessServer.ExecuteTask(Of String, Double)(AddressOf RecuperarDiasRetraso, drRowFV("IdCliente") & String.Empty, services)
                End If

                With data.DtSimulacionTesoreria
                    Dim drRowSim As DataRow = .NewRow
                    drRowSim("ParametrosSimulacion") = data.ParametrosSimulacion
                    drRowSim("EnumTipoDocumento") = enumSimulacionTesoreria.FVNoContabilizada
                    drRowSim("IdOrigen") = drRowFV("IDFactura")
                    drRowSim("IdClienteProveedor") = drRowFV("IDCliente")
                    drRowSim("DescClienteProveedor") = drRowFV("DescCliente")
                    drRowSim("Importe") = drRowFV("ImpVencimientoA")
                    drRowSim("Fecha") = DateAdd(DateInterval.Day, dblDiasRetraso, drRowFV("FechaVencimiento"))
                    drRowSim("IdBancoPropio") = Nz(drRowFV("IdBancoPropio"), data.BancoPropio)
                    drRowSim("IdTipoCobroPago") = System.DBNull.Value ' poner el enumerado
                    drRowSim("IDAgrupacion") = data.AgrupacionIngreso
                    'drRowSim("Empresa") = drRowFC("SyncDB") & String.Empty
                    drRowSim("Modificado") = False
                    drRowSim("FormaPago") = drRowFV("IDFormaPago")

                    .Rows.Add(drRowSim)
                End With
            Next drRowFV
        End If

        Return data.DtSimulacionTesoreria
    End Function

    <Task()> Public Shared Function SimulacionTesoreriaAlbaranesVentas(ByVal data As DataSimulTesALL, ByVal services As ServiceProvider) As DataTable
        Dim objFilter As New Filter
        objFilter.Clear()
        If data.FechaValidez <> cnMinDate AndAlso data.FechaHasta <> cnMinDate Then
            objFilter.Add(New BetweenFilterItem("FechaVencimiento", data.FechaValidez, data.FechaHasta, FilterType.DateTime))
        End If
        Dim dtAV As DataTable = New BE.DataEngine().Filter("vNegSimulacionTesoreriaAV", objFilter)

        objFilter.Clear()
        If data.FechaDesde <> cnMinDate AndAlso data.FechaHasta <> cnMinDate Then
            objFilter.Add(New BetweenFilterItem("FechaVencimiento", data.FechaDesde, data.FechaHasta, FilterType.DateTime))
        End If

        If Not dtAV Is Nothing AndAlso dtAV.Rows.Count > 0 Then


            For Each drRowAV As DataRow In dtAV.Rows

                Dim StData As New DataVto(drRowAV("IDCliente"), String.Empty, drRowAV("ImpVencimientoA"), drRowAV("FechaVencimiento"), _
                                          Nz(drRowAV("IDCondicionPago"), data.CondpagoPred), Nz(drRowAV("IDDiaPago"), data.DiaPagoPred), _
                                          Nz(drRowAV("FactorIVA"), 0), IIf(drRowAV("TieneRE"), drRowAV("FactorRE"), 0), drRowAV("DtoComercial"), _
                                          Nz(drRowAV("FactorDPP"), 0), Nz(drRowAV("FactorRecFinan"), 0), data.DecimalesImpA)
                Dim dtVtoAV As DataTable = ProcessServer.ExecuteTask(Of DataVto, DataTable)(AddressOf VtoAlbaranVenta, StData, services)
                With data.DtSimulacionTesoreria
                    '//Recorremos los Vtos. del los Albaranes q se encuentran en el intervalo indicado.
                    Dim WherePeriodo As String = objFilter.Compose(New AdoFilterComposer)
                    For Each drRowVtoAV As DataRow In dtVtoAV.Select(WherePeriodo)
                        Dim dblDiasRetraso As Double = 0
                        If Length(drRowAV("IdCliente") & String.Empty) > 0 Then
                            dblDiasRetraso = ProcessServer.ExecuteTask(Of String, Double)(AddressOf RecuperarDiasRetraso, drRowAV("IdCliente") & String.Empty, services)
                        End If

                        Dim drRowSim As DataRow = .NewRow
                        drRowSim("ParametrosSimulacion") = data.ParametrosSimulacion
                        drRowSim("EnumTipoDocumento") = enumSimulacionTesoreria.AVNoFacturado
                        drRowSim("IdOrigen") = drRowAV("IDAlbaran")
                        drRowSim("IdClienteProveedor") = drRowAV("IDCliente")
                        drRowSim("DescClienteProveedor") = drRowAV("DescCliente")
                        drRowSim("Importe") = drRowVtoAV("ImpVencimientoA")
                        drRowSim("Fecha") = DateAdd(DateInterval.Day, dblDiasRetraso, drRowVtoAV("FechaVencimiento"))
                        drRowSim("IdBancoPropio") = Nz(drRowAV("IdBancoPropio"), data.BancoPropio)
                        drRowSim("IdTipoCobroPago") = System.DBNull.Value ' poner el enumerado
                        drRowSim("IDAgrupacion") = data.AgrupacionIngreso
                        'drRowSim("Empresa") = drRowAV("SyncDB") & String.Empty
                        drRowSim("Modificado") = False
                        drRowSim("FormaPago") = drRowAV("IDFormaPago")
                        .Rows.Add(drRowSim)
                    Next drRowVtoAV
                End With
            Next
        End If

        Return data.DtSimulacionTesoreria
    End Function

    <Task()> Public Shared Function SimulacionTesoreriaAlbaranesCompras(ByVal data As DataSimulTesALL, ByVal services As ServiceProvider) As DataTable
        Dim objFilter As New Filter
        objFilter.Clear()
        If data.FechaValidez <> cnMinDate AndAlso data.FechaHasta <> cnMinDate Then
            objFilter.Add(New BetweenFilterItem("FechaVencimiento", data.FechaValidez, data.FechaHasta, FilterType.DateTime))
        End If
        Dim dtAC As DataTable = New BE.DataEngine().Filter("vNegSimulacionTesoreriaAC", objFilter)

        objFilter.Clear()
        If data.FechaDesde <> cnMinDate AndAlso data.FechaHasta <> cnMinDate Then
            objFilter.Add(New BetweenFilterItem("FechaVencimiento", data.FechaDesde, data.FechaHasta, FilterType.DateTime))
        End If
        If Not dtAC Is Nothing AndAlso dtAC.Rows.Count > 0 Then

            For Each drRowAC As DataRow In dtAC.Rows
                Dim StData As New DataVto(String.Empty, drRowAC("IDProveedor"), IIf(drRowAC("FactorIRPF") > 0, drRowAC("ImpVencimientoA") * (1 - (drRowAC("FactorIRPF") / 100)), drRowAC("ImpVencimientoA")), _
                                          drRowAC("FechaVencimiento"), Nz(drRowAC("IDCondicionPago"), data.CondpagoPred), Nz(drRowAC("IDDiaPago"), data.DiaPagoPred), _
                                          Nz(drRowAC("FactorIVA"), 0), IIf(data.CalcularImpRE, drRowAC("FactorRE"), 0), drRowAC("DtoComercial"), Nz(drRowAC("FactorDPP"), 0), _
                                          Nz(drRowAC("FactorRecFinan"), 0), data.DecimalesImpA)
                Dim dtVtoAC As DataTable = ProcessServer.ExecuteTask(Of DataVto, DataTable)(AddressOf VtoAlbaranCompra, StData, services)
                With data.DtSimulacionTesoreria
                    Dim drRowSim As DataRow
                    '//Recorremos los Vtos. del los Albaranes q se encuentran en el intervalo indicado.
                    Dim WherePeriodo As String = objFilter.Compose(New AdoFilterComposer)
                    For Each drRowVtoAC As DataRow In dtVtoAC.Select(WherePeriodo)
                        drRowSim = .NewRow
                        drRowSim("ParametrosSimulacion") = data.ParametrosSimulacion
                        drRowSim("EnumTipoDocumento") = enumSimulacionTesoreria.ACNoFacturado
                        drRowSim("IdOrigen") = drRowAC("IDAlbaran")
                        drRowSim("IdClienteProveedor") = drRowAC("IDProveedor")
                        drRowSim("DescClienteProveedor") = drRowAC("DescProveedor")
                        drRowSim("Importe") = drRowVtoAC("ImpVencimientoA")
                        drRowSim("Fecha") = drRowVtoAC("FechaVencimiento")
                        drRowSim("IdBancoPropio") = Nz(drRowAC("IdBancoPropio"), data.BancoPropio)
                        drRowSim("IdTipoCobroPago") = System.DBNull.Value ' poner el enumerado
                        drRowSim("IDAgrupacion") = data.AgrupacionGasto
                        'drRowSim("Empresa") = drRowAV("SyncDB") & String.Empty
                        drRowSim("Modificado") = False
                        drRowSim("FormaPago") = drRowAC("IDFormaPago")
                        .Rows.Add(drRowSim)
                    Next

                    If drRowAC("FactorIRPF") > 0 AndAlso (drRowAC("FechaVencimiento") >= data.FechaDesde AndAlso drRowAC("FechaVencimiento") <= data.FechaHasta) Then
                        drRowSim = .NewRow
                        drRowSim("ParametrosSimulacion") = data.ParametrosSimulacion
                        drRowSim("EnumTipoDocumento") = enumSimulacionTesoreria.ACNoFacturado
                        drRowSim("IdOrigen") = drRowAC("IDAlbaran")
                        drRowSim("IdClienteProveedor") = data.IDProvRetencion
                        drRowSim("DescClienteProveedor") = data.DescProvRetencion
                        Dim dblBaseIRPF As Double = Nz(drRowAC("ImpVencimientoA"), 0)
                        If drRowAC("DtoComercial") > 0 Then dblBaseIRPF = xRound(dblBaseIRPF * (1 - (drRowAC("DtoComercial") / 100)), data.DecimalesImpA)
                        If drRowAC("FactorDPP") > 0 Then dblBaseIRPF = xRound(dblBaseIRPF * (1 - (drRowAC("FactorDPP") / 100)), data.DecimalesImpA)
                        drRowSim("Importe") = IIf(drRowAC("FactorIRPF") > 0, xRound(dblBaseIRPF * drRowAC("FactorIRPF") / 100, data.DecimalesImpA), 0)
                        drRowSim("Fecha") = drRowAC("FechaVencimiento")
                        drRowSim("IdBancoPropio") = Nz(drRowAC("IdBancoPropio"), data.BancoPropio)
                        drRowSim("IdTipoCobroPago") = System.DBNull.Value ' poner el enumerado
                        drRowSim("IDAgrupacion") = data.AgrupacionGasto
                        'drRowSim("Empresa") = drRowAC("SyncDB") & String.Empty
                        drRowSim("Modificado") = False
                        drRowSim("FormaPago") = drRowAC("IDFormaPago")

                        .Rows.Add(drRowSim)
                    End If
                End With
            Next drRowAC

        End If

        Return data.DtSimulacionTesoreria
    End Function

    <Task()> Public Shared Function SimulacionTesoreriaPedidosCompras(ByVal data As DataSimulTesALL, ByVal services As ServiceProvider) As DataTable
        Dim objFilter As New Filter
        objFilter.Clear()
        If data.FechaValidez <> cnMinDate AndAlso data.FechaHasta <> cnMinDate Then
            objFilter.Add(New BetweenFilterItem("FechaVencimiento", data.FechaValidez, data.FechaHasta, FilterType.DateTime))
        End If
        Dim dtPC As DataTable = New BE.DataEngine().Filter("vNegSimulacionTesoreriaPC", objFilter)

        objFilter.Clear()
        If data.FechaDesde <> cnMinDate AndAlso data.FechaHasta <> cnMinDate Then
            objFilter.Add(New BetweenFilterItem("FechaVencimiento", data.FechaDesde, data.FechaHasta, FilterType.DateTime))
        End If

        If Not dtPC Is Nothing AndAlso dtPC.Rows.Count > 0 Then

            For Each drRowPC As DataRow In dtPC.Rows
                Dim StData As New DataVto(String.Empty, drRowPC("IDProveedor"), IIf(drRowPC("FactorIRPF") > 0, drRowPC("ImpVencimientoA") * (1 - (drRowPC("FactorIRPF") / 100)), drRowPC("ImpVencimientoA")), _
                                          drRowPC("FechaVencimiento"), Nz(drRowPC("IDCondicionPago"), data.CondpagoPred), Nz(drRowPC("IDDiaPago"), data.DiaPagoPred), Nz(drRowPC("FactorIVA"), 0), _
                                          IIf(data.CalcularImpRE, drRowPC("FactorRE"), 0), drRowPC("DtoComercial"), Nz(drRowPC("FactorDPP"), 0), Nz(drRowPC("FactorRecFinan"), 0), data.DecimalesImpA)
                Dim dtVtoPC As DataTable = ProcessServer.ExecuteTask(Of DataVto, DataTable)(AddressOf VtoPedidoCompra, StData, services)
                With data.DtSimulacionTesoreria
                    Dim drRowSim As DataRow
                    '//Recorremos los Vtos. del los Albaranes q se encuentran en el intervalo indicado.
                    Dim WherePeriodo As String = objFilter.Compose(New AdoFilterComposer)
                    For Each drRowVtoPC As DataRow In dtVtoPC.Select(WherePeriodo)
                        drRowSim = .NewRow
                        drRowSim("ParametrosSimulacion") = data.ParametrosSimulacion
                        drRowSim("EnumTipoDocumento") = enumSimulacionTesoreria.PCPendiente
                        drRowSim("IdOrigen") = drRowPC("IDPedido")
                        drRowSim("IdClienteProveedor") = drRowPC("IDProveedor")
                        drRowSim("DescClienteProveedor") = drRowPC("DescProveedor")
                        drRowSim("Importe") = drRowVtoPC("ImpVencimientoA")
                        drRowSim("Fecha") = drRowVtoPC("FechaVencimiento")
                        drRowSim("IdBancoPropio") = Nz(drRowPC("IdBancoPropio"), data.BancoPropio)
                        drRowSim("IdTipoCobroPago") = System.DBNull.Value  '0' poner el enumerado
                        drRowSim("IDAgrupacion") = data.AgrupacionGasto
                        'drRowSim("Empresa") = drRowAV("SyncDB") & String.Empty
                        drRowSim("Modificado") = False
                        drRowSim("FormaPago") = drRowPC("IDFormaPago")
                        .Rows.Add(drRowSim)
                    Next

                    If drRowPC("FactorIRPF") > 0 AndAlso (drRowPC("FechaVencimiento") >= data.FechaDesde AndAlso drRowPC("FechaVencimiento") <= data.FechaHasta) Then
                        drRowSim = .NewRow
                        drRowSim("ParametrosSimulacion") = data.ParametrosSimulacion
                        drRowSim("EnumTipoDocumento") = enumSimulacionTesoreria.PCPendiente
                        drRowSim("IdOrigen") = drRowPC("IDPedido")
                        drRowSim("IdClienteProveedor") = data.IDProvRetencion
                        drRowSim("DescClienteProveedor") = data.DescProvRetencion
                        Dim dblBaseIRPF As Double = Nz(drRowPC("ImpVencimientoA"), 0)
                        If drRowPC("DtoComercial") > 0 Then dblBaseIRPF = xRound(dblBaseIRPF * (1 - (drRowPC("DtoComercial") / 100)), data.DecimalesImpA)
                        If drRowPC("FactorDPP") > 0 Then dblBaseIRPF = xRound(dblBaseIRPF * (1 - (drRowPC("FactorDPP") / 100)), data.DecimalesImpA)
                        drRowSim("Importe") = IIf(drRowPC("FactorIRPF") > 0, xRound(dblBaseIRPF * drRowPC("FactorIRPF") / 100, data.DecimalesImpA), 0)
                        drRowSim("Fecha") = drRowPC("FechaVencimiento")
                        drRowSim("IdBancoPropio") = Nz(drRowPC("IdBancoPropio"), data.BancoPropio)
                        drRowSim("IdTipoCobroPago") = System.DBNull.Value '0 ' poner el enumerado
                        drRowSim("IDAgrupacion") = data.AgrupacionGasto
                        'drRowSim("Empresa") = drRowAC("SyncDB") & String.Empty
                        drRowSim("Modificado") = False
                        drRowSim("FormaPago") = drRowPC("IDFormaPago")
                        .Rows.Add(drRowSim)
                    End If
                End With
            Next drRowPC
        End If

        Return data.DtSimulacionTesoreria
    End Function

    <Task()> Public Shared Function SimulacionTesoreriaPedidosVentas(ByVal data As DataSimulTesALL, ByVal services As ServiceProvider) As DataTable
        Dim objFilter As New Filter
        objFilter.Clear()
        If data.FechaValidez <> cnMinDate AndAlso data.FechaHasta <> cnMinDate Then
            objFilter.Add(New BetweenFilterItem("FechaVencimiento", data.FechaValidez, data.FechaHasta, FilterType.DateTime))
        End If
        Dim dtPV As DataTable = New BE.DataEngine().Filter("vNegSimulacionTesoreriaPV", objFilter)

        objFilter.Clear()
        If data.FechaDesde <> cnMinDate AndAlso data.FechaHasta <> cnMinDate Then
            objFilter.Add(New BetweenFilterItem("FechaVencimiento", data.FechaDesde, data.FechaHasta, FilterType.DateTime))
        End If
        If Not dtPV Is Nothing AndAlso dtPV.Rows.Count > 0 Then

            For Each drRowPV As DataRow In dtPV.Rows
                Dim StData As New DataVto(drRowPV("IDCliente"), String.Empty, drRowPV("ImpVencimientoA"), drRowPV("FechaVencimiento"), Nz(drRowPV("IDCondicionPago"), data.CondpagoPred), _
                                          Nz(drRowPV("IDDiaPago"), data.DiaPagoPred), Nz(drRowPV("FactorIVA"), 0), IIf(drRowPV("TieneRE"), drRowPV("FactorRE"), 0), drRowPV("DtoComercial"), _
                                          Nz(drRowPV("FactorDPP"), 0), Nz(drRowPV("FactorRecFinan"), 0), data.DecimalesImpA)
                Dim dtVtoPV As DataTable = ProcessServer.ExecuteTask(Of DataVto, DataTable)(AddressOf VtoPedidoVenta, StData, services)
                Dim WherePeriodo As String = objFilter.Compose(New AdoFilterComposer)
                For Each drRowVtoPV As DataRow In dtVtoPV.Select(WherePeriodo)
                    Dim dblDiasRetraso As Double = 0
                    If Length(drRowPV("IdCliente") & String.Empty) > 0 Then
                        dblDiasRetraso = ProcessServer.ExecuteTask(Of String, Double)(AddressOf RecuperarDiasRetraso, drRowPV("IdCliente") & String.Empty, services)
                    End If

                    With data.DtSimulacionTesoreria
                        Dim drRowSim As DataRow = .NewRow
                        drRowSim("ParametrosSimulacion") = data.ParametrosSimulacion
                        drRowSim("EnumTipoDocumento") = enumSimulacionTesoreria.PVPendiente
                        drRowSim("IdOrigen") = drRowPV("IDPedido")
                        drRowSim("IdClienteProveedor") = drRowPV("IDCliente")
                        drRowSim("DescClienteProveedor") = drRowPV("DescCliente")
                        drRowSim("Importe") = drRowVtoPV("ImpVencimientoA")
                        drRowSim("Fecha") = DateAdd(DateInterval.Day, dblDiasRetraso, drRowVtoPV("FechaVencimiento"))
                        drRowSim("IdBancoPropio") = Nz(drRowPV("IdBancoPropio"), data.BancoPropio)
                        drRowSim("IdTipoCobroPago") = System.DBNull.Value  ' ' poner el enumerado
                        drRowSim("IDAgrupacion") = data.AgrupacionIngreso
                        'drRowSim("Empresa") = drRowAV("SyncDB") & String.Empty
                        drRowSim("Modificado") = False
                        drRowSim("FormaPago") = drRowPV("IDFormaPago")
                        .Rows.Add(drRowSim)
                    End With
                Next drRowVtoPV
            Next drRowPV
        End If

        Return data.DtSimulacionTesoreria
    End Function

    <Task()> Public Shared Function SimulacionTesoreriaBancosPropios(ByVal data As DataSimulTesALL, ByVal services As ServiceProvider) As DataTable
        '//Recuperamos la información de los Bancos Propios, que tengan definida una C.Contable.
        Dim objFilter As New Filter
        objFilter.Clear()
        objFilter.Add(New NumberFilterItem("LEN(CContable)", FilterOperator.GreaterThan, 0))
        Dim objNegBancoPropio As New BancoPropio
        Dim dtBancoPropio As DataTable = objNegBancoPropio.Filter(objFilter)
        '//Recuperamos los saldos de TODOS los B. Propios, para luego operar sobre el DataTable y no estar accediendo a la BD continuamente.
        Dim dtSaldosBP As DataTable = ProcessServer.ExecuteTask(Of Date, DataTable)(AddressOf BancoPropio.SaldosBancosPropiosAFecha, Today.Date, services)
        If Not dtBancoPropio Is Nothing AndAlso dtBancoPropio.Rows.Count > 0 Then
            For Each drRowBancoPropio As DataRow In dtBancoPropio.Rows
                If Length(drRowBancoPropio("CContable") & String.Empty) > 0 Then
                    With data.DtSimulacionTesoreria
                        Dim drRowSim As DataRow = .NewRow
                        drRowSim("ParametrosSimulacion") = data.ParametrosSimulacion
                        drRowSim("EnumTipoDocumento") = enumSimulacionTesoreria.BancoPropio
                        drRowSim("IdClienteProveedor") = drRowBancoPropio("IdBancoPropio")
                        drRowSim("DescClienteProveedor") = drRowBancoPropio("DescBancoPropio")
                        drRowSim("IdBancoPropio") = drRowBancoPropio("IdBancoPropio")
                        drRowSim("Importe") = 0     '//Por defecto
                        drRowSim("Fecha") = Today   '//Por defecto
                        objFilter.Clear()
                        objFilter.Add(New StringFilterItem("IDBancoPropio", drRowBancoPropio("IdBancoPropio")))
                        Dim WhereBancoPropio As String = objFilter.Compose(New AdoFilterComposer)
                        For Each drRowSaldo As DataRow In dtSaldosBP.Select(WhereBancoPropio)
                            drRowSim("Importe") = drRowSaldo("Saldo")
                            drRowSim("Fecha") = drRowSaldo("Fecha")
                        Next drRowSaldo
                        drRowSim("Modificado") = False

                        .Rows.Add(drRowSim)
                    End With
                End If
            Next drRowBancoPropio
        End If

        Return data.DtSimulacionTesoreria
    End Function

    <Task()> Public Shared Function SimulacionTesoreriaObrasHitos(ByVal data As DataSimulTesALL, ByVal services As ServiceProvider) As DataTable
        Dim objFilter As New Filter
        objFilter.Clear()
        If data.FechaValidez <> cnMinDate AndAlso data.FechaHasta <> cnMinDate Then
            objFilter.Add(New BetweenFilterItem("FechaVencimiento", data.FechaValidez, data.FechaHasta, FilterType.DateTime))
        End If
        Dim dtAV As DataTable = New BE.DataEngine().Filter("vNegSimulacionTesoreriaObraHitos", objFilter)
        Dim dtMOD As DataTable = New BE.DataEngine().Filter("vNegSimulacionTesoreriaObraMOD", objFilter)
        Dim dtMaterial As DataTable = New BE.DataEngine().Filter("vNegSimulacionTesoreriaObraMaterial", objFilter)
        Dim dtGasto As DataTable = New BE.DataEngine().Filter("vNegSimulacionTesoreriaObraGasto", objFilter)
        Dim dtCertifica As DataTable = New BE.DataEngine().Filter("vNegSimulacionTesoreriaObraCertificacion", objFilter)
        dtAV.Merge(dtMOD)
        dtAV.Merge(dtMaterial)
        dtAV.Merge(dtGasto)
        dtAV.Merge(dtCertifica)
        objFilter.Clear()
        If data.FechaDesde <> cnMinDate AndAlso data.FechaHasta <> cnMinDate Then
            objFilter.Add(New BetweenFilterItem("FechaVencimiento", data.FechaDesde, data.FechaHasta, FilterType.DateTime))
        End If

        If Not dtAV Is Nothing AndAlso dtAV.Rows.Count > 0 Then


            For Each drRowAV As DataRow In dtAV.Rows
                Dim StData As New DataVto(drRowAV("IDCliente"), String.Empty, Nz(drRowAV("ImpVencimientoA"), 0), drRowAV("FechaVencimiento"), Nz(drRowAV("IDCondicionPago"), data.CondpagoPred), _
                                                           Nz(drRowAV("IDDiaPago"), data.DiaPagoPred), Nz(drRowAV("FactorIVA"), 0), IIf(drRowAV("TieneRE"), drRowAV("FactorRE"), 0), Nz(drRowAV("DtoComercial"), 0), _
                                                           Nz(drRowAV("FactorDPP"), 0), Nz(drRowAV("FactorRecFinan"), 0), data.DecimalesImpA)
                Dim dtVtoAV As DataTable = ProcessServer.ExecuteTask(Of DataVto, DataTable)(AddressOf VtoObraHito, StData, services)
                With data.DtSimulacionTesoreria
                    '//Recorremos los Vtos. del los Albaranes q se encuentran en el intervalo indicado.
                    Dim WherePeriodo As String = objFilter.Compose(New AdoFilterComposer)
                    For Each drRowVtoAV As DataRow In dtVtoAV.Select(WherePeriodo)
                        Dim dblDiasRetraso As Double = 0
                        If Length(drRowAV("IdCliente") & String.Empty) > 0 Then
                            dblDiasRetraso = ProcessServer.ExecuteTask(Of String, Double)(AddressOf RecuperarDiasRetraso, drRowAV("IdCliente") & String.Empty, services)
                        End If

                        Dim drRowSim As DataRow = .NewRow
                        drRowSim("ParametrosSimulacion") = data.ParametrosSimulacion
                        drRowSim("EnumTipoDocumento") = enumSimulacionTesoreria.OVNoFacturada
                        drRowSim("IdOrigen") = drRowAV("IDObra")
                        drRowSim("IdClienteProveedor") = drRowAV("IDCliente")
                        drRowSim("DescClienteProveedor") = drRowAV("DescCliente")
                        drRowSim("Importe") = drRowVtoAV("ImpVencimientoA")
                        drRowSim("Fecha") = DateAdd(DateInterval.Day, dblDiasRetraso, drRowVtoAV("FechaVencimiento"))
                        drRowSim("IdBancoPropio") = Nz(drRowAV("IdBancoPropio"), data.BancoPropio)
                        drRowSim("IdTipoCobroPago") = System.DBNull.Value ' poner el enumerado
                        drRowSim("IDAgrupacion") = data.AgrupacionIngreso
                        'drRowSim("Empresa") = drRowAV("SyncDB") & String.Empty
                        drRowSim("Modificado") = False
                        drRowSim("FormaPago") = drRowAV("IDFormaPago")
                        drRowSim("DescSituacion") = drRowAV("DescSituacion")
                        .Rows.Add(drRowSim)
                    Next drRowVtoAV
                End With
            Next
        End If

        Return data.DtSimulacionTesoreria
    End Function

    <Serializable()> _
    Public Class DataVto
        Public IDCliente As String
        Public IDProveedor As String
        Public ImpVencimiento As Double
        Public FechaVencimiento As Date
        Public IDCondicionPago As String
        Public IDDiaPago As String
        Public FactorIVA As Double
        Public FactorRE As Double
        Public DtoFactura As Boolean
        Public FactorDPP As Double
        Public FactorRecFinan As Double
        Public Decimales As Integer

        Public Sub New()
        End Sub
        Public Sub New(ByVal IDCliente As String, ByVal IDProveedor As String, ByVal ImpVencimiento As Double, ByVal FechaVencimiento As Date, ByVal IDCondicionpago As String, _
                       ByVal IDDiaPago As String, ByVal FactorIVA As Double, ByVal FacorRE As Double, ByVal DtoFactura As Boolean, ByVal FactorDPP As Double, ByVal FactorRecFinan As Double, ByVal Decimales As Integer)
            Me.IDCliente = IDCliente
            Me.IDProveedor = IDProveedor
            Me.ImpVencimiento = ImpVencimiento
            Me.FechaVencimiento = FechaVencimiento
            Me.IDCondicionPago = IDCondicionpago
            Me.IDDiaPago = IDDiaPago
            Me.FactorIVA = FactorIVA
            Me.FactorRE = FactorRE
            Me.DtoFactura = DtoFactura
            Me.FactorDPP = FactorDPP
            Me.FactorRecFinan = FactorRecFinan
            Me.Decimales = Decimales
        End Sub
    End Class

    <Task()> Public Shared Function VtoAlbaranCompra(ByVal data As DataVto, ByVal services As ServiceProvider) As DataTable
        Dim dblTotal As Double
        Dim dblBase As Double
        Dim dblIVA As Double
        Dim dblRE As Double
        Dim dblRecFinan As Double

        '//Calculamos los importes
        dblBase = data.ImpVencimiento
        If data.DtoFactura > 0 Then dblBase = xRound(dblBase * (1 - (CShort(data.DtoFactura) / 100)), data.Decimales)
        If data.FactorDPP > 0 Then dblBase = xRound(dblBase * (1 - (data.FactorDPP / 100)), data.Decimales)
        If data.FactorIVA > 0 Then dblIVA = xRound(dblBase * data.FactorIVA / 100, data.Decimales)
        If data.FactorRE > 0 Then dblRE = xRound(dblBase * data.FactorRE / 100, data.Decimales)
        If data.FactorDPP = 0 And data.FactorRecFinan > 0 Then dblRecFinan = xRound(dblBase * data.FactorRecFinan / 100, data.Decimales)
        dblTotal = dblBase + dblIVA + dblRE + dblRecFinan

        '//Construimos el DataTable a devolver.
        Dim dtVtoAlbaranCompra As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf CrearDtVencimientos, Nothing, services)

        '//Obtenemos las Condiciones de Pago de la línea
        Dim objFilter As New Filter
        objFilter.Clear()
        objFilter.Add(New StringFilterItem("IDCondicionPago", data.IDCondicionPago))

        Dim objNegCondicionLinea As New CondicionPagoLinea
        Dim dtCondicionesLinea As DataTable = objNegCondicionLinea.Filter(objFilter)
        objNegCondicionLinea = Nothing

        '//Calculamos cada uno de los Vtos.
        Dim dblSumaImpVto As Double
        For Each drRowCondLinea As DataRow In dtCondicionesLinea.Rows
            Dim drNewRow As DataRow = dtVtoAlbaranCompra.NewRow
            Dim dataVto As New NegocioGeneral.dataCalculoFechaVencimiento(data.FechaVencimiento, drRowCondLinea("Periodo"), drRowCondLinea("TipoPeriodo"), data.IDDiaPago, False, data.IDProveedor)
            ProcessServer.ExecuteTask(Of NegocioGeneral.dataCalculoFechaVencimiento)(AddressOf NegocioGeneral.CalcularFechaVencimiento, dataVto, New ServiceProvider)
            drNewRow("FechaVencimiento") = dataVto.FechaVencimiento
            drNewRow("ImpVencimientoA") = xRound(dblTotal * drRowCondLinea("Porcentaje") / 100, data.Decimales)
            dblSumaImpVto = dblSumaImpVto + drNewRow("ImpVencimientoA")
            dtVtoAlbaranCompra.Rows.Add(drNewRow)
        Next

        '//Actualizamos el último ImpVto.
        If Not IsNothing(dtVtoAlbaranCompra) AndAlso dtVtoAlbaranCompra.Rows.Count > 0 Then
            For Each drRowVto As DataRow In dtVtoAlbaranCompra.Select(Nothing, "FechaVencimiento DESC")
                If dblSumaImpVto <> dblTotal Then
                    drRowVto("ImpVencimientoA") = drRowVto("ImpVencimientoA") + dblTotal - dblSumaImpVto
                End If
                Exit For
            Next drRowVto
        End If

        If Not IsNothing(dtCondicionesLinea) Then dtCondicionesLinea = Nothing

        Return dtVtoAlbaranCompra
    End Function

    <Task()> Public Shared Function VtoAlbaranVenta(ByVal data As DataVto, ByVal services As ServiceProvider) As DataTable
        Dim dblTotal As Double
        Dim dblBase As Double
        Dim dblIVA As Double
        Dim dblRE As Double
        Dim dblRecFinan As Double

        '//Calculamos los importes
        dblBase = data.ImpVencimiento
        If data.DtoFactura > 0 Then dblBase = xRound(dblBase * (1 - (CShort(data.DtoFactura) / 100)), data.Decimales)
        If data.FactorDPP > 0 Then dblBase = xRound(dblBase * (1 - (data.FactorDPP / 100)), data.Decimales)
        If data.FactorIVA > 0 Then dblIVA = xRound(dblBase * data.FactorIVA / 100, data.Decimales)
        If data.FactorRE > 0 Then dblRE = xRound(dblBase * data.FactorRE / 100, data.Decimales)
        If data.FactorDPP = 0 And data.FactorRecFinan > 0 Then dblRecFinan = xRound(dblBase * data.FactorRecFinan / 100, data.Decimales)
        dblTotal = dblBase + dblIVA + dblRE + dblRecFinan

        '//Creamos la estructura a retornar.
        Dim dtVtoAlbaranVenta As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf CrearDtVencimientos, Nothing, services)

        '//Calculamos los vencimientos
        Dim objFilter As New Filter
        objFilter.Clear()
        objFilter.Add(New StringFilterItem("IDCondicionPago", data.IDCondicionPago))
        Dim objNegCondicionLinea As New CondicionPagoLinea
        Dim dtCondicionesLinea As DataTable = objNegCondicionLinea.Filter(objFilter)
        objNegCondicionLinea = Nothing

        Dim dblSumaImpVto As Double = 0
        For Each drRowCondLinea As DataRow In dtCondicionesLinea.Rows
            Dim drRowNewVto As DataRow = dtVtoAlbaranVenta.NewRow
            Dim dataVto As New NegocioGeneral.dataCalculoFechaVencimiento(data.FechaVencimiento, drRowCondLinea("Periodo"), drRowCondLinea("TipoPeriodo"), data.IDDiaPago, True, data.IDCliente)
            ProcessServer.ExecuteTask(Of NegocioGeneral.dataCalculoFechaVencimiento)(AddressOf NegocioGeneral.CalcularFechaVencimiento, dataVto, services)
            drRowNewVto("FechaVencimiento") = dataVto.FechaVencimiento
            drRowNewVto("ImpVencimientoA") = xRound(dblTotal * drRowCondLinea("Porcentaje") / 100, data.Decimales)
            dblSumaImpVto = dblSumaImpVto + drRowNewVto("ImpVencimientoA")
            dtVtoAlbaranVenta.Rows.Add(drRowNewVto)
        Next drRowCondLinea

        '//Actualizamos el último ImpVto.
        For Each drRowVto As DataRow In dtVtoAlbaranVenta.Select(Nothing, "FechaVencimiento DESC")
            If dblSumaImpVto <> dblTotal Then
                drRowVto("ImpVencimientoA") = drRowVto("ImpVencimientoA") + dblTotal - dblSumaImpVto
            End If
            Exit For
        Next drRowVto

        If Not IsNothing(dtCondicionesLinea) Then dtCondicionesLinea = Nothing

        Return dtVtoAlbaranVenta
    End Function

    <Task()> Public Shared Function VtoPedidoCompra(ByVal data As DataVto, ByVal services As ServiceProvider) As DataTable
        Dim dblTotal As Double
        Dim dblBase As Double
        Dim dblIVA As Double
        Dim dblRE As Double
        Dim dblRecFinan As Double

        '//Calculamos los importes
        dblBase = data.ImpVencimiento
        If data.DtoFactura > 0 Then dblBase = xRound(dblBase * (1 - (CShort(data.DtoFactura) / 100)), data.Decimales)
        If data.FactorDPP > 0 Then dblBase = xRound(dblBase * (1 - (data.FactorDPP / 100)), data.Decimales)
        If data.FactorIVA > 0 Then dblIVA = xRound(dblBase * data.FactorIVA / 100, data.Decimales)
        If data.FactorRE > 0 Then dblRE = xRound(dblBase * data.FactorRE / 100, data.Decimales)
        If data.FactorDPP = 0 And data.FactorRecFinan > 0 Then dblRecFinan = xRound(dblBase * data.FactorRecFinan / 100, data.Decimales)
        dblTotal = dblBase + dblIVA + dblRE + dblRecFinan

        '//Creamos la estructura a retornar.
        Dim dtVtoPedidoCompra As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf CrearDtVencimientos, Nothing, services)

        '//Calculamos los vencimientos
        Dim objFilter As New Filter
        objFilter.Clear()
        objFilter.Add(New StringFilterItem("IDCondicionPago", data.IDCondicionPago))
        Dim objNegCondicionLinea As New CondicionPagoLinea
        Dim dtCondicionesLinea As DataTable = objNegCondicionLinea.Filter(objFilter)
        objNegCondicionLinea = Nothing

        Dim dblSumaImpVto As Double = 0
        For Each drRowCondLinea As DataRow In dtCondicionesLinea.Rows
            Dim drRowNewVto As DataRow = dtVtoPedidoCompra.NewRow
            Dim dataVto As New NegocioGeneral.dataCalculoFechaVencimiento(data.FechaVencimiento, drRowCondLinea("Periodo"), drRowCondLinea("TipoPeriodo"), data.IDDiaPago, True, data.IDProveedor)
            ProcessServer.ExecuteTask(Of NegocioGeneral.dataCalculoFechaVencimiento)(AddressOf NegocioGeneral.CalcularFechaVencimiento, dataVto, New ServiceProvider)
            drRowNewVto("FechaVencimiento") = dataVto.FechaVencimiento
            drRowNewVto("ImpVencimientoA") = xRound(dblTotal * drRowCondLinea("Porcentaje") / 100, data.Decimales)
            dblSumaImpVto = dblSumaImpVto + drRowNewVto("ImpVencimientoA")
            dtVtoPedidoCompra.Rows.Add(drRowNewVto)
        Next drRowCondLinea

        '//Actualizamos el último ImpVto.
        For Each drRowVto As DataRow In dtVtoPedidoCompra.Select(Nothing, "FechaVencimiento DESC")
            If dblSumaImpVto <> dblTotal Then
                drRowVto("ImpVencimientoA") = drRowVto("ImpVencimientoA") + data.ImpVencimiento - dblSumaImpVto
            End If
            Exit For
        Next drRowVto

        If Not IsNothing(dtCondicionesLinea) Then dtCondicionesLinea = Nothing

        Return dtVtoPedidoCompra
    End Function

    <Task()> Public Shared Function VtoObraHito(ByVal data As DataVto, ByVal services As ServiceProvider) As DataTable
        Dim dblTotal As Double
        Dim dblBase As Double
        Dim dblIVA As Double
        Dim dblRE As Double
        Dim dblRecFinan As Double

        '//Calculamos los importes
        dblBase = data.ImpVencimiento
        If data.DtoFactura > 0 Then dblBase = xRound(dblBase * (1 - (CShort(data.DtoFactura) / 100)), data.Decimales)
        If data.FactorDPP > 0 Then dblBase = xRound(dblBase * (1 - (data.FactorDPP / 100)), data.Decimales)
        If data.FactorIVA > 0 Then dblIVA = xRound(dblBase * data.FactorIVA / 100, data.Decimales)
        If data.FactorRE > 0 Then dblRE = xRound(dblBase * data.FactorRE / 100, data.Decimales)
        If data.FactorDPP = 0 And data.FactorRecFinan > 0 Then dblRecFinan = xRound(dblBase * data.FactorRecFinan / 100, data.Decimales)
        dblTotal = dblBase + dblIVA + dblRE + dblRecFinan

        '//Creamos la estructura a retornar.
        Dim dtVtoObraHito As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf CrearDtVencimientos, Nothing, services)

        '//Calculamos los vencimientos
        Dim objFilter As New Filter
        objFilter.Clear()
        objFilter.Add(New StringFilterItem("IDCondicionPago", data.IDCondicionPago))
        Dim objNegCondicionLinea As New CondicionPagoLinea
        Dim dtCondicionesLinea As DataTable = objNegCondicionLinea.Filter(objFilter)
        objNegCondicionLinea = Nothing

        Dim dblSumaImpVto As Double = 0
        For Each drRowCondLinea As DataRow In dtCondicionesLinea.Rows
            Dim drRowNewVto As DataRow = dtVtoObraHito.NewRow
            Dim dataVto As New NegocioGeneral.dataCalculoFechaVencimiento(data.FechaVencimiento, drRowCondLinea("Periodo"), drRowCondLinea("TipoPeriodo"), data.IDDiaPago, True, data.IDCliente)
            ProcessServer.ExecuteTask(Of NegocioGeneral.dataCalculoFechaVencimiento)(AddressOf NegocioGeneral.CalcularFechaVencimiento, dataVto, New ServiceProvider)
            drRowNewVto("FechaVencimiento") = dataVto.FechaVencimiento
            drRowNewVto("ImpVencimientoA") = xRound(dblTotal * drRowCondLinea("Porcentaje") / 100, data.Decimales)
            dblSumaImpVto = dblSumaImpVto + drRowNewVto("ImpVencimientoA")
            dtVtoObraHito.Rows.Add(drRowNewVto)
        Next drRowCondLinea

        '//Actualizamos el último ImpVto.
        For Each drRowVto As DataRow In dtVtoObraHito.Select(Nothing, "FechaVencimiento DESC")
            If dblSumaImpVto <> dblTotal Then
                drRowVto("ImpVencimientoA") = drRowVto("ImpVencimientoA") + dblTotal - dblSumaImpVto
            End If
            Exit For
        Next drRowVto

        If Not IsNothing(dtCondicionesLinea) Then dtCondicionesLinea = Nothing

        Return dtVtoObraHito
    End Function

    <Task()> Public Shared Function VtoPedidoVenta(ByVal data As DataVto, ByVal services As ServiceProvider) As DataTable
        Dim dblBase As Double
        Dim dblTotal As Double
        Dim dblIVA As Double
        Dim dblRE As Double
        Dim dblRecFinan As Double

        '//Calculamos los importes
        dblBase = data.ImpVencimiento
        If data.DtoFactura > 0 Then dblBase = xRound(dblBase * (1 - (CShort(data.DtoFactura) / 100)), data.Decimales)
        If data.FactorDPP > 0 Then dblBase = xRound(dblBase * (1 - (data.FactorDPP / 100)), data.Decimales)
        If data.FactorIVA > 0 Then dblIVA = xRound(dblBase * data.FactorIVA / 100, data.Decimales)
        If data.FactorRE > 0 Then dblRE = xRound(dblBase * data.FactorRE / 100, data.Decimales)
        If data.FactorDPP = 0 And data.FactorRecFinan > 0 Then dblRecFinan = xRound(dblBase * data.FactorRecFinan / 100, data.Decimales)
        dblTotal = dblBase + dblIVA + dblRE + dblRecFinan

        '//Creamos la estructura a retornar.
        Dim dtVtoPedidoVenta As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf CrearDtVencimientos, Nothing, services)

        '//Calculamos los vencimientos
        Dim objFilter As New Filter
        objFilter.Clear()
        objFilter.Add(New StringFilterItem("IDCondicionPago", data.IDCondicionPago))
        Dim dtCondicionesLinea As DataTable = New CondicionPagoLinea().Filter(objFilter)

        Dim dblSumaImpVto As Double = 0
        For Each drRowCondLinea As DataRow In dtCondicionesLinea.Rows
            Dim drRowNewVto As DataRow = dtVtoPedidoVenta.NewRow
            Dim dataVto As New NegocioGeneral.dataCalculoFechaVencimiento(data.FechaVencimiento, drRowCondLinea("Periodo"), drRowCondLinea("TipoPeriodo"), data.IDDiaPago, True, data.IDCliente)
            ProcessServer.ExecuteTask(Of NegocioGeneral.dataCalculoFechaVencimiento)(AddressOf NegocioGeneral.CalcularFechaVencimiento, dataVto, New ServiceProvider)
            drRowNewVto("FechaVencimiento") = dataVto.FechaVencimiento
            drRowNewVto("ImpVencimientoA") = xRound(dblTotal * drRowCondLinea("Porcentaje") / 100, data.Decimales)
            dblSumaImpVto = dblSumaImpVto + drRowNewVto("ImpVencimientoA")
            dtVtoPedidoVenta.Rows.Add(drRowNewVto)
        Next drRowCondLinea

        '//Actualizamos el último ImpVto.
        For Each drRowVto As DataRow In dtVtoPedidoVenta.Select(Nothing, "FechaVencimiento DESC")
            If dblSumaImpVto <> dblTotal Then
                drRowVto("ImpVencimientoA") = drRowVto("ImpVencimientoA") + dblTotal - dblSumaImpVto
            End If
            Exit For
        Next drRowVto

        If Not IsNothing(dtCondicionesLinea) Then dtCondicionesLinea = Nothing

        Return dtVtoPedidoVenta
    End Function

    <Task()> Public Shared Function VtoObra(ByVal data As DataVto, ByVal services As ServiceProvider) As DataTable
        Dim dblTotal As Double
        Dim dblBase As Double
        Dim dblIVA As Double
        Dim dblRE As Double
        Dim dblRecFinan As Double

        '//Calculamos los importes
        dblBase = data.ImpVencimiento
        If data.DtoFactura > 0 Then dblBase = xRound(dblBase * (1 - (CShort(data.DtoFactura) / 100)), data.Decimales)
        If data.FactorDPP > 0 Then dblBase = xRound(dblBase * (1 - (data.FactorDPP / 100)), data.Decimales)
        If data.FactorIVA > 0 Then dblIVA = xRound(dblBase * data.FactorIVA / 100, data.Decimales)
        If data.FactorRE > 0 Then dblRE = xRound(dblBase * data.FactorRE / 100, data.Decimales)
        If data.FactorDPP = 0 And data.FactorRecFinan > 0 Then dblRecFinan = xRound(dblBase * data.FactorRecFinan / 100, data.Decimales)
        dblTotal = dblBase + dblIVA + dblRE + dblRecFinan

        '//Creamos la estructura a retornar.
        Dim dtVtoObraVenta As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf CrearDtVencimientos, Nothing, services)

        '//Calculamos los vencimientos
        Dim objFilter As New Filter
        objFilter.Clear()
        objFilter.Add(New StringFilterItem("IDCondicionPago", data.IDCondicionPago))
        Dim objNegCondicionLinea As New CondicionPagoLinea
        Dim dtCondicionesLinea As DataTable = objNegCondicionLinea.Filter(objFilter)
        objNegCondicionLinea = Nothing

        Dim dblSumaImpVto As Double = 0
        For Each drRowCondLinea As DataRow In dtCondicionesLinea.Rows
            Dim dtRowNewVto = dtVtoObraVenta.NewRow
            Dim dataVto As New NegocioGeneral.dataCalculoFechaVencimiento(data.FechaVencimiento, drRowCondLinea("Periodo"), drRowCondLinea("TipoPeriodo"), data.IDDiaPago, True, data.IDCliente)
            ProcessServer.ExecuteTask(Of NegocioGeneral.dataCalculoFechaVencimiento)(AddressOf NegocioGeneral.CalcularFechaVencimiento, dataVto, New ServiceProvider)
            dtRowNewVto("FechaVencimiento") = dataVto.FechaVencimiento
            dtRowNewVto("ImpVencimientoA") = xRound(dblTotal * drRowCondLinea("Porcentaje") / 100, data.Decimales)
            dblSumaImpVto = dblSumaImpVto + dtRowNewVto("ImpVencimientoA")
            dtVtoObraVenta.Rows.Add(dtRowNewVto)
        Next drRowCondLinea

        '//Actualizamos el último ImpVto.
        For Each drRowVto As DataRow In dtVtoObraVenta.Select(Nothing, "FechaVencimiento DESC")
            If dblSumaImpVto <> dblTotal Then
                drRowVto("ImpVencimientoA") = drRowVto("ImpVencimientoA") + dblTotal - dblSumaImpVto
            End If
            Exit For
        Next drRowVto

        If Not IsNothing(dtCondicionesLinea) Then dtCondicionesLinea = Nothing

        Return dtVtoObraVenta
    End Function


    <Task()> Public Shared Function CrearDtVencimientos(ByVal data As Object, ByVal services As ServiceProvider) As DataTable
        Dim dtVtos As New DataTable
        With dtVtos.Columns
            .Add("ImpVencimientoA", GetType(Double))
            .Add("FechaVencimiento", GetType(Date))
        End With
        Return dtVtos
    End Function

    <Task()> Public Shared Function RecuperarDiasRetraso(ByVal IDCliente As String, ByVal services As ServiceProvider) As Double
        Dim dblDiasRetraso As Double
        Dim objNegCliente As New Cliente
        Dim dtCliente As DataTable = objNegCliente.SelOnPrimaryKey(IDCliente & String.Empty)
        objNegCliente = Nothing
        If Not IsNothing(dtCliente) AndAlso dtCliente.Rows.Count > 0 Then
            dblDiasRetraso = Nz(dtCliente.Rows(0)("DiasRetraso"), 0)
            dtCliente.Rows.Clear()
        End If
        Return dblDiasRetraso
    End Function

#End Region

#Region "Eventos OLAP"

    Private Sub CrearSimulacionOLAP(ByVal dtSimulacionTesoreria As DataTable)
        '//Vaciamos la tabla tbSimulacionOlap.
        DeleteSimulacionOlap()

        '//Traspasamos la Sim.Tesorería a Sim.OLAP
        If Not RellenarSimulacionOlap(dtSimulacionTesoreria) Then
            ApplicationService.GenerateError("Ha ocurrido un error al generar la Simulación OLAP.")
        End If
    End Sub

    Public Overloads Function DeleteSimulacionOlap() As Boolean
        DeleteSimulacionOlap = False
        Dim strProc As String = "DELETE FROM tbSimulacionOlap"
        AdminData.Execute(strProc)
        DeleteSimulacionOlap = True
    End Function

    Public Function RellenarSimulacionOlap(ByVal dtSimTesoreria As DataTable) As Boolean
        RellenarSimulacionOlap = False

        Dim dtSimulacionOLAP As DataTable
        For Each drRowSimTesoreria As DataRow In dtSimTesoreria.Rows
            dtSimulacionOLAP = AdminData.Execute("SELECT * FROM " & "tbSimulacionOlap" & " WHERE 1=2", ExecuteCommand.ExecuteReader)

            Dim drRowSimOLAP As DataRow = dtSimulacionOLAP.NewRow
            drRowSimOLAP("IDSimulacion") = AdminData.GetAutoNumeric
            drRowSimOLAP("Parametro") = drRowSimTesoreria("ParametrosSimulacion")
            drRowSimOLAP("TipoDocumento") = drRowSimTesoreria("EnumTipoDocumento")
            drRowSimOLAP("IdOrigen") = drRowSimTesoreria("IdOrigen")
            drRowSimOLAP("IdClienteProveedor") = drRowSimTesoreria("IdClienteProveedor")
            drRowSimOLAP("DescClienteProveedor") = drRowSimTesoreria("DescClienteProveedor")
            drRowSimOLAP("Importe") = Nz(drRowSimTesoreria("Importe"), 0)
            drRowSimOLAP("Fecha") = drRowSimTesoreria("Fecha")
            drRowSimOLAP("IdBancoPropio") = drRowSimTesoreria("IdBancoPropio")
            drRowSimOLAP("IDAgrupacion") = drRowSimTesoreria("IDAgrupacion")
            drRowSimOLAP("Acumulado") = 0
            'drRowSimOLAP("empresa") = drRowAux("empresa")

            dtSimulacionOLAP.Rows.Add(drRowSimOLAP)
        Next drRowSimTesoreria

        BusinessHelper.UpdateTable(dtSimulacionOLAP)
        RellenarSimulacionOlap = True

    End Function

#End Region

#Region "Datos para informes"

    'Los datos que le llegan a esta función son:
    'Párametros para cálculo de simulación: pFechaDesde,pFechaHasta,pblnBancoPropio, 
    '                                       pblnFactura,pblnAlbaran,pblnPedido,
    '                                       pblnObra
    'Filtros provenientes de la pantalla: Un objeto filter.
    'Información de períodos para el cálculo del informe: plngPeriodo1,plngPeriodo2,
    '                                       plngPeriodo3,plngPeriodo4

    <Serializable()> _
    Public Class DatosInformeTes
        Public FechaDesde As Date
        Public FechaHasta As Date
        Public FechaValidez As Date
        Public BancoPropio As Boolean
        Public Factura As Boolean
        Public Albaran As Boolean
        Public Pedido As Boolean
        Public Obra As Boolean
        Public Promotoras As Boolean
        Public Fil As Filter
        Public IntPeriodo1 As Integer
        Public IntPeriodo2 As Integer
        Public IntPeriodo3 As Integer
        Public IntPeriodo4 As Integer
        Public IntTipoInforme As Integer

        Public Sub New()
        End Sub
        Public Sub New(ByVal FechaDesde As Date, ByVal FechaHasta As Date, ByVal FechaValidez As Date, ByVal BancoPropio As Boolean, ByVal Factura As Boolean, ByVal Albaran As Boolean, _
                       ByVal Pedido As Boolean, ByVal Obra As Boolean, ByVal Promotoras As Boolean, ByVal Fil As Filter, ByVal IntPeriodo1 As Integer, ByVal IntPeriodo2 As Integer, _
                       ByVal IntPeriodo3 As Integer, ByVal IntPeriodo4 As Integer, ByVal IntTipoInforme As Integer)
            Me.FechaDesde = FechaDesde
            Me.FechaHasta = FechaHasta
            Me.FechaValidez = FechaValidez
            Me.BancoPropio = BancoPropio
            Me.Factura = Factura
            Me.Albaran = Albaran
            Me.Pedido = Pedido
            Me.Obra = Obra
            Me.Promotoras = Promotoras
            Me.Fil = Fil
            Me.IntPeriodo1 = IntPeriodo1
            Me.IntPeriodo2 = IntPeriodo2
            Me.IntPeriodo3 = IntPeriodo3
            Me.IntPeriodo4 = IntPeriodo4
            Me.IntTipoInforme = IntTipoInforme
        End Sub
    End Class

    <Task()> Public Shared Function DatosInforme(ByVal data As DatosInformeTes, ByVal services As ServiceProvider) As DataTable
        'Tipo de informe: 1->Previsión de cobros
        '                -1->Antigüedad de la deuda
        'Lanzamos el cálculo de simulación de tesoreria
        Dim StData As New DataSimulTes(data.FechaDesde, data.FechaHasta, data.FechaValidez, data.BancoPropio, _
                                        data.Factura, data.Albaran, data.Pedido, data.Obra, data.Promotoras)
        Dim dtSimulacion As DataTable = ProcessServer.ExecuteTask(Of DataSimulTes, DataTable)(AddressOf SimulacionTesoreria, StData, services)
        If data.IntTipoInforme < 1 Then
            'Construimos un filtro para obtener solo cobros o cobros periódicos
            Dim objFilterCobroOR As New Filter(FilterUnionOperator.Or)
            objFilterCobroOR.Add(New NumberFilterItem("EnumTipoDocumento", enumSimulacionTesoreria.Cobro))
            objFilterCobroOR.Add(New NumberFilterItem("EnumTipoDocumento", enumSimulacionTesoreria.CobroPeriodico))
            data.Fil.Add(objFilterCobroOR)
        End If
        'Filtramos con el objeto Filter
        Dim Where As String = data.Fil.Compose(New AdoFilterComposer)
        Dim drSimulacion() As DataRow = dtSimulacion.Select(Where)
        Dim dtPrevision As DataTable = New NegocioGeneral().CrearDTExportacionCuentas
        If Not drSimulacion Is Nothing Then
            If drSimulacion.Length() > 0 Then
                For Each dr As DataRow In drSimulacion
                    Dim drAlta As DataRow = dtPrevision.NewRow
                    drAlta("EnumTipoDocumento") = dr("EnumTipoDocumento")
                    drAlta("ID") = dr("IDClienteProveedor")
                    drAlta("Descripcion") = dr("DescClienteProveedor")
                    drAlta("NumEfectos") = 1
                    drAlta("InferiorPeriodo") = 0 : drAlta("Periodo1") = 0
                    drAlta("Periodo2") = 0 : drAlta("Periodo3") = 0
                    drAlta("Periodo4") = 0 : drAlta("SuperiorPeriodo") = 0
                    'drAlta("Importe") = dr("Importe")
                    If data.IntTipoInforme < 1 Then
                        'Antigüedad de la deuda
                        If dr("Fecha") > Today.Date Then
                            'El efecto no está vencido
                            drAlta("InferiorPeriodo") = dr("importe")
                        ElseIf dr("Fecha") <= DateAdd(DateInterval.Day, -data.IntPeriodo4, Today.Date) Then
                            'El efecto ha vencido superando el periodo 4
                            drAlta("SuperiorPeriodo") = dr("importe")
                        ElseIf dr("Fecha") > DateAdd(DateInterval.Day, -data.IntPeriodo4, Today.Date) And dr("Fecha") <= DateAdd(DateInterval.Day, -data.IntPeriodo3, Today.Date) Then
                            'El efecto se encuentra entre el periodo 3 y el periodo 4
                            drAlta("Periodo4") = dr("importe")
                        ElseIf dr("Fecha") > DateAdd(DateInterval.Day, -data.IntPeriodo3, Today.Date) And dr("Fecha") <= DateAdd(DateInterval.Day, -data.IntPeriodo2, Today.Date) Then
                            'El efecto se encuentra entre el periodo 2 y el periodo 3
                            drAlta("Periodo3") = dr("importe")
                        ElseIf dr("Fecha") > DateAdd(DateInterval.Day, -data.IntPeriodo2, Today.Date) And dr("Fecha") <= DateAdd(DateInterval.Day, -data.IntPeriodo1, Today.Date) Then
                            'El efecto se encuentra entre el periodo 1 y el periodo 2
                            drAlta("Periodo2") = dr("importe")
                        Else
                            'El efecto se encuentra en el periodo 1
                            drAlta("Periodo1") = dr("importe")
                        End If
                    Else
                        'Previsión de tesorería
                        If dr("Fecha") < Today.Date Then
                            'El efecto esta vencido
                            drAlta("InferiorPeriodo") = dr("Importe")
                        ElseIf DateAdd(DateInterval.Day, -data.IntPeriodo1, dr("Fecha")) < Today.Date Then
                            drAlta("Periodo1") = dr("Importe")
                        ElseIf DateAdd(DateInterval.Day, -data.IntPeriodo2, dr("Fecha")) < Today.Date Then
                            drAlta("Periodo2") = dr("importe")
                        ElseIf DateAdd(DateInterval.Day, -data.IntPeriodo3, dr("Fecha")) < Today.Date Then
                            drAlta("Periodo3") = dr("importe")
                        ElseIf DateAdd(DateInterval.Day, -data.IntPeriodo4, dr("Fecha")) < Today.Date Then
                            drAlta("Periodo4") = dr("importe")
                        Else
                            drAlta("SuperiorPeriodo") = dr("importe")
                        End If
                    End If
                    dtPrevision.Rows.Add(drAlta)
                Next
            End If
        End If
        Return dtPrevision
    End Function

    <Task()> Public Shared Function CrearDTPrevisionTesoreria(ByVal data As Object, ByVal services As ServiceProvider) As DataTable
        Dim dt As New DataTable
        dt.Columns.Add("EnumTipoDocumento", GetType(Integer))
        dt.Columns.Add("ID", GetType(String))
        dt.Columns.Add("Descripcion", GetType(String))
        dt.Columns.Add("NumEfectos", GetType(Integer))
        dt.Columns.Add("InferiorPeriodo", GetType(Double))
        dt.Columns.Add("Periodo1", GetType(Double))
        dt.Columns.Add("Periodo2", GetType(Double))
        dt.Columns.Add("Periodo3", GetType(Double))
        dt.Columns.Add("Periodo4", GetType(Double))
        dt.Columns.Add("SuperiorPeriodo", GetType(Double))
        dt.Columns.Add("Importe", GetType(Double))
        Return dt
    End Function

#End Region

End Class