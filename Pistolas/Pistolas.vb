<Serializable()> _
Public Class Pistolas

    Private Const cn_PEDIDOSCOMPRA_VIEW As String = "vNegPistolasLineasPCPendientes"
    Private Const cn_Almacen_VIEW As String = "vPistolasInventario"

#Region "Recepciones"

    Public Function CrearAlbaranCompraDesdePistola(ByVal pedidosCompra() As DataPistolas.PedidoCompraPistolas_Info, ByVal SuNAlbaran As String, ByVal SuFechaAlbaran As Date) As DataPistolas.ActionResultPistolas

        Dim Result As New DataPistolas.ActionResultPistolas
        Dim aud As AlbaranCompraUpdateData
        If Not IsNothing(pedidosCompra) AndAlso pedidosCompra.Length > 0 Then
            Dim ac_i(pedidosCompra.Length - 1) As CrearAlbaranCompraInfo
            For i As Integer = 0 To pedidosCompra.Length - 1
                ac_i(i) = ObtenerInformacionLinea(pedidosCompra(i))
                ac_i(i).SuAlbaran = SuNAlbaran
                ac_i(i).SuFecha = SuFechaAlbaran
            Next
            If ac_i.Length > 0 Then
                Try
                    'MAP PENDIENTE 
                    ''''aud = New AlbaranCompraCabecera().CrearAlbaranCompra(ac_i)

                Catch e As Exception
                    Result.OK = False
                    Result.Message = e.Message
                End Try
                If Not aud Is Nothing AndAlso aud.IDAlbaran.Length > 0 Then
                    Result.OK = True
                    Dim albaranes As String = String.Join(",", aud.NAlbaran)
                    If aud.IDAlbaran.Length = 1 Then
                        Result.Message = "Se ha generado el albarán " & albaranes
                    Else
                        Result.Message = "Se han generado los albaranes " & albaranes
                    End If
                End If
            End If
        End If
        Return Result
    End Function

    Private Function ObtenerInformacionLinea(ByVal pc_i As DataPistolas.PedidoCompraPistolas_Info) As CrearAlbaranCompraInfo
        If Not IsNothing(pc_i) Then
            Dim ac_i As New CrearAlbaranCompraInfo(pc_i.IDLineaPedido)
            ac_i.IDPedido = pc_i.IDPedido
            ac_i.IDProveedor = pc_i.IDProveedor
            ac_i.IDMoneda = pc_i.IDMoneda
            ac_i.FechaEntregaModificado = Today.Date
            ac_i.Cantidad = pc_i.QServida

            Return ac_i
        End If
    End Function

#End Region

#Region "Expediciones"
    Public Function CrearAlbaranVenta(ByVal preparacion() As DataPistolas.Preparacion_Info) As String

        Dim Result As String = String.Empty

        If Not IsNothing(preparacion) AndAlso preparacion.Length > 0 Then
            Dim AVC As New AlbaranVentaCabecera
            Dim AV_Info(preparacion.Length - 1) As CrearAlbaranVentaInfo2
            For i As Integer = 0 To preparacion.Length - 1
                AV_Info(i) = pasarAAlbaranVentaInfo(preparacion(i))
            Next
            If AV_Info.Length > 0 Then

                'MAP PENDIENTE
                ''''Dim albaran() As AlbaranVentaUpdateData = AVC.CrearAlbaranVentaDesdePistola(AV_Info, Date.Today)
                ''''If albaran(0).IDAlbaran.Length > 0 Then
                ''''    Result = albaran(0).IDAlbaran(0)
                ''''Else
                ''''    Result = String.Empty
                ''''End If
            End If
        End If

        Return Result
    End Function

    Public Function ObtenerContador(ByVal Entity As String) As String

        Dim strResultado As String = String.Empty

        'MAP PENDIENTE
        ''''Dim strIDContador As String = New CentroGestion().GetContadorPredeterminado(CentroGestion.ContadorEntidad.AlbaranVenta)
        ''''If Length(strIDContador) > 0 Then
        ''''    strResultado = New Contador().CounterValueTx(strIDContador).strCounterValue
        ''''Else
        ''''    strResultado = String.Empty
        ''''End If

        Return strResultado
    End Function

#End Region

#Region "Inventarios"

    <Task()> Public Shared Function InventarioDesdePistolas(ByVal Almacenes() As DataPistolas.Almacen_Info, ByVal services As ServiceProvider) As DataPistolas.ActionResultPistolas

        Dim Result As New DataPistolas.ActionResultPistolas
        Dim data(-1) As StockData

        For i As Integer = 0 To Almacenes.Length - 1
            Dim stkData As New StockData(Almacenes(i).IDArticulo, Almacenes(i).Almacen, Almacenes(i).Stock, 0, 0, Today.Date, enumTipoMovimiento.tmInventario)
            ReDim Preserve data(UBound(data) + 1)
            data(UBound(data)) = stkData
        Next

        Result.OK = True
        Result.Message = "Inventario Terminado."

        Try
            'MAP REVISAR
            ''''Dim stk As New Stock
            ''''stk.Inventario(data)

            Dim datInv As New DataTratarStocks(data)
            Dim updateData As StockUpdateData() = ProcessServer.ExecuteTask(Of DataTratarStocks, StockUpdateData())(AddressOf ProcesoStocks.InventarioMasivo, datInv, services)

            Dim Errores(-1) As StockUpdateData
            If Not updateData Is Nothing AndAlso updateData.Length > 0 Then
                For Each upd As StockUpdateData In updateData
                    If upd.Estado = EstadoStock.NoActualizado Then
                        ReDim Preserve Errores(Errores.Length)
                        Errores(Errores.Length - 1) = upd
                    End If
                Next
            End If
            If Errores.Length > 0 AndAlso updateData.Length <> Errores.Length Then
                Result.OK = False
                Result.Message = "Se han producido errores en el Inventario."
                updateData = Errores
            End If

        Catch e As Exception
            Result.OK = False
            Result.Message = e.Message
        End Try

        Return Result
    End Function

#End Region

#Region "Funciones Publicas"

    Public Function RecuperarPedidosCompraPorNPedido(ByVal strNPedido() As String) As DataPistolas.PedidoCompraPistolas_Info()
        If strNPedido.Length > 0 Then
            Dim PedidosCompra() As DataPistolas.PedidoCompraPistolas_Info
            For i As Integer = 0 To strNPedido.Length - 1
                Dim dt As DataTable = New PedidoCompraCabecera().Filter(New StringFilterItem("NPedido", FilterOperator.Equal, strNPedido(i)))
                If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
                    Dim Lineas() As DataPistolas.PedidoCompraPistolas_Info = ObtenerLineasPedidoCompra(dt.Rows(0)("IDPedido"), strNPedido(i))
                    UnirArray(PedidosCompra, Lineas)
                End If
            Next
            Return PedidosCompra
        End If

    End Function

    Public Function RecuperarPedidosCompra(ByVal strIDArticulo As String, ByVal strIDProveedor As String) As DataPistolas.PedidoCompraPistolas_Info()
        '//Recuperar por IDArticulo e IDProveedor las líneas de pedidos de compra pendientes.
        Dim PedidosCompra(-1) As DataPistolas.PedidoCompraPistolas_Info
        Dim _Filter As New Filter
        _Filter.Add("IDProveedor", FilterOperator.Equal, strIDProveedor, FilterType.String)
        _Filter.Add("IDArticulo", FilterOperator.Equal, strIDArticulo, FilterType.String)
        Dim dt As DataTable = New BE.DataEngine().Filter(cn_PEDIDOSCOMPRA_VIEW, _Filter, "IDPedido,NPedido,IDLineaPedido,IDArticulo,QPendiente,DescArticulo,Orden,FechaEntrega,IDProveedor,IDMoneda", "NPedido,Orden ASC")
        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
            ReDim PedidosCompra(dt.Rows.Count - 1)
            For i As Integer = 0 To dt.Rows.Count - 1
                PedidosCompra(i) = New DataPistolas.PedidoCompraPistolas_Info(dt.Rows(i)("IDPedido"), dt.Rows(i)("NPedido"), dt.Rows(i)("IDLineaPedido"), dt.Rows(i)("IDArticulo"), dt.Rows(i)("QPendiente"), dt.Rows(i)("FechaEntrega"), dt.Rows(i)("IDProveedor"), dt.Rows(i)("DescArticulo"))
            Next
        End If
        Return PedidosCompra
    End Function

    Public Function RecuperarInventario(ByVal strIDArticulo As String) As DataPistolas.Almacen_Info()
        '//Recuperar por IDArticulo e IDProveedor las líneas de pedidos de compra pendientes.
        Dim Almacen(-1) As DataPistolas.Almacen_Info
        Dim _Filter As New Filter
        _Filter.Add("IDArticulo", FilterOperator.Equal, strIDArticulo, FilterType.String)
        Dim dt As DataTable = New BE.DataEngine().Filter(cn_Almacen_VIEW, _Filter, "Almacen,Descripcion,Stock,UltInvent,Inventariado,IDArticulo", "Almacen ASC")
        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
            ReDim Almacen(dt.Rows.Count - 1)
            For i As Integer = 0 To dt.Rows.Count - 1
                Almacen(i) = New DataPistolas.Almacen_Info(dt.Rows(i)("Almacen") & String.Empty, dt.Rows(i)("Descripcion") & String.Empty, dt.Rows(i)("Stock") & String.Empty, Nz(dt.Rows(i)("UltInvent"), Date.MinValue), dt.Rows(i)("Inventariado") & String.Empty, dt.Rows(i)("IDArticulo") & String.Empty)
            Next
        End If
        Return Almacen
    End Function

    Public Function pasarAAlmacen_Info(ByVal Almacen As String, ByVal Descripcion As String, ByVal Stock As Double, ByVal UltInvent As String, ByVal Inventariado As String, ByVal IDArticulo As String) As DataPistolas.Almacen_Info
        Dim Almacendato As DataPistolas.Almacen_Info
        Almacendato = New DataPistolas.Almacen_Info(Almacen, Descripcion, Stock, UltInvent, Inventariado, IDArticulo)
        Return Almacendato
    End Function

    Public Function EncuentraArticulo(ByVal strCodigoBarras As String) As DataPistolas.ExisteArticulo

        Dim retorno As DataPistolas.ExisteArticulo
        Dim f As New Filter(FilterUnionOperator.Or)
        Dim f1 As New Filter
        Dim f2 As New Filter
        f.Add(New StringFilterItem("CodigoBarras", strCodigoBarras))
        f.Add(New StringFilterItem("IDArticulo", strCodigoBarras))

        Dim dt As DataTable = New Articulo().Filter(f)
        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
            For Each drFila As DataRow In dt.Rows ' Cambio en la QEmbalaje sino tiene dato o es 0 hay q poner un 1
                If Length(drFila("QEmbalaje")) = 0 OrElse drFila("QEmbalaje") = 0 Then
                    drFila("QEmbalaje") = 1
                End If
            Next
            retorno = New DataPistolas.ExisteArticulo(True, dt.Rows(0)("IDArticulo"), dt.Rows(0)("DescArticulo"), dt.Rows(0)("CodigoBarras") & String.Empty, dt.Rows(0)("QEmbalaje"))
        Else
            retorno = New DataPistolas.ExisteArticulo
        End If

        Return retorno
    End Function

    Public Function ProveedoresPedidosPendientes(ByVal strIDArticulo As String) As DataPistolas.Proveedor_Info()

        Dim ProveedoresEncontrados() As DataPistolas.Proveedor_Info
        Dim dt As DataTable = New BE.DataEngine().Filter("vPistolasProveedoresPCPendientes", New StringFilterItem("IDArticulo", strIDArticulo))
        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
            Dim i As Integer = 0
            ReDim ProveedoresEncontrados(dt.Rows.Count - 1)
            For Each drLinea As DataRow In dt.Rows
                ProveedoresEncontrados(i) = New DataPistolas.Proveedor_Info(drLinea("IDProveedor"), drLinea("DescProveedor"))
                i = i + 1
            Next
        End If

        Return ProveedoresEncontrados
    End Function

    Public Function EncuentraPedidos(ByVal PedidosCompra() As DataPistolas.PedidoCompraPistolas_Info, ByVal strIDArticulo As String) As DataPistolas.PedidoCompraPistolas_Info()

        Dim PedidosEncontrados() As DataPistolas.PedidoCompraPistolas_Info
        For Each pc_i As DataPistolas.PedidoCompraPistolas_Info In PedidosCompra
            If strIDArticulo.Compare(strIDArticulo, pc_i.IDArticulo) = 0 Then
                If IsNothing(PedidosEncontrados) Then
                    ReDim PedidosEncontrados(0)
                Else
                    ReDim Preserve PedidosEncontrados(PedidosEncontrados.Length)
                End If
                PedidosEncontrados(PedidosEncontrados.Length - 1) = New DataPistolas.PedidoCompraPistolas_Info(pc_i.IDArticulo, pc_i.NPedido, pc_i.IDLineaPedido, pc_i.IDArticulo, pc_i.QPendiente, pc_i.DescArticulo)
            End If
        Next
        Return PedidosEncontrados

    End Function

    Public Function EncuentraProveedor(ByVal IDProveedor As String) As DataPistolas.ExisteProveedor

        Dim retorno As DataPistolas.ExisteProveedor
        If Not IsNothing(IDProveedor) AndAlso IDProveedor.Length > 0 Then
            Dim dt As DataTable = New Proveedor().SelOnPrimaryKey(IDProveedor)
            If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
                retorno = New DataPistolas.ExisteProveedor(True, dt.Rows(0)("DescProveedor"))
            Else
                retorno = New DataPistolas.ExisteProveedor
            End If
        End If

        Return retorno
    End Function

    Public Function EsPedidoCompraValido(ByVal strNPedido As String) As Boolean

        If strNPedido.Length > 0 Then
            Dim dt As DataTable = New PedidoCompraCabecera().Filter(New StringFilterItem("NPedido", strNPedido))
            If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
                Return True
            Else
                Return False
            End If
        End If

    End Function

    Public Function ObtenerLineasPedidoCompra(ByVal intIDPedido As Integer, ByVal strNPedido As String) As DataPistolas.PedidoCompraPistolas_Info()

        Dim objFilter As New Filter
        objFilter.Add("IDPedido", FilterOperator.Equal, intIDPedido, FilterType.Numeric)
        objFilter.Add("Estado", FilterOperator.NotEqual, enumpclEstado.pclCerrado, FilterType.Numeric)
        objFilter.Add("Estado", FilterOperator.NotEqual, enumpclEstado.pclservido, FilterType.Numeric)
        Dim dtLineas As DataTable = New PedidoCompraLinea().Filter("IDPedido,IDLineaPedido,IDArticulo,DescArticulo,QPedida,QServida", objFilter.Compose(New AdoFilterComposer), "IDLineaPedido")
        If Not IsNothing(dtLineas) AndAlso dtLineas.Rows.Count > 0 Then
            Dim i As Integer = 0
            Dim LineasPedido(dtLineas.Rows.Count - 1) As DataPistolas.PedidoCompraPistolas_Info
            For Each drLinea As DataRow In dtLineas.Rows
                LineasPedido(i) = New DataPistolas.PedidoCompraPistolas_Info(drLinea("IDPedido"), strNPedido, drLinea("IDLineaPedido"), drLinea("IDArticulo"), drLinea("QPedida") - drLinea("QServida"), , , drLinea("DescArticulo"))
                i = i + 1
            Next
            Return LineasPedido
        End If
    End Function

    Public Sub UnirArray(ByRef PedidosCompra() As DataPistolas.PedidoCompraPistolas_Info, ByVal LineasAdd() As DataPistolas.PedidoCompraPistolas_Info)
        If Not IsNothing(PedidosCompra) AndAlso PedidosCompra.Length > 0 Then
            If Not IsNothing(LineasAdd) AndAlso LineasAdd.Length > 0 Then
                Dim length0 As Integer = PedidosCompra.Length
                ReDim Preserve PedidosCompra(length0 + LineasAdd.Length - 1)
                '//Contador para LineasAdd
                Dim j As Integer = 0
                For i As Integer = length0 To PedidosCompra.Length - 1
                    PedidosCompra(i) = LineasAdd(j)
                    j += 1
                Next
            End If
        Else
            If Not IsNothing(LineasAdd) AndAlso LineasAdd.Length > 0 Then PedidosCompra = LineasAdd
        End If
    End Sub

    Public Function pasarAAlbaranVentaInfo(ByVal preparacion As DataPistolas.Preparacion_Info) As CrearAlbaranVentaInfo2

        Dim AVinfo As New CrearAlbaranVentaInfo2

        Dim filtro As New Filter
        filtro.Add("IDLineaPreparacion", FilterOperator.Equal, preparacion.IDLineaPreparacion)
        Dim dt As DataTable = New BE.DataEngine().Filter("frmMntoExpediciones", filtro)

        AVinfo.Cantidad = preparacion.QExpedir
        AVinfo.IDCliente = preparacion.Cliente
        'MAP REVISAR
        ''''AVinfo.IDLineaPreparacion = preparacion.IDLineaPreparacion
        AVinfo.IDLinea = preparacion.IDLineaPedido
        AVinfo.IDProveedorServicios = String.Empty
        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
            AVinfo.IDPedido = dt.Rows(0)("IDPedido")
            AVinfo.PedidoCliente = dt.Rows(0)("PedidoCliente") & String.Empty
            If IsDate(dt.Rows(0)("FechaEntrega")) Then
                AVinfo.FechaEntregaModificado = dt.Rows(0)("FechaEntrega")
            Else
                AVinfo.FechaEntregaModificado = Date.Today
            End If
            'MAP REVISAR
            ''''If Length(dt.Rows(0)("IDFormaEnvioTransporte") & String.Empty) > 0 Then
            ''''    AVinfo.IDFormaEnvio = dt.Rows(0)("IDFormaEnvioTransporte")
            ''''Else
            ''''    AVinfo.IDFormaEnvio = dt.Rows(0)("IDFormaEnvio")
            ''''End If
            ''''AVinfo.FechaDescarga = Nz(dt.Rows(0)("FechaPrevistaPreparacion"), Date.Today)
            ''''AVinfo.Matricula = dt.Rows(0)("Matricula") & String.Empty
            ''''AVinfo.Remolque = dt.Rows(0)("Remolque") & String.Empty
            ''''AVinfo.Conductor = dt.Rows(0)("Conductor") & String.Empty
            ''''AVinfo.DNIConductor = dt.Rows(0)("NIFConductor") & String.Empty
            ''''AVinfo.HoraLlegada = Nz(dt.Rows(0)("HoraLlegada"), "0:00:00")
            ''''AVinfo.HoraSalida = Nz(dt.Rows(0)("HoraSalida"), "0:00:00")
            ''''AVinfo.IDTipoCamion = dt.Rows(0)("IDTipoCamion") & String.Empty
            ''''AVinfo.IDLineaPreparacion = dt.Rows(0)("IDLineaPreparacion")
            ''''AVinfo.Transportista = dt.Rows(0)("Transportista") & String.Empty
            ''''AVinfo.IDTipoAlbaran = dt.Rows(0)("IDTipoAlbaran") & String.Empty
            ''''AVinfo.IDContador = dt.Rows(0)("IDContador") & String.Empty
        End If
        Return AVinfo
    End Function

    Public Function ComprobarPreparacion(ByVal sPreparacion As String) As DataPistolas.Preparacion_Info()

        Dim PreparacionDatos() As DataPistolas.Preparacion_Info

        Dim dt As DataTable = New BE.DataEngine().Filter("vPistolasTransporteLineas", New StringFilterItem("IDPreparacion", sPreparacion))
        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
            For Each drFila As DataRow In dt.Rows ' Cambio en la QEmbalaje sino tiene dato o es 0 hay q poner un 1
                If Length(drFila("QEmbalaje")) = 0 OrElse drFila("QEmbalaje") = 0 Then
                    drFila("QEmbalaje") = 1
                End If
            Next
            Dim i As Integer = 0
            ReDim PreparacionDatos(dt.Rows.Count - 1)
            For Each drLinea As DataRow In dt.Rows
                If drLinea("Estado") = enumtvcEstadoTransporte.tvcEnCurso Then
                    PreparacionDatos(i) = New DataPistolas.Preparacion_Info(drLinea("IDPreparacion") & String.Empty, "EnCurso" & String.Empty, drLinea("FechaPrevista") & String.Empty, drLinea("DescCliente") & String.Empty, drLinea("ProvinciaEnvio") & String.Empty, drLinea("IDLineaPedido") & String.Empty, drLinea("IDLineaPreparacion") & String.Empty, drLinea("NPedido") & String.Empty, drLinea("IDArticulo") & String.Empty, drLinea("DescArticulo") & String.Empty, drLinea("Cantidad") & String.Empty, drLinea("QExpedir") & String.Empty, drLinea("CodigoBarras") & String.Empty, drLinea("QEmbalaje") & String.Empty, drLinea("DescCliente") & String.Empty, drLinea("ProvinciaEnvio") & String.Empty)
                Else
                    PreparacionDatos(i) = New DataPistolas.Preparacion_Info(drLinea("IDPreparacion") & String.Empty, "Cerrado" & String.Empty, drLinea("FechaPrevista") & String.Empty, drLinea("DescCliente") & String.Empty, drLinea("ProvinciaEnvio") & String.Empty, drLinea("IDLineaPedido") & String.Empty, drLinea("IDLineaPreparacion") & String.Empty, drLinea("NPedido") & String.Empty, drLinea("IDArticulo") & String.Empty, drLinea("DescArticulo") & String.Empty, drLinea("Cantidad") & String.Empty, drLinea("QExpedir") & String.Empty, drLinea("CodigoBarras") & String.Empty, drLinea("QEmbalaje") & String.Empty, drLinea("DescCliente") & String.Empty, drLinea("ProvinciaEnvio") & String.Empty)
                End If
                i = i + 1
            Next
            Return PreparacionDatos
        End If


    End Function

    Public Function pasarAPreparacionInfo(ByVal IDPreparacion As String, ByVal Estado As String, ByVal FechaPrevista As String, ByVal DescCliente As String, ByVal ProvinciaEnvio As String, ByVal IDLineaPedido As String, ByVal IDLineaPreparacion As String, ByVal NPedido As String, ByVal IDArticulo As String, ByVal DescArticulo As String, ByVal Cantidad As String, ByVal QExpedir As String, ByVal CodigoBarras As String, ByVal QEmbalaje As Double) As DataPistolas.Preparacion_Info
        Dim PreparacionDato As DataPistolas.Preparacion_Info
        If Estado = "Cerrado" Then
            PreparacionDato = New DataPistolas.Preparacion_Info(IDPreparacion, 1, FechaPrevista, DescCliente, ProvinciaEnvio, IDLineaPedido, IDLineaPreparacion, NPedido, IDArticulo, DescArticulo, Cantidad, QExpedir, CodigoBarras, QEmbalaje, DescCliente, ProvinciaEnvio)
        Else
            PreparacionDato = New DataPistolas.Preparacion_Info(IDPreparacion, 0, FechaPrevista, DescCliente, ProvinciaEnvio, IDLineaPedido, IDLineaPreparacion, NPedido, IDArticulo, DescArticulo, Cantidad, QExpedir, CodigoBarras, QEmbalaje, DescCliente, ProvinciaEnvio)
        End If
        Return PreparacionDato
    End Function

#End Region

End Class
