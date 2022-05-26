Public Class ProcesoAlbaranVentaPedidos

#Region " Agrupación de Pedidos "

    <Serializable()> _
    Public Class DataColAgrupacionAV
        Public Lineas As DataTable
        Public TipoAgrupacion As enummcAgrupAlbaran
    End Class

    <Task()> Public Shared Function GetGroupColumns(ByVal data As DataColAgrupacionAV, ByVal services As ServiceProvider) As DataColumn()
        Dim columns(9) As DataColumn
        columns(0) = data.Lineas.Columns("IDCliente")
        columns(1) = data.Lineas.Columns("IDMoneda")
        columns(2) = data.Lineas.Columns("IDFormaEnvio")
        columns(3) = data.Lineas.Columns("IDCondicionEnvio")
        columns(4) = data.Lineas.Columns("IDDireccionEnvio")
        columns(5) = data.Lineas.Columns("IDModoTransporte")
        columns(6) = data.Lineas.Columns("IDBancoPropio")
        columns(7) = data.Lineas.Columns("EDI")
        columns(8) = data.Lineas.Columns("Muelle")
        columns(9) = data.Lineas.Columns("PuntoDescarga")
        If data.TipoAgrupacion = enummcAgrupAlbaran.mcPedido Then
            ReDim Preserve columns(10)
            columns(10) = data.Lineas.Columns("IDPedido")
        End If

        Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
        If AppParams.GestionBodegas AndAlso data.Lineas.Columns.Contains("IDDaa") Then
            ReDim Preserve columns(columns.Length)
            columns(columns.Length - 1) = data.Lineas.Columns("IDDaa")
        End If

        If data.Lineas.Columns.Contains("IDClienteDistribuidor") Then
            ReDim Preserve columns(columns.Length)
            columns(columns.Length - 1) = data.Lineas.Columns("IDClienteDistribuidor")
        End If


        Return columns
    End Function

    <Task()> Public Shared Function AgruparPedidos(ByVal data As DataPrcAlbaranar, ByVal services As ServiceProvider) As AlbCabVentaPedido()
        Dim Pedidos() As CrearAlbaranVentaInfo = data.AlbVentaInfo
        Const strViewName As String = "vNegComercialCrearAlbaran"

        Dim ProcInfo As ProcessInfoAV = services.GetService(Of ProcessInfoAV)()

        Dim AppParams As ParametroAlbaranVenta = services.GetService(Of ParametroAlbaranVenta)()

        Dim dtLineas As DataTable
        Dim htLins As New Hashtable

        Dim ids(Pedidos.Length - 1) As Object
        For i As Integer = 0 To Pedidos.Length - 1
            ids(i) = Pedidos(i).IDLinea
            htLins.Add(Pedidos(i).IDLinea, Pedidos(i))
        Next
        If ids.Length > 0 Then
            Dim oFltr As New Filter
            oFltr.Add(New InListFilterItem("IDLineaPedido", ids, FilterType.Numeric))
            oFltr.Add(New NumberFilterItem("Estado", FilterOperator.NotEqual, enumpvlEstado.pvlCerrado))
            ' oFltr.Add(New NumberFilterItem(_PVL.Estado, FilterOperator.NotEqual, enumpvlEstado.pvlServido))
            If ProcInfo.IDTipoAlbaran = AppParams.TipoAlbaranExpDistribuidor Then
                oFltr.Add(New IsNullFilterItem("IDClienteDistribuidor", False))
            End If
            dtLineas = AdminData.GetData(strViewName, oFltr)
        End If

        If Not dtLineas Is Nothing AndAlso dtLineas.Rows.Count > 0 Then
            Dim Agrup As New DataColAgrupacionAV
            Agrup.Lineas = dtLineas
            Agrup.TipoAgrupacion = enummcAgrupAlbaran.mcPedido
            Dim ColsAgrupPedido As DataColumn() = ProcessServer.ExecuteTask(Of DataColAgrupacionAV, DataColumn())(AddressOf GetGroupColumns, Agrup, services)
            Agrup.TipoAgrupacion = enummcAgrupAlbaran.mcCliente
            Dim ColsAgrupClte As DataColumn() = ProcessServer.ExecuteTask(Of DataColAgrupacionAV, DataColumn())(AddressOf GetGroupColumns, Agrup, services)

            Dim oGrprUser As New GroupUserPedidos()
            Dim groupers(1) As GroupHelper
            groupers(enummcAgrupAlbaran.mcPedido) = New GroupHelper(ColsAgrupPedido, oGrprUser)
            groupers(enummcAgrupAlbaran.mcCliente) = New GroupHelper(ColsAgrupClte, oGrprUser)

            For Each rwLin As DataRow In dtLineas.Rows
                groupers(rwLin("AgrupAlbaran")).Group(rwLin)
            Next
            For Each alb As AlbCabVentaPedido In oGrprUser.Albs
                For Each alblin As AlbLinVentaPedido In alb.LineasOrigen
                    alblin.QaServir = DirectCast(htLins(alblin.IDLineaOrigen), CrearAlbaranVentaInfo).Cantidad
                    alblin.Cantidad = DirectCast(htLins(alblin.IDLineaOrigen), CrearAlbaranVentaInfo).CantidadUD
                    If Length(DirectCast(htLins(alblin.IDLineaOrigen), CrearAlbaranVentaInfo).Cantidad2) > 0 Then
                        alblin.Cantidad2 = DirectCast(htLins(alblin.IDLineaOrigen), CrearAlbaranVentaInfo).Cantidad2
                    End If
                    alblin.FechaEntregaModificado = DirectCast(htLins(alblin.IDLineaOrigen), CrearAlbaranVentaInfo).FechaEntregaModificado
                    alblin.Lotes = DirectCast(htLins(alblin.IDLineaOrigen), CrearAlbaranVentaInfo).Lotes
                    alblin.Series = DirectCast(htLins(alblin.IDLineaOrigen), CrearAlbaranVentaInfo).Series
                    alblin.Seguimiento = DirectCast(htLins(alblin.IDLineaOrigen), CrearAlbaranVentaInfo).Seguimiento
                    alblin.ArtCompatibles = DirectCast(htLins(alblin.IDLineaOrigen), CrearAlbaranVentaInfo).ArticulosCompatibles
                Next
            Next

            Return oGrprUser.Albs
        End If
    End Function


   

#End Region

#Region " Crear Lineas Albaran "

    <Task()> Public Shared Sub CrearLineasDesdePedido(ByVal oDocAlb As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        Dim dtOrigen As DataTable = ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta, DataTable)(AddressOf RecuperarDatosPedidos, oDocAlb, services)
        For Each lineaOrigen As DataRow In dtOrigen.Rows

            Dim alblin As AlbLinVenta = Nothing
            For i As Integer = 0 To oDocAlb.Cabecera.LineasOrigen.Length - 1
                If lineaOrigen(oDocAlb.Cabecera.LineasOrigen(i).PrimaryKeyLinOrigen) = oDocAlb.Cabecera.LineasOrigen(i).IDLineaOrigen Then
                    alblin = oDocAlb.Cabecera.LineasOrigen(i)
                    Exit For
                End If
            Next

            If Not alblin Is Nothing Then
                Dim dblCantidad As Double
                If Double.IsNaN(alblin.QaServir) Then
                    dblCantidad = lineaOrigen("QPedida") - lineaOrigen("QServida")
                Else
                    dblCantidad = alblin.QaServir
                End If

                If dblCantidad <> 0 Then
                    Dim NumLineasInsertar As Integer = 1
                    If (alblin.QaServir > 1 OrElse alblin.QaServir < -1) AndAlso Not alblin.Series Is Nothing AndAlso alblin.Series.Rows.Count > 0 Then NumLineasInsertar = alblin.Series.Rows.Count
                    Dim linAlbPed As DataLineasAVDesdeOrigen
                    For i As Integer = NumLineasInsertar - 1 To 0 Step -1
                        Dim linea As DataRow = oDocAlb.dtLineas.NewRow
                        Dim Info As ProcessInfoAV = services.GetService(Of ProcessInfoAV)()

                        linea("IDAlbaran") = oDocAlb.HeaderRow("IDAlbaran")
                        'linea("IDAlmacen") = oDocAlb.HeaderRow("IDAlmacen")

                        linea("FechaAlquiler") = oDocAlb.HeaderRow("FechaAlbaran")

                        ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoAlbaranVenta.AsignarValoresPredeterminadosLinea, linea, services)
                        If Not alblin.Series Is Nothing AndAlso alblin.Series.Rows.Count > 0 Then
                            linAlbPed = New DataLineasAVDesdeOrigen(linea, lineaOrigen, oDocAlb, alblin, dblCantidad, alblin.Series.Rows(i))
                        Else
                            linAlbPed = New DataLineasAVDesdeOrigen(linea, lineaOrigen, oDocAlb, alblin, dblCantidad)
                        End If

                        ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf AsignarDireccionFactura, linAlbPed, services)
                        ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf AsignarDatosPedido, linAlbPed, services)
                        ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf AsignarAlmacenLinea, linAlbPed, services)
                        ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf AsignarDatosArticulo, linAlbPed, services)
                        ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf AsignarCuenta, linAlbPed, services)
                        ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf AsignarDatosEconomicos, linAlbPed, services)
                        ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf AsignarUnidadesCantidades, linAlbPed, services)
                        ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf AsignarUnidadesCantidadesSegundaUnidad, linAlbPed, services)
                        ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf AsignarFechaEntregaModificado, linAlbPed, services)
                        ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf AsignarPreciosDtosImportes, linAlbPed, services)
                        ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf AsignarObservaciones, linAlbPed, services)
                        ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf AsignarNSerie, linAlbPed, services)
                        ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf ProcesoAlbaranVenta.AsignarEstadoStock, linAlbPed, services)

                        Dim ParamsAV As ParametroAlbaranVenta = services.GetService(Of ParametroAlbaranVenta)()
                        Dim ProcInfo As ProcessInfoAV = services.GetService(Of ProcessInfoAV)()
                        If ProcInfo.IDTipoAlbaran <> ParamsAV.TipoAlbaranDeConsumo Then
                            ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf AsignarEmbalajes, linAlbPed, services)
                            ProcessServer.ExecuteTask(Of DataLineasAVDesdeOrigen)(AddressOf AsignarTarifas, linAlbPed, services)
                            linea("IDLineaPadre") = System.DBNull.Value
                        End If

                        Dim c As New DataDocRow(oDocAlb, linea)
                        ProcessServer.ExecuteTask(Of DataDocRow)(AddressOf ValidarArticuloBloqueadoDoc, c, services)
                        oDocAlb.dtLineas.Rows.Add(linea.ItemArray)
                    Next
                End If
            End If
        Next
    End Sub

    <Task()> Public Shared Function RecuperarDatosPedidos(ByVal oDocAlb As DocumentoAlbaranVenta, ByVal services As ServiceProvider) As DataTable
        Dim albCabPed As AlbCabVenta = oDocAlb.Cabecera

        Dim FieldRow As String
        Dim ids(albCabPed.LineasOrigen.Length - 1) As Object
        For i As Integer = 0 To ids.Length - 1
            If Length(FieldRow) = 0 Then FieldRow = albCabPed.LineasOrigen(i).PrimaryKeyLinOrigen
            ids(i) = albCabPed.LineasOrigen(i).IDLineaOrigen
        Next

        Dim oFltr As New Filter
        oFltr.Add(New InListFilterItem(FieldRow, ids, FilterType.Numeric))
        oFltr.Add(New NumberFilterItem("Estado", FilterOperator.NotEqual, enumpvlEstado.pvlCerrado))
        Return New PedidoVentaLinea().Filter(oFltr)
    End Function

    <Task()> Public Shared Sub AsignarDireccionFactura(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        Dim PCCabeceras As EntityInfoCache(Of PedidoVentaCabeceraInfo) = services.GetService(Of EntityInfoCache(Of PedidoVentaCabeceraInfo))()
        Dim CabPedidoInfo As PedidoVentaCabeceraInfo = PCCabeceras.GetEntity(CType(data.Doc.Cabecera, AlbCabVentaPedido).IDOrigen) 'pedido("IDPedido"))
        If data.Doc.HeaderRow.IsNull("IDDireccionFra") And CabPedidoInfo.IDDireccionFra <> 0 Then
            data.Doc.HeaderRow("IDDireccionFra") = CabPedidoInfo.IDDireccionFra
        End If
    End Sub

    <Task()> Public Shared Sub AsignarDatosPedido(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        data.Row(data.AlbLin.PrimaryKeyLinOrigen) = data.Origen(data.AlbLin.PrimaryKeyLinOrigen)
        data.Row(data.Doc.Cabecera.PrimaryKeyCabOrigen) = data.Origen(data.Doc.Cabecera.PrimaryKeyCabOrigen)
        data.Row("IDCentroGestion") = data.Origen("IDCentroGestion")
        If Length(data.Origen("PedidoCliente")) > 0 Then
            data.Row("PedidoCliente") = data.Origen("PedidoCliente")
        ElseIf Length(data.Doc.HeaderRow("PedidoCliente")) > 0 Then
            data.Row("PedidoCliente") = data.Doc.HeaderRow("PedidoCliente")
        End If

        data.Row("IDOrdenLinea") = data.Origen("IdOrdenLinea")
        data.Row("IDTipoLinea") = data.Origen("IDTipoLinea")
    End Sub

    <Task()> Public Shared Sub AsignarAlmacenLinea(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        If Length(data.Doc.HeaderRow("IDTipoAlbaran")) > 0 Then
            Dim TipoAlbInfo As ProcesoAlbaranVenta.TipoAlbaranInfo = ProcessServer.ExecuteTask(Of String, ProcesoAlbaranVenta.TipoAlbaranInfo)(AddressOf ProcesoAlbaranVenta.TipoDeAlbaran, data.Doc.HeaderRow("IDTipoAlbaran"), services)
            If TipoAlbInfo.Tipo <> enumTipoAlbaran.Consumo Then
                data.Row("IDAlmacen") = data.Origen("IDAlmacen")
            Else
                data.Row("IDAlmacen") = data.Doc.HeaderRow("IDAlmacenDeposito")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarDatosArticulo(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
        Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.Origen("IDArticulo"))
        If ArtInfo.Configurable Then ApplicationService.GenerateError("No se puede expedir un artículo configurable {0}.", Quoted(data.Origen("IDArticulo")))
        data.Row("IDArticulo") = data.Origen("IDArticulo")
        data.Row("DescArticulo") = data.Origen("DescArticulo")
        'linea("PrecioUltimaCompra") = pedido("PrecioUltimaCompra")
        data.Row("RefCliente") = data.Origen("RefCliente")
        data.Row("DescRefCliente") = data.Origen("DescRefCliente")
        data.Row("Revision") = data.Origen("Revision")
        If Length(data.Doc.HeaderRow("IDTipoAlbaran")) > 0 Then
            Dim TipoAlbInfo As ProcesoAlbaranVenta.TipoAlbaranInfo = ProcessServer.ExecuteTask(Of String, ProcesoAlbaranVenta.TipoAlbaranInfo)(AddressOf ProcesoAlbaranVenta.TipoDeAlbaran, data.Doc.HeaderRow("IDTipoAlbaran"), services)
            If TipoAlbInfo.Tipo = enumTipoAlbaran.ExpedDistribuidor OrElse TipoAlbInfo.Tipo = enumTipoAlbaran.AbonoDistribuidor Then
                data.Row("EstadoStock") = CInt(enumavlEstadoStock.avlSinGestion)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarCuenta(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        Dim AppParamsConta As ParametroContabilidadVenta = services.GetService(Of ParametroContabilidadVenta)()
        If AppParamsConta.Contabilidad Then
            data.Row("CContable") = data.Origen("CContable")
        End If
    End Sub

    <Task()> Public Shared Sub AsignarDatosEconomicos(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        Dim PCCabeceras As EntityInfoCache(Of PedidoVentaCabeceraInfo) = services.GetService(Of EntityInfoCache(Of PedidoVentaCabeceraInfo))()
        Dim CabPedidoInfo As PedidoVentaCabeceraInfo = PCCabeceras.GetEntity(data.Origen("IDPedido"))

        If Length(CabPedidoInfo.IDCondicionPago) > 0 Then
            data.Row("IDCondicionPago") = CabPedidoInfo.IDCondicionPago
        Else
            ApplicationService.GenerateError("La Condición Pago no existe.")
        End If
        If Length(CabPedidoInfo.IDFormaPago) > 0 Then
            data.Row("IDFormaPago") = CabPedidoInfo.IDFormaPago
        Else
            ApplicationService.GenerateError("La Forma Pago no existe.")
        End If
        If CabPedidoInfo.IDDireccionFra <> 0 Then
            data.Row("IDDireccionFra") = CabPedidoInfo.IDDireccionFra
        End If
        data.Row("Dto") = CabPedidoInfo.DtoPedido
        data.Row("IDTipoIva") = data.Origen("IDTipoIva")
        If CabPedidoInfo.IDClienteBanco <> 0 Then
            data.Row("IDClienteBanco") = CabPedidoInfo.IDClienteBanco
        Else
            Dim f As New Filter
            f.Add(New StringFilterItem("IDCliente", data.Doc.HeaderRow("IDCliente")))
            f.Add(New BooleanFilterItem("Predeterminado", True))
            Dim dtClienteBanco As DataTable = New ClienteBanco().Filter(f)
            If dtClienteBanco.Rows.Count > 0 Then
                data.Row("IDClienteBanco") = dtClienteBanco.Rows(0)("IDClienteBanco")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarUnidadesCantidades(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        data.Row("IDUdMedida") = data.Origen("IDUdMedida")
        data.Row("IDUDExpedicion") = Nz(data.Origen("IDUDExpedicion"), String.Empty)
        data.Row("IDUdInterna") = data.Origen("IDUdInterna")
        data.Row("UdValoracion") = data.Origen("UdValoracion")
        data.Row("QServida") = data.Cantidad
        If TypeOf data.AlbLin Is AlbLinVentaPedido Then
            If data.Row("QServida") <> 0 Then
                If Length(data.NSerie) > 0 Then
                    data.Row("Factor") = data.Origen("Factor")
                    data.Row("QInterna") = data.Cantidad * data.Origen("Factor")
                Else
                    data.Row("QInterna") = CType(data.AlbLin, AlbLinVentaPedido).Cantidad
                    data.Row("Factor") = data.Row("QInterna") / data.Row("QServida")
                End If
            Else
                data.Row("QInterna") = 0
                data.Row("Factor") = 0
            End If
        Else
            data.Row("Factor") = data.Origen("Factor")
            data.Row("QInterna") = data.Row("QServida") * data.Origen("Factor")
            If data.Origen("QPedida") <> 0 Then
                If data.Row("QInterna") <> (data.Row("QServida") * (data.Origen("QInterna") / data.Origen("QPedida"))) Then
                    data.Row("QInterna") = data.Row("QServida") * (data.Origen("QInterna") / data.Origen("QPedida"))
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarUnidadesCantidadesSegundaUnidad(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        If ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, data.Row("IDArticulo"), services) Then
            If TypeOf data.AlbLin Is AlbLinVentaPedido Then
                If Length(CType(data.AlbLin, AlbLinVentaPedido).Cantidad2) = 0 Then
                    ApplicationService.GenerateError("El Articulo {0} se gestiona con Doble Unidad. Debe indicar la misma.", Quoted(data.Row("IDArticulo")))
                Else
                    data.Row("QInterna2") = CType(data.AlbLin, AlbLinVentaPedido).Cantidad2
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarFechaEntregaModificado(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        If CDate(Nz(data.Origen("FechaEntregaModificado"), cnMinDate)) <> cnMinDate Then
            data.Row("FechaEntregaModificado") = data.Origen("FechaEntregaModificado")
        ElseIf Not data.AlbLin Is Nothing AndAlso TypeOf data.AlbLin Is AlbLinVentaPedido Then
            If Length(CType(data.AlbLin, AlbLinVentaPedido).FechaEntregaModificado) > 0 AndAlso Nz(CType(data.AlbLin, AlbLinVentaPedido).FechaEntregaModificado, cnMinDate) <> cnMinDate Then
                data.Row("FechaEntregaModificado") = CType(data.AlbLin, AlbLinVentaPedido).FechaEntregaModificado
            Else : data.Row("FechaEntregaModificado") = Today
            End If
        Else : data.Row("FechaEntregaModificado") = Today
        End If
    End Sub

    <Task()> Public Shared Sub AsignarPreciosDtosImportes(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        Dim AVL As New AlbaranVentaLinea
        Dim context As New BusinessData(data.Doc.HeaderRow)
        AVL.ApplyBusinessRule("Precio", data.Origen("Precio"), data.Row, context)
        data.Row("PrecioCosteA") = data.Origen("PrecioCosteA")
        data.Row("PrecioCosteB") = data.Origen("PrecioCosteB")
        data.Row("Dto1") = data.Origen("Dto1")
        data.Row("Dto2") = data.Origen("Dto2")
        data.Row("Dto3") = data.Origen("Dto3")
        data.Row("Dto") = data.Origen("Dto")
        data.Row("DtoProntoPago") = data.Origen("DtoProntoPago")
        data.Row("Regalo") = data.Origen("Regalo")
        If data.Origen("Regalo") Then
            data.Row("Importe") = 0
            data.Row("ImporteA") = 0
            data.Row("ImporteB") = 0
        End If
        Dim ValAyB As New ValoresAyB(New DataRowPropertyAccessor(data.Row), data.Doc.IDMoneda, data.Doc.CambioA, data.Doc.CambioB)
        ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)
    End Sub

    <Task()> Public Shared Sub AsignarObservaciones(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        data.Row("Texto") = data.Origen("Texto")
    End Sub

    <Task()> Public Shared Sub AsignarNSerie(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        If Length(data.NSerie) > 0 Then data.Row("Lote") = data.NSerie
        If Length(data.IDEstadoActivo) > 0 Then data.Row("IDEstadoActivo") = data.IDEstadoActivo
        If Length(data.IDOperario) > 0 Then data.Row("IDOperario") = data.IDOperario
        If Length(data.Ubicacion) > 0 Then data.Row("Ubicacion") = data.Ubicacion
    End Sub

    <Task()> Public Shared Sub AsignarEmbalajes(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        'Gestión de Etiquetas
        'Primero miramos que no haya nada configurado de embalajes o contenedores,
        'a nivel de relación Articulo - Cliente
        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
        Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.Origen("IDArticulo"))
        Dim ArtCltes As EntityInfoCache(Of ArticuloClienteInfo) = services.GetService(Of EntityInfoCache(Of ArticuloClienteInfo))()
        Dim ArtClteInfo As ArticuloClienteInfo = ArtCltes.GetEntity(data.Doc.HeaderRow("IDCliente"), data.Origen("IDArticulo"))
        If IsNothing(ArtClteInfo) OrElse (ArtClteInfo.QContenedor = 0 AndAlso ArtClteInfo.QEmbalaje = 0) Then
            'Si no existen datos a nivel de relación Artículo - Cliente,
            'miramos los datos configurados a nivel de cabecera de articulo
            If IsNothing(ArtClteInfo) Then ArtClteInfo = New ArticuloClienteInfo
            ArtClteInfo.IDArticulo = ArtInfo.IDArticulo
            ArtClteInfo.IDArticuloContenedor = ArtInfo.IDArticuloContenedor
            ArtClteInfo.QContenedor = ArtInfo.QContenedor
            ArtClteInfo.IDArticuloEmbalaje = ArtInfo.IDArticuloEmbalaje
            ArtClteInfo.QEmbalaje = ArtInfo.QEmbalaje
        End If

        If Not IsNothing(ArtClteInfo) Then
            Dim QContenedor As Double
            Dim QEmbalaje As Double

            data.Row("IDArticuloContenedor") = IIf(Length(ArtClteInfo.IDArticuloContenedor) > 0, ArtClteInfo.IDArticuloContenedor, DBNull.Value)
            data.Row("IDArticuloEmbalaje") = IIf(Length(ArtClteInfo.IDArticuloEmbalaje) > 0, ArtClteInfo.IDArticuloEmbalaje, DBNull.Value)
            QContenedor = IIf(ArtClteInfo.QContenedor > 0, ArtClteInfo.QContenedor, 0)
            QEmbalaje = IIf(ArtClteInfo.QEmbalaje > 0, ArtClteInfo.QEmbalaje, 0)
            Dim entero As Integer
            Dim resto As Integer
            If QContenedor > 0 Then
                entero = Int(data.Row("QServida") / QContenedor)
                resto = data.Row("QServida") Mod QContenedor
                If resto > 0 Then
                    data.Row("QEtiContenedor") = entero + 1
                Else
                    data.Row("QEtiContenedor") = entero
                End If
            End If
            If QEmbalaje > 0 Then
                entero = Int(data.Row("QServida") / QEmbalaje)
                resto = data.Row("QServida") Mod QEmbalaje
                If resto > 0 Then
                    data.Row("QEtiEmbalaje") = entero + 1
                Else
                    data.Row("QEtiEmbalaje") = entero
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarTarifas(ByVal data As DataLineasAVDesdeOrigen, ByVal services As ServiceProvider)
        data.Row("IDPromocionLinea") = data.Origen("IDPromocionLinea")
        data.Row("IDPromocion") = data.Origen("IDPromocion")
        data.Row("IDTarifa") = data.Origen("IDTarifa")
        data.Row("SeguimientoTarifa") = data.Origen("SeguimientoTarifa")
        data.Row("IDLineaOfertaDetalle") = data.Origen("IDLineaOfertaDetalle")
    End Sub

    <Task()> Public Shared Sub ValidarArticuloBloqueadoDoc(ByVal data As DataDocRow, ByVal services As ServiceProvider)
        If Length(data.Row("IDArticulo")) > 0 AndAlso Length(data.Doc.HeaderRow("IDCliente")) > 0 Then
            Dim StDatos As New Cliente.DataBloqArtClie
            StDatos.IDArticulo = data.Row("IDArticulo") : StDatos.IDCliente = data.Doc.HeaderRow("IDCliente")
            If ProcessServer.ExecuteTask(Of Cliente.DataBloqArtClie, Boolean)(AddressOf Cliente.ComprobarBloqueoArticuloCliente, StDatos, services) Then
                ApplicationService.GenerateError("El Artículo está Bloqueado para este Cliente.")
            End If
        End If
    End Sub

#End Region

#Region " Actualización de Pedidos "

    Private Shared MyIDPedido As Integer

    <Task()> Public Shared Sub ActualizarPedidoDesdeAlbaranTPV(ByVal Doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        If Doc Is Nothing OrElse Doc.dtLineas Is Nothing Then Exit Sub

        Dim PedidosTPVCerrar As List(Of Object) = (From c In Doc.dtLineas _
                                                    Where Not c.IsNull("IDPedido") _
                                                    Select c("IDPedido") Distinct).ToList()
        If Not PedidosTPVCerrar Is Nothing Then
            Dim PVL As New PedidoVentaLinea
            For Each IDPedido As Integer In PedidosTPVCerrar
                Dim dtLineasPedido As DataTable = PVL.Filter(New NumberFilterItem("IDPedido", IDPedido))
                For Each drLinea As DataRow In dtLineasPedido.Rows
                    drLinea("Estado") = enumpvlEstado.pvlServido
                    drLinea("QServida") = drLinea("QPedida")
                Next
                BusinessHelper.UpdateTable(dtLineasPedido)
            Next
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarPedidoDesdeAlbaran(ByVal Doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        If Doc Is Nothing OrElse Doc.dtLineas Is Nothing Then Exit Sub

        'If Length(Doc.HeaderRow("IDTPV")) > 0 Then
        '    ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf ActualizarPedidoDesdeAlbaranTPV, Doc, services)
        'Else
        Dim fLineasPedido As New Filter
        fLineasPedido.Add(New IsNullFilterItem("IDPedido", False))
        fLineasPedido.Add(New IsNullFilterItem("IDLineaPedido", False))
        Dim strLineasPedido As String = fLineasPedido.Compose(New AdoFilterComposer)
        Dim Pedidos As New System.Collections.Generic.List(Of DataTable)
        For Each lineaAlbaran As DataRow In Doc.dtLineas.Select(strLineasPedido, "IDPedido,IDLineaPedido")
            Dim DataActua As New DataActuaPedidos(Doc, lineaAlbaran, Pedidos)
            ProcessServer.ExecuteTask(Of DataActuaPedidos)(AddressOf ActualizarLineasPedido, DataActua, services)
        Next
        If Not Pedidos Is Nothing AndAlso Pedidos.Count > 0 Then
            For Each Pedido As DataTable In Pedidos
                BusinessHelper.UpdateTable(Pedido)
            Next
        End If
        'End If
    End Sub

    <Task()> Public Shared Sub ActualizarAlbaranClteDesdeAlbaran(ByVal Doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        If Not Doc.Cabecera Is Nothing Then
            If Not Doc.HeaderRow Is Nothing AndAlso Length(Doc.HeaderRow("IDTipoAlbaran")) > 0 Then
                Dim TipoAlbInfo As ProcesoAlbaranVenta.TipoAlbaranInfo = ProcessServer.ExecuteTask(Of String, ProcesoAlbaranVenta.TipoAlbaranInfo)(AddressOf ProcesoAlbaranVenta.TipoDeAlbaran, Doc.HeaderRow("IDTipoAlbaran"), services)
                If TipoAlbInfo.Tipo = enumTipoAlbaran.AbonoDistribuidor Then
                    Dim IDAlbaran As Integer = Doc.Cabecera.IDOrigen
                    If IDAlbaran <> 0 Then
                        Dim dtAVC As DataTable = New AlbaranVentaCabecera().SelOnPrimaryKey(IDAlbaran)
                        If dtAVC.Rows.Count > 0 Then
                            dtAVC.Rows(0)("IDAlbaranAbono") = Doc.HeaderRow("IDAlbaran")
                            BusinessHelper.UpdateTable(dtAVC)
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    <Serializable()> _
    Public Class DataActuaPedidos
        Public Doc As DocumentoAlbaranVenta
        Public LineaAlbaran As DataRow
        Public Pedidos As New System.Collections.Generic.List(Of DataTable)

        Public Sub New()
        End Sub
        Public Sub New(ByVal Doc As DocumentoAlbaranVenta, ByVal LineaAlbaran As DataRow, ByVal Pedidos As System.Collections.Generic.List(Of DataTable))
            Me.Doc = Doc
            Me.LineaAlbaran = LineaAlbaran
            Me.Pedidos = Pedidos
        End Sub
    End Class

    <Task()> Public Shared Sub ActualizarLineasPedido(ByVal data As DataActuaPedidos, ByVal services As ServiceProvider)
        If Length(data.LineaAlbaran("IDPedido")) > 0 Then 'AndAlso Length(data.LineaAlbaran("IDLineaPedido")) > 0 Then
            Dim BlnFound As Boolean = False
            Dim DtPedCab As DataTable
            Dim DtPedLineas As DataTable
            If Not data.Pedidos Is Nothing AndAlso data.Pedidos.Count > 0 Then
                MyIDPedido = data.LineaAlbaran("IDPedido")
                Dim FindIndexCab As Integer = data.Pedidos.FindIndex(AddressOf FindPedido)
                If FindIndexCab <> -1 Then DtPedCab = data.Pedidos(FindIndexCab)
                Dim FindIndexLin As Integer = data.Pedidos.FindIndex(AddressOf FindLinPedido)
                If FindIndexLin <> -1 Then DtPedLineas = data.Pedidos(FindIndexLin)
                If FindIndexCab <> -1 AndAlso FindIndexLin <> -1 Then
                    BlnFound = True
                Else
                    BlnFound = False
                    DtPedCab = New PedidoVentaCabecera().SelOnPrimaryKey(data.LineaAlbaran("IDPedido"))
                    DtPedLineas = New PedidoVentaLinea().Filter(New FilterItem("IDPedido", data.LineaAlbaran("IDPedido")))
                End If
            Else
                BlnFound = False
                DtPedCab = New PedidoVentaCabecera().SelOnPrimaryKey(data.LineaAlbaran("IDPedido"))
                DtPedLineas = New PedidoVentaLinea().Filter(New FilterItem("IDPedido", data.LineaAlbaran("IDPedido")))
            End If
            If Not DtPedCab Is Nothing AndAlso DtPedCab.Rows.Count > 0 AndAlso Not DtPedLineas Is Nothing AndAlso DtPedLineas.Rows.Count > 0 Then
                Dim TipoAlbaranes As ParametroAlbaranVenta = services.GetService(Of ParametroAlbaranVenta)()
                Dim fLineaPedido As New Filter

                '//Lineas de Pedido del TPV, pueden venir agrupadas en una única línea, de ahí la validación que vena el IDLineaPedido relleno. 
                '//Para saber que son este tipo de pedidos.
                If Nz(data.LineaAlbaran("IDLineaPedido"), 0) <> 0 Then fLineaPedido.Add(New NumberFilterItem("IDLineaPedido", data.LineaAlbaran("IDLineaPedido")))
                For Each lineaPedido As DataRow In DtPedLineas.Select(fLineaPedido.Compose(New AdoFilterComposer))
                    'Actualizar QServidaLineaPedido
                    If Nz(data.LineaAlbaran("TipoLineaAlbaran"), enumavlTipoLineaAlbaran.avlNormal) <> enumavlTipoLineaAlbaran.avlComponente OrElse _
                       (Nz(data.LineaAlbaran("TipoLineaAlbaran"), enumavlTipoLineaAlbaran.avlNormal) = enumavlTipoLineaAlbaran.avlComponente AndAlso _
                        Length(data.LineaAlbaran("IDLineaPadre")) = 0) Then
                        If data.LineaAlbaran.RowState <> DataRowState.Modified OrElse data.LineaAlbaran("QServida") <> data.LineaAlbaran("QServida", DataRowVersion.Original) Then
                            If TipoAlbaranes.TipoAlbaranDeConsumo <> data.Doc.HeaderRow("IDTipoAlbaran") Then
                                Dim OriginalQServida As Double
                                Dim ProposedQServida As Double = Nz(data.LineaAlbaran("QServida"), 0)
                                If data.LineaAlbaran.RowState = DataRowState.Modified Then
                                    OriginalQServida = data.LineaAlbaran("QServida", DataRowVersion.Original)
                                End If
                                lineaPedido("QServida") = Nz(lineaPedido("QServida"), 0) + (ProposedQServida - OriginalQServida)
                                ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoPedidoVenta.AsignarEstadoLinea, lineaPedido, services)
                            End If
                        End If
                    End If
                    'Actualizar QAlbaranLineaPedido
                    If data.Doc.HeaderRow("IDTipoAlbaran") <> TipoAlbaranes.TipoAlbaranDeConsumo _
                    And data.Doc.HeaderRow("IDTipoAlbaran") <> TipoAlbaranes.TipoAlbaranDeIntercambio Then
                        lineaPedido("QAlbaran") = 0
                    End If
                    'Actualizar ConfirmadoLineaPedido
                    If TipoAlbaranes.TipoAlbaranDeConsumo <> data.Doc.HeaderRow("IDTipoAlbaran") Then
                        lineaPedido("Confirmado") = False
                    End If
                    'Actualizar DepositoLineaPedido
                    If TipoAlbaranes.TipoAlbaranDeDeposito = data.Doc.HeaderRow("IDTipoAlbaran") Then
                        lineaPedido("Deposito") = (TipoAlbaranes.TipoAlbaranDeDeposito = data.Doc.HeaderRow("IDTipoAlbaran"))
                    End If
                    'Actualizar QFacturadaLineaPedido
                    If data.LineaAlbaran.RowState <> DataRowState.Modified OrElse data.LineaAlbaran("QServida") <> data.LineaAlbaran("QServida", DataRowVersion.Original) Then
                        If TipoAlbaranes.TipoAlbaranDeDeposito <> data.Doc.HeaderRow("IDTipoAlbaran") AndAlso data.LineaAlbaran("TipoLineaAlbaran") <> enumavlTipoLineaAlbaran.avlComponente Then
                            Dim OriginalQServida As Double
                            Dim ProposedQServida As Double = Nz(data.LineaAlbaran("QServida"), 0)
                            If data.LineaAlbaran.RowState = DataRowState.Modified Then
                                OriginalQServida = data.LineaAlbaran("QServida", DataRowVersion.Original)
                            End If
                            lineaPedido("QFacturada") = Nz(lineaPedido("QFacturada"), 0) + (ProposedQServida - OriginalQServida)
                        End If
                    End If
                Next
                If Not BlnFound Then
                    data.Pedidos.Add(DtPedCab)
                    data.Pedidos.Add(DtPedLineas)
                End If
            End If
        End If
    End Sub

    Private Shared Function FindPedido(ByVal Ped As DataTable) As Boolean
        If Not Ped Is Nothing AndAlso Ped.Rows.Count > 0 Then
            If Not Ped.Columns.Contains("IDLineaPedido") Then
                Dim DrFind() As DataRow = Ped.Select("IDPedido = " & MyIDPedido)
                If DrFind.Length > 0 Then
                    Return True
                Else : Return False
                End If
            Else : Return False
            End If
        Else : Return False
        End If
    End Function

    Private Shared Function FindLinPedido(ByVal Ped As DataTable) As Boolean
        If Not Ped Is Nothing AndAlso Ped.Rows.Count > 0 Then
            If Ped.Columns.Contains("IDLineaPedido") Then
                Dim DrFind() As DataRow = Ped.Select("IDPedido = " & MyIDPedido)
                If DrFind.Length > 0 Then
                    Return True
                Else : Return False
                End If
            Else : Return False
            End If
        Else : Return False
        End If
    End Function

    <Task()> Public Shared Sub ActualizarLineaPedido(ByVal lineaAlbaran As DataRow, ByVal services As ServiceProvider)
        If Length(lineaAlbaran("IDPedido")) > 0 AndAlso Length(lineaAlbaran("IDLineaPedido")) > 0 Then
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf ActualizarQServidaLineaPedido, lineaAlbaran, services)
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf ActualizarQAlbaranLineaPedido, lineaAlbaran, services)
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf ActualizarConfirmadoLineaPedido, lineaAlbaran, services)
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf ActualizarDepositoLineaPedido, lineaAlbaran, services)
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf ActualizarQFacturadaLineaPedido, lineaAlbaran, services)
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarQServidaLineaPedido(ByVal lineaAlbaran As DataRow, ByVal services As ServiceProvider)
        If Nz(lineaAlbaran("TipoLineaAlbaran"), enumavlTipoLineaAlbaran.avlNormal) <> enumavlTipoLineaAlbaran.avlComponente Then
            If lineaAlbaran.RowState <> DataRowState.Modified OrElse lineaAlbaran("QServida") <> lineaAlbaran("QServida", DataRowVersion.Original) Then
                Dim Albaranes As DocumentInfoCache(Of DocumentoAlbaranVenta) = services.GetService(Of DocumentInfoCache(Of DocumentoAlbaranVenta))()
                Dim DocAlb As DocumentoAlbaranVenta = Albaranes.GetDocument(lineaAlbaran("IDAlbaran"))
                Dim TipoAlbaranes As ParametroAlbaranVenta = services.GetService(Of ParametroAlbaranVenta)()
                If TipoAlbaranes.TipoAlbaranDeConsumo <> DocAlb.HeaderRow("IDTipoAlbaran") Then
                    Dim Pedidos As DocumentInfoCache(Of DocumentoPedidoVenta) = services.GetService(Of DocumentInfoCache(Of DocumentoPedidoVenta))()
                    Dim DocPed As DocumentoPedidoVenta = Pedidos.GetDocument(lineaAlbaran("IDPedido"))

                    Dim OriginalQServida As Double
                    Dim ProposedQServida As Double = Nz(lineaAlbaran("QServida"), 0)
                    If lineaAlbaran.RowState = DataRowState.Modified Then
                        OriginalQServida = lineaAlbaran("QServida", DataRowVersion.Original)
                    End If

                    DocPed.SetQServida(lineaAlbaran("IDLineaPedido"), ProposedQServida - OriginalQServida, services)
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarQAlbaranLineaPedido(ByVal lineaAlbaran As DataRow, ByVal services As ServiceProvider)
        Dim Albaranes As DocumentInfoCache(Of DocumentoAlbaranVenta) = services.GetService(Of DocumentInfoCache(Of DocumentoAlbaranVenta))()
        Dim DocAlb As DocumentoAlbaranVenta = Albaranes.GetDocument(lineaAlbaran("IDAlbaran"))
        Dim TipoAlbaranes As ParametroAlbaranVenta = services.GetService(Of ParametroAlbaranVenta)()
        If DocAlb.HeaderRow("IDTipoAlbaran") <> TipoAlbaranes.TipoAlbaranDeConsumo _
        And DocAlb.HeaderRow("IDTipoAlbaran") <> TipoAlbaranes.TipoAlbaranDeIntercambio Then
            Dim Pedidos As DocumentInfoCache(Of DocumentoPedidoVenta) = services.GetService(Of DocumentInfoCache(Of DocumentoPedidoVenta))()
            Dim DocPed As DocumentoPedidoVenta = Pedidos.GetDocument(lineaAlbaran("IDPedido"))
            DocPed.SetQAlbaran(lineaAlbaran("IDLineaPedido"), 0, services)
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarConfirmadoLineaPedido(ByVal lineaAlbaran As DataRow, ByVal services As ServiceProvider)
        Dim Albaranes As DocumentInfoCache(Of DocumentoAlbaranVenta) = services.GetService(Of DocumentInfoCache(Of DocumentoAlbaranVenta))()
        Dim DocAlb As DocumentoAlbaranVenta = Albaranes.GetDocument(lineaAlbaran("IDAlbaran"))
        Dim TipoAlbaranes As ParametroAlbaranVenta = services.GetService(Of ParametroAlbaranVenta)()
        If TipoAlbaranes.TipoAlbaranDeConsumo <> DocAlb.HeaderRow("IDTipoAlbaran") Then
            Dim Pedidos As DocumentInfoCache(Of DocumentoPedidoVenta) = services.GetService(Of DocumentInfoCache(Of DocumentoPedidoVenta))()
            Dim DocPed As DocumentoPedidoVenta = Pedidos.GetDocument(lineaAlbaran("IDPedido"))
            DocPed.SetConfirmado(lineaAlbaran("IDLineaPedido"), False, services)
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarDepositoLineaPedido(ByVal lineaAlbaran As DataRow, ByVal services As ServiceProvider)
        Dim Albaranes As DocumentInfoCache(Of DocumentoAlbaranVenta) = services.GetService(Of DocumentInfoCache(Of DocumentoAlbaranVenta))()
        Dim DocAlb As DocumentoAlbaranVenta = Albaranes.GetDocument(lineaAlbaran("IDAlbaran"))
        Dim TipoAlbaranes As ParametroAlbaranVenta = services.GetService(Of ParametroAlbaranVenta)()

        Dim Pedidos As DocumentInfoCache(Of DocumentoPedidoVenta) = services.GetService(Of DocumentInfoCache(Of DocumentoPedidoVenta))()
        Dim DocPed As DocumentoPedidoVenta = Pedidos.GetDocument(lineaAlbaran("IDPedido"))
        If TipoAlbaranes.TipoAlbaranDeDeposito = DocAlb.HeaderRow("IDTipoAlbaran") Then
            DocPed.SetDeposito(lineaAlbaran("IDLineaPedido"), TipoAlbaranes.TipoAlbaranDeDeposito = DocAlb.HeaderRow("IDTipoAlbaran"), services)
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarQFacturadaLineaPedido(ByVal lineaAlbaran As DataRow, ByVal services As ServiceProvider)
        If lineaAlbaran.RowState <> DataRowState.Modified OrElse lineaAlbaran("QServida") <> lineaAlbaran("QServida", DataRowVersion.Original) Then
            Dim Albaranes As DocumentInfoCache(Of DocumentoAlbaranVenta) = services.GetService(Of DocumentInfoCache(Of DocumentoAlbaranVenta))()
            Dim DocAlb As DocumentoAlbaranVenta = Albaranes.GetDocument(lineaAlbaran("IDAlbaran"))
            Dim TipoAlbaranes As ParametroAlbaranVenta = services.GetService(Of ParametroAlbaranVenta)()
            If TipoAlbaranes.TipoAlbaranDeDeposito <> DocAlb.HeaderRow("IDTipoAlbaran") Then
                Dim Pedidos As DocumentInfoCache(Of DocumentoPedidoVenta) = services.GetService(Of DocumentInfoCache(Of DocumentoPedidoVenta))()
                Dim DocPed As DocumentoPedidoVenta = Pedidos.GetDocument(lineaAlbaran("IDPedido"))

                Dim OriginalQServida As Double
                Dim ProposedQServida As Double = Nz(lineaAlbaran("QServida"), 0)
                If lineaAlbaran.RowState = DataRowState.Modified Then
                    OriginalQServida = lineaAlbaran("QServida", DataRowVersion.Original)
                End If

                DocPed.SetQFacturada(lineaAlbaran("IDLineaPedido"), ProposedQServida - OriginalQServida, services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub GrabarPedidos(ByVal data As Object, ByVal services As ServiceProvider)
        AdminData.BeginTx()
        Dim Pedidos As DocumentInfoCache(Of DocumentoPedidoVenta) = services.GetService(Of DocumentInfoCache(Of DocumentoPedidoVenta))()

        For Each key As Integer In Pedidos.Keys
            Dim DocPed As DocumentoPedidoVenta = Pedidos.GetDocument(key)
            DocPed.SetData()
        Next
        Pedidos.Clear()
    End Sub

#End Region

#Region " Analítica y representantes "

    <Task()> Public Shared Sub CopiarRepresentantes(ByVal Doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        If Doc.HeaderRow("Estado") <> enumavcEstadoFactura.avcFacturado Then
            ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoComercial.CopiarRepresentantes, Doc, services)
        End If
    End Sub

    <Task()> Public Shared Sub CalcularRepresentantes(ByVal Doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        If Doc.HeaderRow("Estado") <> enumavcEstadoFactura.avcFacturado Then
            ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoComercial.CalcularRepresentantes, Doc, services)
        End If
    End Sub

    <Task()> Public Shared Sub CopiarAnalitica(ByVal Doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        If Doc.HeaderRow("Estado") <> enumavcEstadoFactura.avcFacturado Then
            Dim AppParams As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
            If Not AppParams.Contabilidad OrElse Not AppParams.Analitica.AplicarAnalitica Then Exit Sub

            Dim IDOrigen(-1) As Object
            Dim f As New Filter
            f.Add(New IsNullFilterItem("IDLineaPedido", False))
            Dim WhereNotNullLineaPedido As String = f.Compose(New AdoFilterComposer)
            For Each linea As DataRow In Doc.dtLineas.Select(WhereNotNullLineaPedido)
                ReDim Preserve IDOrigen(IDOrigen.Length)
                IDOrigen(IDOrigen.Length - 1) = linea("IDLineaPedido")
            Next
            If IDOrigen.Length > 0 Then
                Dim dtAnaliticaOrigen As DataTable = New PedidoVentaAnalitica().Filter(New InListFilterItem("IDLineaPedido", IDOrigen, FilterType.Numeric))
                Dim datosCopia As New NegocioGeneral.DataCopiarAnalitica(dtAnaliticaOrigen, Doc)
                ProcessServer.ExecuteTask(Of NegocioGeneral.DataCopiarAnalitica)(AddressOf NegocioGeneral.CopiarAnalitica, datosCopia, services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CalcularAnalitica(ByVal Doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        If Doc.HeaderRow("Estado") <> enumavcEstadoFactura.avcFacturado Then
            ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf AnaliticaLineasNSerie, Doc, services)
            ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf NegocioGeneral.CalcularAnalitica, Doc, services)
        End If
    End Sub

    <Task()> Public Shared Sub AnaliticaLineasNSerie(ByVal Doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        If Doc.HeaderRow("Estado") <> enumavcEstadoFactura.avcFacturado Then
            '//Altas de NºSerie. Ver si hay analítica en el Pedido para que se recalcule. Copiamos la Analitica sólo de las lineas de NºSerie
            Dim AppParams As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
            If Not AppParams.Contabilidad OrElse Not AppParams.Analitica.AplicarAnalitica Then Exit Sub

            Dim IDOrigen(-1) As Object
            Dim f As New Filter
            f.Add(New IsNullFilterItem("IDLineaPedido", False))
            f.Add(New IsNullFilterItem("Lote", False))
            f.Add(New IsNullFilterItem("Ubicacion", False))
            f.Add(New NumberFilterItem("QServida", 1))
            f.Add(New NumberFilterItem("QInterna", 1))
            Dim WhereNotNullLineaPedido As String = f.Compose(New AdoFilterComposer)
            For Each linea As DataRow In Doc.dtLineas.Select(WhereNotNullLineaPedido, Nothing, DataViewRowState.Added)
                ReDim Preserve IDOrigen(IDOrigen.Length)
                IDOrigen(IDOrigen.Length - 1) = linea("IDLineaPedido")
            Next
            If IDOrigen.Length > 0 Then
                Dim dtAnaliticaOrigen As DataTable = New PedidoVentaAnalitica().Filter(New InListFilterItem("IDLineaPedido", IDOrigen, FilterType.Numeric))
                Dim datosCopia As New NegocioGeneral.DataCopiarAnalitica(dtAnaliticaOrigen, Doc)
                ProcessServer.ExecuteTask(Of NegocioGeneral.DataCopiarAnalitica)(AddressOf NegocioGeneral.CopiarAnalitica, datosCopia, services)
            End If
            '//Fin Altas Nº Serie.
        End If
    End Sub

#End Region


End Class
