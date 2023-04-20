Imports System.Collections.Generic

Public Class PrcPropuestaPedidoCompraDesdePedidoVenta
    Inherits Process(Of List(Of DataPrcPropuestaPedidoCompraDesdePedidoVenta), DataTable)

    Public Overrides Sub RegisterTasks()
        Me.AddTask(Of List(Of DataPrcPropuestaPedidoCompraDesdePedidoVenta), DataTable)(AddressOf GetDisponibilidadPedidosVenta)
        Me.AddTask(Of DataTable, DataTable)(AddressOf PropuestaPedidoCompra)
    End Sub

    <Task()> Public Shared Function GetDisponibilidadPedidosVenta(ByVal pedidosVenta As List(Of DataPrcPropuestaPedidoCompraDesdePedidoVenta), ByVal services As ServiceProvider) As DataTable
        Dim IDs(pedidosVenta.Count - 1) As Integer
        For i As Integer = 0 To pedidosVenta.Count - 1
            IDs(i) = pedidosVenta(i).IDPedido
        Next

        Dim IDPedidosVentaObject(IDs.Length) As Object
        IDs.CopyTo(IDPedidosVentaObject, 0)
        Dim fPedidos As New Filter
        fPedidos.Add(New InListFilterItem("IDPedido", IDPedidosVentaObject, FilterType.Numeric))
        Dim cabeceras As DataTable = New PedidoVentaCabecera().Filter(fPedidos)
        If cabeceras.Rows.Count > 0 Then
            Dim dtDisponibilidad As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf CrearEstructuraDisponibilidad, Nothing, services)
            cabeceras.DefaultView.Sort = "IDPedido"
            For Each pedido As DataPrcPropuestaPedidoCompraDesdePedidoVenta In pedidosVenta
                Dim i As Integer = cabeceras.DefaultView.Find(pedido.IDPedido)
                If i >= 0 Then
                    Dim cabecera As DataRow = cabeceras.DefaultView(i).Row
                    Dim datosBase As DataRow = dtDisponibilidad.NewRow()
                    datosBase("IDPedido") = cabecera("IDPedido")
                    datosBase("NPedido") = cabecera("NPedido")
                    datosBase("IDCliente") = cabecera("IDCliente")
                    datosBase("IDDireccionEnvio") = cabecera("IDDireccionEnvio")
                    datosBase("IDCentroGestion") = cabecera("IDCentroGestion")

                    If Not cabecera.IsNull("IDDireccionEnvio") Then
                        Dim direccion As DataTable = New ClienteDireccion().SelOnPrimaryKey(cabecera("IDDireccionEnvio"))
                        If direccion.Rows.Count > 0 Then
                            datosBase("IDAlmacenEnvio") = direccion.Rows(0)("IDAlmacen")
                        End If
                    End If

                    Dim f1 As New Filter
                    f1.Add(New NumberFilterItem("IDPedido", cabecera("IDPedido")))
                    Dim f2 As New Filter(FilterUnionOperator.Or)
                    f2.Add(New NumberFilterItem("Estado", enumpvlEstado.pvlPedido))
                    f2.Add(New NumberFilterItem("Estado", enumpvlEstado.pvlParcServido))
                    Dim f3 As New Filter
                    f3.Add(f1)
                    f3.Add(f2)
                    Dim lineas As DataTable = New PedidoVentaLinea().Filter(f3)
                    If lineas.Rows.Count > 0 Then
                        If pedido.PedidoCompleto Then
                            For Each linea As DataRow In lineas.Rows
                                Dim nr As DataRow = dtDisponibilidad.NewRow()
                                nr.ItemArray = datosBase.ItemArray
                                nr("IDArticulo") = linea("IDArticulo")
                                nr("QPedida") = linea("QPedida")
                                nr("QServida") = linea("QServida")
                                nr("Precio") = linea("Precio")
                                nr("Dto1") = linea("Dto1")
                                nr("Dto2") = linea("Dto2")
                                nr("Dto3") = linea("Dto3")
                                nr("IDLineaPedido") = linea("IDLineaPedido")
                                nr("FechaEntrega") = linea("FechaEntrega")
                                nr("IDAlmacen") = linea("IDAlmacen")
                                nr("Qinterna2") = linea("Qinterna2")
                                nr("Factor") = linea("Factor")
                                dtDisponibilidad.Rows.Add(nr)
                            Next
                        Else
                            lineas.DefaultView.Sort = "IDLineaPedido"
                            For Each lineaPedido As DataLineaPrcPropuestaPedidoCompraDesdePedidoVenta In pedido.Lineas
                                Dim j As Integer = lineas.DefaultView.Find(lineaPedido.IDLineaPedido)
                                If j >= 0 Then
                                    Dim linea As DataRow = lineas.DefaultView(j).Row

                                    Dim nr As DataRow = dtDisponibilidad.NewRow()
                                    nr.ItemArray = datosBase.ItemArray
                                    nr("IDArticulo") = linea("IDArticulo")
                                    nr("QPedida") = lineaPedido.QPedida
                                    'pendiente revisar las cantidades
                                    nr("QServida") = linea("QServida")
                                    nr("Precio") = linea("Precio")
                                    nr("Dto1") = linea("Dto1")
                                    nr("Dto2") = linea("Dto2")
                                    nr("Dto3") = linea("Dto3")
                                    nr("IDLineaPedido") = linea("IDLineaPedido")
                                    nr("FechaEntrega") = linea("FechaEntrega")
                                    nr("IDAlmacen") = linea("IDAlmacen")
                                    nr("Qinterna2") = linea("Qinterna2")
                                    nr("Factor") = linea("Factor")
                                    dtDisponibilidad.Rows.Add(nr)
                                End If
                            Next
                        End If
                    End If
                End If
            Next
            Return dtDisponibilidad
        End If
    End Function

    <Task()> Public Shared Function CrearEstructuraDisponibilidad(ByVal data As Object, ByVal services As ServiceProvider) As DataTable
        Dim disponibilidad As New DataTable
        disponibilidad.RemotingFormat = SerializationFormat.Binary
        disponibilidad.Columns.Add("IDPedido", GetType(Integer))
        disponibilidad.Columns.Add("NPedido", GetType(String))
        disponibilidad.Columns.Add("IDCliente", GetType(String))
        disponibilidad.Columns.Add("IDDireccionEnvio", GetType(Integer))
        disponibilidad.Columns.Add("IDAlmacenEnvio", GetType(String))
        disponibilidad.Columns.Add("IDCentroGestion", GetType(String))
        disponibilidad.Columns.Add("IDArticulo", GetType(String))
        disponibilidad.Columns.Add("QPedida", GetType(Double))
        disponibilidad.Columns.Add("QServida", GetType(Double))
        disponibilidad.Columns.Add("Precio", GetType(Double))
        disponibilidad.Columns.Add("Dto1", GetType(Double))
        disponibilidad.Columns.Add("Dto2", GetType(Double))
        disponibilidad.Columns.Add("Dto3", GetType(Double))
        disponibilidad.Columns.Add("IDLineaPedido", GetType(Integer))
        disponibilidad.Columns.Add("FechaEntrega", GetType(Date))
        disponibilidad.Columns.Add("IDAlmacen", GetType(String))
        disponibilidad.Columns.Add("QInterna2", GetType(Double))
        disponibilidad.Columns.Add("Factor", GetType(Double))
        Return disponibilidad
    End Function

    <Task()> Public Shared Function PropuestaPedidoCompra(ByVal DisponibilidadPedidoVenta As DataTable, ByVal services As ServiceProvider) As DataTable
        '//Utilizada desde:
        '//1.la consulta de disponibilidad de pedidos de venta
        '//2.desde la otra sobrecarga
        If Not DisponibilidadPedidoVenta Is Nothing Then
            Dim propuestas As DataTable = New PedidoCompraLinea().AddNew
            ProcessServer.ExecuteTask(Of DataTable, DataTable)(AddressOf AddCamposPropuesta, propuestas, services)

            For Each lineaPedido As DataRow In DisponibilidadPedidoVenta.Rows
                '//Controlar que no se genere otro pedido de compra si un pedido ya ha generado una linea de compra
                '//(independientemente de que el pedido sea EntregaGrupo o no)
                Dim control As DataTable = New GRPPedidoVentaCompraLinea().TrazaPVLPrincipal(lineaPedido("IDLineaPedido"))
                If control.Rows.Count = 0 OrElse control.Rows(0).IsNull("IDPCPrincipal") Then
                    Dim drNewRow As DataRow = propuestas.NewRow

                    Dim datCopia As New DataDocRowOrigen(Nothing, lineaPedido, drNewRow)
                    ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarDatosProveedorPrincipal, datCopia, services)
                    ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarValoresPredeterminados, datCopia, services)
                    ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarArticulo, datCopia, services)
                    ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarCantidad, datCopia, services)
                    ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarCantidadSegundaUnidad, datCopia, services)
                    ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarPrecioPedidoVenta, datCopia, services)
                    ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarPrecio, datCopia, services)
                    ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarFechaEntrega, datCopia, services)
                    ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarCentroGestion, datCopia, services)
                    ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarAlmacen, datCopia, services)
                    ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarCamposPropuesta, datCopia, services)
                    propuestas.Rows.Add(drNewRow)
                End If
            Next

            Return propuestas
        End If
    End Function

    <Task()> Public Shared Function AddCamposPropuesta(ByVal data As DataTable, ByVal services As ServiceProvider) As DataTable
        If Not data Is Nothing Then
            data.Columns.Add("IDProveedor", GetType(String))
            data.Columns.Add("DescProveedor", GetType(String))
            data.Columns.Add("EmpresaGrupo", GetType(Boolean))
            data.Columns.Add("EntregaProveedor", GetType(Boolean))
            data.Columns.Add("BaseDatos", GetType(Guid))
            data.Columns.Add("IDPedidoVenta", GetType(Integer))
            data.Columns.Add("NPedidoVenta", GetType(String))
            data.Columns.Add("IDCliente", GetType(String))
            data.Columns.Add("IDDireccion", GetType(Integer))
            data.Columns.Add("IDLineaPedidoVenta", GetType(Integer))
            data.Columns.Add("PrecioPedidoVenta", GetType(Double))
            data.Columns.Add("Dto1PedidoVenta", GetType(Double))
            data.Columns.Add("Dto2PedidoVenta", GetType(Double))
            data.Columns.Add("Dto3PedidoVenta", GetType(Double))
            data.Columns.Add("DescAlmacen", GetType(String))
            data.Columns.Add("IDUDInterna2", GetType(String))
        End If
    End Function

    <Task()> Public Shared Sub AsignarValoresPredeterminados(ByVal data As DataDocRowOrigen, ByVal services As ServiceProvider)
        data.RowDestino("Precio") = 0
        data.RowDestino("Dto1") = 0
        data.RowDestino("Dto2") = 0
        data.RowDestino("Dto3") = 0
        data.RowDestino("Dto") = 0
        data.RowDestino("DtoProntoPago") = 0
        data.RowDestino("Estado") = enumpclEstado.pclpedido
    End Sub

    <Task()> Public Shared Sub AsignarArticulo(ByVal data As DataDocRowOrigen, ByVal services As ServiceProvider)
        Dim context As New BusinessData
        context("Fecha") = Today
        context("IDProveedor") = data.RowDestino("IDProveedor")
        context("MensajeFaltaProveedor") = False
        Dim PCL As New PedidoCompraLinea
        PCL.ApplyBusinessRule("IDArticulo", data.RowOrigen("IDArticulo"), data.RowDestino, context)
    End Sub

    <Task()> Public Shared Sub AsignarCantidad(ByVal data As DataDocRowOrigen, ByVal services As ServiceProvider)
        Dim context As New BusinessData
        context("Fecha") = Today
        context("IDProveedor") = data.RowDestino("IDProveedor")

        Dim QPendiente As Double = data.RowOrigen("QPedida") - data.RowOrigen("QServida")
        If data.RowOrigen("QPedida") > 0 Then
            If QPendiente < 0 Then
                QPendiente = 0
            End If
        Else
            If QPendiente > 0 Then
                QPendiente = 0
            End If
        End If
        '//Convertimos la QPendiente en Unidad de Venta a QPedida en Unidad de Compra
        Dim QInterna As Double = QPendiente * Nz(data.RowOrigen("Factor"), 1)
        Dim QPedidaCompra As Double = QInterna / Nz(data.RowDestino("Factor"), 1)
        Dim PCL As New PedidoCompraLinea
        PCL.ApplyBusinessRule("QPedida", QPedidaCompra, data.RowDestino, context)
    End Sub

    <Task()> Public Shared Sub AsignarCantidadSegundaUnidad(ByVal data As DataDocRowOrigen, ByVal services As ServiceProvider)
        If ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, data.RowDestino("IDArticulo"), services) Then
            If Length(data.RowOrigen("QInterna2")) = 0 Then
                ApplicationService.GenerateError("El Articulo {0} se gestiona con Doble Unidad. Debe indicar la misma.", Quoted(data.RowDestino("IDArticulo")))
            Else
                data.RowDestino("QInterna2") = data.RowOrigen("QInterna2")
            End If
        End If
    End Sub


    <Task()> Public Shared Sub AsignarPrecioPedidoVenta(ByVal data As DataDocRowOrigen, ByVal services As ServiceProvider)
        data.RowDestino("PrecioPedidoVenta") = data.RowOrigen("Precio")
        data.RowDestino("Dto1PedidoVenta") = data.RowOrigen("Dto1")
        data.RowDestino("Dto2PedidoVenta") = data.RowOrigen("Dto2")
        data.RowDestino("Dto3PedidoVenta") = data.RowOrigen("Dto3")
    End Sub

    <Task()> Public Shared Sub AsignarPrecio(ByVal data As DataDocRowOrigen, ByVal services As ServiceProvider)
        Dim context As New BusinessData
        context("Fecha") = Today
        context("IDProveedor") = data.RowDestino("IDProveedor")
        Dim PCL As New PedidoCompraLinea
        If data.RowDestino("EmpresaGrupo") AndAlso data.RowDestino("EntregaProveedor") Then
            PCL.ApplyBusinessRule("Precio", Nz(data.RowDestino("PrecioPedidoVenta"), 0), data.RowDestino, context)
            PCL.ApplyBusinessRule("Dto1", Nz(data.RowDestino("Dto1PedidoVenta"), 0), data.RowDestino, context)
            PCL.ApplyBusinessRule("Dto2", Nz(data.RowDestino("Dto2PedidoVenta"), 0), data.RowDestino, context)
            PCL.ApplyBusinessRule("Dto3", Nz(data.RowDestino("Dto3PedidoVenta"), 0), data.RowDestino, context)
        ElseIf data.RowDestino("EmpresaGrupo") AndAlso Not data.RowDestino("EntregaProveedor") Then
            data.RowDestino("Dto1") = 0
            data.RowDestino("Dto2") = 0
            data.RowDestino("Dto3") = 0
            data.RowDestino("Dto") = 0
            data.RowDestino("DtoProntoPago") = 0
            PCL.ApplyBusinessRule("Precio", 0, data.RowDestino, context)
        End If
    End Sub

    <Task()> Public Shared Sub AsignarPrecioTarifa(ByVal data As DataDocRowOrigen, ByVal services As ServiceProvider)
        Dim context As New BusinessData
        context("Fecha") = Today
        context("IDProveedor") = data.RowDestino("IDProveedor")
        Dim PCL As New PedidoCompraLinea
        PCL.ApplyBusinessRule("QPedida", Nz(data.RowDestino("QPedida"), 0), data.RowDestino, context)
    End Sub


    <Task()> Public Shared Sub AsignarDatosProveedorPrincipal(ByVal data As DataDocRowOrigen, ByVal services As ServiceProvider)
        Dim f As New Filter
        f.Add(New StringFilterItem("IDArticulo", data.RowOrigen("IDArticulo")))
        f.Add(New BooleanFilterItem("Principal", True))
        Dim proveedorPrincipal As DataTable = New ArticuloProveedor().Filter(f)
        If proveedorPrincipal.Rows.Count > 0 Then
            Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
            Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(proveedorPrincipal.Rows(0)("IDProveedor"))
            data.RowDestino("IDProveedor") = ProvInfo.IDProveedor
            data.RowDestino("DescProveedor") = ProvInfo.DescProveedor
            data.RowDestino("EmpresaGrupo") = ProvInfo.EmpresaGrupo
            '//este campo en un principio toma el mismo valor que EmpresaGrupo
            data.RowDestino("EntregaProveedor") = ProvInfo.EmpresaGrupo
            data.RowDestino("BaseDatos") = ProvInfo.BaseDatos
        Else
            data.RowDestino("EmpresaGrupo") = False
            data.RowDestino("EntregaProveedor") = False
        End If
    End Sub

    <Task()> Public Shared Sub AsignarCamposPropuesta(ByVal data As DataDocRowOrigen, ByVal services As ServiceProvider)
        data.RowDestino("IDPedidoVenta") = data.RowOrigen("IDPedido")
        data.RowDestino("NPedidoVenta") = data.RowOrigen("NPedido")
        data.RowDestino("IDCliente") = data.RowOrigen("IDCliente")
        data.RowDestino("IDDireccion") = data.RowOrigen("IDDireccionEnvio")
        data.RowDestino("IDLineaPedidoVenta") = data.RowOrigen("IDLineaPedido")
        If Length(data.RowDestino("IDAlmacen")) > 0 Then
            Dim Almacenes As EntityInfoCache(Of AlmacenInfo) = services.GetService(Of EntityInfoCache(Of AlmacenInfo))()
            Dim AlmInfo As AlmacenInfo = Almacenes.GetEntity(data.RowDestino("IDAlmacen"))
            data.RowDestino("DescAlmacen") = AlmInfo.DescAlmacen
        End If
    End Sub

    <Task()> Public Shared Sub AsignarFechaEntrega(ByVal data As DataDocRowOrigen, ByVal services As ServiceProvider)
        data.RowDestino("FechaEntrega") = data.RowOrigen("FechaEntrega")
    End Sub

    <Task()> Public Shared Sub AsignarCentroGestion(ByVal data As DataDocRowOrigen, ByVal services As ServiceProvider)
        data.RowDestino("IDCentroGestion") = data.RowOrigen("IDCentroGestion")
    End Sub

    <Task()> Public Shared Sub AsignarAlmacen(ByVal data As DataDocRowOrigen, ByVal services As ServiceProvider)
        data.RowDestino("IDAlmacen") = data.RowOrigen("IDAlmacen")
    End Sub

End Class
