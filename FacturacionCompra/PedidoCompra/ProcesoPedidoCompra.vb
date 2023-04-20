Public Class ProcesoPedidoCompra

#Region " Creación de instancias de documento "

    <Task()> Public Shared Function CrearDocumentoPedidoCompra(ByVal data As PedCabCompra, ByVal services As ServiceProvider) As DocumentoPedidoCompra
        Dim Doc As DocumentoPedidoCompra
        Select Case data.Origen
            Case enumOrigenPedidoCompra.Programa
                If Length(CType(data, PedCabCompraProgramaCompra).IDPedido) = 0 Then
                    Doc = New DocumentoPedidoCompra(data, services)
                Else
                    Doc = New DocumentoPedidoCompra(CType(data, PedCabCompraProgramaCompra).IDPedido)
                    Doc.Cabecera = CType(data, PedCabCompraProgramaCompra) 'Debemos indicar la Cabecera en este caso
                End If
                'Case enumOrigenPedidoCompra.OfertaCompra
                '    Doc = New DocumentoPedidoCompra(data, services)
            Case Else
                Doc = New DocumentoPedidoCompra(data, services)
        End Select
        Return Doc
    End Function
    <Task()> Public Shared Function CrearDocumento(ByVal UpdtCtx As UpdatePackage, ByVal services As ServiceProvider) As DocumentoPedidoCompra
        Return New DocumentoPedidoCompra(UpdtCtx)
    End Function

#End Region


#Region " Agrupación "

    <Task()> Public Shared Function GetGroupColumns(ByVal table As DataTable, ByVal services As ServiceProvider) As DataColumn()
        '//Se definen las columnas que nos permitirán abrir un pedido nuevo
        Dim columns(3) As DataColumn
        columns(0) = table.Columns("IDProveedor")
        columns(1) = table.Columns("IdMoneda")
        columns(2) = table.Columns("IdFormaPago")
        columns(3) = table.Columns("IdCondicionPago")
        If table.Columns.Contains("IDPedido") Then
            ReDim Preserve columns(columns.Length)
            columns(columns.Length - 1) = table.Columns("IDPedido")
        End If
        If table.Columns.Contains("IdCondicionEnvio") Then
            ReDim Preserve columns(columns.Length)
            columns(columns.Length - 1) = table.Columns("IdCondicionEnvio")
        End If
        If table.Columns.Contains("IdFormaEnvio") Then
            ReDim Preserve columns(columns.Length)
            columns(columns.Length - 1) = table.Columns("IdFormaEnvio")
        End If
        If table.Columns.Contains("IdOperario") Then
            ReDim Preserve columns(columns.Length)
            columns(columns.Length - 1) = table.Columns("IdOperario")
        End If

        Return columns
    End Function

#Region " Origen - Planificaciones "

    <Task()> Public Shared Function AgruparPlanificaciones(ByVal data As DataPrcCrearPedidoCompraPlanificacion, ByVal services As ServiceProvider) As PedCabCompraPlanif()
        If Not data.Planificaciones Is Nothing AndAlso data.Planificaciones.Rows.Count > 0 Then
            '//Rellenamos las clases que nos servirán de semilla para crear el pedido
            Dim oGrprUser As New GroupUserPlanificaciones()
            Dim GroupCols As DataColumn() = ProcessServer.ExecuteTask(Of DataTable, DataColumn())(AddressOf GetGroupColumnsPlanif, data.Planificaciones, services)

            '//Se crean los agrupadores
            Dim groupers(0) As GroupHelper
            groupers(0) = New GroupHelper(GroupCols, oGrprUser)

            '//A través de los agrupadores 
            For Each drOfertas As DataRow In data.Planificaciones.Select(Nothing, "IDProveedor")
                groupers(0).Group(drOfertas)
            Next

            Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
            For Each ped As PedCabCompraPlanif In oGrprUser.Pedidos
                Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(ped.IDProveedor)
                ped.IDMoneda = ProvInfo.IDMoneda
                ped.IDFormaPago = ProvInfo.IDFormaPago
                ped.IDCondicionPago = ProvInfo.IDCondicionPago

                For Each pedlin As PedLinCompraPlanificacion In ped.LineasOrigen
                    Dim f As New Filter
                    f.Add(New StringFilterItem("IDProveedor", ped.IDProveedor))
                    f.Add(New StringFilterItem("IDMarca", pedlin.IDMarca)) '//IDArticulo o IDArticuloIDAlmacen
                    Dim WhereProvArticuloAlmacen As String = f.Compose(New AdoFilterComposer)
                    Dim adr() As DataRow = data.Planificaciones.Select(WhereProvArticuloAlmacen)
                    If Not adr Is Nothing AndAlso adr.Length > 0 Then
                        pedlin.Cantidad = Nz(adr(0)("Cantidad"), 0)

                        If ped.DatosOrigen Is Nothing Then ped.DatosOrigen = data.Planificaciones.Clone
                        ped.DatosOrigen.ImportRow(adr(0))
                    End If
                Next
            Next

            Return oGrprUser.Pedidos
        End If
    End Function

    <Task()> Public Shared Function GetGroupColumnsPlanif(ByVal table As DataTable, ByVal services As ServiceProvider) As DataColumn()
        '//Se definen las columnas que nos permitirán abrir un pedido nuevo
        Dim columns(-1) As DataColumn
        'columns(0) = table.Columns("IdMoneda")
        'columns(1) = table.Columns("IdFormaPago")
        'columns(2) = table.Columns("IdCondicionPago")
        Dim ProcInfo As ProcessInfoPlanif = services.GetService(Of ProcessInfoPlanif)()
        If Not ProcInfo Is Nothing AndAlso ProcInfo.AgruparPorProveedor Then
            ReDim Preserve columns(columns.Length)
            columns(columns.Length - 1) = table.Columns("IDProveedor")
        Else
            ReDim Preserve columns(columns.Length)
            columns(columns.Length - 1) = table.Columns("IDMarca")
        End If

        Return columns
    End Function

#End Region

#Region " Origen - Ofertas de Compra "

    <Task()> Public Shared Function AgruparOfertasCompra(ByVal data As DataPrcCrearPedidoCompraOfertaCompra, ByVal services As ServiceProvider) As PedCabCompraOfertaCompra()
        If Not data.Ofertas Is Nothing AndAlso data.Ofertas.Length > 0 Then
            Dim strViewName As String = "vctlConsOfertasCompraOfertasCompra"

            Dim htLins As New Hashtable
            Dim values(data.Ofertas.Length - 1) As Object
            For i As Integer = 0 To data.Ofertas.Length - 1
                values(i) = data.Ofertas(i).IDLineaOferta
                htLins.Add(data.Ofertas(i).IDLineaOferta, data.Ofertas(i))
            Next

            Dim f As New Filter
            f.Add(New InListFilterItem("IDLineaOferta", values, FilterType.Numeric))
            f.Add(New NumberFilterItem("Estado", enumOfertaCabecera.ocAdjudicada))
            f.Add(New BooleanFilterItem("Adjudicado", True))
            f.Add(New IsNullFilterItem("FechaEntrega", False))
            Dim dtOfertas As DataTable = New BE.DataEngine().Filter(strViewName, f)

            '//Rellenamos las clases que nos servirán de semilla para crear el pedido
            Dim oGrprUser As New GroupUserOfertaCompra()

            '//Se crean los agrupadores
            Dim ColGroups() As DataColumn = ProcessServer.ExecuteTask(Of DataTable, DataColumn())(AddressOf GetGroupColumns, dtOfertas, services)
            Dim groupers(0) As GroupHelper
            groupers(0) = New GroupHelper(ColGroups, oGrprUser)

            '//A través de los agrupadores 
            For Each drOfertas As DataRow In dtOfertas.Select(Nothing, "IDProveedor, IDOferta")
                groupers(0).Group(drOfertas)
            Next

            For Each ped As PedCabCompraOfertaCompra In oGrprUser.Pedidos
                For Each pedlin As PedLinCompraOfertaCompra In ped.LineasOrigen
                    pedlin.Cantidad = DirectCast(htLins(pedlin.IDLineaOrigen), DataOfertaCompra).QOferta
                Next
            Next

            Return oGrprUser.Pedidos
        End If
    End Function

#End Region

#Region " Origen - Agrupar Programas "

    <Task()> Public Shared Function AgruparProgramas(ByVal data As DataPrcCrearPedidoCompraPrograma, ByVal services As ServiceProvider) As PedCabCompraProgramaCompra()
        If Not data.Programas Is Nothing AndAlso data.Programas.Length > 0 Then
            '//se seleccionan todas las lineas de programa a confirmar
            Dim strViewName As String = "vfrmConfirmacionProgramaCompra"

            Dim htLins As New Hashtable
            Dim values(data.Programas.Length - 1) As Object
            For i As Integer = 0 To data.Programas.Length - 1
                values(i) = data.Programas(i).IDLineaPrograma
                htLins.Add(data.Programas(i).IDLineaPrograma, data.Programas(i))
            Next


            Dim oFltr As New Filter
            oFltr.Add(New InListFilterItem("IDLineaPrograma", values, FilterType.Numeric))
            ' oFltr.Add(New NumberFilterItem("Confirmada", enumplEstadoLinea.plNoConfirmada))
            Dim dtProgramas As DataTable = New BE.DataEngine().Filter(strViewName, oFltr)

            '//Rellenamos las clases que nos servirán de semilla para crear el pedido
            Dim oGrprUser As New GroupUserProgramaCompra()

            '//Se crean los agrupadores
            Dim ColGroups() As DataColumn = ProcessServer.ExecuteTask(Of DataTable, DataColumn())(AddressOf GetGroupColumns, dtProgramas, services)
            Dim groupers(0) As GroupHelper
            groupers(0) = New GroupHelper(ColGroups, oGrprUser)

            '//A través de los agrupadores 
            For Each rwLin As DataRow In dtProgramas.Select(Nothing, "IDProveedor, IDPrograma")
                If rwLin("Confirmada") = enumplEstadoLinea.plNoConfirmada OrElse rwLin("QPrevista") > rwLin("QConfirmada") OrElse rwLin("QPrevista") > rwLin("QPendiente") Then
                    groupers(0).Group(rwLin)
                End If
            Next

            For Each ped As PedCabCompraProgramaCompra In oGrprUser.Peds
                For Each pedlin As PedLinCompraProgramaCompra In ped.LineasOrigen
                    pedlin.QConfirmada = DirectCast(htLins(pedlin.IDLineaOrigen), DataProgramaCompra).QConfirmada
                    pedlin.FechaConfirmacion = DirectCast(htLins(pedlin.IDLineaOrigen), DataProgramaCompra).FechaConfirmacion
                    pedlin.Cantidad = DirectCast(htLins(pedlin.IDLineaOrigen), DataProgramaCompra).Cantidad
                Next
            Next

            Return oGrprUser.Peds
        End If
    End Function

#End Region

#Region " Origen - Subcontrataciones "

    <Task()> Public Shared Function AgruparSubcontrataciones(ByVal data As DataPrcCrearPedidoCompraSubcontratacion, ByVal services As ServiceProvider) As PedCabCompraSubcontratacion()
        If Not data.Subcontrataciones Is Nothing AndAlso data.Subcontrataciones.Length > 0 Then
            Dim strViewName As String = "vCTLCIEnvioASubcontratacion"

            Dim htLins As New Hashtable
            Dim values(data.Subcontrataciones.Length - 1) As Object
            For i As Integer = 0 To data.Subcontrataciones.Length - 1
                values(i) = data.Subcontrataciones(i).IDOrdenRuta
                htLins.Add(data.Subcontrataciones(i).IDOrdenRuta, data.Subcontrataciones(i))
            Next

            Dim f As New Filter
            f.Add(New InListFilterItem("IDOrdenRuta", values, FilterType.Numeric))
            f.Add(New NumberFilterItem("Estado", FilterOperator.NotEqual, enumofEstado.ofeOfAnulada))
            f.Add(New NumberFilterItem("Estado", FilterOperator.NotEqual, enumofEstado.ofeTerminada))
            ' f.Add(New NumberFilterItem("QPedida", FilterOperator.NotEqual, 0))
            Dim dtSubcontrataciones As DataTable = New BE.DataEngine().Filter(strViewName, f)
            If dtSubcontrataciones.Rows.Count > 0 Then
                '//Rellenamos las clases que nos servirán de semilla para crear el pedido
                Dim AppParams As ParametroCompra = services.GetService(Of ParametroCompra)()
                Dim oGrprUser As New GroupUserSubcontrataciones(AppParams.TipoCompraSubcontratacion)

                '//Se crean los agrupadores
                Dim oDataGroup As New DataAgrupSubcontratacion(dtSubcontrataciones, Not data.AgruparPorProveedor)
                Dim cols() As DataColumn = ProcessServer.ExecuteTask(Of DataAgrupSubcontratacion, DataColumn())(AddressOf GetGroupColumnsSubcontratacion, oDataGroup, services)
                Dim groupers(0) As GroupHelper
                groupers(0) = New GroupHelper(cols, oGrprUser)

                '//A través de los agrupadores 
                Dim Orden As String
                If data.AgruparPorProveedor Then
                    Orden = "IDProveedor"
                Else
                    Orden = "IDOrden,IDProveedor"
                End If
                dtSubcontrataciones.TableName = "OrdenRuta"
                For Each drSubcontratacion As DataRow In dtSubcontrataciones.Select(Nothing, Orden)
                    groupers(0).Group(drSubcontratacion)
                Next

                For Each ped As PedCabCompraSubcontratacion In oGrprUser.Pedidos
                    For Each pedlin As PedLinCompraSubcontratacion In ped.LineasOrigen
                        pedlin.Cantidad = DirectCast(htLins(pedlin.IDLineaOrigen), DataSubcontratacion).QPedida
                        pedlin.QInterna = DirectCast(htLins(pedlin.IDLineaOrigen), DataSubcontratacion).QInterna
                        pedlin.IDUDInterna = DirectCast(htLins(pedlin.IDLineaOrigen), DataSubcontratacion).IDUDInterna
                        pedlin.IDUDProduccion = DirectCast(htLins(pedlin.IDLineaOrigen), DataSubcontratacion).IDUDProduccion
                        pedlin.FechaEntrega = DirectCast(htLins(pedlin.IDLineaOrigen), DataSubcontratacion).FechaEntrega
                    Next
                Next

                Return oGrprUser.Pedidos
            End If

        End If
    End Function

    Public Class DataAgrupSubcontratacion
        Public Datos As DataTable
        Public AgruparPorOF As Boolean

        Public Sub New(ByVal Datos As DataTable, ByVal AgruparPorOF As Boolean)
            Me.Datos = Datos
            Me.AgruparPorOF = AgruparPorOF
        End Sub
    End Class
    <Task()> Public Shared Function GetGroupColumnsSubcontratacion(ByVal data As DataAgrupSubcontratacion, ByVal services As ServiceProvider) As DataColumn()
        '//Se definen las columnas que nos permitirán abrir un pedido nuevo
        Dim columns(3) As DataColumn
        columns(0) = data.Datos.Columns("IDProveedor")
        columns(1) = data.Datos.Columns("IdMoneda")
        columns(2) = data.Datos.Columns("IdFormaPago")
        columns(3) = data.Datos.Columns("IdCondicionPago")
        'columns(4) = data.Datos.Columns("IDPedido")
        If data.AgruparPorOF Then
            ReDim Preserve columns(4)
            columns(4) = data.Datos.Columns("IDOrden")
        End If

        Return columns
    End Function

#End Region

#Region " Origen - Solicitudes de Compra "

    <Task()> Public Shared Function AgruparSolicitudesCompra(ByVal data As DataPrcCrearPedidoCompraSolicitudCompra, ByVal services As ServiceProvider) As PedCabCompraSolicitudCompra()
        If Not data.Solicitudes Is Nothing AndAlso data.Solicitudes.Length > 0 Then
            Dim htLins As New Hashtable
            Dim values(data.Solicitudes.Length - 1) As Object
            For i As Integer = 0 To data.Solicitudes.Length - 1
                values(i) = data.Solicitudes(i).IDLineaSolicitud
                htLins.Add(data.Solicitudes(i).IDLineaSolicitud, data.Solicitudes(i))
            Next

            Dim f As New Filter
            f.Add(New InListFilterItem("IDLineaSolicitud", values, FilterType.Numeric))
            f.Add(New NumberFilterItem("Estado", FilterOperator.NotEqual, enumscEstado.scCerrada))
            f.Add(New NumberFilterItem("Estado", FilterOperator.NotEqual, enumscEstado.scDenegada))
            Dim dtSolicitudes As DataTable = New BE.DataEngine().Filter("vCIGestionDeSolicitudes", f)
            If dtSolicitudes.Rows.Count > 0 Then
                '//Rellenamos las clases que nos servirán de semilla para crear el pedido
                Dim oGrprUser As New GroupUserSolicitudes

                '//Se crean los agrupadores
                Dim cols() As DataColumn = ProcessServer.ExecuteTask(Of DataTable, DataColumn())(AddressOf GetGroupColumns, dtSolicitudes, services)
                Dim groupers(0) As GroupHelper
                groupers(0) = New GroupHelper(cols, oGrprUser)

                '//A través de los agrupadores 
                For Each drSolicitud As DataRow In dtSolicitudes.Select(Nothing, "IDProveedor,IDSolicitud")
                    groupers(0).Group(drSolicitud)
                Next

                For Each ped As PedCabCompraSolicitudCompra In oGrprUser.Pedidos
                    For Each pedlin As PedLinCompraSolicitudCompra In ped.LineasOrigen
                        pedlin.Cantidad = DirectCast(htLins(pedlin.IDLineaOrigen), DataSolicitudCompra).QSolicitar
                    Next
                Next

                Return oGrprUser.Pedidos
            End If

        End If
    End Function

#End Region

#Region " Origen - Obras "

    <Task()> Public Shared Function AgruparObras(ByVal data As DataPrcCrearPedidoCompraObras, ByVal services As ServiceProvider) As PedCabCompra()
        If Not data.Obras Is Nothing AndAlso data.Obras.Length > 0 Then
            Dim htLins As New Hashtable
            Dim values(data.Obras.Length - 1) As Object
            For i As Integer = 0 To data.Obras.Length - 1
                values(i) = data.Obras(i).IDOrigen
                htLins.Add(data.Obras(i).IDOrigen, data.Obras(i))
            Next

            Dim ViewName As String : Dim FieldName As String
            If data.PorTrabajo Then
                FieldName = "IDTrabajo"
                ViewName = "vFrmMntoObraGeneraCompraTrabajo"
            ElseIf data.PorMaterial Then
                FieldName = "IDLineaMaterial"
                ViewName = "vFrmMntoObraGeneraCompra"
            End If
            Dim f As New Filter
            f.Add(New InListFilterItem(FieldName, values, FilterType.Numeric))
            Dim dtObras As DataTable = New BE.DataEngine().Filter(ViewName, f)
            If dtObras.Rows.Count > 0 Then
                '//Rellenamos las clases que nos servirán de semilla para crear el pedido
                Dim oGrprUser As New GroupUserCompraObras(data.PorMaterial, data.PorTrabajo)

                '//Se crean los agrupadores
                Dim cols() As DataColumn = ProcessServer.ExecuteTask(Of DataTable, DataColumn())(AddressOf GetGroupColumns, dtObras, services)
                Dim groupers(0) As GroupHelper
                groupers(0) = New GroupHelper(cols, oGrprUser)


                '//A través de los agrupadores 
                For Each drObra As DataRow In dtObras.Select(Nothing, "IDProveedor")
                    groupers(0).Group(drObra)
                Next

                For Each ped As PedCabCompraObra In oGrprUser.Pedidos
                    For Each pedlin As PedLinCompra In ped.LineasOrigen
                        pedlin.Cantidad = DirectCast(htLins(pedlin.IDLineaOrigen), DataOrigenPC).QPedir
                    Next
                Next

                Return oGrprUser.Pedidos
            End If

        End If
    End Function

#End Region

#Region " Origen - Mantenimiento "

    <Task()> Public Shared Function AgruparMantenimiento(ByVal data As DataPrcCrearPedidoCompraMantenimiento, ByVal services As ServiceProvider) As PedCabCompraMantenimiento()
        If Not data.Preventivos Is Nothing AndAlso data.Preventivos.Length > 0 Then
            Dim htLins As New Hashtable
            Dim values(data.Preventivos.Length - 1) As Object
            For i As Integer = 0 To data.Preventivos.Length - 1
                values(i) = data.Preventivos(i).IDOrigen
                htLins.Add(data.Preventivos(i).IDOrigen, data.Preventivos(i))
            Next

            Dim f As New Filter
            f.Add(New InListFilterItem("IDMntoOTPrev", values, FilterType.Numeric))
            Dim dtPreventivos As DataTable = New BE.DataEngine().Filter("vFrmMntoOTGeneraCompra", f)
            If dtPreventivos.Rows.Count > 0 Then
                '//Rellenamos las clases que nos servirán de semilla para crear el pedido
                Dim oGrprUser As New GroupUserMantenimiento

                '//Se crean los agrupadores
                Dim cols() As DataColumn = ProcessServer.ExecuteTask(Of DataTable, DataColumn())(AddressOf GetGroupColumnsMantenimiento, dtPreventivos, services)
                Dim groupers(0) As GroupHelper
                groupers(0) = New GroupHelper(cols, oGrprUser)

                '//A través de los agrupadores 
                For Each drPreventivo As DataRow In dtPreventivos.Select(Nothing, "IDProveedor")
                    groupers(0).Group(drPreventivo)
                Next

                For Each ped As PedCabCompraMantenimiento In oGrprUser.Pedidos
                    For Each pedlin As PedLinCompraMantenimiento In ped.LineasOrigen
                        pedlin.Cantidad = DirectCast(htLins(pedlin.IDLineaOrigen), DataOrigenPC).QPedir
                    Next
                Next

                Return oGrprUser.Pedidos
            End If

        End If
    End Function
    <Task()> Public Shared Function GetGroupColumnsMantenimiento(ByVal table As DataTable, ByVal services As ServiceProvider) As DataColumn()
        '//Se definen las columnas que nos permitirán abrir un pedido nuevo
        Dim columns(0) As DataColumn
        columns(0) = table.Columns("IDProveedor")

        Return columns
    End Function

#End Region

#Region " Origen - Ofertas Comerciales  "

    <Task()> Public Shared Function AgruparOfertasComerciales(ByVal data As DataPrcCrearPedidoOfertaComercial, ByVal services As ServiceProvider) As PedCabCompraOfertaComercial()
        If data.Detalle Then
            Return ProcessServer.ExecuteTask(Of DataPrcCrearPedidoOfertaComercial, PedCabCompraOfertaComercial())(AddressOf AgruparOfertaComercialDetalle, data, services)
        Else
            Return ProcessServer.ExecuteTask(Of DataPrcCrearPedidoOfertaComercial, PedCabCompraOfertaComercial())(AddressOf AgruparOfertaComercial, data, services)
        End If
    End Function
    <Task()> Public Shared Function AgruparOfertaComercial(ByVal data As DataPrcCrearPedidoOfertaComercial, ByVal services As ServiceProvider) As PedCabCompraOfertaComercial()
        If Not data.Ofertas Is Nothing AndAlso data.Ofertas.Length > 0 Then
            Dim htLins As New Hashtable
            Dim values(data.Ofertas.Length - 1) As Object
            For i As Integer = 0 To data.Ofertas.Length - 1
                values(i) = data.Ofertas(i).IDOfertaComercial
                htLins.Add(values(i), data.Ofertas(i))
            Next

            Dim oFltr As New Filter
            oFltr.Add(New InListFilterItem("IDOfertaComercial", values, FilterType.Numeric))
            oFltr.Add(New IsNullFilterItem("IDArticulo", False))
            oFltr.Add(New BooleanFilterItem("LanzarCompra", True))
            oFltr.Add(New NumberFilterItem("EstadoCompra", enumocdEstadoCompraVenta.ecvPendiente))

            Dim strViewName As String = "vfrmOfertaComercialTratamientoCompra"
            Dim dtOfertas As DataTable = New BE.DataEngine().Filter(strViewName, oFltr)
            ProcessServer.ExecuteTask(Of DataTable)(AddressOf ValidarDatosOfertas, dtOfertas, services)

            '//Rellenamos las clases que nos servirán de semilla para crear el pedido
            Dim oGrprUser As New GroupUserCompraOfertasComerciales()

            '//Se crean los agrupadores
            Dim GroupCols() As DataColumn = ProcessServer.ExecuteTask(Of DataTable, DataColumn())(AddressOf GetGroupColumns, dtOfertas, services)
            Dim groupers(0) As GroupHelper
            groupers(enummcAgrupPedido.mcCliente) = New GroupHelper(GroupCols, oGrprUser)

            If Not dtOfertas Is Nothing AndAlso dtOfertas.Rows.Count > 0 Then
                For Each lineaDetalle As DataRow In dtOfertas.Select(Nothing, "IDProveedor, IDOfertaComercial")
                    If Length(lineaDetalle("IDArticulo")) > 0 Then
                        groupers(enummcAgrupPedido.mcCliente).Group(lineaDetalle)
                    End If
                Next
            End If

            Return oGrprUser.Peds
        End If
    End Function
    <Task()> Public Shared Function AgruparOfertaComercialDetalle(ByVal data As DataPrcCrearPedidoOfertaComercial, ByVal services As ServiceProvider) As PedCabCompraOfertaComercial()
        If Not data.Ofertas Is Nothing AndAlso data.Ofertas.Length > 0 Then
            Dim htLins As New Hashtable
            Dim values(data.Ofertas.Length - 1) As Object
            For i As Integer = 0 To data.Ofertas.Length - 1
                values(i) = data.Ofertas(i).IDLineaOfertaDetalle
                htLins(data.Ofertas(i).IDOfertaComercial) = data.Ofertas(i)
            Next

            Dim strViewName As String = "vfrmOfertaComercialTratamientoCompra"
            Dim oFltr As New Filter
            oFltr.Add(New InListFilterItem("IDLineaOfertaDetalle", values, FilterType.Numeric))
            oFltr.Add(New IsNullFilterItem("IDArticulo", False))
            oFltr.Add(New BooleanFilterItem("LanzarCompra", True))
            oFltr.Add(New NumberFilterItem("EstadoCompra", enumocdEstadoCompraVenta.ecvPendiente))
            Dim dtOfertas As DataTable = New BE.DataEngine().Filter(strViewName, oFltr)
            ProcessServer.ExecuteTask(Of DataTable)(AddressOf ValidarDatosOfertas, dtOfertas, services)

            '//Rellenamos las clases que nos servirán de semilla para crear el pedido
            Dim oGrprUser As New GroupUserCompraOfertasComerciales()

            '//Se crean los agrupadores
            Dim GroupCols() As DataColumn = ProcessServer.ExecuteTask(Of DataTable, DataColumn())(AddressOf GetGroupColumns, dtOfertas, services)

            Dim groupers(0) As GroupHelper
            groupers(enummcAgrupPedido.mcCliente) = New GroupHelper(GroupCols, oGrprUser)

            If Not dtOfertas Is Nothing AndAlso dtOfertas.Rows.Count > 0 Then
                For Each ofertaDetalle As DataRow In dtOfertas.Select(Nothing, "IDProveedor, IDOfertaComercial")
                    groupers(enummcAgrupPedido.mcCliente).Group(ofertaDetalle)
                Next
            End If


            Return oGrprUser.Peds
        End If
    End Function

#End Region

#Region " Origen - Pedido Venta "

    <Task()> Public Shared Function AgruparPedidosVenta(ByVal data As DataPrcCrearPedidoCompraDesdePedidoVenta, ByVal services As ServiceProvider) As PedCabCompraPedidoVenta()
        Dim htLin As New Hashtable : Dim IDLineasPedidoVenta(-1) As Object

        For Each propuesta As DataRow In data.Propuestas.Select
            ReDim Preserve IDLineasPedidoVenta(IDLineasPedidoVenta.Length)
            IDLineasPedidoVenta(IDLineasPedidoVenta.Length - 1) = propuesta("IDLineaPedidoVenta")
        Next

        Dim fLineasPV As New Filter
        fLineasPV.Add(New InListFilterItem("IDLineaPedido", IDLineasPedidoVenta, FilterType.Numeric))
        Dim dtOrigenDatos As DataTable = New BE.DataEngine().Filter("vNegDisponibilidadPedidoVenta", fLineasPV)
        dtOrigenDatos.Columns.Add("EmpresaGrupo", GetType(Boolean))
        dtOrigenDatos.Columns.Add("EntregaProveedor", GetType(Boolean))
        dtOrigenDatos.Columns.Add("IDProveedor", GetType(String))
        dtOrigenDatos.Columns.Add("BaseDatos", GetType(Guid))
        dtOrigenDatos.Columns.Add("QInterna2", GetType(Double))

        For Each linea As DataRow In dtOrigenDatos.Rows
            Dim adrPropuesta() As DataRow = data.Propuestas.Select("IDLineaPedidoVenta=" & linea("IDLineaPedido"))
            If Not adrPropuesta Is Nothing AndAlso adrPropuesta.Length > 0 Then
                linea("EmpresaGrupo") = adrPropuesta(0)("EmpresaGrupo")
                linea("EntregaProveedor") = adrPropuesta(0)("EntregaProveedor")
                linea("IDProveedor") = adrPropuesta(0)("IDProveedor")
                linea("BaseDatos") = adrPropuesta(0)("BaseDatos")
                linea("QInterna2") = adrPropuesta(0)("QInterna2")
                linea("IDUDMedida") = adrPropuesta(0)("IDUDMedida")
                linea("Factor") = adrPropuesta(0)("Factor") 'Factor Compra 
                linea("QPedida") = adrPropuesta(0)("QPedida")
                linea("Precio") = adrPropuesta(0)("Precio")
                linea("Dto1") = adrPropuesta(0)("Dto1")
                linea("Dto2") = adrPropuesta(0)("Dto2")
                linea("Dto3") = adrPropuesta(0)("Dto3")
                linea("SeguimientoTarifa") = adrPropuesta(0)("SeguimientoTarifa")
                Dim datPrecio As New DataAsignarPrecioPedidoCompra(adrPropuesta(0)("EmpresaGrupo"), adrPropuesta(0)("EntregaProveedor"), linea)
                ProcessServer.ExecuteTask(Of DataAsignarPrecioPedidoCompra)(AddressOf AsignarPrecioPedidoCompra, datPrecio, services)
            End If
            htLin(linea("IDLineaPedido")) = linea
        Next

        '//Pedidos de compra que se agrupan por proveedor-direccion
        Dim oGrprUser As New GroupUserPCPedidosVenta
        Dim f As New Filter
        f.Add(New BooleanFilterItem("EmpresaGrupo", True))
        f.Add(New BooleanFilterItem("EntregaProveedor", True))
        Dim datProvDir As New DataGetGroupColumnsPV(dtOrigenDatos, True)
        Dim grpProvDir As DataColumn() = ProcessServer.ExecuteTask(Of DataGetGroupColumnsPV, DataColumn())(AddressOf GetGroupColumnsPV, datProvDir, services)
        Dim groupers(0) As GroupHelper
        groupers(0) = New GroupHelper(grpProvDir, oGrprUser)
        Dim WhereProvDir As String = f.Compose(New AdoFilterComposer)
        For Each linea As DataRow In dtOrigenDatos.Select(WhereProvDir)
            groupers(0).Group(linea)
        Next

        '//Pedidos de compra que se agrupan unicamente por proveedor
        Dim f1 As New Filter(FilterUnionOperator.Or)
        f1.Add(New BooleanFilterItem("EmpresaGrupo", False))
        f1.Add(New BooleanFilterItem("EntregaProveedor", False))
        Dim datProv As New DataGetGroupColumnsPV(dtOrigenDatos, False)
        Dim grpProv As DataColumn() = ProcessServer.ExecuteTask(Of DataGetGroupColumnsPV, DataColumn())(AddressOf GetGroupColumnsPV, datProv, services)
        groupers(0) = New GroupHelper(grpProv, oGrprUser)
        Dim WhereProv As String = f1.Compose(New AdoFilterComposer)
        For Each linea As DataRow In dtOrigenDatos.Select(WhereProv)
            groupers(0).Group(linea)
        Next

        For Each ped As PedCabCompraPedidoVenta In oGrprUser.Pedidos
            For Each pedlin As PedLinCompraPedidoVenta In ped.LineasOrigen
                pedlin.Cantidad = Nz(htLin(pedlin.IDLineaOrigen)("QPedida"), 0)
                pedlin.Cantidad2 = CDbl(Nz(htLin(pedlin.IDLineaOrigen)("QInterna2"), 0))
                If ped.DatosOrigen Is Nothing Then ped.DatosOrigen = dtOrigenDatos.Clone
                ped.DatosOrigen.ImportRow(htLin(pedlin.IDLineaOrigen))
            Next
        Next

        Return oGrprUser.Pedidos
    End Function
    Public Class DataGetGroupColumnsPV
        Public Propuestas As DataTable
        Public AgruparPorDireccion As Boolean
        Public Sub New(ByVal Propuestas As DataTable, ByVal AgruparPorDireccion As Boolean)
            Me.Propuestas = Propuestas
            Me.AgruparPorDireccion = AgruparPorDireccion
        End Sub
    End Class
    <Task()> Public Shared Function GetGroupColumnsPV(ByVal data As DataGetGroupColumnsPV, ByVal services As ServiceProvider) As DataColumn()
        Dim columns(0) As DataColumn
        columns(0) = data.Propuestas.Columns("IDProveedor")
        If data.AgruparPorDireccion Then
            ReDim Preserve columns(columns.Length)
            columns(1) = data.Propuestas.Columns("IDDireccionEnvio")
        End If
        Return columns
    End Function

    Public Class DataAsignarPrecioPedidoCompra
        Public EmpresaGrupo As Boolean
        Public EntregaProveedor As Boolean
        Public Row As DataRow
        Public Sub New(ByVal EmpresaGrupo As Boolean, ByVal EntregaProveedor As Boolean, ByVal Row As DataRow)
            Me.EmpresaGrupo = EmpresaGrupo
            Me.EntregaProveedor = EntregaProveedor
            Me.Row = Row
        End Sub
    End Class
    <Task()> Public Shared Sub AsignarPrecioPedidoCompra(ByVal data As DataAsignarPrecioPedidoCompra, ByVal services As ServiceProvider)
        Dim PCL As New PedidoCompraLinea
        Dim dblPrecio As Double
        If data.EmpresaGrupo AndAlso data.EntregaProveedor Then
            dblPrecio = data.Row("Precio")
            PCL.ApplyBusinessRule("Precio", dblPrecio, data.Row, New BusinessData(data.Row))
        ElseIf data.EmpresaGrupo AndAlso Not data.EntregaProveedor Then
            dblPrecio = 0
            data.Row("Dto1") = 0
            data.Row("Dto2") = 0
            data.Row("Dto3") = 0
            data.Row("Dto") = 0
            data.Row("DtoProntoPago") = 0
            PCL.ApplyBusinessRule("QPedida", data.Row("QPedida"), data.Row, New BusinessData(data.Row))
        End If
    End Sub

    <Task()> Public Shared Sub AsignarPrecioPedidoCompraTarifa(ByVal data As DataAsignarPrecioPedidoCompra, ByVal services As ServiceProvider)
        Dim PCL As New PedidoCompraLinea
        PCL.ApplyBusinessRule("QPedida", data.Row("QPedida"), data.Row, New BusinessData(data.Row))
    End Sub

#End Region

#End Region

#Region " Ordenar Pedidos "
    'Ordena los pedidos teniendo en cuenta el proveedor, moneda y condiciones de Pago
    <Task()> Public Shared Sub Ordenar(ByVal data As PedCabCompra(), ByVal services As ServiceProvider)
        If data IsNot Nothing Then Array.Sort(data, New OrdenPedidoCompra)
    End Sub
#End Region

#Region "Analítica "

    <Task()> Public Shared Sub CalcularAnalitica(ByVal Doc As DocumentoPedidoCompra, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DocumentoPedidoCompra)(AddressOf NegocioGeneral.CalcularAnalitica, Doc, services)
    End Sub
#End Region

#Region " Asignar datos "

    <Task()> Public Shared Sub AsignarValoresPredeterminadosLinea(ByVal row As DataRow, ByVal services As ServiceProvider)
        row("IdLineaPedido") = AdminData.GetAutoNumeric
        row("Estado") = enumpclEstado.pclpedido
    End Sub

    <Task()> Public Shared Sub AsignarValoresPredeterminadosGenerales(ByVal Doc As DocumentoPedidoCompra, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoComunes.AsignarIdentificadorPedido, Doc.HeaderRow, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoComunes.AsignarFechaPedido, Doc.HeaderRow, services)
        Dim HeaderRow As IPropertyAccessor = New DataRowPropertyAccessor(Doc.HeaderRow)
        ProcessServer.ExecuteTask(Of DocumentoPedidoCompra)(AddressOf AsignarFechaEntregaPedido, Doc, services)
        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ProcesoComunes.AsignarEjercicioContablePedido, HeaderRow, services)
    End Sub


    <Task()> Public Shared Sub AsignarFechaEntregaPedido(ByVal Doc As DocumentoPedidoCompra, ByVal services As ServiceProvider)
        If Not Doc.HeaderRow Is Nothing AndAlso (Not Doc.dtLineas Is Nothing AndAlso Doc.dtLineas.Rows.Count > 0) Then
            If Doc.HeaderRow.IsNull("FechaEntrega") AndAlso IsDate(Doc.dtLineas.Rows(0)("FechaEntrega")) Then
                Doc.HeaderRow("FechaEntrega") = Doc.dtLineas.Rows(0)("FechaEntrega")
            End If
        End If
    End Sub


    <Task()> Public Shared Sub AsignarNumeroPedido(ByVal Doc As DocumentoPedidoCompra, ByVal services As ServiceProvider)
        If Doc.HeaderRow.RowState = DataRowState.Added Then
            If Not Doc.HeaderRow.IsNull("IDContador") Then
                Dim StDatos As New Contador.DatosCounterValue
                StDatos.IDCounter = Doc.HeaderRow("IDContador")
                StDatos.TargetClass = New PedidoCompraCabecera
                StDatos.TargetField = "NPedido"
                StDatos.DateField = "FechaPedido"
                StDatos.DateValue = Doc.HeaderRow("FechaPedido")
                StDatos.IDEjercicio = Doc.HeaderRow("IDEjercicio") & String.Empty
                Doc.HeaderRow("NPedido") = ProcessServer.ExecuteTask(Of Contador.DatosCounterValue, String)(AddressOf Contador.CounterValue, StDatos, services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarDatosProveedor(ByVal Doc As DocumentoPedidoCompra, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DocumentoPedidoCompra)(AddressOf ProcesoCompra.AsignarDatosProveedor, Doc, services)

        If Doc.Proveedor Is Nothing Then
            Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
            Doc.Proveedor = Proveedores.GetEntity(Doc.HeaderRow("IDProveedor"))
        End If
        If Doc.Cabecera.Origen = enumOrigenPedidoCompra.Programa Then
            If Length(CType(Doc.Cabecera, PedCabCompraProgramaCompra).IDFormaEnvio) > 0 Then
                Doc.HeaderRow("IDFormaEnvio") = CType(Doc.Cabecera, PedCabCompraProgramaCompra).IDFormaEnvio
            End If
            If Length(CType(Doc.Cabecera, PedCabCompraProgramaCompra).IDCondicionEnvio) > 0 Then
                Doc.HeaderRow("IDCondicionEnvio") = CType(Doc.Cabecera, PedCabCompraProgramaCompra).IDCondicionEnvio
            End If
        End If
        If Doc.HeaderRow.IsNull("IDFormaEnvio") Then Doc.HeaderRow("IDFormaEnvio") = Doc.Proveedor.IDFormaEnvio
        If Doc.HeaderRow.IsNull("IDCondicionEnvio") Then Doc.HeaderRow("IDCondicionEnvio") = Doc.Proveedor.IDCondicionEnvio
        If Doc.HeaderRow.IsNull("IDDiaPago") Then Doc.HeaderRow("IDDiaPago") = Doc.Proveedor.IDDiaPago
        If Doc.HeaderRow.IsNull("Dto") Then Doc.HeaderRow("Dto") = Doc.Proveedor.DtoComercial
    End Sub

    <Task()> Public Shared Sub AsignarDireccionEnvio(ByVal Doc As DocumentoPedidoCompra, ByVal services As ServiceProvider)
        If Doc.HeaderRow.IsNull("IDDireccion") Then
            Dim strProveedor As String = Doc.IDProveedor
            Dim stDatosDirec As New ProveedorDireccion.DataDirecEnvio
            stDatosDirec.IDProveedor = strProveedor
            stDatosDirec.TipoDireccion = enumpdTipoDireccion.pdDireccionPedido
            Dim dtDireccion As DataTable = ProcessServer.ExecuteTask(Of ProveedorDireccion.DataDirecEnvio, DataTable)(AddressOf ProveedorDireccion.ObtenerDireccionEnvio, stDatosDirec, services)
            If Not dtDireccion Is Nothing AndAlso dtDireccion.Rows.Count > 0 Then
                Doc.HeaderRow("IDDireccion") = dtDireccion.Rows(0)("IDDireccion")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarTipoCompra(ByVal Doc As DocumentoPedidoCompra, ByVal services As ServiceProvider)
        Dim AppParamsCompra As ParametroCompra = services.GetService(Of ParametroCompra)()
        If Doc.HeaderRow.IsNull("IDTipoCompra") Then
            Select Case Doc.Cabecera.Origen
                Case enumOrigenPedidoCompra.Subcontratacion
                    Doc.HeaderRow("IDTipoCompra") = AppParamsCompra.TipoCompraSubcontratacion
                Case enumOrigenPedidoCompra.PedidoVenta
                    Doc.HeaderRow("IDTipoCompra") = AppParamsCompra.TipoCompraPedidoVenta
                Case Else
                    Doc.HeaderRow("IDTipoCompra") = AppParamsCompra.TipoCompraNormal
            End Select
        End If
    End Sub

    <Task()> Public Shared Sub AsignarEmpresaGrupo(ByVal Doc As DocumentoPedidoCompra, ByVal services As ServiceProvider)
        If Doc Is Nothing OrElse Doc.Cabecera Is Nothing Then Exit Sub
        Select Case Doc.Cabecera.Origen
            Case enumOrigenPedidoCompra.PedidoVenta
                Doc.HeaderRow("EmpresaGrupo") = CType(Doc.Cabecera, PedCabCompraPedidoVenta).Multiempresa
        End Select
    End Sub

    <Task()> Public Shared Sub AsignarEntregaProveedor(ByVal Doc As DocumentoPedidoCompra, ByVal services As ServiceProvider)
        If Doc Is Nothing OrElse Doc.Cabecera Is Nothing Then Exit Sub
        Select Case Doc.Cabecera.Origen
            Case enumOrigenPedidoCompra.PedidoVenta
                Doc.HeaderRow("EntregaProveedor") = CType(Doc.Cabecera, PedCabCompraPedidoVenta).EntregaProveedor
        End Select
    End Sub

    <Task()> Public Shared Sub AsignarPedidoVenta(ByVal Doc As DocumentoPedidoCompra, ByVal services As ServiceProvider)
        If Doc Is Nothing OrElse Doc.Cabecera Is Nothing Then Exit Sub
        Select Case Doc.Cabecera.Origen
            Case enumOrigenPedidoCompra.PedidoVenta
                Doc.HeaderRow("IDPedidoVenta") = CType(Doc.Cabecera, PedCabCompraPedidoVenta).IDOrigen
        End Select
    End Sub


    <Task()> Public Shared Sub AsignarCentroGestion(ByVal Doc As DocumentoPedidoCompra, ByVal services As ServiceProvider)
        Dim AppParamsCompra As ParametroCompra = services.GetService(Of ParametroCompra)()
        If Doc.HeaderRow.IsNull("IDCentroGestion") AndAlso Length(AppParamsCompra.General.CentroGestion) > 0 Then
            Doc.HeaderRow("IDCentroGestion") = AppParamsCompra.General.CentroGestion
        End If
    End Sub

    <Task()> Public Shared Sub AsignarOperario(ByVal Doc As DocumentoPedidoCompra, ByVal services As ServiceProvider)
        If TypeOf Doc.Cabecera Is PedCabCompraObra OrElse TypeOf Doc.Cabecera Is PedCabCompraMantenimiento Then
            Dim Info As ProcessInfoPC = services.GetService(Of ProcessInfoPC)()
            If Len(Info.IDOperario) > 0 Then
                Doc.HeaderRow("IDOperario") = Info.IDOperario
            End If
        Else
            If TypeOf Doc.Cabecera Is PedCabCompraOfertaCompra AndAlso Length(CType(Doc.Cabecera, PedCabCompraOfertaCompra).IDOperario) > 0 Then
                Doc.HeaderRow("IDOperario") = CType(Doc.Cabecera, PedCabCompraOfertaCompra).IDOperario
            ElseIf TypeOf Doc.Cabecera Is PedCabCompraProgramaCompra AndAlso Length(CType(Doc.Cabecera, PedCabCompraProgramaCompra).IDOperario) > 0 Then
                Doc.HeaderRow("IDOperario") = CType(Doc.Cabecera, PedCabCompraProgramaCompra).IDOperario
            ElseIf TypeOf Doc.Cabecera Is PedCabCompraSolicitudCompra AndAlso Length(CType(Doc.Cabecera, PedCabCompraSolicitudCompra).IDOperario) > 0 Then
                Doc.HeaderRow("IDOperario") = CType(Doc.Cabecera, PedCabCompraSolicitudCompra).IDOperario
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarDiaPago(ByVal Doc As DocumentoPedidoCompra, ByVal services As ServiceProvider)
        If Length(Doc.HeaderRow("IDDiaPago")) = 0 Then
            If TypeOf Doc.Cabecera Is PedCabCompraOfertaCompra AndAlso Length(CType(Doc.Cabecera, PedCabCompraOfertaCompra).IDDiaPago) > 0 Then
                Doc.HeaderRow("IDDiaPago") = CType(Doc.Cabecera, PedCabCompraOfertaCompra).IDDiaPago
            End If
        End If
    End Sub

    <Serializable()> _
    Public Class DataEstadoLinea
        Public IDProveedor As String
        Public Linea As DataRow
    End Class

    <Task()> Public Shared Sub AsignarEstadoLinea(ByVal data As DataEstadoLinea, ByVal services As ServiceProvider)
        If Length(data.IDProveedor) = 0 Then Exit Sub

        Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
        Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(data.IDProveedor)

        Dim QPedida As Double = Nz(data.Linea("QPedida"), 0)
        Dim QServida As Double = Nz(data.Linea("QServida"), 0)
        If ProvInfo.PorcentajeTolCierre <> 0 Then
            Dim dblPorcenResul As Double = (QServida * 100) / QPedida
            Dim dblResul As Double = dblPorcenResul - (100 - ProvInfo.PorcentajeTolCierre)
            If QServida <= 0 Then
                data.Linea("Estado") = enumpvlEstado.pvlPedido
            ElseIf dblResul >= 0 Then
                data.Linea("Estado") = enumpvlEstado.pvlCerrado
            ElseIf dblResul < 0 Then
                data.Linea("Estado") = enumpclEstado.pclparcservido
            End If
        Else
            If QServida <= 0 Then
                data.Linea("Estado") = enumpvlEstado.pvlPedido
            ElseIf QServida >= QPedida Then
                data.Linea("Estado") = enumpvlEstado.pvlServido
            ElseIf QServida < QPedida Then
                data.Linea("Estado") = enumpclEstado.pclparcservido
            End If
        End If
    End Sub


#End Region

#Region " Cálculos de Importes, Totales y Monedas "

    <Task()> Public Shared Sub ActualizarCambiosMoneda(ByVal Doc As DocumentoPedidoCompra, ByVal services As ServiceProvider)
        If Doc.HeaderRow.RowState = DataRowState.Modified Then
            If Doc.HeaderRow("IDMoneda") <> Nz(Doc.HeaderRow("IDMoneda", DataRowVersion.Original), Nothing) Then
                Dim pcl As New PedidoCompraLinea
                Dim context As New BusinessData
                context("IDMoneda") = Doc.HeaderRow("IDMoneda")
                context("Fecha") = Doc.Fecha
                For Each row As DataRow In Doc.dtLineas.Rows
                    pcl.ApplyBusinessRule("Precio", row("Precio"), row, context)
                Next
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CalcularImporteLineasPedido(ByVal Doc As DocumentoPedidoCompra, ByVal services As ServiceProvider)
        '//NO utilizar el CalcularImporteLineas de ProcesoComunes. Hay que pasar la cantidad a la QPedida y viceversa.
        For Each linea As DataRow In Doc.dtLineas.Rows
            Dim ILinea As IPropertyAccessor = New DataRowPropertyAccessor(linea)
            ILinea("Cantidad") = linea("QPedida")
            ILinea("IDMoneda") = Doc.IDMoneda
            ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf NegocioGeneral.CalcularPrecioImporte, ILinea, services)
            Dim lineaIProperty As New ValoresAyB(New DataRowPropertyAccessor(linea), Doc.IDMoneda, Doc.CambioA, Doc.CambioB)
            ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, lineaIProperty, services)
        Next
    End Sub

    <Task()> Public Shared Sub CalcularBasesImponibles(ByVal Doc As DocumentoPedidoCompra, ByVal services As ServiceProvider)
        If Not IsNothing(Doc.dtLineas) Then
            Dim desglose() As DataBaseImponible = ProcessServer.ExecuteTask(Of DocumentoPedidoCompra, DataBaseImponible())(AddressOf ProcesoComunes.DesglosarImporte, Doc, services)
            Dim datosCalculo As New ProcesoComunes.DataCalculoTotalesCab
            datosCalculo.Doc = Doc
            datosCalculo.BasesImponibles = desglose
            ProcessServer.ExecuteTask(Of ProcesoComunes.DataCalculoTotalesCab)(AddressOf CalcularTotalesCabecera, datosCalculo, services)
        End If
    End Sub

    <Task()> Public Shared Sub CalcularTotalesCabecera(ByVal data As ProcesoComunes.DataCalculoTotalesCab, ByVal services As ServiceProvider)
        Dim BaseImponibleTotal As Double = 0 : Dim ImporteLineas As Double = 0
        Dim ImporteIVATotal As Double = 0 : Dim ImporteRETotal As Double = 0
        If Not IsNothing(data.Doc.HeaderRow) Then
            If Not IsNothing(data.BasesImponibles) AndAlso data.BasesImponibles.Length > 0 Then
                Dim CondsPago As EntityInfoCache(Of CondicionPagoInfo) = services.GetService(Of EntityInfoCache(Of CondicionPagoInfo))()
                Dim CondPagoInfo As CondicionPagoInfo = CondsPago.GetEntity(data.Doc.HeaderRow("IDCondicionPago"))
                If Not IsNothing(CondPagoInfo) Then
                    data.Doc.HeaderRow("DtoProntoPago") = CondPagoInfo.DtoProntoPago
                    data.Doc.HeaderRow("RecFinan") = CondPagoInfo.RecFinan
                End If

                Dim AppParamsCompra As ParametroCompra = services.GetService(Of ParametroCompra)()
                Dim TiposIVA As EntityInfoCache(Of TipoIvaInfo) = services.GetService(Of EntityInfoCache(Of TipoIvaInfo))()
                Dim Proveedors As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
                Dim ProvInfo As ProveedorInfo = Proveedors.GetEntity(data.Doc.HeaderRow("IDProveedor"))
                For Each BI As DataBaseImponible In data.BasesImponibles
                    ImporteLineas = ImporteLineas + BI.BaseImponible
                    If Length(BI.IDTipoIva) > 0 Then
                        Dim factor As Double = 0
                        BaseImponibleTotal = BaseImponibleTotal + BI.BaseImponible
                        ' HistoricoTipoIVA
                        Dim TIVAInfo As TipoIvaInfo = TiposIVA.GetEntity(BI.IDTipoIva, data.Doc.Fecha)

                        If Length(TIVAInfo.IDTipoIVA) > 0 Then
                            '//valor por defecto
                            factor = TIVAInfo.Factor

                            '//Para los ivas especiales que no se repercuten
                            '//If TIVAInfo.SinRepercutir Then factor = TIVAInfo.IVASinRepercutir
                        End If

                        ' Dim Base As Double = BI.BaseImponibleNormal + BI.BaseImponibleEspecial
                        Dim Base As Double = BI.BaseImponible
                        ImporteIVATotal = ImporteIVATotal + Base * factor / 100
                        If AppParamsCompra.EmpresaConRecargoEquivalencia Then
                            ImporteRETotal = ImporteRETotal + Base * TIVAInfo.IVARE / 100
                        End If
                    End If
                Next
            End If
        End If

        If Nz(data.Doc.HeaderRow("RecFinan"), 0) > 0 Then
            Dim ImpLineasNormales As Double = 0
            If Not data.Doc.dtLineas Is Nothing AndAlso data.Doc.dtLineas.Rows.Count > 0 Then
                Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
                For Each linea As DataRow In data.Doc.dtLineas.Rows
                    Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(linea("IDArticulo"))
                    If Not ArtInfo.Especial Then
                        ImpLineasNormales += Nz(linea("Importe"), 0)
                    End If
                Next
            End If

            Dim ValAyBImpLin As New ValoresAyB(ImpLineasNormales, data.Doc.IDMoneda, data.Doc.CambioA, data.Doc.CambioB)
            Dim fImpLineasNormales As fImporte = ProcessServer.ExecuteTask(Of ValoresAyB, fImporte)(AddressOf NegocioGeneral.MantenimientoValoresImporteAyB, ValAyBImpLin, services)

            data.Doc.HeaderRow("ImpRecFinan") = xRound(fImpLineasNormales.Importe * data.Doc.HeaderRow("RecFinan") / 100, data.Doc.Moneda.NDecimalesImporte)
            data.Doc.HeaderRow("ImpRecFinanA") = xRound(fImpLineasNormales.ImporteA * data.Doc.HeaderRow("RecFinan") / 100, data.Doc.MonedaA.NDecimalesImporte)
            data.Doc.HeaderRow("ImpRecFinanB") = xRound(fImpLineasNormales.ImporteB * data.Doc.HeaderRow("RecFinan") / 100, data.Doc.MonedaB.NDecimalesImporte)
        End If

        data.Doc.HeaderRow("BaseImponible") = BaseImponibleTotal
        data.Doc.HeaderRow("ImpIVA") = ImporteIVATotal
        data.Doc.HeaderRow("ImpRE") = ImporteRETotal
        data.Doc.HeaderRow("ImpPedido") = ImporteLineas

        Dim ValAyB As New ValoresAyB(New DataRowPropertyAccessor(data.Doc.HeaderRow), data.Doc.IDMoneda, data.Doc.CambioA, data.Doc.CambioB)
        ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)
    End Sub

#End Region

#Region " Crear Lineas de Pedido "

    Public Class DataLineasDesdeOrigen
        Public Row As DataRow
        Public Origen As DataRow
        Public Cantidad As Double
        Public CantidadUd As Double
        Public IDUDInterna As String
        Public IDUDMedida As String
        Public Doc As DocumentoPedidoCompra
        Public PedLin As PedLinCompra

        Public Sub New(ByVal Row As DataRow, ByVal Origen As DataRow, ByVal Doc As DocumentoPedidoCompra, ByVal PedLin As PedLinCompra, ByVal Cantidad As Double)
            Me.Row = Row
            Me.Origen = Origen
            Me.Doc = Doc
            Me.PedLin = PedLin
            Me.Cantidad = Cantidad
        End Sub
    End Class

    <Task()> Public Shared Sub CrearLineasPedidoDesdeOrigen(ByVal Doc As DocumentoPedidoCompra, ByVal services As ServiceProvider)
        Dim dtOrigenPedido As DataTable = ProcessServer.ExecuteTask(Of DocumentoPedidoCompra, DataTable)(AddressOf RecuperarDatosOrigenPedido, Doc, services)
        If dtOrigenPedido Is Nothing OrElse dtOrigenPedido.Rows.Count = 0 Then Exit Sub

        Dim IntOrdenLinea As Integer = 0
        Dim PedCab As PedCabCompra = Doc.Cabecera
        For Each drOrigen As DataRow In dtOrigenPedido.Rows
            Dim PedLin As PedLinCompra = Nothing

            Select Case Doc.Cabecera.Origen
                Case enumOrigenPedidoCompra.Planificacion
                    For i As Integer = 0 To PedCab.LineasOrigen.Length - 1
                        If drOrigen("IDMarca") = CType(PedCab.LineasOrigen(i), PedLinCompraPlanificacion).IDMarca Then
                            PedLin = PedCab.LineasOrigen(i)
                            Exit For
                        End If
                    Next
                Case enumOrigenPedidoCompra.PedidoVenta
                    For i As Integer = 0 To PedCab.LineasOrigen.Length - 1
                        If drOrigen("IDLineaPedido") = PedCab.LineasOrigen(i).IDLineaOrigen Then
                            PedLin = PedCab.LineasOrigen(i)
                            Exit For
                        End If
                    Next
                Case Else
                    For i As Integer = 0 To PedCab.LineasOrigen.Length - 1
                        If drOrigen(PedCab.LineasOrigen(i).PrimaryKeyLinOrigen) = PedCab.LineasOrigen(i).IDLineaOrigen Then
                            PedLin = PedCab.LineasOrigen(i)
                            Exit For
                        End If
                    Next
            End Select

            Dim dblCantidad As Double
            If Not Double.IsNaN(PedLin.Cantidad) Then
                dblCantidad = PedLin.Cantidad
            Else
                dblCantidad = PedLin.QConfirmada
            End If

            Dim oPCL As New PedidoCompraLinea
            If Doc.dtLineas Is Nothing Then
                Dim dtLineas As DataTable = oPCL.AddNew
                Doc.Add(GetType(PedidoCompraLinea).Name, dtLineas)
            End If

            If dblCantidad <> 0 Then
                Dim drLinea As DataRow = Doc.dtLineas.NewRow
                ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarValoresPredeterminadosLinea, drLinea, services)

                Dim LinPedido As New DataLineasDesdeOrigen(drLinea, drOrigen, Doc, PedLin, dblCantidad)
                drLinea("IDPedido") = Doc.HeaderRow("IDPedido")

                ProcessServer.ExecuteTask(Of DataLineasDesdeOrigen)(AddressOf EstablecerEnlaceConOrigen, LinPedido, services)
                ProcessServer.ExecuteTask(Of DataLineasDesdeOrigen)(AddressOf AsignarArticulo, LinPedido, services)
                ProcessServer.ExecuteTask(Of DataLineasDesdeOrigen)(AddressOf AsignarDatosArticuloDesdeOrigen, LinPedido, services)
                ProcessServer.ExecuteTask(Of DataLineasDesdeOrigen)(AddressOf AsignarDatosObras, LinPedido, services)
                ProcessServer.ExecuteTask(Of DataLineasDesdeOrigen)(AddressOf AsignarCantidadesyUnidades, LinPedido, services)
                ProcessServer.ExecuteTask(Of DataLineasDesdeOrigen)(AddressOf AsignarCantidadesyUnidadesSegundaUnidad, LinPedido, services)
                ProcessServer.ExecuteTask(Of DataLineasDesdeOrigen)(AddressOf AsignarFechaEntrega, LinPedido, services)
                ProcessServer.ExecuteTask(Of DataLineasDesdeOrigen)(AddressOf AsignarAlmacen, LinPedido, services)
                ProcessServer.ExecuteTask(Of DataLineasDesdeOrigen)(AddressOf AsignarPrecio, LinPedido, services)
                ProcessServer.ExecuteTask(Of DataLineasDesdeOrigen)(AddressOf AsignarDtos, LinPedido, services)
                ProcessServer.ExecuteTask(Of DataLineasDesdeOrigen)(AddressOf AsignarContratoCompra, LinPedido, services)
                ProcessServer.ExecuteTask(Of DataLineasDesdeOrigen)(AddressOf AsignarCentroGestionDesdeOrigen, LinPedido, services)
                ProcessServer.ExecuteTask(Of DataLineasDesdeOrigen)(AddressOf AsignarTipoLinea, LinPedido, services)
                ProcessServer.ExecuteTask(Of DataLineasDesdeOrigen)(AddressOf CrearTrazabilidad, LinPedido, services)

                If drOrigen.Table.Columns.Contains("SeguimientoTarifa") AndAlso Length(drOrigen("SeguimientoTarifa")) > 0 Then drLinea("SeguimientoTarifa") = drOrigen("SeguimientoTarifa")

                If drOrigen.Table.Columns.Contains("IDOrdenLinea") Then
                    If Length(drOrigen("IDOrdenLinea")) > 0 Then
                        drLinea("IDOrdenLinea") = drOrigen("IDOrdenLinea")
                    Else
                        IntOrdenLinea += 1
                        drLinea("IDOrdenLinea") = IntOrdenLinea
                    End If
                Else
                    IntOrdenLinea += 1
                    drLinea("IDOrdenLinea") = IntOrdenLinea
                End If

                Doc.dtLineas.Rows.Add(drLinea)

                Dim datSubcont As New DataDocRow(Doc, drLinea)
                ProcessServer.ExecuteTask(Of DataDocRow)(AddressOf CrearComponentesSubcontratacion, datSubcont, services)

            End If
        Next
    End Sub

    <Task()> Public Shared Sub CrearTrazabilidad(ByVal data As DataLineasDesdeOrigen, ByVal services As ServiceProvider)
        If data.Doc Is Nothing OrElse data.Doc.dtLineas Is Nothing OrElse data.Doc.Cabecera Is Nothing Then Exit Sub

        Select Case data.Doc.Cabecera.Origen
            Case enumOrigenPedidoCompra.PedidoVenta
                Dim Cabecera As PedCabCompraPedidoVenta = CType(data.Doc.Cabecera, PedCabCompraPedidoVenta)
                Dim BBDDActual As Guid = AdminData.GetSessionInfo().DataBase.DataBaseID

                If Length(data.Row("PedidoVentaOrigen")) > 0 Then
                    Dim traza As DataRow = data.Doc.dtTrazabilidad.NewRow
                    traza("IDPVLinea") = AdminData.GetAutoNumeric
                    traza("IDPVPrincipal") = Cabecera.IDOrigen
                    traza("NPVPrincipal") = Cabecera.NOrigen
                    traza("IDClientePrincipal") = Cabecera.IDCliente
                    traza("IDLineaPVPrincipal") = data.Origen("IDLineaPedido")
                    traza("IDPCPrincipal") = data.Doc.HeaderRow("IDPedido")
                    traza("NPCPrincipal") = data.Doc.HeaderRow("NPedido")
                    traza("IDLineaPCPrincipal") = data.Row("IDLineaPedido")
                    traza("EntregaProveedor") = Cabecera.EntregaProveedor
                    traza("IDBDPrincipal") = BBDDActual
                    data.Doc.dtTrazabilidad.Rows.Add(traza)
                End If
        End Select
    End Sub

    <Task()> Public Shared Function RecuperarDatosOrigenPedido(ByVal Doc As DocumentoPedidoCompra, ByVal services As ServiceProvider) As DataTable
        Dim dtOrigenPedido As DataTable
        Select Case Doc.Cabecera.Origen
            Case enumOrigenPedidoCompra.Planificacion
                dtOrigenPedido = CType(Doc.Cabecera, PedCabCompraPlanif).DatosOrigen
            Case enumOrigenPedidoCompra.PedidoVenta
                dtOrigenPedido = CType(Doc.Cabecera, PedCabCompraPedidoVenta).DatosOrigen
            Case Else
                dtOrigenPedido = ProcessServer.ExecuteTask(Of DocumentoPedidoCompra, DataTable)(AddressOf RecuperarDatosOrigenGenerico, Doc, services)
        End Select
        Return dtOrigenPedido
    End Function

    <Task()> Public Shared Sub AsignarArticulo(ByVal LinPedido As DataLineasDesdeOrigen, ByVal services As ServiceProvider)
        Dim oPCL As New PedidoCompraLinea
        '//Para que no coja una Tarifa, metemos en el contexto el IDOrdenRuta
        Dim context As New BusinessData(LinPedido.Doc.HeaderRow)
        If LinPedido.Origen.Table.Columns.Contains("IDOrdenRuta") Then
            If Length(LinPedido.Origen("IDOrdenRuta")) > 0 Then
                context("IDOrdenRuta") = LinPedido.Origen("IDOrdenRuta")
            End If
        End If
        oPCL.ApplyBusinessRule("IDArticulo", LinPedido.Origen("IDArticulo"), LinPedido.Row, context)
    End Sub

    <Task()> Public Shared Sub AsignarCantidadesyUnidades(ByVal LinPedido As DataLineasDesdeOrigen, ByVal services As ServiceProvider)
        '//Para que no coja una Tarifa, metemos en el contexto el IDOrdenRuta
        Dim context As New BusinessData(LinPedido.Doc.HeaderRow)
        If LinPedido.Origen.Table.Columns.Contains("IDOrdenRuta") Then
            If Length(LinPedido.Origen("IDOrdenRuta")) > 0 Then
                context("IDOrdenRuta") = LinPedido.Origen("IDOrdenRuta")
            End If
        End If
        Dim oPCL As New PedidoCompraLinea
        Select Case LinPedido.Doc.Cabecera.Origen
            Case enumOrigenPedidoCompra.Obras
                Dim datosFactor As New ArticuloUnidadAB.DatosFactorConversion(LinPedido.Row("IDArticulo"), LinPedido.Row("IDUDMedida"), LinPedido.Row("IDUDInterna"))
                LinPedido.Row("Factor") = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, datosFactor, services)
                If LinPedido.Row("Factor") = 0 Then LinPedido.Row("Factor") = 1
                LinPedido.Row("QPedida") = LinPedido.Cantidad
                LinPedido.Row("QInterna") = LinPedido.Cantidad
                oPCL.ApplyBusinessRule("QPedida", LinPedido.Row("QPedida") / LinPedido.Row("Factor"), LinPedido.Row, context)
            Case enumOrigenPedidoCompra.Subcontratacion
                Dim PedLinSub As PedLinCompraSubcontratacion = CType(LinPedido.PedLin, PedLinCompraSubcontratacion)
                oPCL.ApplyBusinessRule("IDUDInterna", PedLinSub.IDUDInterna, LinPedido.Row, context)
                oPCL.ApplyBusinessRule("IDUDMedida", PedLinSub.IDUDProduccion, LinPedido.Row, context)
                oPCL.ApplyBusinessRule("QPedida", PedLinSub.QInterna, LinPedido.Row, context)
                oPCL.ApplyBusinessRule("QInterna", LinPedido.Cantidad, LinPedido.Row, context)
                'oPCL.ApplyBusinessRule("Factor", Nz(LinPedido.Origen("FactorProduccion"), 1), LinPedido.Row, context)
            Case enumOrigenPedidoCompra.OfertaComercial
                LinPedido.Row("IDUDInterna") = New Articulo().GetItemRow(LinPedido.Row("IDArticulo"))("IDUDInterna")
                If Length(LinPedido.Origen("IDUDVenta")) = 0 Then
                    ApplicationService.GenerateError("Hay líneas de oferta con artículos sin unidad de venta definida. Debe indicar la misma.", Quoted(LinPedido.Origen("IDUDVenta")))
                Else
                    LinPedido.Row("IDUDMedida") = LinPedido.Origen("IDUDVenta")
                End If
                LinPedido.Row("QPedida") = LinPedido.Cantidad
                Dim datosFactor As New ArticuloUnidadAB.DatosFactorConversion(LinPedido.Row("IDArticulo"), LinPedido.Row("IDUDMedida"), LinPedido.Row("IDUDInterna"))
                LinPedido.Row("Factor") = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, datosFactor, services)
                If LinPedido.Row("Factor") = 0 Then LinPedido.Row("Factor") = 1
                oPCL.ApplyBusinessRule("QInterna", LinPedido.Cantidad * LinPedido.Row("Factor"), LinPedido.Row, context)
                'Qué calculos o qué demás cosas tengo que hacer en este punto para que se calculen bien todos los datos.
                'Comentar con Aintzane.
            Case enumOrigenPedidoCompra.OfertaCompra
                Dim PedLin As PedLinCompraOfertaCompra = DirectCast(LinPedido.PedLin, PedLinCompraOfertaCompra)
                If Length(PedLin.IDUDCompra) > 0 Then
                    oPCL.ApplyBusinessRule("IDUDMedida", PedLin.IDUDCompra, LinPedido.Row, context)
                End If
                oPCL.ApplyBusinessRule("QPedida", LinPedido.Cantidad, LinPedido.Row, context)
                oPCL.ApplyBusinessRule("QInterna", PedLin.QInterna, LinPedido.Row, context)
            Case enumOrigenPedidoCompra.Programa
                If LinPedido.Origen.Table.Columns.Contains("Factor") Then oPCL.ApplyBusinessRule("Factor", Nz(LinPedido.Origen("Factor"), 1), LinPedido.Row, context)
                oPCL.ApplyBusinessRule("QPedida", LinPedido.Cantidad, LinPedido.Row, context)
            Case Else
                oPCL.ApplyBusinessRule("QPedida", LinPedido.Cantidad, LinPedido.Row, context)
        End Select
    End Sub

    <Task()> Public Shared Sub AsignarCantidadesyUnidadesSegundaUnidad(ByVal LinPedido As DataLineasDesdeOrigen, ByVal services As ServiceProvider)
        If ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, LinPedido.Row("IDArticulo"), services) Then
            If TypeOf LinPedido.PedLin Is PedLinCompraPedidoVenta Then
                If Length(CType(LinPedido.PedLin, PedLinCompraPedidoVenta).Cantidad2) = 0 Then
                    ApplicationService.GenerateError("El Articulo {0} se gestiona con Doble Unidad. Debe indicar la misma.", Quoted(LinPedido.Row("IDArticulo")))
                Else
                    LinPedido.Row("QInterna2") = CType(LinPedido.PedLin, PedLinCompraPedidoVenta).Cantidad2
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarFechaEntrega(ByVal LinPedido As DataLineasDesdeOrigen, ByVal services As ServiceProvider)
        Select Case LinPedido.Doc.Cabecera.Origen
            Case enumOrigenPedidoCompra.Subcontratacion
                LinPedido.Row("FechaEntrega") = CType(LinPedido.PedLin, PedLinCompraSubcontratacion).FechaEntrega
            Case enumOrigenPedidoCompra.Obras
                If LinPedido.Origen.Table.Columns.Contains("FechaEntrega") AndAlso IsDate(LinPedido.Origen("FechaEntrega")) Then
                    LinPedido.Row("FechaEntrega") = LinPedido.Origen("FechaEntrega")
                ElseIf Length(LinPedido.Doc.HeaderRow("FechaEntrega")) > 0 AndAlso LinPedido.Doc.HeaderRow("FechaEntrega") <> cnMinDate Then
                    LinPedido.Row("FechaEntrega") = LinPedido.Doc.HeaderRow("FechaEntrega")
                Else
                    Dim Info As ProcessInfoPC = services.GetService(Of ProcessInfoPC)()
                    If Len(Info.FechaEntrega) > 0 Then
                        LinPedido.Row("FechaEntrega") = Info.FechaEntrega
                    End If
                End If
            Case Else
                    If LinPedido.Origen.Table.Columns.Contains("FechaEntrega") AndAlso IsDate(LinPedido.Origen("FechaEntrega")) Then
                        LinPedido.Row("FechaEntrega") = LinPedido.Origen("FechaEntrega")
                    ElseIf Length(LinPedido.Doc.HeaderRow("FechaEntrega")) > 0 AndAlso LinPedido.Doc.HeaderRow("FechaEntrega") <> cnMinDate Then
                        LinPedido.Row("FechaEntrega") = LinPedido.Doc.HeaderRow("FechaEntrega")
                    Else
                        LinPedido.Row("FechaEntrega") = Today
                    End If
        End Select
    End Sub

    <Task()> Public Shared Sub AsignarDatosArticuloDesdeOrigen(ByVal LinPedido As DataLineasDesdeOrigen, ByVal services As ServiceProvider)
        Select Case LinPedido.Doc.Cabecera.Origen
            Case enumOrigenPedidoCompra.Subcontratacion
                Dim ORuta As BusinessHelper = BusinessHelper.CreateBusinessObject("OrdenRuta")
                Dim dtOperacion As DataTable = ORuta.SelOnPrimaryKey(LinPedido.Origen("IDOrdenRuta"))
                If dtOperacion.Rows.Count > 0 Then
                    LinPedido.Row("DescArticulo") = dtOperacion.Rows(0)("DescOperacion")
                    LinPedido.Row("DescRefProveedor") = dtOperacion.Rows(0)("DescOperacion")
                    Dim context As New BusinessData(LinPedido.Doc.HeaderRow)
                    If LinPedido.Origen.Table.Columns.Contains("IDOrdenRuta") Then
                        If Length(LinPedido.Origen("IDOrdenRuta")) > 0 Then
                            context("IDOrdenRuta") = LinPedido.Origen("IDOrdenRuta")
                        End If
                    End If
                    Dim PCL As New PedidoCompraLinea
                    PCL.ApplyBusinessRule("Precio", dtOperacion.Rows(0)("CosteOperacionE"), LinPedido.Row, context)
                    If Length(dtOperacion.Rows(0)("IDCContable")) > 0 Then
                        LinPedido.Row("CContable") = dtOperacion.Rows(0)("IDCContable")
                    Else
                        ApplicationService.GenerateError("La operación {0} no tiene indicada una Cuenta Contable.", Quoted(dtOperacion.Rows(0)("DescOperacion")))
                    End If

                    LinPedido.Row("UdValoracion") = dtOperacion.Rows(0)("UdValoracion")
                End If
            Case enumOrigenPedidoCompra.OfertaCompra
                If LinPedido.Origen.Table.Columns.Contains("DescOferta") AndAlso Not LinPedido.Origen.IsNull("DescOferta") Then
                    LinPedido.Row("DescArticulo") = LinPedido.Origen("DescOferta")
                End If
            Case enumOrigenPedidoCompra.PedidoVenta
                If LinPedido.Origen.Table.Columns.Contains("DescArticulo") AndAlso Not LinPedido.Origen.IsNull("DescArticulo") Then
                    LinPedido.Row("DescArticulo") = LinPedido.Origen("DescArticulo")
                End If
                If LinPedido.Origen.Table.Columns.Contains("IDUDInterna") AndAlso Not LinPedido.Origen.IsNull("IDUDInterna") Then
                    LinPedido.Row("IDUDInterna") = LinPedido.Origen("IDUDInterna")
                End If
                If LinPedido.Origen.Table.Columns.Contains("IDUDMedida") AndAlso Not LinPedido.Origen.IsNull("IDUDMedida") Then
                    LinPedido.Row("IDUDMedida") = LinPedido.Origen("IDUDMedida")
                End If
                If LinPedido.Origen.Table.Columns.Contains("UdValoracion") AndAlso Not LinPedido.Origen.IsNull("UdValoracion") Then
                    LinPedido.Row("UdValoracion") = LinPedido.Origen("UdValoracion")
                End If
            Case Else
                If LinPedido.Origen.Table.Columns.Contains("DescArticulo") AndAlso Not LinPedido.Origen.IsNull("DescArticulo") Then
                    LinPedido.Row("DescArticulo") = LinPedido.Origen("DescArticulo")
                End If
                If LinPedido.Origen.Table.Columns.Contains("IDCContable") AndAlso Not LinPedido.Origen.IsNull("IDCContable") Then
                    LinPedido.Row("CContable") = LinPedido.Origen("IDCContable")
                End If
                If LinPedido.Origen.Table.Columns.Contains("CContable") AndAlso Not LinPedido.Origen.IsNull("CContable") Then
                    LinPedido.Row("CContable") = LinPedido.Origen("CContable")
                End If
                If LinPedido.Origen.Table.Columns.Contains("IDUDInterna") AndAlso Not LinPedido.Origen.IsNull("IDUDInterna") Then
                    LinPedido.Row("IDUDInterna") = LinPedido.Origen("IDUDInterna")
                End If
                If LinPedido.Origen.Table.Columns.Contains("IDUDMedida") AndAlso Not LinPedido.Origen.IsNull("IDUDMedida") Then
                    LinPedido.Row("IDUDMedida") = LinPedido.Origen("IDUDMedida")
                End If
                If LinPedido.Origen.Table.Columns.Contains("UdValoracion") AndAlso Not LinPedido.Origen.IsNull("UdValoracion") Then
                    LinPedido.Row("UdValoracion") = LinPedido.Origen("UdValoracion")
                End If
                If LinPedido.Origen.Table.Columns.Contains("DescDetalle") AndAlso Not LinPedido.Origen.IsNull("DescDetalle") Then
                    LinPedido.Row("DescArticulo") = LinPedido.Origen("DescDetalle")
                End If
        End Select

    End Sub

    <Task()> Public Shared Sub AsignarDatosObras(ByVal LinPedido As DataLineasDesdeOrigen, ByVal services As ServiceProvider)
        If LinPedido.Origen.Table.Columns.Contains("IDObra") AndAlso Not LinPedido.Origen.IsNull("IDObra") Then
            LinPedido.Row("IDObra") = LinPedido.Origen("IDObra")
        End If
        If LinPedido.Origen.Table.Columns.Contains("IDTrabajo") AndAlso Not LinPedido.Origen.IsNull("IDTrabajo") Then
            LinPedido.Row("IDTrabajo") = LinPedido.Origen("IDTrabajo")
        End If
        If LinPedido.Origen.Table.Columns.Contains("IDLineaMaterial") AndAlso Not LinPedido.Origen.IsNull("IDLineaMaterial") Then
            LinPedido.Row("IDLineaMaterial") = LinPedido.Origen("IDLineaMaterial")
        End If
        If LinPedido.Origen.Table.Columns.Contains("CCCompra") AndAlso Not LinPedido.Origen.IsNull("CCCompra") Then
            LinPedido.Row("CContable") = LinPedido.Origen("CCCompra")
        End If
    End Sub

    <Task()> Public Shared Sub AsignarAlmacen(ByVal LinPedido As DataLineasDesdeOrigen, ByVal services As ServiceProvider)
        If LinPedido.Origen.Table.Columns.Contains("IDAlmacen") AndAlso Not LinPedido.Origen.IsNull("IDAlmacen") Then
            LinPedido.Row("IDAlmacen") = LinPedido.Origen("IDAlmacen")
        End If
    End Sub

    <Task()> Public Shared Sub AsignarPrecio(ByVal LinPedido As DataLineasDesdeOrigen, ByVal services As ServiceProvider)
        '//Para que no coja una Tarifa, metemos en el contexto el IDOrdenRuta
        Dim context As New BusinessData(LinPedido.Doc.HeaderRow)
        If LinPedido.Origen.Table.Columns.Contains("IDOrdenRuta") Then
            If Length(LinPedido.Origen("IDOrdenRuta")) > 0 Then
                context("IDOrdenRuta") = LinPedido.Origen("IDOrdenRuta")
            End If
        End If

        Dim oPCL As New PedidoCompraLinea
        If LinPedido.Doc.Cabecera.Origen <> enumOrigenPedidoCompra.Subcontratacion Then
            Dim dblPrecio As Double
            If LinPedido.Origen.Table.Columns.Contains("Precio") Then
                dblPrecio = LinPedido.Origen("Precio")
            ElseIf LinPedido.Origen.Table.Columns.Contains("PrecioPrevMatA") Then
                If LinPedido.Row("IDUDInterna") = LinPedido.Row("IDUDMedida") Then
                    dblPrecio = LinPedido.Origen("PrecioPrevMatA")
                Else
                    Dim datosFactor As New ArticuloUnidadAB.DatosFactorConversion(LinPedido.Row("IDArticulo"), LinPedido.Row("IDUDMedida"), LinPedido.Row("IDUDInterna"))
                    Dim Factor As Double = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, datosFactor, services)
                    dblPrecio = LinPedido.Origen("PrecioPrevMatA") * Factor
                End If
            ElseIf LinPedido.Origen.Table.Columns.Contains("TasaPrevMatA") Then
                dblPrecio = LinPedido.Origen("TasaPrevMatA")
            End If
            If LinPedido.Doc.Cabecera.Origen = enumOrigenPedidoCompra.OfertaComercial Then
                If LinPedido.Origen.Table.Columns.Contains("CosteProveedor") Then
                    dblPrecio = LinPedido.Origen("CosteProveedor")
                End If
            End If

            oPCL.ApplyBusinessRule("Precio", dblPrecio, LinPedido.Row, context)
        End If
    End Sub

    <Task()> Public Shared Sub AsignarDtos(ByVal LinPedido As DataLineasDesdeOrigen, ByVal services As ServiceProvider)
        '//Para que no coja una Tarifa, metemos en el contexto el IDOrdenRuta
        Dim context As New BusinessData(LinPedido.Doc.HeaderRow)
        If LinPedido.Origen.Table.Columns.Contains("IDOrdenRuta") Then
            If Length(LinPedido.Origen("IDOrdenRuta")) > 0 Then
                context("IDOrdenRuta") = LinPedido.Origen("IDOrdenRuta")
            End If
        End If
        Dim oPCL As New PedidoCompraLinea
        If LinPedido.Origen.Table.Columns.Contains("Dto1") Then
            oPCL.ApplyBusinessRule("Dto1", LinPedido.Origen("Dto1"), LinPedido.Row, context)
        End If
        If LinPedido.Origen.Table.Columns.Contains("Dto2") Then
            oPCL.ApplyBusinessRule("Dto2", LinPedido.Origen("Dto2"), LinPedido.Row, context)
        End If
        If LinPedido.Origen.Table.Columns.Contains("Dto3") Then
            oPCL.ApplyBusinessRule("Dto3", LinPedido.Origen("Dto3"), LinPedido.Row, context)
        End If
        oPCL.ApplyBusinessRule("Dto", Nz(LinPedido.Doc.HeaderRow("Dto"), 0), LinPedido.Row, context)

        If Length(LinPedido.Doc.HeaderRow("IDCondicionPago")) > 0 Then
            Dim CondicionesPago As EntityInfoCache(Of CondicionPagoInfo) = services.GetService(Of EntityInfoCache(Of CondicionPagoInfo))()
            Dim CondPago As CondicionPagoInfo = CondicionesPago.GetEntity(LinPedido.Doc.HeaderRow("IDCondicionPago"))
            oPCL.ApplyBusinessRule("DtoProntoPago", CondPago.DtoProntoPago, LinPedido.Row, context)
        End If
    End Sub

    <Task()> Public Shared Sub AsignarContratoCompra(ByVal LinPedido As DataLineasDesdeOrigen, ByVal services As ServiceProvider)
        If LinPedido.Origen.Table.Columns.Contains("IDContrato") AndAlso Not LinPedido.Origen.IsNull("IDContrato") Then
            LinPedido.Row("IDContrato") = LinPedido.Origen("IDContrato")
            LinPedido.Row("IDLineaContrato") = LinPedido.Origen("IDLineaContrato")
        End If
    End Sub

    <Task()> Public Shared Sub AsignarTipoLinea(ByVal LinPedido As DataLineasDesdeOrigen, ByVal services As ServiceProvider)
        'If Length(LinPedido.Row("TipoLineaCompra")) = 0 Then
        Select Case LinPedido.Doc.Cabecera.Origen
            Case enumOrigenPedidoCompra.Subcontratacion
                LinPedido.Row("TipoLineaCompra") = enumaclTipoLineaAlbaran.aclSubcontratacion
            Case Else
                LinPedido.Row("TipoLineaCompra") = enumaclTipoLineaAlbaran.aclNormal
        End Select
        'End If
    End Sub

    <Task()> Public Shared Sub AsignarCentroGestionDesdeOrigen(ByVal LinPedido As DataLineasDesdeOrigen, ByVal services As ServiceProvider)
        If LinPedido.Origen.Table.Columns.Contains("IDCentroGestion") Then
            LinPedido.Row("IDCentroGestion") = LinPedido.Origen("IDCentroGestion")
        Else
            LinPedido.Row("IDCentroGestion") = LinPedido.Doc.HeaderRow("IDCentroGestion")
        End If
        If Length(LinPedido.Row("IDCentroGestion")) = 0 Then
            Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
            LinPedido.Row("IDCentroGestion") = AppParams.CentroGestion
        End If
    End Sub

    <Task()> Public Shared Sub EstablecerEnlaceConOrigen(ByVal LinPedido As DataLineasDesdeOrigen, ByVal services As ServiceProvider)
        Select Case LinPedido.Doc.Cabecera.Origen
            Case enumOrigenPedidoCompra.Solicitud
                If LinPedido.Origen.Table.Columns.Contains("IdSolicitud") Then LinPedido.Row("IdSolicitud") = LinPedido.Origen("IdSolicitud")
                If LinPedido.Origen.Table.Columns.Contains("IdLineaSolicitud") Then LinPedido.Row("IdLineaSolicitud") = LinPedido.Origen("IdLineaSolicitud")
                If LinPedido.Origen.Table.Columns.Contains("IDMntoOTPrev") Then LinPedido.Row("IDMntoOTPrev") = LinPedido.Origen("IDMntoOTPrev")
            Case enumOrigenPedidoCompra.Programa
                If LinPedido.Origen.Table.Columns.Contains("IDPrograma") Then LinPedido.Row("IdPrograma") = LinPedido.Origen("IDPrograma")
                If LinPedido.Origen.Table.Columns.Contains("IDLineaPrograma") Then LinPedido.Row("IdLineaPrograma") = LinPedido.Origen("IDLineaPrograma")
            Case enumOrigenPedidoCompra.Subcontratacion
                If LinPedido.Origen.Table.Columns.Contains("IDOrdenRuta") Then LinPedido.Row("IDOrdenRuta") = LinPedido.Origen("IDOrdenRuta")
            Case enumOrigenPedidoCompra.Obras
                If LinPedido.Origen.Table.Columns.Contains("IDLineaMaterial") Then LinPedido.Row("IDLineaMaterial") = LinPedido.Origen("IDLineaMaterial")
                If LinPedido.Origen.Table.Columns.Contains("IDTrabajo") Then LinPedido.Row("IDTrabajo") = LinPedido.Origen("IDTrabajo")
                If LinPedido.Origen.Table.Columns.Contains("IDObra") Then LinPedido.Row("IDObra") = LinPedido.Origen("IDObra")
            Case enumOrigenPedidoCompra.Mnto
                If LinPedido.Origen.Table.Columns.Contains("IDMntoOTPrev") Then LinPedido.Row("IDMntoOTPrev") = LinPedido.Origen("IDMntoOTPrev")
            Case enumOrigenPedidoCompra.OfertaComercial
                'If LinPedido.Origen.Table.Columns.Contains("IDContrato") Then LinPedido.Row("IDContrato") = LinPedido.Origen("IDContrato")
                'If LinPedido.Origen.Table.Columns.Contains("IDLineaContrato") Then LinPedido.Row("IDLineaContrato") = LinPedido.Origen("IDLineaContrato")
                If LinPedido.Origen.Table.Columns.Contains("IDLineaOfertaDetalle") Then LinPedido.Row("IDLineaOfertaDetalle") = LinPedido.Origen("IDLineaOfertaDetalle")
                If LinPedido.Origen.Table.Columns.Contains("IDOrdenLinea") Then LinPedido.Row("IDOrdenLinea") = LinPedido.Origen("IDOrdenLinea")
            Case enumOrigenPedidoCompra.OfertaCompra
                If LinPedido.Origen.Table.Columns.Contains("IDLineaSolicitud") Then
                    If Length(LinPedido.Origen("IDLineaSolicitud")) > 0 Then LinPedido.Row("IDLineaSolicitud") = LinPedido.Origen("IDLineaSolicitud")
                    If Length(LinPedido.Origen("IDSolicitud")) > 0 Then LinPedido.Row("IDSolicitud") = LinPedido.Origen("IDSolicitud")
                End If
                If LinPedido.Origen.Table.Columns.Contains("IdOferta") Then LinPedido.Row("IdOferta") = LinPedido.Origen("IdOferta")
                If LinPedido.Origen.Table.Columns.Contains("IDMntoOTPrev") Then LinPedido.Row("IDMntoOTPrev") = LinPedido.Origen("IDMntoOTPrev")
            Case enumOrigenPedidoCompra.PedidoVenta
                LinPedido.Row("PedidoVentaOrigen") = LinPedido.Origen("IDPedido")
        End Select
    End Sub

    '<Task()> Public Shared Function RecuperarDatosObras(ByVal Doc As DocumentoPedidoCompra, ByVal services As ServiceProvider) As DataTable
    '    Dim ids(CType(Doc.Cabecera, PedCabCompraObra).LineasOrigen.Length - 1) As Object
    '    Dim oFltr As New Filter
    '    For i As Integer = 0 To ids.Length - 1
    '        ids(i) = CType(Doc.Cabecera, PedCabCompraObra).LineasOrigen(i).IDLineaOrigen
    '    Next
    '    Dim FieldName As String
    '    Dim ViewName As String
    '    If CType(Doc.Cabecera, PedCabCompraObra).PorMateriales Then
    '        FieldName = "IDTrabajo"
    '        ViewName = "vFrmMntoObraGeneraCompra"
    '    ElseIf CType(Doc.Cabecera, PedCabCompraObra).PorTrabajos Then
    '        FieldName = "IDLineaMaterial"
    '        ViewName = "vFrmMntoObraGeneraCompraTrabajo"
    '    End If
    '    If Length(FieldName) > 0 AndAlso Length(ViewName) > 0 Then
    '        oFltr.Add(New InListFilterItem(FieldName, ids, FilterType.Numeric))
    '        Return New BE.DataEngine().Filter(ViewName, oFltr)
    '    End If
    'End Function

    <Task()> Public Shared Function RecuperarDatosOrigenGenerico(ByVal Doc As DocumentoPedidoCompra, ByVal services As ServiceProvider) As DataTable
        If Not Doc.Cabecera Is Nothing Then
            Select Case Doc.Cabecera.Origen
                'Case enumOrigenPedidoCompra.PedidoVenta
                '    Return CType(Doc.Cabecera, PedCabCompraPedidoVenta).DatosOrigen
                Case Else
                    Dim ids(CType(Doc.Cabecera, PedCabCompra).LineasOrigen.Length - 1) As Object
                    Dim FieldName As String
                    Dim oFltr As New Filter
                    For i As Integer = 0 To ids.Length - 1
                        If i = 0 Then FieldName = CType(Doc.Cabecera, PedCabCompra).LineasOrigen(i).PrimaryKeyLinOrigen
                        ids(i) = CType(Doc.Cabecera, PedCabCompra).LineasOrigen(i).IDLineaOrigen
                    Next
                    oFltr.Add(New InListFilterItem(FieldName, ids, FilterType.Numeric))

                    Return New BE.DataEngine().Filter(CType(Doc.Cabecera, PedCabCompra).ViewName, oFltr)
            End Select
        End If
    End Function

    <Task()> Public Shared Sub LineasDeRegalo(ByVal Doc As DocumentoPedidoVenta, ByVal services As ServiceProvider)
        Dim f1 As New Filter
        f1.Add(New StringFilterItem("IDCliente", Doc.HeaderRow("IDCliente")))
        Dim f2 As New Filter(FilterUnionOperator.Or)
        For Each linea As DataRow In Doc.dtLineas.Select
            f2.Add(New StringFilterItem("IDArticulo", FilterOperator.Equal, linea("IDArticulo")))
        Next
        f1.Add(f2)

        Dim pvl As New PedidoVentaLinea
        Dim f As New Filter
        Dim regalos As DataTable = New BE.DataEngine().Filter("vNegArticulosRegaloPorPromocion", f1)
        For Each regalo As DataRow In regalos.Rows
            f.Clear()
            f.Add(New NumberFilterItem("IDPedido", FilterOperator.Equal, Doc.HeaderRow("IDPedido")))
            f.Add(New StringFilterItem("IDArticulo", FilterOperator.Equal, regalo("IDArticulo")))
            Dim WhereArticuloRegalo As String = f.Compose(New AdoFilterComposer)
            Dim lineapedido() As DataRow = Doc.dtLineas.Select(WhereArticuloRegalo)
            If UBound(lineapedido) >= 0 Then
                If regalo("QPromocionada") + regalo("QRegalo") < regalo("QMaxPromocionable") Then
                    Dim newrow As DataRow = Doc.dtLineas.NewRow
                    'copiar todos los datos de la linea del pedido
                    newrow.ItemArray = lineapedido(0).ItemArray

                    'sobreescribir parte de los datos copiados
                    newrow("IDLineaPedido") = AdminData.GetAutoNumeric
                    newrow("Estado") = enumpvlEstado.pvlCerrado
                    newrow("IDArticulo") = regalo("IDArticuloRegalo")
                    newrow("DescArticulo") = regalo("DescArticuloRegalo")
                    newrow("QPedida") = 0
                    If regalo("QPedida") <> 0 Then
                        newrow("QPedida") = Fix(lineapedido(0)("QPedida") / regalo("QPedida")) * regalo("QRegalo")
                    End If
                    newrow("IDPromocionLinea") = regalo("IDPromocionLinea")
                    newrow("Regalo") = True

                    pvl.ApplyBusinessRule("IDUdMedida", newrow("IDUdMedida"), newrow)
                    pvl.ApplyBusinessRule("Regalo", newrow("Regalo"), newrow)

                    If newrow("QPedida") <> 0 Then
                        Doc.dtLineas.Rows.Add(newrow.ItemArray)
                    End If
                End If
            End If
        Next
    End Sub

#End Region

#Region " Validaciones "

    <Task()> Public Shared Sub ValidarDatosOfertas(ByVal data As DataTable, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataTable)(AddressOf ValidarOfertaExistenDatosProcesar, data, services)
        ProcessServer.ExecuteTask(Of DataTable)(AddressOf ValidarOfertaEmpresaProveedor, data, services)
        ProcessServer.ExecuteTask(Of DataTable)(AddressOf ValidarOfertaProveedor, data, services)
        ProcessServer.ExecuteTask(Of DataTable)(AddressOf ValidarOfertaFechaEntrega, data, services)
        ProcessServer.ExecuteTask(Of DataTable)(AddressOf ValidarOfertaPresupuesto, data, services)
    End Sub

    <Task()> Public Shared Sub ValidarOfertaExistenDatosProcesar(ByVal data As DataTable, ByVal services As ServiceProvider)
        If data Is Nothing OrElse data.Rows.Count = 0 Then
            ApplicationService.GenerateError("Las líneas seleccionadas no pueden generar Pedidos de Compra. Compruebe el/los artículo/s, si pueden generar Pedidos de Compra y su estado de actualización.")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarOfertaEmpresaProveedor(ByVal data As DataTable, ByVal services As ServiceProvider)
        If Not data Is Nothing Then
            Dim f As New Filter
            f.Add(New IsNullFilterItem("IDProveedor"))
            f.Add(New IsNullFilterItem("IDEmpresa"))
            Dim WhereNullEmpresaProveedor As String = f.Compose(New AdoFilterComposer)
            Dim adr() As DataRow = data.Select(WhereNullEmpresaProveedor)
            If Not adr Is Nothing AndAlso adr.Length > 0 Then
                ApplicationService.GenerateError("Proveedor y Empresa no pueden estar vacíos a la vez. Revise las Ofertas seleccionadas.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarOfertaProveedor(ByVal data As DataTable, ByVal services As ServiceProvider)
        If Not data Is Nothing Then
            Dim f As New Filter
            f.Add(New IsNullFilterItem("IDArticulo", False))
            f.Add(New IsNullFilterItem("IDProveedor"))
            Dim WhereNullProveedor As String = f.Compose(New AdoFilterComposer)
            Dim adr() As DataRow = data.Select(WhereNullProveedor)
            If Not adr Is Nothing AndAlso adr.Length > 0 Then
                ApplicationService.GenerateError("El código de Proveedor es obligatorio para todas las líneas susceptibles de crear un Pedido de Compra.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarOfertaFechaEntrega(ByVal data As DataTable, ByVal services As ServiceProvider)
        If Not data Is Nothing Then
            Dim f As New Filter
            f.Add(New IsNullFilterItem("IDArticulo", False))
            f.Add(New IsNullFilterItem("FechaEntrega"))
            Dim WhereNullFechaEntrega As String = f.Compose(New AdoFilterComposer)
            Dim adr() As DataRow = data.Select(WhereNullFechaEntrega)
            If Not adr Is Nothing AndAlso adr.Length > 0 Then
                ApplicationService.GenerateError("La Fecha de Entrega es obligatoria para todas las líneas susceptibles de crear un Pedido de Compra.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarOfertaPresupuesto(ByVal data As DataTable, ByVal services As ServiceProvider)
        If Not data Is Nothing Then
            Dim f As New Filter
            f.Add(New IsNullFilterItem("IDPresupuesto", False))
            f.Add(New IsNullFilterItem("IDArticulo"))
            Dim WhereNullArticulo As String = f.Compose(New AdoFilterComposer)
            Dim adr() As DataRow = data.Select(WhereNullArticulo)
            If Not adr Is Nothing AndAlso adr.Length > 0 Then
                ApplicationService.GenerateError("Es necesario asociar los Presupuestos de la Oferta con Artículos. Revise las ofertas seleccionadas.")
            End If
        End If
    End Sub


    <Task()> Public Shared Sub ValidacionesContabilidad(ByVal data As DataRow, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf NegocioGeneral.ValidarEjercicioContablePedido, data, services)
    End Sub

    <Task()> Public Shared Sub ValidarDocumento(ByVal Doc As DocumentoPedidoCompra, ByVal services As ServiceProvider)
        Dim pcc As New PedidoCompraCabecera
        pcc.Validate(Doc.HeaderRow.Table)
        Dim pcl As New PedidoCompraLinea
        pcl.Validate(Doc.dtLineas)
    End Sub

#End Region

#Region " Actualizacion de Entidades relacionadas"

    <Task()> Public Shared Sub ActualizarEntidadesDependientesUpdate(ByVal Doc As DocumentoPedidoCompra, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DocumentoPedidoCompra)(AddressOf ActualizarPrograma, Doc, services)
        ProcessServer.ExecuteTask(Of DocumentoPedidoCompra)(AddressOf ActualizarOfertaCompra, Doc, services)
        ProcessServer.ExecuteTask(Of DocumentoPedidoCompra)(AddressOf ActualizarOrdenRutaDesdePedidoCompra, Doc, services)
        ProcessServer.ExecuteTask(Of DocumentoPedidoCompra)(AddressOf ActualizarOfertaComercial, Doc, services)
        ProcessServer.ExecuteTask(Of DocumentoPedidoCompra)(AddressOf ActualizarSolicitud, Doc, services)
        ProcessServer.ExecuteTask(Of DocumentoPedidoCompra)(AddressOf ActualizarMantenimiento, Doc, services)
        ProcessServer.ExecuteTask(Of DocumentoPedidoCompra)(AddressOf ActualizarObras, Doc, services)
    End Sub

    <Task()> Public Shared Sub ActualizarEntidadesDependientes(ByVal Doc As DocumentoPedidoCompra, ByVal services As ServiceProvider)
        Select Case Doc.Cabecera.Origen
            Case enumOrigenPedidoCompra.Programa
                ProcessServer.ExecuteTask(Of DocumentoPedidoCompra)(AddressOf ActualizarPrograma, Doc, services)
            Case enumOrigenPedidoCompra.OfertaCompra
                ProcessServer.ExecuteTask(Of DocumentoPedidoCompra)(AddressOf ActualizarOfertaCompra, Doc, services)
            Case enumOrigenPedidoCompra.Subcontratacion
                ProcessServer.ExecuteTask(Of DocumentoPedidoCompra)(AddressOf ActualizarOrdenRutaDesdePedidoCompra, Doc, services)
            Case enumOrigenPedidoCompra.OfertaComercial
                ProcessServer.ExecuteTask(Of DocumentoPedidoCompra)(AddressOf ActualizarOfertaComercial, Doc, services)
            Case enumOrigenPedidoCompra.Solicitud
                ProcessServer.ExecuteTask(Of DocumentoPedidoCompra)(AddressOf ActualizarSolicitud, Doc, services)
            Case enumOrigenPedidoCompra.Mnto
                ProcessServer.ExecuteTask(Of DocumentoPedidoCompra)(AddressOf ActualizarMantenimiento, Doc, services)
                'Case enumOrigenPedidoCompra.Obras
                '    ProcessServer.ExecuteTask(Of DocumentoPedidoCompra)(AddressOf ActualizarObras, Doc, services)
            Case enumOrigenPedidoCompra.PedidoVenta
                ProcessServer.ExecuteTask(Of DocumentoPedidoCompra)(AddressOf ActualizarPedidoVenta, Doc, services)
        End Select

        ProcessServer.ExecuteTask(Of DocumentoPedidoCompra)(AddressOf ActualizarObras, Doc, services)
    End Sub

#Region " Actualización de Programas de compra "

    <Task()> Public Shared Sub ActualizarPrograma(ByVal Doc As DocumentoPedidoCompra, ByVal services As ServiceProvider)
        If Doc Is Nothing Then Exit Sub
        For Each lineapedido As DataRow In Doc.dtLineas.Select(Nothing, "IDPrograma,IDLineaPrograma", DataViewRowState.CurrentRows)
            Dim datosActProg As New DataActualizarProgramaLinea(lineapedido)
            ProcessServer.ExecuteTask(Of DataActualizarProgramaLinea)(AddressOf ProcesoPedidoCompra.ActualizarProgramaLinea, datosActProg, services)
        Next
    End Sub
    Public Class DataActualizarProgramaLinea
        Public LineaPedido As DataRow
        Public DeletingRow As Boolean
        Public FechaConfirmacion As Date?

        Public Sub New(ByVal LineaPedido As DataRow, Optional ByVal DeletingRow As Boolean = False, Optional ByVal FechaConfirmacion As Date = cnMinDate)
            Me.LineaPedido = LineaPedido
            Me.DeletingRow = DeletingRow
            If FechaConfirmacion <> cnMinDate Then Me.FechaConfirmacion = FechaConfirmacion
        End Sub
    End Class
    <Task()> Public Shared Sub ActualizarProgramaLinea(ByVal data As DataActualizarProgramaLinea, ByVal services As ServiceProvider)
        Dim CambioLineaProg As Boolean
        Dim CambioQPedida As Boolean
        Dim CambioFechaEntrega As Boolean

        If data.LineaPedido.RowState = DataRowState.Modified Then
            CambioLineaProg = (Nz(data.LineaPedido("IDLineaPrograma"), 0) <> Nz(data.LineaPedido("IDLineaPrograma", DataRowVersion.Original), 0))
            CambioQPedida = (data.LineaPedido("QPedida") <> data.LineaPedido("QPedida", DataRowVersion.Original))
            CambioFechaEntrega = (data.LineaPedido("FechaEntrega") <> data.LineaPedido("FechaEntrega", DataRowVersion.Original))
        End If

        If data.LineaPedido.RowState = DataRowState.Added OrElse CambioLineaProg OrElse CambioQPedida OrElse CambioFechaEntrega OrElse data.DeletingRow Then
            Dim pl As New ProgramaCompraLinea
            Dim Programa As DataTable = pl.SelOnPrimaryKey(data.LineaPedido("IDLineaPrograma"))
            If Not IsNothing(Programa) AndAlso Programa.Rows.Count > 0 Then
                If data.DeletingRow Then
                    Programa.Rows(0)("Confirmada") = CBool(enumplEstadoLinea.plNoConfirmada)
                    Programa.Rows(0)("QConfirmada") -= data.LineaPedido("QPedida")
                    If Programa.Rows(0)("QConfirmada") < 0 Then Programa.Rows(0)("QConfirmada") = 0
                    Programa.Rows(0)("FechaConfirmacion") = DBNull.Value
                Else
                    Dim dblQModificada As Integer
                    If data.LineaPedido.RowState = DataRowState.Modified Then
                        dblQModificada = data.LineaPedido("QPedida", DataRowVersion.Original)
                    End If
                    Programa.Rows(0)("QConfirmada") = Nz(Programa.Rows(0)("QConfirmada"), 0) + (data.LineaPedido("QPedida") - dblQModificada)
                    Programa.Rows(0)("Confirmada") = CBool(enumplEstadoLinea.plConfirmada)
                    If data.FechaConfirmacion <> cnMinDate Then
                        Programa.Rows(0)("FechaConfirmacion") = data.FechaConfirmacion
                    Else
                        Programa.Rows(0)("FechaConfirmacion") = Today
                    End If
                    Programa.Rows(0)("FechaEntrega") = data.LineaPedido("FechaEntrega")
                End If
                BusinessHelper.UpdateTable(Programa)
            End If
        End If
    End Sub

#End Region

#Region " Actualización de Solicitud de Compra "

    <Task()> Public Shared Sub ActualizarSolicitud(ByVal Doc As DocumentoPedidoCompra, ByVal services As ServiceProvider)
        If Doc Is Nothing Then Exit Sub
        For Each lineapedido As DataRow In Doc.dtLineas.Select(Nothing, "IDSolicitud,IDLineaSolicitud")
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoPedidoCompra.ActualizarSolicitudLinea, lineapedido, services)
        Next
    End Sub
    <Task()> Public Shared Sub ActualizarSolicitudLinea(ByVal lineapedido As DataRow, ByVal services As ServiceProvider)
        If Not IsNothing(lineapedido) Then
            Dim CambioLineaSolic As Boolean
            Dim CambioQPedida As Boolean

            If lineapedido.RowState = DataRowState.Modified Then
                CambioLineaSolic = (Nz(lineapedido("IDLineaSolicitud"), 0) <> Nz(lineapedido("IDLineaSolicitud", DataRowVersion.Original), 0))
                CambioQPedida = (lineapedido("QPedida") <> lineapedido("QPedida", DataRowVersion.Original))
            End If

            If lineapedido.RowState = DataRowState.Added OrElse CambioLineaSolic OrElse CambioQPedida Then
                If Length(lineapedido("IDLineaSolicitud")) > 0 Then
                    Dim sl As BusinessHelper = BusinessHelper.CreateBusinessObject("SolicitudCompraLinea")
                    Dim dtSolicitud As DataTable = sl.SelOnPrimaryKey(lineapedido("IDLineaSolicitud"))
                    If Not IsNothing(dtSolicitud) AndAlso dtSolicitud.Rows.Count > 0 Then
                        If lineapedido("QPedida") = 0 AndAlso Length(lineapedido("IDOferta")) = 0 Then
                            'Borrando registro
                            dtSolicitud.Rows(0)("FechaEstado") = Date.Today
                            dtSolicitud.Rows(0)("QTramitada") = dtSolicitud.Rows(0)("QTramitada") - lineapedido("QPedida", DataRowVersion.Original)
                            If dtSolicitud.Rows(0)("QTramitada") = 0 Then
                                dtSolicitud.Rows(0)("Estado") = enumscEstado.scSolicitado
                            End If
                        ElseIf lineapedido("QPedida") = 0 AndAlso Length(lineapedido("IDOferta")) <> 0 Then
                            'Borrando registro
                            dtSolicitud.Rows(0)("Estado") = enumscEstado.scSolicitaOferta
                        Else
                            dtSolicitud.Rows(0)("FechaEstado") = Date.Today
                            If Length(lineapedido("IDOferta")) = 0 Then
                                Dim dblQModificada As Integer
                                If lineapedido.RowState = DataRowState.Modified Then
                                    dblQModificada = lineapedido("QPedida", DataRowVersion.Original)
                                End If
                                dtSolicitud.Rows(0)("QTramitada") = dtSolicitud.Rows(0)("QTramitada") + (lineapedido("QPedida") - dblQModificada)
                            End If
                            dtSolicitud.Rows(0)("Estado") = enumscEstado.scPedido
                        End If
                        BusinessHelper.UpdateTable(dtSolicitud)
                    End If
                End If
            End If
        End If
    End Sub

#End Region

#Region " Actualización de Oferta de Compra "

    <Task()> Public Shared Sub ActualizarOfertaCompra(ByVal Doc As DocumentoPedidoCompra, ByVal services As ServiceProvider)
        If Doc Is Nothing Then Exit Sub
        For Each lineapedido As DataRow In Doc.dtLineas.Select(Nothing, "IDOferta,IDLineaOferta")
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoPedidoCompra.ActualizarOfertaCompraLinea, lineapedido, services)
        Next
    End Sub
    <Task()> Public Shared Sub ActualizarOfertaCompraLinea(ByVal lineaPedido As DataRow, ByVal services As ServiceProvider)
        If Not IsNothing(lineaPedido) Then
            Dim CambioOferta As Boolean
            Dim CambioCantidad As Boolean
            If lineaPedido.RowState = DataRowState.Modified Then
                CambioOferta = (lineaPedido("IDOferta") & String.Empty <> lineaPedido("IDOferta", DataRowVersion.Original) & String.Empty)
                CambioCantidad = (lineaPedido("QPedida") & String.Empty <> lineaPedido("QPedida", DataRowVersion.Original) & String.Empty)
            End If

            If lineaPedido.RowState = DataRowState.Added OrElse CambioOferta OrElse CambioCantidad Then
                If Length(lineaPedido("IDOferta")) > 0 Then
                    Dim of1 As BusinessHelper = BusinessHelper.CreateBusinessObject("OfertaCabecera")
                    Dim dtOferta As DataTable = of1.SelOnPrimaryKey(lineaPedido("IDOferta"))
                    If Not IsNothing(dtOferta) AndAlso dtOferta.Rows.Count > 0 Then
                        If Nz(lineaPedido("QPedida"), 0) = 0 Then
                            dtOferta.Rows(0)("Estado") = enumOfertaCabecera.ocAdjudicada
                            dtOferta.Rows(0)("IDPedido") = System.DBNull.Value
                        Else
                            dtOferta.Rows(0)("Estado") = enumOfertaCabecera.ocCerrada
                            dtOferta.Rows(0)("IDPedido") = lineaPedido("IDPedido")
                        End If
                        BusinessHelper.UpdateTable(dtOferta)
                    End If
                End If
            End If
        End If
    End Sub

#End Region

#Region " Actualización de Oferta Comercial "

    <Task()> Public Shared Sub ActualizarOfertaComercial(ByVal Doc As DocumentoPedidoCompra, ByVal services As ServiceProvider)
        If Doc Is Nothing Then Exit Sub
        For Each lineapedido As DataRow In Doc.dtLineas.Select()
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoPedidoCompra.ActualizarOfertaDetalle, lineapedido, services)
        Next
    End Sub

    <Task()> Public Shared Sub ActualizarOfertaDetalle(ByVal lineaPedido As DataRow, ByVal services As ServiceProvider)
        If Length(lineaPedido("IDLineaOfertaDetalle")) > 0 Then
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf ActualizarEstadoCompra, lineaPedido, services)
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarEstadoCompra(ByVal lineaPedido As DataRow, ByVal services As ServiceProvider)
        Dim CambioLineaOfertaDetalle As Boolean

        If lineaPedido.RowState = DataRowState.Modified Then
            CambioLineaOfertaDetalle = (lineaPedido("IDLineaOfertaDetalle") <> lineaPedido("IDLineaOfertaDetalle", DataRowVersion.Original))
        End If

        If lineaPedido.RowState = DataRowState.Added OrElse CambioLineaOfertaDetalle Then
            If Length(lineaPedido("IDLineaOfertaDetalle")) > 0 Then
                Dim Entidad As BusinessHelper = BusinessHelper.CreateBusinessObject("OfertaComercialDetalle")
                Dim dtDetalle As DataTable = Entidad.SelOnPrimaryKey(lineaPedido("IDLineaOfertaDetalle"))
                If dtDetalle.Rows.Count > 0 Then
                    dtDetalle.Rows(0)("EstadoCompra") = enumocdEstadoCompraVenta.ecvLanzado
                End If
                BusinessHelper.UpdateTable(dtDetalle)
            End If
        End If
    End Sub

#End Region

#Region " Actualización de Obras "

    <Task()> Public Shared Sub ActualizarObras(ByVal Doc As DocumentoPedidoCompra, ByVal services As ServiceProvider)
        If Doc Is Nothing OrElse Doc.dtLineas Is Nothing Then Exit Sub
        Dim f As New Filter(FilterUnionOperator.Or)
        f.Add(New IsNullFilterItem("IDLineaMaterial", False))
        f.Add(New IsNullFilterItem("IDTrabajo", False))
        Dim WhereNotNullTrabajoMaterial As String = f.Compose(New AdoFilterComposer)
        For Each lineapedido As DataRow In Doc.dtLineas.Select(WhereNotNullTrabajoMaterial)
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf ActualizarLineasObra, lineapedido, services)
        Next
    End Sub

    <Task()> Public Shared Sub ActualizarLineasObra(ByVal lineapedido As DataRow, ByVal services As ServiceProvider)
        If Length(lineapedido("IDLineaMaterial")) > 0 Then
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf ActualizarObraMaterial, lineapedido, services)
        ElseIf Length(lineapedido("IDTrabajo")) > 0 Then
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf ActualizarObraTrabajo, lineapedido, services)
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarObraMaterial(ByVal lineapedido As DataRow, ByVal services As ServiceProvider)
        If lineapedido("IDLineaMaterial") > 0 Then
            Dim CambioIDLineaMaterial As Boolean
            Dim CambioQInterna As Boolean
            Dim CambioFechaEntrega As Boolean

            If lineapedido.RowState = DataRowState.Modified Then
                CambioIDLineaMaterial = (Nz(lineapedido("IDLineaMaterial"), 0) <> Nz(lineapedido("IDLineaMaterial", DataRowVersion.Original), 0))
                CambioQInterna = (lineapedido("QInterna") <> lineapedido("QInterna", DataRowVersion.Original))
                CambioFechaEntrega = (lineapedido("FechaEntrega") <> lineapedido("FechaEntrega", DataRowVersion.Original))
            End If

            If lineapedido.RowState = DataRowState.Added OrElse CambioIDLineaMaterial OrElse CambioQInterna OrElse CambioFechaEntrega Then
                Dim OM As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraMaterial")
                Dim dt As DataTable = OM.Filter(New NumberFilterItem("IDLineaMaterial", lineapedido("IDLineaMaterial")))
                If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                    Dim OriginalQInterna As Double
                    Dim ProposedQInterna As Double = Nz(lineapedido("QInterna"), 0)
                    If lineapedido.RowState = DataRowState.Modified Then
                        OriginalQInterna = lineapedido("QInterna", DataRowVersion.Original)
                    End If

                    Dim DiferenciaQInterna As Double = ProposedQInterna - OriginalQInterna
                    dt.Rows(0)("QPedida") = dt.Rows(0)("QPedida") + DiferenciaQInterna

                    If Length(dt.Rows(0)("FechaEntrega")) = 0 Then dt.Rows(0)("FechaEntrega") = lineapedido("FechaEntrega")
                    BusinessHelper.UpdateTable(dt)
                End If
            End If
        End If

    End Sub

    <Task()> Public Shared Sub ActualizarObraTrabajo(ByVal lineapedido As DataRow, ByVal services As ServiceProvider)
        If Length(lineapedido("IDTrabajo")) > 0 Then
            Dim CambioIDTrabajo As Boolean
            Dim CambioQInterna As Boolean
            Dim CambioFechaEntrega As Boolean

            If lineapedido.RowState = DataRowState.Modified Then
                CambioIDTrabajo = (Nz(lineapedido("IDTrabajo"), 0) <> Nz(lineapedido("IDTrabajo", DataRowVersion.Original), 0))
                CambioQInterna = (lineapedido("QInterna") <> lineapedido("QInterna", DataRowVersion.Original))
                CambioFechaEntrega = (lineapedido("FechaEntrega") <> lineapedido("FechaEntrega", DataRowVersion.Original))
            End If

            If lineapedido.RowState = DataRowState.Added OrElse CambioIDTrabajo OrElse CambioQInterna OrElse CambioFechaEntrega Then
                Dim ovl As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraTrabajo")
                Dim dt As DataTable = ovl.Filter(New NumberFilterItem("IDTrabajo", lineapedido("IDTrabajo")))
                If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                    Dim OriginalQInterna As Double
                    Dim ProposedQInterna As Double = Nz(lineapedido("QInterna"), 0)
                    If lineapedido.RowState = DataRowState.Modified Then
                        OriginalQInterna = lineapedido("QInterna", DataRowVersion.Original)
                    End If

                    Dim DiferenciaQInterna As Double = ProposedQInterna - OriginalQInterna
                    dt.Rows(0)("QPedida") = dt.Rows(0)("QPedida") + DiferenciaQInterna

                    If Length(dt.Rows(0)("FechaEntrega")) = 0 Then dt.Rows(0)("FechaEntrega") = lineapedido("FechaEntrega")

                    BusinessHelper.UpdateTable(dt)
                End If
            End If
        End If

    End Sub

#End Region

#Region " Actualización de OT  (Mantenimiento) "

    <Task()> Public Shared Sub ActualizarMantenimiento(ByVal Doc As DocumentoPedidoCompra, ByVal services As ServiceProvider)
        If Doc Is Nothing Then Exit Sub
        For Each lineapedido As DataRow In Doc.dtLineas.Select()
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoPedidoCompra.ActualizarLineaMantenimiento, lineapedido, services)
        Next
    End Sub

    <Task()> Public Shared Sub ActualizarLineaMantenimiento(ByVal lineapedido As DataRow, ByVal services As ServiceProvider)
        If Not IsNothing(lineapedido) AndAlso Length(lineapedido("IDMntoOTPrev")) > 0 Then
            Dim CambioIDMntoOT As Boolean
            Dim CambioQPedida As Boolean

            If lineapedido.RowState = DataRowState.Modified Then
                CambioIDMntoOT = (lineapedido("IDMntoOTPrev") <> lineapedido("IDMntoOTPrev", DataRowVersion.Original))
                CambioQPedida = (lineapedido("QPedida") <> lineapedido("QPedida", DataRowVersion.Original))
            End If

            If lineapedido.RowState = DataRowState.Added OrElse CambioIDMntoOT OrElse CambioQPedida Then
                Dim OTPrev As BusinessHelper = BusinessHelper.CreateBusinessObject("MntoOTPrevLinea")
                Dim dtOT As DataTable = OTPrev.SelOnPrimaryKey(lineapedido("IDMntoOTPrev"))
                If Not dtOT Is Nothing AndAlso dtOT.Rows.Count > 0 Then
                    'If lineapedido("QPedida") = 0 Then
                    '    dtOT.Rows(0)("QPedida") = dtOT.Rows(0)("QPedida") - lineapedido("QPedida")
                    'Else
                    Dim dblQModificada As Integer
                    If lineapedido.RowState = DataRowState.Modified Then
                        dblQModificada = lineapedido("QPedida", DataRowVersion.Original)
                    End If
                    dtOT.Rows(0)("QPedida") = Nz(dtOT.Rows(0)("QPedida"), 0) + (lineapedido("QPedida") - dblQModificada)
                    'End If

                    BusinessHelper.UpdateTable(dtOT)
                End If
            End If
        End If

    End Sub

#End Region

#Region " Actualización de Subcontrataciones (Orden Ruta)"

    <Task()> Public Shared Sub ActualizarOrdenRutaDesdePedidoCompra(ByVal Doc As DocumentoPedidoCompra, ByVal services As ServiceProvider)
        If Doc Is Nothing OrElse Doc.dtLineas Is Nothing Then Exit Sub
        Dim f As New Filter
        f.Add(New IsNullFilterItem("IDOrdenRuta", False))
        Dim WhereNotNullOrdenRuta As String = f.Compose(New AdoFilterComposer)
        For Each lineaPedido As DataRow In Doc.dtLineas.Select(WhereNotNullOrdenRuta, "IDPedido,IDLineaPedido")
            If lineaPedido("TipoLineaCompra") = enumaclTipoLineaAlbaran.aclSubcontratacion Then
                ProcessServer.ExecuteTask(Of Object)(AddressOf ActualizarQEnviadaOrdenRuta, lineaPedido, services)
            End If
        Next
    End Sub
    <Task()> Public Shared Sub ActualizarQEnviadaOrdenRuta(ByVal lineapedido As DataRow, ByVal services As ServiceProvider)
        If IsNumeric(lineapedido("IDOrdenRuta")) AndAlso lineapedido("IDOrdenRuta") <> 0 AndAlso lineapedido("TipoLineaCompra") = enumaclTipoLineaAlbaran.aclSubcontratacion Then
            Dim CambioIDOrdenRuta As Boolean
            Dim CambioQPedida As Boolean
            Dim CambioQServida As Boolean

            If lineapedido.RowState = DataRowState.Modified Then
                CambioIDOrdenRuta = (lineapedido("IDOrdenRuta") <> Nz(lineapedido("IDOrdenRuta", DataRowVersion.Original), -1))
                CambioQPedida = (lineapedido("QPedida") <> lineapedido("QPedida", DataRowVersion.Original))
                CambioQServida = (lineapedido("QServida") <> lineapedido("QServida", DataRowVersion.Original))
            End If

            If lineapedido.RowState = DataRowState.Added OrElse CambioIDOrdenRuta OrElse CambioQPedida OrElse CambioQServida Then
                Dim IDOrdenRuta As Integer = lineapedido("IDOrdenRuta")
                Dim QPedida As Double = Nz(lineapedido("QPedida"), 0)
                Dim QServida As Double = Nz(lineapedido("QServida"), 0)
                Dim OC As BusinessHelper = BusinessHelper.CreateBusinessObject(AdminData.GetEntityInfo("OrdenRuta"))
                Dim dt As DataTable = OC.SelOnPrimaryKey(IDOrdenRuta)
                If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                    Dim Factor As Double = Nz(lineapedido("Factor"), 1)
                    If Factor <= 0 Then Factor = 1
                    Select Case lineapedido.RowState
                        Case DataRowState.Added
                            dt.Rows(0)("QEnviada") = dt.Rows(0)("QEnviada") + (Factor * QPedida)
                        Case DataRowState.Modified
                            Dim QPedidaOriginal As Double = lineapedido("QPedida", DataRowVersion.Original)
                            Dim QServidaOriginal As Double = lineapedido("QServida", DataRowVersion.Original)
                            Dim IncQPedida As Double = QPedida - QPedidaOriginal
                            Dim IncQServida As Double = QServida - QServidaOriginal
                            dt.Rows(0)("QEnviada") = dt.Rows(0)("QEnviada") + (Factor * (IncQPedida - IncQServida))
                            If dt.Rows(0)("QEnviada") < 0 Then dt.Rows(0)("QEnviada") = 0
                        Case DataRowState.Deleted
                            dt.Rows(0)("QEnviada") -= (QPedida - QServida)
                            If dt.Rows(0)("QEnviada") < 0 Then dt.Rows(0)("QEnviada") = 0
                    End Select
                    BusinessHelper.UpdateTable(dt)
                End If
            End If
        End If
    End Sub

#End Region

#Region " Actualización de Pedido de Venta "

    <Task()> Public Shared Sub ActualizarPedidoVenta(ByVal Doc As DocumentoPedidoCompra, ByVal services As ServiceProvider)
        Dim dtPedidoVentaOrigen As DataTable = New PedidoVentaCabecera().SelOnPrimaryKey(Doc.Cabecera.IDOrigen)
        If dtPedidoVentaOrigen.Rows.Count > 0 Then
            Dim Cabecera As PedCabCompraPedidoVenta = CType(Doc.Cabecera, PedCabCompraPedidoVenta)
            dtPedidoVentaOrigen.Rows(0)("EmpresaGrupo") = Cabecera.Multiempresa
            dtPedidoVentaOrigen.Rows(0)("EntregaProveedor") = Cabecera.EntregaProveedor
            BusinessHelper.UpdateTable(dtPedidoVentaOrigen)
        End If
    End Sub

#End Region

#End Region

#Region " Subcontratación "

    '<Task()> Public Shared Function CrearComponentesSubcontratacion(ByVal data As DocumentoPedidoCompra, ByVal services As ServiceProvider) As DataTable
    '    If data.Doc.Cabecera.Origen <> enumOrigenPedidoCompra.Subcontratacion Then Exit Function
    '    Dim datSubcont As New DataDocRow(data.doc, data.linea)
    '    ProcessServer.ExecuteTask(Of DataDocRow)(AddressOf CrearComponentesSubcontratacionPC, datSubCon, services)
    '    Return data.Doc.dtLineas
    'End Function

    <Task()> Public Shared Sub CrearComponentesSubcontratacionPC(ByVal Doc As DocumentoPedidoCompra, ByVal services As ServiceProvider)
        If Not Doc.dtLineas Is Nothing AndAlso Doc.dtLineas.Rows.Count > 0 Then
            Dim fSubcontratacion As New Filter
            fSubcontratacion.Add(New NumberFilterItem("TipoLineaCompra", enumaclTipoLineaAlbaran.aclSubcontratacion))
            Dim LineasSubcontr As String = fSubcontratacion.Compose(New AdoFilterComposer)
            For Each linea As DataRow In Doc.dtLineas.Select(LineasSubcontr, Nothing, DataViewRowState.Added)
                Dim datSubcon As New DataDocRow(Doc, linea)
                ProcessServer.ExecuteTask(Of DataDocRow)(AddressOf CrearComponentesSubcontratacion, datSubcon, services)
            Next
        End If
    End Sub

    <Task()> Public Shared Sub CrearComponentesSubcontratacion(ByVal data As DataDocRow, ByVal services As ServiceProvider)
        Dim doc As DocumentoPedidoCompra = CType(data.Doc, DocumentoPedidoCompra)
        '//Si viene de un proceso, comprobamos que se genera desde subcontrataciones
        If Not doc.Cabecera Is Nothing AndAlso doc.Cabecera.Origen <> enumOrigenPedidoCompra.Subcontratacion Then Exit Sub
        '//data.Row es la línea de subcontratación
        Dim AppParamsCompra As ParametroCompra = services.GetService(Of ParametroCompra)()
        If data.Row.RowState = DataRowState.Added AndAlso data.Doc.HeaderRow("IDTipoCompra") = AppParamsCompra.TipoCompraSubcontratacion AndAlso data.Row("TipoLineaCompra") = enumaclTipoLineaAlbaran.aclSubcontratacion Then
            Dim dtComponentes As DataTable = ProcessServer.ExecuteTask(Of DataRow, DataTable)(AddressOf ComponentesDeSubcontratacion, data.Row, services)
            If Not dtComponentes Is Nothing AndAlso dtComponentes.Rows.Count > 0 Then
                Dim pcl As New PedidoCompraLinea
                For Each drComponente As DataRow In dtComponentes.Rows
                    Dim context As New BusinessData(doc.HeaderRow)
                    If data.Row.Table.Columns.Contains("IDOrdenRuta") Then context("IDOrdenRuta") = data.Row("IDOrdenRuta")
                    If data.Row.Table.Columns.Contains("TipoLineaCompra") Then context("TipoLineaCompra") = enumaclTipoLineaAlbaran.aclComponente
                    Dim drLinea As DataRow = pcl.AddNewForm.Rows(0)
                    drLinea("IDLineaPedido") = AdminData.GetAutoNumeric
                    drLinea("IDPedido") = data.Doc.HeaderRow("IDPedido")
                    pcl.ApplyBusinessRule("IDArticulo", drComponente("IDComponente"), drLinea, context)

                    If data.Row.Table.Columns.Contains("IDOrdenRuta") Then
                        drLinea("IDOrdenRuta") = data.Row("IDOrdenRuta")
                        If Length(drComponente("IDAlmacen")) > 0 Then
                            drLinea("IDAlmacen") = drComponente("IDAlmacen")
                        End If
                    End If

                    drLinea("IDCentroGestion") = data.Row("IDCentroGestion")
                    drLinea("FechaEntrega") = data.Row("FechaEntrega")
                    drLinea("QInterna") = (Nz(data.Row("QInterna"), 0) * Nz(drComponente("Cantidad"), 0)) * (1 + (Nz(drComponente("Merma"), 0)) / 100)
                    drLinea("Precio") = 0
                    drLinea("Importe") = 0
                    drLinea("ImporteA") = 0
                    drLinea("ImporteB") = 0
                    drLinea("TipoLineaCompra") = enumaclTipoLineaAlbaran.aclComponente
                    drLinea("IDLineaPadre") = data.Row("IDLineaPedido") 'Linea Padre
                    If Length(drComponente("DescComponente")) Then
                        drLinea("DescArticulo") = drComponente("DescComponente") & String.Empty
                    End If
                    drLinea("Factor") = Nz(drComponente("FactorProduccion"), 1)
                    If drLinea("Factor") > 0 Then
                        drLinea("QPedida") = drLinea("QInterna") / drLinea("Factor")
                    Else
                        drLinea("QPedida") = drLinea("QInterna")
                    End If
                    drLinea("IDUDInterna") = drComponente("IDUDInterna")
                    drLinea("IDUDMedida") = drComponente("IdUdCompra")
                    drLinea("IDOrdenLinea") = data.Row("IDOrdenLinea")
                    If Length(drComponente("IDCContable")) > 0 Then
                        drLinea("CContable") = drComponente("IDCContable")
                    End If
                    doc.dtLineas.Rows.Add(drLinea.ItemArray)
                Next
            End If
        End If
    End Sub

    <Task()> Public Shared Function ComponentesDeSubcontratacion(ByVal Origen As DataRow, ByVal services As ServiceProvider) As DataTable
        Dim dtComponentes As DataTable
        If Length(Origen("IDArticulo")) > 0 And Origen.Table.Columns.Contains("IDOrdenRuta") AndAlso Length(Origen("IDOrdenRuta")) > 0 Then
            Dim ClsOP As BusinessHelper = BusinessHelper.CreateBusinessObject("OrdenRuta")
            Dim DtOp As DataTable = ClsOP.SelOnPrimaryKey(Origen("IDOrdenRuta"))
            Dim f As New Filter

            If Not DtOp Is Nothing AndAlso DtOp.Rows.Count > 0 AndAlso Nz(DtOp.Rows(0)("IDOrden"), 0) > 0 Then
                f.Add(New NumberFilterItem("IDOrden", FilterOperator.Equal, DtOp.Rows(0)("IDOrden")))
            End If
            If Not DtOp Is Nothing AndAlso DtOp.Rows.Count > 0 AndAlso Nz(DtOp.Rows(0)("Secuencia"), 0) > 0 Then
                f.Add(New NumberFilterItem("Secuencia", FilterOperator.Equal, DtOp.Rows(0)("Secuencia")))
            End If
            f.Add(New StringFilterItem("IDArticulo", FilterOperator.Equal, Origen("IDArticulo")))
            dtComponentes = New BE.DataEngine().Filter("vNegComponentesSubcontratacion", f)
        Else
            Dim f As New Filter
            f.Add(New StringFilterItem("IDArticulo", FilterOperator.Equal, Origen("IDArticulo")))
            dtComponentes = New BE.DataEngine().Filter("vNegComponentesSubcontratacionArticulo", f)
        End If
        Return dtComponentes
    End Function

    <Task()> Public Shared Sub ActualizarComponentesSubcontratacion(ByVal Doc As DocumentoPedidoCompra, ByVal services As ServiceProvider)
        For Each LineaPedido As DataRow In Doc.dtLineas.Select
            If (LineaPedido.RowState = DataRowState.Modified) AndAlso (LineaPedido("TipoLineaCompra") = enumaclTipoLineaAlbaran.aclSubcontratacion) Then
                Dim datos As New DataDocRow(Doc, LineaPedido)
                Dim f As New Filter
                f.Add(New NumberFilterItem("IDPedido", LineaPedido("IDPedido")))
                f.Add(New NumberFilterItem("IDLineaPadre", LineaPedido("IDLineaPedido")))
                f.Add(New NumberFilterItem("TipoLineaCompra", enumaclTipoLineaAlbaran.aclComponente))
                Dim WhereComponentes As String = f.Compose(New AdoFilterComposer)
                Dim Componentes() As DataRow = CType(Doc, DocumentCabLin).dtLineas.Select(WhereComponentes)
                If Not Componentes Is Nothing AndAlso Componentes.Length > 0 Then
                    If Nz(LineaPedido("QInterna", DataRowVersion.Original), 0) <> 0 Then
                        Dim factorVariacion As Double = LineaPedido("QInterna") / LineaPedido("QInterna", DataRowVersion.Original)
                        For Each componente As DataRow In Componentes
                            componente("QPedida") = componente("QPedida") * factorVariacion
                            componente("QInterna") = componente("QInterna") * factorVariacion
                        Next
                    End If
                End If

            End If
        Next
    End Sub

#End Region

#Region " Crear Cabecera de Pedido "

    <Task()> Public Shared Function GeneraPedidoCompraAbierto(ByVal data As String, ByVal services As ServiceProvider) As CreateElement
        Dim dtPrograma As DataTable = New ProgramaCompraCabecera().SelOnPrimaryKey(data)
        If Not IsNothing(dtPrograma) AndAlso dtPrograma.Rows.Count > 0 Then
            Dim drPrograma As DataRow = dtPrograma.Rows(0)
            Dim PCC As New PedidoCompraCabecera()
            Dim dtPCC As DataTable = PCC.AddNewForm()
            If Not IsNothing(dtPCC) AndAlso dtPCC.Rows.Count > 0 Then
                Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
                Dim drRowPCC As DataRow = dtPCC.Rows(0)
                drRowPCC("IDProveedor") = drPrograma("IDProveedor")
                drRowPCC("FechaPedido") = Today

                '//Asignar Dirección
                drRowPCC("IDDireccion") = System.DBNull.Value
                If Length(drPrograma("IDDireccionPedido")) > 0 Then
                    drRowPCC("IDDireccion") = drPrograma("IDDireccionPedido")
                Else
                    Dim dataDir As New ProveedorDireccion.DataDirecEnvio
                    dataDir.IDProveedor = drRowPCC("IDProveedor")
                    dataDir.TipoDireccion = enumpdTipoDireccion.pdDireccionPedido
                    Dim dtDireccion As DataTable = ProcessServer.ExecuteTask(Of ProveedorDireccion.DataDirecEnvio, DataTable)(AddressOf ProveedorDireccion.ObtenerDireccionEnvio, dataDir, services)
                    If Not dtDireccion Is Nothing AndAlso dtDireccion.Rows.Count > 0 Then
                        drRowPCC("IDDireccion") = dtDireccion.Rows(0)("IDDireccion")
                    End If
                End If

                drRowPCC("IDOperario") = drPrograma("IDOperario")
                drRowPCC("IDCentroGestion") = drPrograma("IDCentroGestion")
                drRowPCC("IDFormaPago") = drPrograma("IDFormaPago")
                drRowPCC("IDCondicionPago") = drPrograma("IDCondicionPago")
                drRowPCC("IDFormaEnvio") = drPrograma("IDFormaEnvio")
                drRowPCC("IDCondicionEnvio") = drPrograma("IDCondicionEnvio")
                drRowPCC("IDAlmacen") = drPrograma("IDAlmacen")
                PCC.ApplyBusinessRule("IDMoneda", drPrograma("IDMoneda"), drRowPCC, Nothing)

                Dim StDatos As New Contador.DatosCounterValue(drRowPCC("IDContador"), New PedidoCompraCabecera, "NPedido", "FechaPedido", drRowPCC("FechaPedido"))
                StDatos.IDEjercicio = drRowPCC("IDEjercicio") & String.Empty
                drRowPCC("NPedido") = ProcessServer.ExecuteTask(Of Contador.DatosCounterValue, String)(AddressOf Contador.CounterValue, StDatos, services)

                dtPCC = PCC.Update(dtPCC)
                If Not dtPCC Is Nothing AndAlso dtPCC.Rows.Count > 0 Then
                    Dim e As New CreateElement
                    e.IDElement = dtPCC.Rows(0)("IDPedido")
                    e.NElement = dtPCC.Rows(0)("NPedido")
                    Return e
                End If
            End If
        End If


    End Function

#End Region

#Region " Multiempresa "

    Public Class DataValidarProveedorAsociado
        Public IDProveedor As String
        Public IDCliente As String

        Public Sub New(ByVal IDProveedor As String)
            Me.IDProveedor = IDProveedor
        End Sub
    End Class
    <Task()> Public Shared Sub ValidarProveedorAsociado(ByVal data As DataValidarProveedorAsociado, ByVal services As ServiceProvider)
        If Length(data.IDProveedor) > 0 Then
            '//NOTA: El campo IDProveedorAsociado de la tabla tbMaestroCliente, es compartido por todas las BBDD del grupo.
            Dim dtCliente As DataTable = New Cliente().Filter(New StringFilterItem("IDProveedorAsociado", data.IDProveedor))
            If dtCliente.Rows.Count = 0 Then
                ApplicationService.GenerateError("El Proveedor del grupo no está asociado a ninguna Empresa.")
            ElseIf dtCliente.Rows.Count > 1 Then
                ApplicationService.GenerateError("El Proveedor del grupo está asociado a más de una Empresa.")
            Else
                data.IDCliente = dtCliente.Rows(0)("IDCliente")
            End If
        End If
    End Sub

#End Region

End Class
