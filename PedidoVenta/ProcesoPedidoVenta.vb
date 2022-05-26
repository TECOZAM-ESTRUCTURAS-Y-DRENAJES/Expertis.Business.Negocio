Public Class ProcesoPedidoVenta

#Region " Creación de instancias de documento "

    <Task()> Public Shared Function CrearDocumentoPedidoVenta(ByVal data As PedCab, ByVal services As ServiceProvider) As DocumentoPedidoVenta
        Dim Doc As DocumentoPedidoVenta
        Select Case data.Origen
            Case enumOrigenPedido.Programa
                If Length(CType(data, PedCabPrograma).IDPedido) = 0 OrElse CType(data, PedCabPrograma).IDPedido = 0 Then
                    Doc = New DocumentoPedidoVenta(data, services)
                Else
                    Doc = New DocumentoPedidoVenta(CType(data, PedCabPrograma).IDPedido)
                    Doc.Cabecera = data 'Debemos indicar la Cabecera en este caso
                End If
            Case enumOrigenPedido.Oferta, enumOrigenPedido.PedidoCompra, enumOrigenPedido.Copia
                Doc = New DocumentoPedidoVenta(data, services)
            Case enumOrigenPedido.EDI
                If Not CType(data, PedCabEDI).IDPedido.HasValue Then
                    Doc = New DocumentoPedidoVenta(data, services)
                Else
                    Doc = New DocumentoPedidoVenta(CType(data, PedCabEDI).IDPedido)
                    Doc.Cabecera = data 'Debemos indicar la Cabecera en este caso
                End If
        End Select
        Return Doc
    End Function

    <Task()> Public Shared Function CrearDocumento(ByVal UpdtCtx As UpdatePackage, ByVal services As ServiceProvider) As DocumentoPedidoVenta
        Return New DocumentoPedidoVenta(UpdtCtx)
    End Function

#End Region

#Region " Agrupación "

    <Task()> Public Shared Function AgruparOfertas(ByVal data As DataPrcCrearPedidoOfertaComercial, ByVal services As ServiceProvider) As PedCabOfertaComercial()
        If data.Detalle Then
            Return ProcessServer.ExecuteTask(Of DataPrcCrearPedidoOfertaComercial, PedCabOfertaComercial())(AddressOf AgruparOfertaDetalle, data, services)
        Else
            Return ProcessServer.ExecuteTask(Of DataPrcCrearPedidoOfertaComercial, PedCabOfertaComercial())(AddressOf AgruparOferta, data, services)
        End If
    End Function
    <Task()> Public Shared Function AgruparOferta(ByVal data As DataPrcCrearPedidoOfertaComercial, ByVal services As ServiceProvider) As PedCabOfertaComercial()
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
            oFltr.Add(New BooleanFilterItem("LanzarVenta", True))
            oFltr.Add(New NumberFilterItem("EstadoVenta", enumocdEstadoCompraVenta.ecvPendiente))

            Dim strViewName As String = "vfrmOfertaComercialTratamiento"
            Dim dtOfertas As DataTable = New BE.DataEngine().Filter(strViewName, oFltr)
            ProcessServer.ExecuteTask(Of DataTable)(AddressOf ValidarDatosOfertas, dtOfertas, services)

            '//Rellenamos las clases que nos servirán de semilla para crear el pedido
            Dim oGrprUser As New GroupUserOfertasComerciales()

            '//Se crean los agrupadores
            Dim dataColsAgrup As New DataGetGroupColumns(dtOfertas, enummcAgrupPedido.mcCliente)
            Dim GroupCols() As DataColumn = ProcessServer.ExecuteTask(Of DataGetGroupColumns, DataColumn())(AddressOf GetGroupColumns, dataColsAgrup, services)

            Dim groupers(0) As GroupHelper
            groupers(enummcAgrupPedido.mcCliente) = New GroupHelper(GroupCols, oGrprUser)

            If Not dtOfertas Is Nothing AndAlso dtOfertas.Rows.Count > 0 Then
                For Each lineaDetalle As DataRow In dtOfertas.Select(Nothing, "IDOfertaComercial,IDLineaOfertaDetalle")
                    If Length(lineaDetalle("IDArticulo")) > 0 Then
                        groupers(enummcAgrupPedido.mcCliente).Group(lineaDetalle)
                    End If
                Next
            End If

            For Each ped As PedCabOfertaComercial In oGrprUser.Peds
                If Not DirectCast(htLins(ped.IDOferta), DataOfertaComercial).FechaEntrega Is Nothing AndAlso DirectCast(htLins(ped.IDOferta), DataOfertaComercial).FechaEntrega <> cnMinDate Then
                    ped.FechaEntrega = DirectCast(htLins(ped.IDOferta), DataOfertaComercial).FechaEntrega
                End If
                If Length(DirectCast(htLins(ped.IDOferta), DataOfertaComercial).PedidoCliente) > 0 Then
                    ped.PedidoCliente = DirectCast(htLins(ped.IDOferta), DataOfertaComercial).PedidoCliente
                End If
            Next

            Return oGrprUser.Peds
        End If
    End Function
    <Task()> Public Shared Function AgruparOfertaDetalle(ByVal data As DataPrcCrearPedidoOfertaComercial, ByVal services As ServiceProvider) As PedCabOfertaComercial()
        If Not data.Ofertas Is Nothing AndAlso data.Ofertas.Length > 0 Then
            Dim htLins As New Hashtable
            Dim values(data.Ofertas.Length - 1) As Object
            For i As Integer = 0 To data.Ofertas.Length - 1
                values(i) = data.Ofertas(i).IDLineaOfertaDetalle
                htLins(data.Ofertas(i).IDOfertaComercial) = data.Ofertas(i)
            Next

            Dim strViewName As String = "vfrmOfertaComercialTratamiento"
            Dim oFltr As New Filter
            oFltr.Add(New InListFilterItem("IDLineaOfertaDetalle", values, FilterType.Numeric))
            oFltr.Add(New IsNullFilterItem("IDArticulo", False))
            oFltr.Add(New BooleanFilterItem("LanzarVenta", True))
            oFltr.Add(New NumberFilterItem("EstadoVenta", enumocdEstadoCompraVenta.ecvPendiente))
            Dim dtOfertas As DataTable = New BE.DataEngine().Filter(strViewName, oFltr)
            ProcessServer.ExecuteTask(Of DataTable)(AddressOf ValidarDatosOfertas, dtOfertas, services)

            '//Rellenamos las clases que nos servirán de semilla para crear el pedido
            Dim oGrprUser As New GroupUserOfertasComerciales()

            '//Se crean los agrupadores
            Dim dataColsAgrup As New DataGetGroupColumns(dtOfertas, enummcAgrupPedido.mcCliente)
            Dim GroupCols() As DataColumn = ProcessServer.ExecuteTask(Of DataGetGroupColumns, DataColumn())(AddressOf GetGroupColumns, dataColsAgrup, services)

            Dim groupers(0) As GroupHelper
            groupers(enummcAgrupPedido.mcCliente) = New GroupHelper(GroupCols, oGrprUser)

            If Not dtOfertas Is Nothing AndAlso dtOfertas.Rows.Count > 0 Then
                For Each ofertaDetalle As DataRow In dtOfertas.Rows
                    groupers(enummcAgrupPedido.mcCliente).Group(ofertaDetalle)
                Next
            End If

            For Each ped As PedCabOfertaComercial In oGrprUser.Peds
                If Not DirectCast(htLins(ped.IDOferta), DataOfertaComercial).FechaEntrega Is Nothing AndAlso DirectCast(htLins(ped.IDOferta), DataOfertaComercial).FechaEntrega <> cnMinDate Then
                    ped.FechaEntrega = DirectCast(htLins(ped.IDOferta), DataOfertaComercial).FechaEntrega
                End If
                If Length(DirectCast(htLins(ped.IDOferta), DataOfertaComercial).PedidoCliente) > 0 Then
                    ped.PedidoCliente = DirectCast(htLins(ped.IDOferta), DataOfertaComercial).PedidoCliente
                End If
            Next

            Return oGrprUser.Peds
        End If
    End Function
    <Task()> Public Shared Sub ValidarDatosOfertas(ByVal data As DataTable, ByVal services As ServiceProvider)
        If data Is Nothing OrElse data.Rows.Count = 0 Then
            ApplicationService.GenerateError("Las líneas seleccionadas no pueden generar Pedidos de Venta. Compruebe el/los artículo/s, si pueden generar Pedidos de Venta y su estado de actualización.")
        End If

        Dim f As New Filter
        f.Add(New IsNullFilterItem("IDCliente"))
        f.Add(New IsNullFilterItem("IDEmpresa"))
        Dim WhereNullEmpresaCliente As String = f.Compose(New AdoFilterComposer)
        Dim adr() As DataRow = data.Select(WhereNullEmpresaCliente)
        If Not adr Is Nothing AndAlso adr.Length > 0 Then
            ApplicationService.GenerateError("Cliente y Empresa no pueden estar vacíos a la vez. Revise las ofertas seleccionadas.")
        End If

        f.Clear()
        f.Add(New IsNullFilterItem("IDPresupuesto", False))
        f.Add(New IsNullFilterItem("IDArticulo"))
        Dim WhereNullArticulo As String = f.Compose(New AdoFilterComposer)
        adr = data.Select(WhereNullArticulo)
        If Not adr Is Nothing AndAlso adr.Length > 0 Then
            ApplicationService.GenerateError("Es necesario asociar los Presupuestos de la Oferta con Artículos. Revise las ofertas seleccionadas.")
        End If

    End Sub

    <Task()> Public Shared Function AgruparProgramas(ByVal data As DataPrcCrearPedidoVentaPrograma, ByVal services As ServiceProvider) As PedCabPrograma()
        If Not data.Programas Is Nothing AndAlso data.Programas.Length > 0 Then
            '//se seleccionan todas las lineas de programa a confirmar
            Dim strViewName As String = "vfrmConfirmacionPrograma"

            Dim htLins As New Hashtable
            Dim values(data.Programas.Length - 1) As Object
            For i As Integer = 0 To data.Programas.Length - 1
                values(i) = data.Programas(i).IDLineaPrograma
                htLins.Add(data.Programas(i).IDLineaPrograma, data.Programas(i))
            Next

            Dim oFltr As New Filter
            oFltr.Add(New InListFilterItem("IDLineaPrograma", values, FilterType.Numeric))
            'oFltr.Add(New NumberFilterItem("Confirmada", enumplEstadoLinea.plNoConfirmada))
            Dim dtProgramas As DataTable = New BE.DataEngine().Filter(strViewName, oFltr)

            '//Rellenamos las clases que nos servirán de semilla para crear el pedido
            Dim oGrprUser As New GroupUserProgramas()

            '//Se crean los agrupadores
            Dim dataColsAgrup As New DataGetGroupColumns(dtProgramas, enummcAgrupPedido.mcPrograma)
            Dim GroupCols() As DataColumn = ProcessServer.ExecuteTask(Of DataGetGroupColumns, DataColumn())(AddressOf GetGroupColumns, dataColsAgrup, services)
            Dim groupers(1) As GroupHelper
            groupers(enummcAgrupPedido.mcPrograma) = New GroupHelper(GroupCols, oGrprUser)

            dataColsAgrup = New DataGetGroupColumns(dtProgramas, enummcAgrupPedido.mcCliente)
            GroupCols = ProcessServer.ExecuteTask(Of DataGetGroupColumns, DataColumn())(AddressOf GetGroupColumns, dataColsAgrup, services)
            groupers(enummcAgrupPedido.mcCliente) = New GroupHelper(GroupCols, oGrprUser)

            '//A través de los agrupadores 
            For Each rwLin As DataRow In dtProgramas.Rows
                If rwLin("Confirmada") = enumplEstadoLinea.plNoConfirmada OrElse rwLin("QPrevista") > rwLin("QConfirmada") OrElse rwLin("QPrevista") > rwLin("QPendiente") Then
                    groupers(rwLin("AgrupPedido")).Group(rwLin)
                End If
            Next

            For Each ped As PedCabPrograma In oGrprUser.Peds
                For Each pedlin As PedLinPrograma In ped.Lineas
                    pedlin.QConfirmada = DirectCast(htLins(pedlin.IDLineaPrograma), DataPrograma).QConfirmada
                    pedlin.FechaConfirmacion = DirectCast(htLins(pedlin.IDLineaPrograma), DataPrograma).FechaConfirmacion
                    pedlin.Cantidad = DirectCast(htLins(pedlin.IDLineaPrograma), DataPrograma).QConfirmar
                Next
            Next

            Return oGrprUser.Peds
        End If
    End Function

    Public Class DataGetGroupColumns
        Public Agrupacion As enummcAgrupPedido
        Public Datos As DataTable

        Public Sub New(ByVal Datos As DataTable, ByVal Agrupacion As enummcAgrupPedido)
            Me.Datos = Datos
            Me.Agrupacion = Agrupacion
        End Sub
    End Class

    <Task()> Public Shared Function GetGroupColumns(ByVal data As DataGetGroupColumns, ByVal services As ServiceProvider) As DataColumn()
        '//Se definen las columnas que nos permitirán abrir un pedido nuevo
        Dim columns(1) As DataColumn
        columns(0) = data.Datos.Columns("IDCliente")
        columns(1) = data.Datos.Columns("IdMoneda")
        'columns(2) = table.Columns("EDI")
        If data.Agrupacion = enummcAgrupPedido.mcPrograma Then
            ReDim Preserve columns(3)
            columns(2) = data.Datos.Columns("IDPrograma")
            columns(3) = data.Datos.Columns("IDPedido")
            'Else
            '    ReDim Preserve columns(2)
            '    columns(2) = table.Columns("IDPedido")
        End If

        Return columns
    End Function

#End Region

#Region " Ordenar Pedidos (de Programa)"
    'Ordena los pedidos teniendo en cuenta el cliente, moneda y programa
    <Task()> Public Shared Sub Ordenar(ByVal data As PedCabPrograma(), ByVal services As ServiceProvider)
        If data IsNot Nothing Then Array.Sort(data, New OrdenPrograma)
    End Sub
#End Region
#Region "Analítica y representantes"

    <Task()> Public Shared Sub CalcularRepresentantes(ByVal Doc As DocumentoPedidoVenta, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DocumentoPedidoVenta)(AddressOf ProcesoComercial.CalcularRepresentantes, Doc, services)
    End Sub
    <Task()> Public Shared Sub CalcularAnalitica(ByVal Doc As DocumentoPedidoVenta, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DocumentoPedidoVenta)(AddressOf NegocioGeneral.CalcularAnalitica, Doc, services)
    End Sub
#End Region

#Region " Asignar datos "

    <Task()> Public Shared Sub AsignarValoresPredeterminadosGenerales(ByVal Doc As DocumentoPedidoVenta, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoComunes.AsignarIdentificadorPedido, Doc.HeaderRow, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoComunes.AsignarFechaPedido, Doc.HeaderRow, services)
        ProcessServer.ExecuteTask(Of DataRowPropertyAccessor)(AddressOf ProcesoComunes.AsignarEjercicioContablePedido, New DataRowPropertyAccessor(Doc.HeaderRow), services)
        ProcessServer.ExecuteTask(Of DocumentoPedidoVenta)(AddressOf AsignarFechaEntregaDesdeOC, Doc, services)
        ProcessServer.ExecuteTask(Of DocumentoPedidoVenta)(AddressOf AsignarFechaEntrega, Doc, services)
        ProcessServer.ExecuteTask(Of DocumentoPedidoVenta)(AddressOf AsignarResponsable, Doc, services)
    End Sub

    <Task()> Public Shared Sub AsignarFechaEntregaDesdeOC(ByVal Doc As DocumentoPedidoVenta, ByVal services As ServiceProvider)
        If Not Doc.HeaderRow Is Nothing And Not Doc.Cabecera Is Nothing Then
            If Doc.HeaderRow.IsNull("FechaEntrega") AndAlso Doc.Cabecera.Origen = enumOrigenPedido.Oferta Then
                Doc.HeaderRow("FechaEntrega") = CType(Doc.Cabecera, PedCabOfertaComercial).FechaEntrega
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarPedidoCliente(ByVal Doc As DocumentoPedidoVenta, ByVal services As ServiceProvider)
        If Not Doc.HeaderRow Is Nothing And Not Doc.Cabecera Is Nothing Then
            Doc.HeaderRow("PedidoCliente") = Doc.Cabecera.PedidoCliente
        End If
    End Sub

    <Task()> Public Shared Sub AsignarResponsable(ByVal Doc As DocumentoPedidoVenta, ByVal services As ServiceProvider)
        If Not Doc.HeaderRow Is Nothing And Not Doc.Cabecera Is Nothing Then
            Dim strIDOper As String = ProcessServer.ExecuteTask(Of Object, String)(AddressOf Operario.ObtenerIDOperarioUsuario, Nothing, services)
            If Len(strIDOper) > 0 Then Doc.HeaderRow("Responsable") = strIDOper
        End If
    End Sub

    <Task()> Public Shared Sub AsignarFechaEntrega(ByVal Doc As DocumentoPedidoVenta, ByVal services As ServiceProvider)
        If Not Doc.HeaderRow Is Nothing AndAlso (Not Doc.dtLineas Is Nothing AndAlso Doc.dtLineas.Rows.Count > 0) Then
            If Doc.HeaderRow.IsNull("FechaEntrega") AndAlso IsDate(Doc.dtLineas.Rows(0)("FechaEntrega")) Then
                Doc.HeaderRow("FechaEntrega") = Doc.dtLineas.Rows(0)("FechaEntrega")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarFechaEntregaOrigen(ByVal data As DataDocRowOrigen, ByVal services As ServiceProvider)
        'If data.IsNull("FechaEntrega") Then data("FechaEntrega") = data("FechaPedido")
        Select Case CType(data.Doc, DocumentoPedidoVenta).Cabecera.Origen
            Case enumOrigenPedido.Programa, enumOrigenPedido.PedidoCompra
                If IsDate(data.RowOrigen("FechaEntrega")) Then
                    data.RowDestino("FechaEntrega") = data.RowOrigen("FechaEntrega")
                Else
                    data.RowDestino("FechaEntrega") = Today
                End If
            Case enumOrigenPedido.Oferta
                data.RowDestino("FechaEntrega") = CType(data.Doc, DocumentoPedidoVenta).HeaderRow("FechaEntrega")
            Case enumOrigenPedido.EDI
                If IsDate(data.RowOrigen("FechaEntrega")) Then
                    data.RowDestino("FechaEntrega") = data.RowOrigen("FechaEntrega")
                End If
        End Select
    End Sub

    <Task()> Public Shared Sub AsignarNumeroPedido(ByVal Doc As DocumentoPedidoVenta, ByVal services As ServiceProvider)
        If Doc.HeaderRow.RowState = DataRowState.Added Then
            If Not Doc.HeaderRow.IsNull("IDContador") Then
                Dim StDatos As New Contador.DatosCounterValue
                StDatos.IDCounter = Doc.HeaderRow("IDContador")
                StDatos.TargetClass = New PedidoVentaCabecera
                StDatos.TargetField = "NPedido"
                StDatos.DateField = "FechaPedido"
                StDatos.DateValue = Doc.HeaderRow("FechaPedido")
                StDatos.IDEjercicio = Doc.HeaderRow("IDEjercicio") & String.Empty
                Doc.HeaderRow("NPedido") = ProcessServer.ExecuteTask(Of Contador.DatosCounterValue, String)(AddressOf Contador.CounterValue, StDatos, services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarDatosCabeceraEDI(ByVal Doc As DocumentoPedidoVenta, ByVal services As ServiceProvider)
        If Doc.HeaderRow.RowState = DataRowState.Added Then
            Select Case Doc.Cabecera.Origen
                Case enumOrigenPedido.EDI
                    Dim p As PedCabEDI = Doc.Cabecera
                    If p.DepartamentoEDI.Length > 0 Then
                        Doc.HeaderRow("DepartamentoEDI") = p.DepartamentoEDI
                    End If
                    If p.SeccionEDI.Length > 0 Then
                        Doc.HeaderRow("SeccionEDI") = p.SeccionEDI
                    End If
            End Select
        End If
    End Sub

    <Task()> Public Shared Sub AsignarDatosCliente(ByVal Doc As DocumentoPedidoVenta, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DocumentoPedidoVenta)(AddressOf ProcesoComercial.AsignarDatosCliente, Doc, services)

        If Doc.Cliente Is Nothing Then
            Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
            Doc.Cliente = Clientes.GetEntity(Doc.HeaderRow("IDCliente"))
        End If

        If Doc.Cliente.Bloqueado Then ApplicationService.GenerateError("El Cliente está Bloqueado.")

        If Doc.HeaderRow.IsNull("IDFormaEnvio") Then Doc.HeaderRow("IDFormaEnvio") = Doc.Cliente.FormaEnvio
        If Doc.HeaderRow.IsNull("IDCondicionEnvio") Then Doc.HeaderRow("IDCondicionEnvio") = Doc.Cliente.CondicionEnvio
        If Doc.HeaderRow.IsNull("DtoPedido") Then Doc.HeaderRow("DtoPedido") = Doc.Cliente.DtoComercial
        If Doc.HeaderRow.IsNull("Prioridad") Then Doc.HeaderRow("Prioridad") = Doc.Cliente.Prioridad
    End Sub

    <Task()> Public Shared Sub AsignarDireccionEnvio(ByVal Doc As DocumentoPedidoVenta, ByVal services As ServiceProvider)
        Select Case Doc.Cabecera.Origen
            Case enumOrigenPedido.Programa
                If CType(Doc.Cabecera, PedCabPrograma).IDDireccionEnvio <> 0 Then Doc.HeaderRow("IDDireccionEnvio") = CType(Doc.Cabecera, PedCabPrograma).IDDireccionEnvio
            Case enumOrigenPedido.PedidoCompra
                If CType(Doc.Cabecera, PedCabVentaPedidoCompra).IDDireccionEnvio <> 0 Then Doc.HeaderRow("IDDireccionEnvio") = CType(Doc.Cabecera, PedCabVentaPedidoCompra).IDDireccionEnvio
            Case enumOrigenPedido.Oferta
                If CType(Doc.Cabecera, PedCabOfertaComercial).IDDireccionCliente <> 0 Then Doc.HeaderRow("IDDireccionEnvio") = CType(Doc.Cabecera, PedCabOfertaComercial).IDDireccionCliente
            Case enumOrigenPedido.EDI
                If CType(Doc.Cabecera, PedCabEDI).IDDireccionEnvio <> 0 Then Doc.HeaderRow("IDDireccionEnvio") = CType(Doc.Cabecera, PedCabEDI).IDDireccionEnvio
        End Select
        If Doc.HeaderRow.IsNull("IDDireccionEnvio") Then
            Dim strCliente As String = Doc.IDCliente
            If Doc.Cliente.GrupoDireccion Then strCliente = Doc.Cliente.GrupoCliente
            Dim StDatosDirec As New ClienteDireccion.DataDirecEnvio(strCliente, enumcdTipoDireccion.cdDireccionEnvio)
            Dim dtDireccion As DataTable = ProcessServer.ExecuteTask(Of ClienteDireccion.DataDirecEnvio, DataTable)(AddressOf ClienteDireccion.ObtenerDireccionEnvio, StDatosDirec, services)
            If Not dtDireccion Is Nothing AndAlso dtDireccion.Rows.Count > 0 Then
                Doc.HeaderRow("IDDireccionEnvio") = dtDireccion.Rows(0)("IDDireccion")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarFechaAviso(ByVal Doc As DocumentoPedidoVenta, ByVal services As ServiceProvider)
        If Doc.HeaderRow.RowState = DataRowState.Added Then
            Dim AppParamsVenta As ParametroVenta = services.GetService(Of ParametroVenta)()
            If AppParamsVenta.PermisoExpedicion AndAlso Doc.HeaderRow.IsNull("FechaAviso") Then Doc.HeaderRow("FechaAviso") = Doc.HeaderRow("FechaPedido")
        End If
    End Sub

    <Task()> Public Shared Sub AsignarEstadoPedido(ByVal Doc As DocumentoPedidoVenta, ByVal services As ServiceProvider)
        If Doc.HeaderRow.RowState = DataRowState.Added Then
            If Doc.HeaderRow.IsNull("Estado") Then Doc.HeaderRow("Estado") = enumpvcEstado.pvcPedido
        End If
    End Sub

    <Task()> Public Shared Sub AsignarCentroGestionPedido(ByVal Doc As DocumentoPedidoVenta, ByVal services As ServiceProvider)
        Dim AppParamsVenta As ParametroVenta = services.GetService(Of ParametroVenta)()
        If Doc.HeaderRow.IsNull("IDCentroGestion") AndAlso Length(AppParamsVenta.General.CentroGestion) > 0 Then
            Doc.HeaderRow("IDCentroGestion") = AppParamsVenta.General.CentroGestion
        End If
    End Sub

    <Task()> Public Shared Sub AsignarEstadoLinea(ByVal linea As DataRow, ByVal services As ServiceProvider)
        Dim QPedida As Double = Nz(linea("QPedida"), 0)
        Dim QServida As Double = Nz(linea("QServida"), 0)

        If QPedida < 0 Then
            If QServida <= QPedida Then
                linea("Estado") = enumpvlEstado.pvlServido
            ElseIf QServida > QPedida AndAlso QServida <> 0 Then
                linea("Estado") = enumpvlEstado.pvlParcServido
            Else
                If QServida <= 0 Then
                    linea("Estado") = enumpvlEstado.pvlPedido
                    linea("QServida") = 0
                ElseIf QServida >= QPedida Then
                    linea("Estado") = enumpvlEstado.pvlServido
                ElseIf QServida < QPedida Then
                    linea("Estado") = enumpvlEstado.pvlParcServido
                End If
            End If
        Else
            If QServida <= 0 Then
                linea("Estado") = enumpvlEstado.pvlPedido
                linea("QServida") = 0
            ElseIf QServida >= QPedida Then
                linea("Estado") = enumpvlEstado.pvlServido
            ElseIf QServida < QPedida Then
                linea("Estado") = enumpvlEstado.pvlParcServido
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarFechaPreparacion(ByVal Doc As DocumentoPedidoVenta, ByVal services As ServiceProvider)
        If Doc.HeaderRow.RowState = DataRowState.Added Then
            If Not Doc.HeaderRow.IsNull("FechaPreparacion") Then
                For Each row As DataRow In Doc.dtLineas.Rows
                    row("FechaPreparacion") = Doc.HeaderRow("FechaPreparacion")
                Next
            End If
        ElseIf Doc.HeaderRow.RowState = DataRowState.Modified Then
            If Nz(Doc.HeaderRow("FechaPreparacion"), Date.MinValue) <> Nz(Doc.HeaderRow("FechaPreparacion", DataRowVersion.Original), Date.MinValue) Then
                For Each row As DataRow In Doc.dtLineas.Rows
                    row("FechaPreparacion") = Doc.HeaderRow("FechaPreparacion")
                Next
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarConfirmacionLineas(ByVal Doc As DocumentoPedidoVenta, ByVal services As ServiceProvider)
        Dim AppParams As ParametroVenta = services.GetService(Of ParametroVenta)()
        If AppParams.ExpedirPedidosConfirmados Then
            If AppParams.PedidoConfirmado Then
                Dim dvAdded As New DataView(Doc.dtLineas, Nothing, Nothing, DataViewRowState.Added)
                For Each drv As DataRowView In dvAdded
                    drv.Row("Confirmado") = True
                    drv.Row("QAlbaran") = Nz(drv.Row("QPedida"), 0) - Nz(drv.Row("QServida"), 0)
                Next

                Dim dvModified As New DataView(Doc.dtLineas, "Confirmado=1", Nothing, DataViewRowState.ModifiedCurrent)
                For Each drv As DataRowView In dvModified
                    If Nz(drv.Row("QPedida"), 0) <> Nz(drv.Row("QPedida", DataRowVersion.Original), 0) Then
                        drv.Row("QAlbaran") = Nz(drv.Row("QPedida"), 0) - Nz(drv.Row("QServida"), 0)
                    End If
                Next
            End If
        End If
    End Sub

#End Region

#Region " Promociones "

    <Task()> Public Shared Sub TratarPromocionesLineas(ByVal Doc As DocumentoPedidoVenta, ByVal services As ServiceProvider)
        If Length(Doc.HeaderRow("IDTPV")) = 0 Then
            For Each linea As DataRow In Doc.dtLineas.Select
                If Nz(linea("Regalo"), 0) = 0 Then

                    '10. Quitamos la información anterior
                    If linea.RowState = DataRowState.Modified Then
                        If Nz(linea("IDPromocionLinea", DataRowVersion.Original), 0) <> 0 Then
                            If (linea("QPedida") <> linea("QPedida", DataRowVersion.Original) OrElse linea("IdArticulo") <> linea("IdArticulo", DataRowVersion.Original)) Then
                                Dim Dt As DataTable = Doc.dtLineas.Clone
                                Dt.ImportRow(linea)
                                Dim datActPromo As New PromocionLinea.DatosActuaLinPromoDr(Dt, True, Doc)
                                ProcessServer.ExecuteTask(Of PromocionLinea.DatosActuaLinPromoDr)(AddressOf PromocionLinea.ActualizarLineaPromocion, datActPromo, services)
                            End If
                        End If
                    End If

                    '20. Insertamos la nueva información
                    If Nz(linea("IDPromocionLinea"), 0) <> 0 Then
                        If linea.RowState = DataRowState.Added OrElse _
                          (linea("QPedida") <> linea("QPedida", DataRowVersion.Original) OrElse linea("IdArticulo") <> linea("IdArticulo", DataRowVersion.Original)) Then
                            Dim pl As New PromocionLinea
                            Dim dtPromLinea As DataTable = pl.SelOnPrimaryKey(linea("IDPromocionLinea"))
                            If Not IsNothing(dtPromLinea) AndAlso dtPromLinea.Rows.Count > 0 Then
                                If linea("QPedida") >= dtPromLinea.Rows(0)("QMinPedido") Then
                                    Dim Dt As DataTable = Doc.dtLineas.Clone
                                    Dt.ImportRow(linea)
                                    Dim datActPromo As New PromocionLinea.DatosActuaLinPromoDr(Dt, False, Doc)
                                    ProcessServer.ExecuteTask(Of PromocionLinea.DatosActuaLinPromoDr)(AddressOf PromocionLinea.ActualizarLineaPromocion, datActPromo, services)

                                    Dim datosRegalo As New DataNuevaLineaRegalo(Doc, linea, dtPromLinea.Rows(0))
                                    ProcessServer.ExecuteTask(Of DataNuevaLineaRegalo)(AddressOf NuevaLineaRegalo, datosRegalo, services)
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarQLineasPromociones(ByVal Doc As DocumentoPedidoVenta, ByVal services As ServiceProvider)
        For Each linea As DataRow In Doc.dtLineas.Select
            If Nz(linea("Regalo"), 0) = 0 Then
                '10. Quitamos la información anterior
                If linea.RowState = DataRowState.Modified Then
                    If Nz(linea("IDPromocionLinea", DataRowVersion.Original), 0) <> 0 Then
                        If (linea("QPedida") <> linea("QPedida", DataRowVersion.Original) OrElse linea("IdArticulo") <> linea("IdArticulo", DataRowVersion.Original)) Then
                            ProcessServer.ExecuteTask(Of Integer)(AddressOf PromocionLinea.ActualizarLineaPromocionQ, linea("IDPromocionLinea", DataRowVersion.Original), services)
                        End If
                    End If
                End If

                '30. Actualizamos en función de la QServida al cerrar la línea.
                If linea.RowState = DataRowState.Modified Then
                    If Nz(linea("IDPromocionLinea"), 0) <> 0 AndAlso Nz(linea("QPedida"), 0) <> Nz(linea("QPedida", DataRowVersion.Original), 0) Then
                        ProcessServer.ExecuteTask(Of Integer)(AddressOf PromocionLinea.ActualizarLineaPromocionQ, linea("IDPromocionLinea"), services)
                    ElseIf Nz(linea("IDPromocionLinea"), 0) <> 0 AndAlso (linea("Estado") = enumpvlEstado.pvlCerrado OrElse linea("Estado", DataRowVersion.Original) = enumpvlEstado.pvlCerrado) Then
                        '//Si ha cambiado el estado a cerrado hay que comprobar si existe alguna diferencia entre la cantidad pedida y la cantidad servida.
                        If linea("Estado", DataRowVersion.Original) <> linea("Estado") AndAlso Nz(linea("QPedida"), 0) <> Nz(linea("QServida"), 0) Then
                            '//Hay que actualizar la cantidad promocionada
                            ProcessServer.ExecuteTask(Of Integer)(AddressOf PromocionLinea.ActualizarLineaPromocionQ, linea("IDPromocionLinea"), services)
                        End If
                    End If
                ElseIf linea.RowState = DataRowState.Added Then
                    If Nz(linea("IDPromocionLinea"), 0) <> 0 Then
                        ProcessServer.ExecuteTask(Of Integer)(AddressOf PromocionLinea.ActualizarLineaPromocionQ, linea("IDPromocionLinea"), services)
                    End If
                End If
            End If
        Next
    End Sub

    Public Class DataNuevaLineaRegalo
        Public Doc As DocumentoPedidoVenta
        Public Row As DataRow
        Public RowPromocion As DataRow
        Public ActualizarPromo As Boolean

        Public Sub New(ByVal Doc As DocumentoPedidoVenta, ByVal Row As DataRow, ByVal RowPromocion As DataRow, Optional ByVal ActualizarPromo As Boolean = True)
            Me.Doc = Doc
            Me.Row = Row
            Me.RowPromocion = RowPromocion
            Me.ActualizarPromo = ActualizarPromo
        End Sub
    End Class

    <Task()> Public Shared Sub NuevaLineaRegalo(ByVal data As DataNuevaLineaRegalo, ByVal services As ServiceProvider)
        If Not IsNothing(data.Row) AndAlso Length(data.Row("IDPromocionLinea")) > 0 Then
            Dim dblQServida As Double
            If data.Row("QPedida") > data.RowPromocion("QMaxPedido") Then
                dblQServida = data.RowPromocion("QMaxPedido")
            Else
                dblQServida = data.Row("QPedida")
            End If

            Dim f As New Filter
            f.Add(New NumberFilterItem("IDPromocionLinea", data.Row("IDPromocionLinea")))
            f.Add(New StringFilterItem("IDArticulo", data.Row("IDArticulo")))

            Dim dtArticuloRegalo As DataTable = AdminData.GetData("vNegPromocionArticulosRegaloPedido", f)
            If Not IsNothing(dtArticuloRegalo) AndAlso dtArticuloRegalo.Rows.Count > 0 Then
                Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
                Dim strAlmacenPred As String = AppParams.Almacen
                Dim PVL As New PedidoVentaLinea
                Dim strTipoLinea As String = ProcessServer.ExecuteTask(Of Object, String)(AddressOf TipoLinea.TipoLineaRegalo, Nothing, services)
                Dim intOrden As Integer = Nz(data.Doc.dtLineas.Compute("MAX(IDOrdenLinea)", Nothing), 0)

                Dim context As New BusinessData(data.Doc.HeaderRow)
                f.Clear()
                f.Add(New NumberFilterItem("IDPedido", data.Row("IDPedido")))
                For Each drArticuloRegalo As DataRow In dtArticuloRegalo.Rows
                    'Nuevo registro
                    Dim drPVL As DataRow = data.Doc.dtLineas.NewRow
                    drPVL("IDLineaPedido") = AdminData.GetAutoNumeric
                    drPVL("Estado") = enumpvlEstado.pvlPedido
                    drPVL("IDTipoLinea") = strTipoLinea
                    drPVL("IDPedido") = data.Doc.HeaderRow("IDPedido")
                    drPVL("IDCentroGestion") = data.Doc.HeaderRow("IDCentroGestion")

                    drPVL = PVL.ApplyBusinessRule("IDArticulo", drArticuloRegalo("IDArticuloRegalo"), drPVL, context)
                    drPVL("FechaEntrega") = IIf(Length(data.Doc.HeaderRow("FechaEntrega")) > 0, data.Doc.HeaderRow("FechaEntrega"), data.Doc.HeaderRow("FechaPedido"))
                    context("Fecha") = drPVL("FechaEntrega")
                    drPVL("Regalo") = True

                    'En el campo Cantidad guardamos la Cantidad indicada con el ArticuloRegalo
                    drPVL("QPedida") = Fix((dblQServida / drArticuloRegalo("QPedida"))) * drArticuloRegalo("QRegalo")
                    If drPVL("QPedida") = 0 Then
                        drPVL("QPedida") = drArticuloRegalo("QRegalo")
                    End If

                    'Se incrementa el IDOrden para cada linea de regalo generada
                    intOrden = intOrden + 1
                    drPVL("IDOrdenLinea") = intOrden

                    drPVL("IDPromocion") = data.Row("IDPromocion")
                    drPVL("IDPromocionLinea") = data.Row("IDPromocionLinea")
                    drPVL = PVL.ApplyBusinessRule("QPedida", drPVL("QPedida"), drPVL, context)
                    data.Doc.dtLineas.Rows.Add(drPVL)
                Next
                'If data.ActualizarPromo Then
                '    'Actualización QPromocionada
                '    Dim PL As New PromocionLinea
                '    Dim drPromocionLinea As DataRow = PL.GetItemRow(data.Row("IDPromocionLinea"))
                '    drPromocionLinea("QPromocionada") = drPromocionLinea("QPromocionada") + dblQServida
                '    BusinessHelper.UpdateTable(drPromocionLinea.Table)
                'End If
            End If
        End If
    End Sub

#End Region

#Region " Cálculos de Importes, Totales y Monedas "

    <Task()> Public Shared Sub ActualizarCambiosMoneda(ByVal Doc As DocumentoPedidoVenta, ByVal services As ServiceProvider)
        If Doc.HeaderRow.RowState = DataRowState.Modified Then
            If Doc.HeaderRow("IDMoneda") <> Nz(Doc.HeaderRow("IDMoneda", DataRowVersion.Original), Nothing) Then
                Dim pvl As New PedidoVentaLinea
                Dim context As New BusinessData
                context("IDMoneda") = Doc.HeaderRow("IDMoneda")
                context("Fecha") = Doc.Fecha
                For Each row As DataRow In Doc.dtLineas.Rows
                    ' pvl.ApplyBusinessRule("IDMoneda", Doc.HeaderRow("IDMoneda"), row, context)
                    pvl.ApplyBusinessRule("Precio", row("Precio"), row, context)
                Next
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CalcularImporteLineasPedido(ByVal Doc As DocumentoPedidoVenta, ByVal services As ServiceProvider)
        'ProcessServer.ExecuteTask(Of DocumentCabLin)(AddressOf ProcesoComercial.RecuperarTiposIVADireccionEnvio, Doc, services)
        '//NO utilizar el CalcularImporteLineas de ProcesoComunes. Hay que pasar la cantidad a la QPedida y viceversa.
        For Each linea As DataRow In Doc.dtLineas.Rows
            If linea.RowState <> DataRowState.Deleted Then
                Dim ILinea As IPropertyAccessor = New DataRowPropertyAccessor(linea)
                ILinea("Cantidad") = linea("QPedida")
                ILinea("IDMoneda") = Doc.IDMoneda
                ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf NegocioGeneral.CalcularPrecioImporte, ILinea, services)
                Dim lineaIProperty As New ValoresAyB(New DataRowPropertyAccessor(linea), Doc.IDMoneda, Doc.CambioA, Doc.CambioB)
                ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, lineaIProperty, services)
            End If
        Next
    End Sub

    <Task()> Public Shared Sub CalcularBasesImponibles(ByVal Doc As DocumentoPedidoVenta, ByVal services As ServiceProvider)
        If Not IsNothing(Doc.dtLineas) Then
            Dim desglose() As DataBaseImponible = ProcessServer.ExecuteTask(Of DocumentoPedidoVenta, DataBaseImponible())(AddressOf ProcesoComunes.DesglosarImporte, Doc, services)
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

                Dim TiposIVA As EntityInfoCache(Of TipoIvaInfo) = services.GetService(Of EntityInfoCache(Of TipoIvaInfo))()
                Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
                Dim ClteInfo As ClienteInfo = Clientes.GetEntity(data.Doc.HeaderRow("IDCliente"))
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
                            If TIVAInfo.SinRepercutir Then factor = TIVAInfo.IVASinRepercutir
                        End If

                        ' Dim Base As Double = BI.BaseImponibleNormal + BI.BaseImponibleEspecial
                        Dim Base As Double = BI.BaseImponible
                        ImporteIVATotal = ImporteIVATotal + Base * factor / 100
                        If ClteInfo.TieneRE Then
                            ImporteRETotal = ImporteRETotal + Base * TIVAInfo.IVARE / 100
                        End If
                    End If
                Next
            End If
        End If
        data.Doc.HeaderRow("BaseImponible") = BaseImponibleTotal
        data.Doc.HeaderRow("ImpIVA") = ImporteIVATotal
        data.Doc.HeaderRow("ImpRE") = ImporteRETotal
        data.Doc.HeaderRow("ImpPedido") = ImporteLineas

        If Nz(data.Doc.HeaderRow("RecFinan"), 0) > 0 Then
            Dim Total As Double = 0
            Total = BaseImponibleTotal + ImporteIVATotal + ImporteRETotal
            data.Doc.HeaderRow("ImpRecFinan") = xRound(Total * data.Doc.HeaderRow("RecFinan") / 100, data.Doc.Moneda.NDecimalesImporte)
        Else : data.Doc.HeaderRow("ImpRecFinan") = 0
        End If

        Dim ValAyB As New ValoresAyB(New DataRowPropertyAccessor(data.Doc.HeaderRow), data.Doc.IDMoneda, data.Doc.CambioA, data.Doc.CambioB)
        ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)
    End Sub

#End Region

#Region " Crear Lineas de Pedido desde un Origen "

    <Task()> Public Shared Sub CrearLineasDesdeOrigen(ByVal Doc As DocumentoPedidoVenta, ByVal services As ServiceProvider)
        Dim lineas As DataTable = Doc.dtLineas
        Dim oPVL As New PedidoVentaLinea
        If lineas Is Nothing Then
            lineas = oPVL.AddNew
            Doc.Add(GetType(PedidoVentaLinea).Name, lineas)
        End If

        Dim f As New Filter
        Dim dtOrigen As DataTable = ProcessServer.ExecuteTask(Of DocumentoPedidoVenta, DataTable)(AddressOf RecuperarDatosOrigen, Doc, services)
        For Each lineaOrigen As DataRow In dtOrigen.Rows
            Dim linea As DataRow = lineas.NewRow
            Dim datDesdeOrigen As New DataDocRowOrigen(Doc, lineaOrigen, linea)
            Dim dblCantidad As Double = ProcessServer.ExecuteTask(Of DataDocRowOrigen, Double)(AddressOf GetCantidad, datDesdeOrigen, services)
            If dblCantidad <> 0 Then
                Dim datDesdeCab As New DataDocRow(Doc, linea)
                ProcessServer.ExecuteTask(Of DataDocRow)(AddressOf AsignarDatosLineaDesdeCabecera, datDesdeCab, services)
                ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarValoresPredeterminadosLinea, linea, services)
                ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarFechaEntregaOrigen, datDesdeOrigen, services)
                ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarArticuloLineaPedido, datDesdeOrigen, services)

                ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarUnidadesMedida, datDesdeOrigen, services)
                Dim context As New BusinessData(Doc.HeaderRow)
                context("Fecha") = context("FechaPedido")
                If Doc.Cabecera.Origen = enumOrigenPedido.Oferta Then
                    oPVL.ApplyBusinessRule("QPedida", datDesdeOrigen.RowDestino("QPedida"), linea, context)
                Else
                    oPVL.ApplyBusinessRule("QPedida", dblCantidad, linea, context)
                End If
                ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarPrecios, datDesdeOrigen, services)
                ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarCentroGestionLineaPedido, datDesdeOrigen, services)
                ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarAlmacenLineaPedido, datDesdeOrigen, services)
                ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarEnlaceConOrigen, datDesdeOrigen, services)
                ProcessServer.ExecuteTask(Of DataDocRowOrigen)(AddressOf AsignarDatosLineasEDI, datDesdeOrigen, services)

                lineas.Rows.Add(linea)
            End If
        Next
    End Sub

    <Task()> Public Shared Sub AsignarValoresPredeterminadosLinea(ByVal data As DataRow, ByVal services As ServiceProvider)
        data("IDLineaPedido") = AdminData.GetAutoNumeric
        data("Estado") = enumpvlEstado.pvlPedido
        '//Campo estandar utilizado en EDI
        data("Eliminar") = False
    End Sub

    <Task()> Public Shared Function GetCantidad(ByVal data As DataDocRowOrigen, ByVal services As ServiceProvider) As Double
        Dim dblCantidad As Double
        Select Case CType(data.Doc, DocumentoPedidoVenta).Cabecera.Origen
            Case enumOrigenPedido.Programa
                Dim PedLin As PedLinPrograma = Nothing
                For i As Integer = 0 To CType(data.Doc.Cabecera, PedCabPrograma).Lineas.Length - 1
                    If data.RowOrigen("IDLineaPrograma") = CType(data.Doc.Cabecera, PedCabPrograma).Lineas(i).IDLineaPrograma Then
                        PedLin = CType(data.Doc.Cabecera, PedCabPrograma).Lineas(i)
                        Exit For
                    End If
                Next

                If Not Double.IsNaN(PedLin.Cantidad) Then
                    dblCantidad = PedLin.Cantidad
                Else
                    dblCantidad = PedLin.QConfirmada
                End If
            Case enumOrigenPedido.Oferta
                dblCantidad = data.RowOrigen("QEstimadaConsumo")
            Case enumOrigenPedido.PedidoCompra
                dblCantidad = data.RowOrigen("QPedida")
            Case enumOrigenPedido.EDI
                dblCantidad = data.RowOrigen("QProgramada")
        End Select
        Return dblCantidad
    End Function

    <Task()> Public Shared Sub AsignarDatosLineaDesdeCabecera(ByVal data As DataDocRow, ByVal services As ServiceProvider)
        If Length(data.Row("IDPedido")) = 0 Then data.Row("IDPedido") = data.Doc.HeaderRow("IDPedido")
        If Length(data.Row("FechaEntrega")) = 0 Then data.Row("FechaEntrega") = data.Doc.HeaderRow("FechaEntrega")
        If Length(data.Row("PedidoCliente")) = 0 Then data.Row("PedidoCliente") = data.Doc.HeaderRow("PedidoCliente")
        If Length(data.Row("Prioridad")) = 0 Then data.Row("Prioridad") = data.Doc.HeaderRow("Prioridad")
        data.Row("DtoProntoPago") = data.Doc.HeaderRow("DtoProntoPago")
    End Sub

    <Task()> Public Shared Sub AsignarArticuloLineaPedido(ByVal data As DataDocRowOrigen, ByVal services As ServiceProvider)
        Dim context As New BusinessData(data.Doc.HeaderRow)
        Dim PVL As New PedidoVentaLinea
        PVL.ApplyBusinessRule("IDArticulo", data.RowOrigen("IDArticulo"), data.RowDestino, context)
        If CType(data.Doc, DocumentoPedidoVenta).Cabecera.Origen = enumOrigenPedido.Programa OrElse CType(data.Doc, DocumentoPedidoVenta).Cabecera.Origen = enumOrigenPedido.Oferta Then
            If Length(data.RowOrigen("IDTipoIva")) > 0 Then data.RowDestino("IDTipoIva") = data.RowOrigen("IDTipoIva")
        End If
        If data.RowOrigen.Table.Columns.Contains("DescDetalle") AndAlso Length(data.RowOrigen("DescDetalle")) > 0 Then data.RowDestino("DescArticulo") = data.RowOrigen("DescDetalle")
        If data.RowOrigen.Table.Columns.Contains("DescRefCliente") AndAlso Length(data.RowOrigen("DescRefCliente")) > 0 Then data.RowDestino("DescRefCliente") = data.RowOrigen("DescRefCliente")
    End Sub

    <Task()> Public Shared Sub AsignarUnidadesMedida(ByVal data As DataDocRowOrigen, ByVal services As ServiceProvider)
        Select Case CType(data.Doc, DocumentoPedidoVenta).Cabecera.Origen
            Case enumOrigenPedido.Oferta
                If Length(data.RowOrigen("IDUDVenta")) > 0 Then
                    'Factor de conversión
                    Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
                    Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.RowOrigen("IDArticulo"))
                    Dim StFactor As New ArticuloUnidadAB.DatosFactorConversion(data.RowOrigen("IDArticulo"), data.RowOrigen("IDUDVenta"), ArtInfo.IDUDVenta)
                    Dim DblFactor As Double = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, StFactor, services)
                    data.RowDestino("IDUDMedida") = data.RowOrigen("IDUDVenta")
                    data.RowDestino("QPedida") = data.RowOrigen("QEstimadaConsumo")
                    data.RowDestino("factor") = DblFactor
                End If
                If Length(data.RowOrigen("UDValoracion")) > 0 Then data.RowDestino("UDValoracion") = data.RowOrigen("UDValoracion")
            Case enumOrigenPedido.EDI
                If IsNumeric(data.RowOrigen("Factor")) AndAlso data.RowOrigen("Factor") <> 0 Then
                    data.RowDestino("Factor") = data.RowOrigen("Factor")
                End If
                If Length(data.RowOrigen("IDUDMedida")) > 0 Then
                    data.RowDestino("IDUDMedida") = data.RowOrigen("IDUDMedida")
                End If
            Case Else
                data.RowDestino("Factor") = data.RowOrigen("Factor")
                data.RowDestino("UdValoracion") = data.RowOrigen("UdValoracion")
                data.RowDestino("IDUDMedida") = data.RowOrigen("IDUDMedida")
                data.RowDestino("IDUDInterna") = data.RowOrigen("IDUDInterna")
        End Select
    End Sub

    <Task()> Public Shared Sub AsignarPrecios(ByVal data As DataDocRowOrigen, ByVal services As ServiceProvider)
        If CType(data.Doc, DocumentoPedidoVenta).Cabecera.Origen = enumOrigenPedido.EDI Then
            data.RowDestino("Dto1") = 0
            data.RowDestino("Dto2") = 0
            data.RowDestino("Dto3") = 0
            If data.RowOrigen("Precio") = 0 Then data.RowOrigen("Precio") = data.RowDestino("Precio")
        Else
            data.RowDestino("Dto1") = data.RowOrigen("Dto1")
            data.RowDestino("Dto2") = data.RowOrigen("Dto2")
            data.RowDestino("Dto3") = data.RowOrigen("Dto3")
        End If
        Dim context As New BusinessData(data.Doc.HeaderRow)
        Dim PVL As New PedidoVentaLinea
        If data.RowOrigen.Table.Columns.Contains("Dto") Then data.RowDestino("Dto") = data.RowOrigen("Dto")


        PVL.ApplyBusinessRule("Precio", data.RowOrigen("Precio"), data.RowDestino, context)


        Dim dataTarifa As New DataCalculoTarifaComercial(data.RowOrigen("IDArticulo"), data.RowDestino("QInterna"), data.RowDestino("FechaEntrega"))
        ProcessServer.ExecuteTask(Of DataCalculoTarifaComercial)(AddressOf ProcesoComercial.TarifaCosteArticulo, dataTarifa, services)
        If Not dataTarifa.DatosTarifa Is Nothing Then
            data.RowDestino("PrecioCosteA") = dataTarifa.DatosTarifa.PrecioCosteA
        End If
        If data.RowOrigen.Table.Columns.Contains("SeguimientoTarifa") Then data.RowDestino("SeguimientoTarifa") = data.RowOrigen("SeguimientoTarifa")

    End Sub

    <Task()> Public Shared Sub AsignarCentroGestionLineaPedido(ByVal data As DataDocRowOrigen, ByVal services As ServiceProvider)
        If CType(data.Doc, DocumentoPedidoVenta).Cabecera.Origen = enumOrigenPedido.EDI OrElse Not data.RowOrigen.Table.Columns.Contains("IDCentroGestion") OrElse Length(data.RowOrigen("IDCentroGestion")) = 0 Then
            data.RowDestino("IDCentroGestion") = data.Doc.HeaderRow("IDCentroGestion")
        Else
            data.RowDestino("IDCentroGestion") = data.RowOrigen("IDCentroGestion")
        End If
    End Sub

    <Task()> Public Shared Sub AsignarAlmacenLineaPedido(ByVal data As DataDocRowOrigen, ByVal services As ServiceProvider)
        If data.RowOrigen.Table.Columns.Contains("IDAlmacen") AndAlso CType(data.Doc, DocumentoPedidoVenta).Cabecera.Origen <> enumOrigenPedido.PedidoCompra Then
            If Length(data.RowOrigen("IDAlmacen")) > 0 Then
                data.RowDestino("IDAlmacen") = data.RowOrigen("IDAlmacen")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarEnlaceConOrigen(ByVal data As DataDocRowOrigen, ByVal services As ServiceProvider)
        Select Case CType(data.Doc, DocumentoPedidoVenta).Cabecera.Origen
            Case enumOrigenPedido.Oferta
                data.RowDestino("IDLineaOfertaDetalle") = data.RowOrigen("IDLineaOfertaDetalle")
                data.RowDestino("Revision") = data.RowOrigen("Revision")
                data.RowDestino("IDEvaluador") = data.RowOrigen("IDEvaluador")
                data.RowDestino("IDOrdenLinea") = data.RowOrigen("IDOrdenLinea")
            Case enumOrigenPedido.Programa
                data.RowDestino("IDLineaPrograma") = data.RowOrigen("IDLineaPrograma")
                data.RowDestino("IDPrograma") = data.RowOrigen("IDPrograma")
                data.RowDestino("PedidoCliente") = data.RowOrigen("ProgramaCliente")
                data.RowDestino("IDOrdenLinea") = data.RowOrigen("IDOrdenLinea")
                data.RowDestino("Muelle") = data.RowOrigen("Muelle")
                data.RowDestino("PuntoDescarga") = data.RowOrigen("PuntoDescarga")
                If data.RowOrigen.Table.Columns.Contains("TextoLin") Then data.RowDestino("Texto") = data.RowOrigen("TextoLin")
            Case enumOrigenPedido.PedidoCompra
                '//Completamos la trazabilidad
                Dim DocumentosPC As DocumentInfoCache(Of DocumentoPedidoCompra) = services.GetService(Of DocumentInfoCache(Of DocumentoPedidoCompra))()
                Dim DocPC As DocumentoPedidoCompra = DocumentosPC.GetDocument(data.RowOrigen("IDPedido"))
                If DocPC.dtTrazabilidad.Rows.Count > 0 Then
                    Dim f As New Filter
                    f.Add(New NumberFilterItem("IDPCPrincipal", data.RowOrigen("IDPedido")))
                    f.Add(New NumberFilterItem("IDLineaPCPrincipal", data.RowOrigen("IDLineaPedido")))
                    Dim WhereTrazaLineaPedido As String = f.Compose(New AdoFilterComposer)
                    Dim TrazaLineaPedido() As DataRow = DocPC.dtTrazabilidad.Select(WhereTrazaLineaPedido)
                    If TrazaLineaPedido.Length > 0 Then
                        TrazaLineaPedido(0)("IDPVSecundaria") = data.Doc.HeaderRow("IDPedido")
                        TrazaLineaPedido(0)("NPVSecundaria") = data.Doc.HeaderRow("NPedido")
                        TrazaLineaPedido(0)("IDLineaPVSecundaria") = data.RowDestino("IDLineaPedido")
                        Dim datBBDD As DataBasesDatosMultiempresa = services.GetService(Of DataBasesDatosMultiempresa)()
                        TrazaLineaPedido(0)("IDBDSecundaria") = datBBDD.IDBaseDatosSecundaria
                    End If
                    'DocPC.SetData()
                End If
        End Select
    End Sub

    <Task()> Public Shared Sub AsignarDatosLineasEDI(ByVal data As DataDocRowOrigen, ByVal services As ServiceProvider)
        Select Case CType(data.Doc, DocumentoPedidoVenta).Cabecera.Origen
            Case enumOrigenPedido.EDI
                data.RowDestino("Muelle") = data.RowOrigen("Muelle")
                data.RowDestino("PuntoDescarga") = data.RowOrigen("PuntoDescarga")
        End Select
    End Sub

    <Task()> Public Shared Function RecuperarDatosOrigen(ByVal Doc As DocumentoPedidoVenta, ByVal services As ServiceProvider) As DataTable
        Dim oFltr As New Filter
        Dim ViewName As String
        Select Case Doc.Cabecera.Origen
            Case enumOrigenPedido.Programa
                Dim ids(CType(Doc.Cabecera, PedCabPrograma).Lineas.Length - 1) As Object
                For i As Integer = 0 To ids.Length - 1
                    ids(i) = CType(Doc.Cabecera, PedCabPrograma).Lineas(i).IDLineaPrograma
                Next
                ViewName = "vfrmConfirmacionPrograma"
                oFltr.Add(New InListFilterItem("IDLineaPrograma", ids, FilterType.Numeric))
            Case enumOrigenPedido.Oferta
                Dim ids(CType(Doc.Cabecera, PedCabOfertaComercial).LineaOfertaDetalle.Length - 1) As Object
                For i As Integer = 0 To ids.Length - 1
                    ids(i) = CType(Doc.Cabecera, PedCabOfertaComercial).LineaOfertaDetalle(i)
                Next
                oFltr.Add(New InListFilterItem("IDLineaOfertaDetalle", ids, FilterType.Numeric))
                ViewName = "vfrmOfertaComercialTratamiento"
            Case enumOrigenPedido.PedidoCompra
                Dim Cab As PedCabVentaPedidoCompra = CType(Doc.Cabecera, PedCabVentaPedidoCompra)

                Return Cab.DatosOrigen
            Case enumOrigenPedido.EDI
                Dim Cab As PedCabEDI = CType(Doc.Cabecera, PedCabEDI)
                oFltr.Add(New NumberFilterItem("IDPedidoEDI", Cab.IDPedidoEDI))
                oFltr.Add(New NumberFilterItem("Equivalencia", enumEquivalencia.eqvPedido))
                oFltr.Add(New BooleanFilterItem("Procesar", True))
                ViewName = "VCTLCIPedidosLineasEDI"
        End Select

        Return New BE.DataEngine().Filter(ViewName, oFltr)

    End Function

    <Task()> Public Shared Sub LineasDeRegalo(ByVal Doc As DocumentoPedidoVenta, ByVal services As ServiceProvider)

        If CType(Doc, DocumentoPedidoVenta).Cabecera.Origen = enumOrigenPedido.Oferta Then Exit Sub

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

    <Task()> Public Shared Sub ValidacionesContabilidad(ByVal data As DataRow, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf NegocioGeneral.ValidarEjercicioContablePedido, data, services)
    End Sub

    <Task()> Public Shared Sub ValidarDocumento(ByVal Doc As DocumentoPedidoVenta, ByVal services As ServiceProvider)
        Dim PVC As New PedidoVentaCabecera
        PVC.Validate(Doc.HeaderRow.Table)

        Dim PVL As New PedidoVentaLinea
        PVL.Validate(Doc.dtLineas)

        Dim PVR As New PedidoVentaRepresentante
        PVR.Validate(Doc.dtVentaRepresentante)
    End Sub

#End Region

#Region " Grabar Documento "

    <Task()> Public Shared Sub GrabarDocumento(ByVal Doc As DocumentoPedidoVenta, ByVal services As ServiceProvider)
        'ProcessServer.ExecuteTask(Of Object)(AddressOf Business.General.Comunes.BeginTransaction, Nothing, services)
        AdminData.BeginTx()
        ProcessServer.ExecuteTask(Of DocumentoPedidoVenta)(AddressOf Business.General.Comunes.UpdateDocument, Doc, services)
        ProcessServer.ExecuteTask(Of DocumentoPedidoVenta)(AddressOf ProcesoPedidoVenta.ActualizarPrograma, Doc, services)
        ProcessServer.ExecuteTask(Of DocumentoPedidoVenta)(AddressOf ProcesoPedidoVenta.ActualizarOfertaComercial, Doc, services)
        'ProcessServer.ExecuteTask(Of Boolean)(AddressOf Business.General.Comunes.CommitTransaction, True, services)
        AdminData.CommitTx(True)
    End Sub

#End Region

    '#Region " Actualización de Programas "

    '    <Task()> Public Shared Sub ActualizarPrograma(ByVal Doc As DocumentoPedidoVenta, ByVal services As ServiceProvider)
    '        If Doc Is Nothing OrElse Doc.dtLineas Is Nothing Then Exit Sub
    '        'If Doc.Cabecera.Origen <> enumOrigenPedido.Programa Then Exit Sub
    '        Dim PL As New ProgramaLinea
    '        For Each lineaPedido As DataRow In Doc.dtLineas.Rows
    '            Dim PedLin As PedLinPrograma
    '            If Not CType(Doc.Cabecera, PedCabPrograma) Is Nothing Then
    '                For Each linea As PedLinPrograma In CType(Doc.Cabecera, PedCabPrograma).Lineas
    '                    If linea.IDLineaPrograma = lineaPedido("IDLineaPrograma") Then
    '                        PedLin = linea
    '                    End If
    '                Next
    '            End If
    '            Dim Programa As DataTable = PL.SelOnPrimaryKey(lineaPedido("IDLineaPrograma"))
    '            If Not Programa Is Nothing AndAlso Programa.Rows.Count > 0 Then
    '                If lineaPedido.RowState = DataRowState.Deleted Then
    '                    Dim DblConfir As Double = Programa.Rows(0)("QConfirmada") - lineaPedido("QPedida")
    '                    Programa.Rows(0)("QConfirmada") = DblConfir
    '                    Programa.Rows(0)("Confirmada") = IIf(DblConfir > 0, True, False)
    '                    Programa.Rows(0)("FechaConfirmacion") = IIf(DblConfir > 0, Programa.Rows(0)("FechaConfirmacion"), DBNull.Value)
    '                Else
    '                    Dim dblQModificada As Integer
    '                    If lineaPedido.RowState = DataRowState.Modified Then
    '                        dblQModificada = lineaPedido("QPedida", DataRowVersion.Original)
    '                    End If
    '                    Programa.Rows(0)("QConfirmada") = Nz(Programa.Rows(0)("QConfirmada"), 0) + (lineaPedido("QPedida") - dblQModificada)
    '                    Programa.Rows(0)("Confirmada") = True

    '                    If Not PedLin Is Nothing AndAlso Length(PedLin.FechaConfirmacion) > 0 AndAlso PedLin.FechaConfirmacion <> cnMinDate Then
    '                        Programa.Rows(0)("FechaConfirmacion") = PedLin.FechaConfirmacion
    '                    Else
    '                        Programa.Rows(0)("FechaConfirmacion") = Today
    '                    End If

    '                    Programa.Rows(0)("FechaEntrega") = lineaPedido("FechaEntrega")
    '                End If
    '            End If
    '            AdminData.SetData(Programa)
    '        Next
    '    End Sub

    '#End Region
#Region " Actualización de Programas  "

    <Task()> Public Shared Sub ActualizarPrograma(ByVal Doc As DocumentoPedidoVenta, ByVal services As ServiceProvider)
        If Doc Is Nothing Then Exit Sub
        For Each lineapedido As DataRow In Doc.dtLineas.Select(Nothing, "IDPrograma,IDLineaPrograma", DataViewRowState.CurrentRows)
            Dim datosActProg As New DataActualizarProgramaLinea(lineapedido)
            ProcessServer.ExecuteTask(Of DataActualizarProgramaLinea)(AddressOf ProcesoPedidoVenta.ActualizarProgramaLinea, datosActProg, services)
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
        Dim pl As New ProgramaLinea
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

    End Sub

#End Region
#Region " Actualización de Ofertas Comerciales "

    <Task()> Public Shared Sub ActualizarOfertaComercial(ByVal Doc As DocumentoPedidoVenta, ByVal services As ServiceProvider)
        If Doc Is Nothing OrElse Doc.dtLineas Is Nothing Then Exit Sub
        For Each lineaPedido As DataRow In Doc.dtLineas.Select()
            If Length(lineaPedido("IDLineaOfertaDetalle")) > 0 Then
                ProcessServer.ExecuteTask(Of DataRow)(AddressOf ActualizarOfertaDetalle, lineaPedido, services)
                Dim datActOferta As New DataDocRow(Doc, lineaPedido)
                ProcessServer.ExecuteTask(Of DataDocRow)(AddressOf ActualizarClienteOferta, datActOferta, services)
            End If
        Next
    End Sub

    <Task()> Public Shared Sub ActualizarOfertaDetalle(ByVal lineaPedido As DataRow, ByVal services As ServiceProvider)
        If Length(lineaPedido("IDLineaOfertaDetalle")) > 0 Then
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf ActualizarEstadoVenta, lineaPedido, services)
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarEstadoVenta(ByVal lineaPedido As DataRow, ByVal services As ServiceProvider)
        If Length(lineaPedido("IDLineaOfertaDetalle")) > 0 Then
            Dim Entidad As BusinessHelper = BusinessHelper.CreateBusinessObject("OfertaComercialDetalle")
            Dim dtDetalle As DataTable = Entidad.SelOnPrimaryKey(lineaPedido("IDLineaOfertaDetalle"))
            If dtDetalle.Rows.Count > 0 Then
                dtDetalle.Rows(0)("EstadoVenta") = enumocdEstadoCompraVenta.ecvLanzado
            End If
            BusinessHelper.UpdateTable(dtDetalle)
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarClienteOferta(ByVal data As DataDocRow, ByVal services As ServiceProvider)
        If Not CType(data.Doc, DocumentoPedidoVenta).Cabecera Is Nothing AndAlso CType(data.Doc, DocumentoPedidoVenta).Cabecera.Origen = enumOrigenPedido.Oferta Then
            Dim Entidad As BusinessHelper = BusinessHelper.CreateBusinessObject("OfertaComercialCabecera")
            Dim dtOferta As DataTable = Entidad.SelOnPrimaryKey(CType(CType(data.Doc, DocumentoPedidoVenta).Cabecera, PedCabOfertaComercial).IDOferta) 'data.Row("IDOfertaComercial"))
            If Length(dtOferta.Rows(0)("IDEmpresa")) > 0 Then
                If Length(dtOferta.Rows(0)("IDCliente")) > 0 AndAlso dtOferta.Rows(0)("IDCliente") <> data.Doc.HeaderRow("IDCliente") Then
                    dtOferta.Rows(0)("IDCliente") = data.Doc.HeaderRow("IDCliente")

                    Dim dataDir As New ClienteDireccion.DataDirecEnvio(dtOferta.Rows(0)("IDCliente"), enumcdTipoDireccion.cdDireccionEnvio)
                    Dim dtDireccion As DataTable = ProcessServer.ExecuteTask(Of ClienteDireccion.DataDirecEnvio, DataTable)(AddressOf ClienteDireccion.ObtenerDireccionEnvio, dataDir, services)
                    If Not dtDireccion Is Nothing AndAlso dtDireccion.Rows.Count > 0 Then
                        dtOferta.Rows(0)("IDDireccionCliente") = dtDireccion.Rows(0)("IDDireccion")
                    End If
                End If
                BusinessHelper.UpdateTable(dtOferta)
            End If
        End If
    End Sub
#End Region


#Region " Crear Cabecera de Pedido "

    <Serializable()> _
    Public Class DataGeneraPedidoVentaAbierto
        Public IDCliente As String
        Public IDAlmacen As String
        Public IDMoneda As String
        Public IDDireccionEnvio As Integer?
        Public IDCentroGestion As String
        Public PedidoCliente As String

        Public EDI As Boolean
        Public FechaPedido As Date


        Public Sub New(ByVal IDCliente As String, ByVal IDAlmacen As String, ByVal IDMoneda As String, ByVal IDCentroGestion As String, ByVal PedidoCliente As String, ByVal EDI As Boolean, Optional ByVal IDDireccionEnvio As Integer = 0)
            Me.IDCliente = IDCliente
            If Length(IDAlmacen) > 0 Then Me.IDAlmacen = IDAlmacen
            If Length(IDMoneda) > 0 Then Me.IDMoneda = IDMoneda
            Me.IDCentroGestion = IDCentroGestion
            If Length(PedidoCliente) > 0 Then Me.PedidoCliente = PedidoCliente
            Me.EDI = EDI
            If Length(IDDireccionEnvio) > 0 Then Me.IDDireccionEnvio = IDDireccionEnvio

            Me.FechaPedido = Today
        End Sub
    End Class

    <Task()> Public Shared Function GeneraPedidoVentaAbierto(ByVal data As DataGeneraPedidoVentaAbierto, ByVal services As ServiceProvider) As CreateElement
        Dim PVC As New PedidoVentaCabecera()
        Dim dtPVC As DataTable = PVC.AddNewForm()
        If Not IsNothing(dtPVC) AndAlso dtPVC.Rows.Count > 0 Then
            Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
            Dim drRowPVC As DataRow = dtPVC.Rows(0)
            drRowPVC("IDCliente") = data.IDCliente
            If Length(data.IDAlmacen) > 0 Then drRowPVC("IDAlmacen") = data.IDAlmacen
            If Length(data.IDMoneda) > 0 Then
                PVC.ApplyBusinessRule("IDMoneda", data.IDMoneda, drRowPVC, Nothing)
            End If
            If Length(data.IDDireccionEnvio) > 0 Then
                drRowPVC("IDDireccionEnvio") = data.IDDireccionEnvio
            Else
                '//Recuperamos la Dirección del Cliente.
                Dim dataDir As New ClienteDireccion.DataDirecEnvio(data.IDCliente, enumcdTipoDireccion.cdDireccionEnvio)
                Dim dtDireccion As DataTable = ProcessServer.ExecuteTask(Of ClienteDireccion.DataDirecEnvio, DataTable)(AddressOf ClienteDireccion.ObtenerDireccionEnvio, dataDir, services)
                If Not dtDireccion Is Nothing AndAlso dtDireccion.Rows.Count > 0 Then
                    drRowPVC("IDDireccionEnvio") = dtDireccion.Rows(0)("IDDireccion")
                End If
            End If

            drRowPVC("IDCentroGestion") = data.IDCentroGestion
            If Length(data.PedidoCliente) > 0 Then drRowPVC("PedidoCliente") = data.PedidoCliente
            drRowPVC("EDI") = data.EDI
            drRowPVC("FechaPedido") = data.FechaPedido

            Dim ClteInfo As ClienteInfo = Clientes.GetEntity(drRowPVC("IDCliente"))
            drRowPVC("IDFormaPago") = ClteInfo.FormaPago
            drRowPVC("IDCondicionPago") = ClteInfo.CondicionPago
            drRowPVC("IDFormaEnvio") = ClteInfo.FormaEnvio
            drRowPVC("IDCondicionEnvio") = ClteInfo.CondicionEnvio

            'Dim StDatos As New Contador.DatosCounterValue(drRowPVC("IDContador"), New PedidoVentaCabecera, "NPedido", "FechaPedido", drRowPVC("FechaPedido"))
            'StDatos.IDEjercicio = drRowPVC("IDEjercicio") & String.Empty
            'drRowPVC("NPedido") = ProcessServer.ExecuteTask(Of Contador.DatosCounterValue, String)(AddressOf Contador.CounterValue, StDatos, services)

            dtPVC = PVC.Update(dtPVC)
            If Not dtPVC Is Nothing AndAlso dtPVC.Rows.Count > 0 Then
                Dim e As New CreateElement
                e.IDElement = dtPVC.Rows(0)("IDPedido")
                e.NElement = dtPVC.Rows(0)("NPedido")
                Return e
            End If
        End If
    End Function

#End Region


End Class


Public Class LineasPedidoEliminadas
    Public IDLineas As Hashtable

    Public Sub New()
        IDLineas = New Hashtable
    End Sub
End Class