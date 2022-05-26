Public Class PrcCrearPedidoVentaEnBDSecundaria
    Inherits Process(Of DataPrcCrearPedidoVentaEnBDSecundaria, DataResultadoMultiempresaPC)

    Public Overrides Sub RegisterTasks()
        Me.AddTask(Of DataPrcCrearPedidoVentaEnBDSecundaria)(AddressOf ValidacionesPrevias)
        Me.AddTask(Of DataPrcCrearPedidoVentaEnBDSecundaria)(AddressOf PrepararInformacionProceso)
        Me.AddTask(Of DataPrcCrearPedidoVentaEnBDSecundaria, PedCabVentaPedidoCompra())(AddressOf AgruparPedidoCompra)
        Me.AddTask(Of PedCabVentaPedidoCompra())(AddressOf ProcesoComunes.EstablecerEmpresaSecundaria)
        Me.AddForEachTask(Of PrcCrearPedidoVenta)(OnExceptionBehaviour.NextLoop)
        Me.AddTask(Of Object)(AddressOf ProcesoComunes.EstablecerEmpresaPrincipal)
        Me.AddTask(Of Object, DataResultadoMultiempresaPC)(AddressOf GetResultadoMultiempresa)
    End Sub

    <Task()> Public Shared Sub ValidacionesPrevias(ByVal data As DataPrcCrearPedidoVentaEnBDSecundaria, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataPrcCrearPedidoVentaEnBDSecundaria)(AddressOf ValidarPedidoVentaGenerado, data, services)
        ProcessServer.ExecuteTask(Of DataPrcCrearPedidoVentaEnBDSecundaria)(AddressOf ValidarProveedorEmpresaGrupo, data, services)

        '//Validar Proveedor Asociado
        Dim datValProv As New ProcesoPedidoCompra.DataValidarProveedorAsociado(data.Doc.HeaderRow("IDProveedor"))
        ProcessServer.ExecuteTask(Of ProcesoPedidoCompra.DataValidarProveedorAsociado)(AddressOf ProcesoPedidoCompra.ValidarProveedorAsociado, datValProv, services)
        data.IDCliente = datValProv.IDCliente

        ProcessServer.ExecuteTask(Of DataPrcCrearPedidoVentaEnBDSecundaria)(AddressOf ValidarEmpresasObligatorias, data, services)
    End Sub

    <Task()> Public Shared Sub ValidarPedidoVentaGenerado(ByVal data As DataPrcCrearPedidoVentaEnBDSecundaria, ByVal services As ServiceProvider)
        Dim grp As New GRPPedidoVentaCompraLinea
        Dim control As DataTable = grp.TrazaPCPrincipal(data.Doc.HeaderRow("IDPedido"))
        If control.Rows.Count > 0 AndAlso Not control.Rows(0).IsNull("IDPVSecundaria") Then
            ApplicationService.GenerateError("El pedido de compra Nº {0} ya ha generado un pedido de venta. Consultar la ficha de seguimiento de Compra/Venta.", Quoted(control.Rows(0)("NPCPrincipal")))
        End If
    End Sub

    <Task()> Public Shared Sub ValidarProveedorEmpresaGrupo(ByVal data As DataPrcCrearPedidoVentaEnBDSecundaria, ByVal services As ServiceProvider)
        Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
        Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(data.Doc.HeaderRow("IDProveedor"))

        If Not ProvInfo.EmpresaGrupo Then
            ApplicationService.GenerateError("El proveedor {0} no es empresa de grupo.", Quoted(data.Doc.HeaderRow("IDProveedor")))
        ElseIf Length(ProvInfo.BaseDatos) = 0 Then
            ApplicationService.GenerateError("El proveedor {0} no tiene asignado una base de datos válida.", Quoted(data.Doc.HeaderRow("IDProveedor")))
        Else
            data.IDBaseDatosSecundaria = ProvInfo.BaseDatos
        End If
    End Sub

    <Task()> Public Shared Sub ValidarEmpresasObligatorias(ByVal data As DataPrcCrearPedidoVentaEnBDSecundaria, ByVal services As ServiceProvider)
        If Length(data.IDBaseDatosPrincipal) = 0 OrElse Length(data.IDBaseDatosSecundaria) = 0 Then
            ApplicationService.GenerateError("Debe indicarse una Empresa Principal y un Empresa Secundaria. Revise sus datos.")
        End If
    End Sub

    <Task()> Public Shared Sub PrepararInformacionProceso(ByVal data As DataPrcCrearPedidoVentaEnBDSecundaria, ByVal services As ServiceProvider)
        '//Ponemos el contador a nothing, para que al venir de la creación de PC no traiga el contador, que se ha elegido para el
        '// Pedido de Compra y pueda coger el del Pedido de Venta.
        Dim Info As ProcessInfo = services.GetService(Of ProcessInfo)()
        Info.IDContador = Nothing

        services.RegisterService(New ProcessInfoPVBBDDSec(data.IDCliente))

        Dim dataBD As New DataBasesDatosMultiempresa(data.IDBaseDatosPrincipal, data.IDBaseDatosSecundaria)
        ProcessServer.ExecuteTask(Of DataBasesDatosMultiempresa)(AddressOf ProcesoComunes.GetDescripcionBasesDatosMultiempresa, dataBD, services)
        services.RegisterService(dataBD)

        Dim DocumentosPC As DocumentInfoCache(Of DocumentoPedidoCompra) = services.GetService(Of DocumentInfoCache(Of DocumentoPedidoCompra))()
        DocumentosPC.AddDocument(DocumentosPC.GetKey(data.Doc.HeaderRow("IDPedido")), data.Doc)
        ' ProcessServer.ExecuteTask(Of DataPrcCrearPedidoVentaEnBDSecundaria)(AddressOf PrepararInformacionPedidoCompraOrigen, data, services)
    End Sub

    <Task()> Public Shared Sub PrepararInformacionPedidoCompraOrigen(ByVal data As DataPrcCrearPedidoVentaEnBDSecundaria, ByVal services As ServiceProvider)
        Dim PedidosCompra As DataResultadoMultiempresaPC = services.GetService(Of DataResultadoMultiempresaPC)()
        If Not PedidosCompra.Items.ContainsKey(data.Doc.HeaderRow("IDPedido")) Then
            Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
            Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(data.Doc.HeaderRow("IDProveedor"))

            Dim PCInfoNew As New GeneracionPedidosCompraInfo
            PCInfoNew.Proveedor = String.Concat(ProvInfo.IDProveedor, " - ", ProvInfo.DescProveedor)
            PCInfoNew.EmpresaGrupo = ProvInfo.EmpresaGrupo
            PCInfoNew.EntregaProveedor = data.Doc.HeaderRow("EntregaProveedor")
            PCInfoNew.IDPedidoCompra = data.Doc.HeaderRow("IDPedido")
            PCInfoNew.NPedidoCompra = data.Doc.HeaderRow("NPedido")
            PCInfoNew.BaseDatos1 = AdminData.GetSessionInfo.DataBase.DataBaseDescription
            PedidosCompra.Add(PCInfoNew)
        End If
    End Sub

    <Task()> Public Shared Function AgruparPedidoCompra(ByVal data As DataPrcCrearPedidoVentaEnBDSecundaria, ByVal services As ServiceProvider) As PedCabVentaPedidoCompra()
        Dim ProcInfo As ProcessInfoPVBBDDSec = services.GetService(Of ProcessInfoPVBBDDSec)()
        Dim IDCliente As String = ProcInfo.IDCliente

        Dim IDLineasPedidoCompra(-1) As Object
        For Each traza As DataRow In data.Doc.dtLineas.Rows
            ReDim Preserve IDLineasPedidoCompra(IDLineasPedidoCompra.Length)
            IDLineasPedidoCompra(IDLineasPedidoCompra.Length - 1) = traza("IDLineaPedido")
        Next

        Dim IDDireccionEntrega As Integer
        If data.Doc.dtTrazabilidad.Rows.Count = 0 Then
            '//Cuando no tenemos líneas de trazabilidad, es que estamos introduciendo un Pedido de Compra manual. No viene de un Pedido de Venta.
            For Each linea As DataRow In data.Doc.dtLineas.Rows
                Dim traza As DataRow = data.Doc.dtTrazabilidad.NewRow
                traza("IDPVLinea") = AdminData.GetAutoNumeric()
                traza("IDPCPrincipal") = data.Doc.HeaderRow("IDPedido")
                traza("NPCPrincipal") = data.Doc.HeaderRow("NPedido")
                traza("IDLineaPCPrincipal") = linea("IDLineaPedido")
                traza("EntregaProveedor") = False
                traza("IDBDPrincipal") = data.IDBaseDatosPrincipal
                data.Doc.dtTrazabilidad.Rows.Add(traza)

                'ReDim Preserve IDLineasPedidoCompra(IDLineasPedidoCompra.Length)
                'IDLineasPedidoCompra(IDLineasPedidoCompra.Length - 1) = traza("IDLineaPCPrincipal")
            Next
        ElseIf data.Doc.dtTrazabilidad.Rows.Count > 0 Then
            For Each linea As DataRow In data.Doc.dtLineas.Rows
                Dim DrFind() As DataRow = data.Doc.dtTrazabilidad.Select("IDPCPrincipal = " & data.Doc.HeaderRow("IDPedido") & " AND IDLineaPCPrincipal = " & linea("IDLineaPedido"))
                If DrFind.Length <= 0 Then
                    Dim traza As DataRow = data.Doc.dtTrazabilidad.NewRow
                    traza("IDPVLinea") = AdminData.GetAutoNumeric()
                    traza("IDPCPrincipal") = data.Doc.HeaderRow("IDPedido")
                    traza("NPCPrincipal") = data.Doc.HeaderRow("NPedido")
                    traza("IDLineaPCPrincipal") = linea("IDLineaPedido")
                    traza("EntregaProveedor") = False
                    traza("IDBDPrincipal") = data.IDBaseDatosPrincipal
                    data.Doc.dtTrazabilidad.Rows.Add(traza)
                End If
            Next
            If Length(data.Doc.dtTrazabilidad.Rows(0)("IDPVPrincipal")) Then
                Dim dtPedidoVenta As DataTable = New PedidoVentaCabecera().SelOnPrimaryKey(data.Doc.dtTrazabilidad.Rows(0)("IDPVPrincipal"))
                If Not dtPedidoVenta Is Nothing AndAlso dtPedidoVenta.Rows.Count > 0 Then
                    IDDireccionEntrega = dtPedidoVenta.Rows(0)("IDDireccionEnvio")
                End If
            End If
        End If

        Dim f As New Filter
        f.Add(New InListFilterItem("IDLineaPedido", IDLineasPedidoCompra, FilterType.Numeric))
        Dim dtDatosOrigen As DataTable = New BE.DataEngine().Filter("vDisponibilidadPedidoCompraGrupo", f)
        Dim dtLineasPC As DataTable = New PedidoCompraLinea().Filter(f, , "IDArticulo, Precio, QPedida, Dto1, Dto2, Dto3, Dto, IDLineaPedido, IDPedido,IDUDMedida, FechaEntrega,Factor,UdValoracion,IDUDInterna")

        Dim oGrprUser As New GroupUserPVPedidosCompra
        Dim grpPC(0) As DataColumn
        grpPC(0) = dtDatosOrigen.Columns("IDPedido")

        Dim groupers(0) As GroupHelper
        groupers(0) = New GroupHelper(grpPC, oGrprUser)

        For Each dr As DataRow In dtDatosOrigen.Rows
            groupers(0).Group(dr)
        Next

        Dim fPedido As New Filter
        For Each ped As PedCabVentaPedidoCompra In oGrprUser.Pedidos
            ped.IDCliente = IDCliente
            If ped.EntregaProveedor Then ped.IDDireccionEnvio = IDDireccionEntrega
            fPedido.Clear()
            fPedido.Add(New NumberFilterItem("IDPedido", ped.IDPedido))
            Dim WherePedido As String = fPedido.Compose(New AdoFilterComposer)
            Dim adr() As DataRow = dtLineasPC.Select(WherePedido)
            If Not adr Is Nothing AndAlso adr.Length > 0 Then

                If ped.DatosOrigen Is Nothing Then ped.DatosOrigen = dtLineasPC.Clone
                For i As Integer = 0 To adr.Length - 1
                    ped.DatosOrigen.ImportRow(adr(i))
                Next
            End If
        Next
        Return oGrprUser.Pedidos
    End Function

    <Task()> Public Shared Function GetResultadoMultiempresa(ByVal data As Object, ByVal services As ServiceProvider) As DataResultadoMultiempresaPC
        Return services.GetService(Of DataResultadoMultiempresaPC)()
    End Function

    Protected Overrides Function OnException(ByVal exceptionArgs As Engine.BE.BusinessProcesses.ProcessExceptionArgs) As Engine.BE.BusinessProcesses.OnExceptionBehaviour
        '//Si se produce un error, volvemos a la BBDD Origen.
        AdminData.RollBackTx(True)

        Dim BDInfo As DataBasesDatosMultiempresa = exceptionArgs.Services.GetService(Of DataBasesDatosMultiempresa)()
        If Length(BDInfo.IDBaseDatosPrincipal) > 0 AndAlso BDInfo.IDBaseDatosPrincipal <> AdminData.GetSessionInfo.DataBase.DataBaseID Then
            AdminData.SetCurrentConnection(BDInfo.IDBaseDatosPrincipal)
        End If
        Return MyBase.OnException(exceptionArgs)
    End Function

End Class
