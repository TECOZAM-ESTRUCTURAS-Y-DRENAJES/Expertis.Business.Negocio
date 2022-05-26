Imports Solmicro.Expertis.Business.Negocio.NegocioGeneral

Public Class _PedidoCompraCabecera
    Public Const IDPedido As String = "IDPedido"
    Public Const IDContador As String = "IDContador"
    Public Const FechaPedido As String = "FechaPedido"
    Public Const IDProveedor As String = "IDProveedor"
    Public Const IDDireccion As String = "IDDireccion"
    Public Const IDOperario As String = "IDOperario"
    Public Const IDCentroGestion As String = "IDCentroGestion"
    Public Const IDDiaPago As String = "IDDiaPago"
    Public Const IDFormaPago As String = "IDFormaPago"
    Public Const IDCondicionPago As String = "IDCondicionPago"
    Public Const IDFormaEnvio As String = "IDFormaEnvio"
    Public Const IDCondicionEnvio As String = "IDCondicionEnvio"
    Public Const IDAlmacen As String = "IDAlmacen"
    Public Const IDMoneda As String = "IDMoneda"
    Public Const CambioA As String = "CambioA"
    Public Const CambioB As String = "CambioB"
    Public Const Dto As String = "Dto"
    Public Const ImpPedido As String = "ImpPedido"
    Public Const ImpPedidoA As String = "ImpPedidoA"
    Public Const ImpPedidoB As String = "ImpPedidoB"
    Public Const Texto As String = "Texto"
    Public Const IDPedidoVenta As String = "IDPedidoVenta"
    Public Const IDTipoCompra As String = "IDTipoCompra"
    Public Const FechaCreacionAudi As String = "FechaCreacionAudi"
    Public Const FechaModificacionAudi As String = "FechaModificacionAudi"
    Public Const UsuarioAudi As String = "UsuarioAudi"
    Public Const SyncDB As String = "SyncDB"
    Public Const FechaEntrega As String = "FechaEntrega"
End Class

<Serializable()> _
Public Class PedidoCompraInfo
    Public Fields As New BusinessData
    Public Lineas(-1) As PedidoCompraLineaInfo
End Class

<Serializable()> _
Public Class PedidoCompraLineaInfo
    Public Fields As New BusinessData
    Public Analitica As DataTable
End Class

Public Class PedidoCompraCabeceraInfo
    Inherits ClassEntityInfo

    Public IDPedido As Integer
    Public IDCondicionPago As String
    Public IDFormaPago As String
    Public IDMoneda As String
    Public Dto As Double

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal data As DataRow)
        MyBase.New(data)
    End Sub

    Public Overloads Overrides Sub Fill(ByVal ParamArray PrimaryKey() As Object)
        Dim dtPCCInfo As DataTable = New PedidoCompraCabecera().SelOnPrimaryKey(PrimaryKey(0))
        If dtPCCInfo.Rows.Count > 0 Then
            Me.Fill(dtPCCInfo.Rows(0))
        End If
    End Sub

End Class

<Serializable()> _
Public Class PedidoCompraUpdateData
    Public Pedidos As DataTable
    Public ProgramasError(-1) As String
    Public ProveedoresError(-1) As String
    Public ArticulosError(-1) As String
    Public OFsError(-1) As String
    Public MensajeError As String
End Class

Public Class PedidoCompraCabecera
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper
    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbPedidoCompraCabecera"
    Private Const cnParametroProduccion As Integer = 3

#Region " RegisterAddNewTasks "

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoComunes.AsignarIdentificadorPedido, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf NegocioGeneral.AsignarCentroGestion, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarContador, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarNumeroPedidoProvisional, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf NegocioGeneral.AsignarAlmacen, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoComunes.AsignarFechaPedido, data, services)
        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ProcesoComunes.AsignarEjercicioContablePedido, New DataRowPropertyAccessor(data), services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoCompra.AsignarTipoCompra, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarOperario, data, services)
    End Sub

    <Task()> Public Shared Sub AsignarContador(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim CE As New CentroEntidad
        CE.CentroGestion = data("IDCentroGestion") & String.Empty
        CE.ContadorEntidad = CentroGestion.ContadorEntidad.PedidoCompra
        data("IDContador") = ProcessServer.ExecuteTask(Of CentroEntidad, String)(AddressOf New CentroGestion().GetContadorPredeterminado, CE, services)
    End Sub

    <Task()> Public Shared Sub AsignarNumeroPedidoProvisional(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDContador")) > 0 Then
            Dim dtContadores As DataTable = ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf Contador.CounterDt, GetType(PedidoCompraCabecera).Name, services)
            Dim adr As DataRow() = dtContadores.Select("IDContador = " & Quoted(data("IDContador")))
            If Not IsNothing(adr) AndAlso adr.Length > 0 Then
                data("NPedido") = adr(0)("ValorProvisional")
            Else
                'Si no está bien configurado el Contador de Pedidos de Compra en el Centro de Gestión,
                'cogemos el Contador por defecto de la entidad Pedido CompraCabecera.
                Dim dtContadorPred As DataTable = ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf Contador.CounterDefault, GetType(PedidoCompraCabecera).Name, services)
                If Not dtContadorPred Is Nothing AndAlso dtContadorPred.Rows.Count > 0 Then
                    data("IDContador") = dtContadorPred.Rows(0)("IDContador")
                    adr = dtContadores.Select("IDContador = " & Quoted(data("IDContador")))
                    If Not IsNothing(adr) AndAlso adr.Length > 0 Then
                        data("NPedido") = adr(0)("ValorProvisional")
                    End If
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarOperario(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim strIDOperario As String = ProcessServer.ExecuteTask(Of Object, String)(AddressOf Operario.ObtenerIDOperarioUsuario, Nothing, services)
        If Len(strIDOperario) > 0 Then
            data("IDOperario") = strIDOperario
        End If
    End Sub

#End Region
#Region " Validate "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarProveedorObligatorio)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarFechaPedidoObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarMonedaObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarFormaPagoObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarCondicionPagoObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf ProcesoPedidoCompra.ValidacionesContabilidad)
    End Sub


#End Region
#Region " Update "
    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of UpdatePackage, DocumentoPedidoCompra)(AddressOf ProcesoPedidoCompra.CrearDocumento)
        updateProcess.AddTask(Of DocumentoPedidoCompra)(AddressOf ProcesoPedidoCompra.AsignarValoresPredeterminadosGenerales)
        updateProcess.AddTask(Of DocumentoPedidoCompra)(AddressOf ProcesoPedidoCompra.AsignarNumeroPedido)
        updateProcess.AddTask(Of DocumentoPedidoCompra)(AddressOf ProcesoPedidoCompra.AsignarCentroGestion)
        updateProcess.AddTask(Of DocumentoPedidoCompra)(AddressOf ProcesoPedidoCompra.ActualizarCambiosMoneda)
        updateProcess.AddTask(Of DocumentoPedidoCompra)(AddressOf ProcesoPedidoCompra.CrearComponentesSubcontratacionPC)
        updateProcess.AddTask(Of DocumentoPedidoCompra)(AddressOf ProcesoPedidoCompra.ActualizarComponentesSubcontratacion)
        updateProcess.AddTask(Of DocumentoPedidoCompra)(AddressOf ProcesoPedidoCompra.CalcularImporteLineasPedido)
        updateProcess.AddTask(Of DocumentoPedidoCompra)(AddressOf ProcesoPedidoCompra.CalcularAnalitica)
        updateProcess.AddTask(Of DocumentoPedidoCompra)(AddressOf ProcesoPedidoCompra.CalcularBasesImponibles)
        updateProcess.AddTask(Of DocumentoPedidoCompra)(AddressOf ProcesoComunes.TotalDocumento)
        updateProcess.AddTask(Of DocumentoPedidoCompra)(AddressOf Business.General.Comunes.BeginTransaction)
        updateProcess.AddTask(Of DocumentoPedidoCompra)(AddressOf ProcesoPedidoCompra.ActualizarEntidadesDependientesUpdate)
        updateProcess.AddTask(Of DocumentoPedidoCompra)(AddressOf Business.General.Comunes.UpdateDocument)
        updateProcess.AddTask(Of DocumentoPedidoCompra)(AddressOf Business.General.Comunes.MarcarComoActualizado)
    End Sub
#End Region

#Region " Delete "

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf ActualizarProgramaDelete)
        deleteProcess.AddTask(Of DataRow)(AddressOf ActualizarPedidosMultiEmpresaDelete)
    End Sub

    <Task()> Public Shared Sub ActualizarProgramaDelete(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim dt As DataTable = New ProgramaCompraCabecera().Filter(New NumberFilterItem("IDPedido", data("IDPedido")))
        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            For Each dr As DataRow In dt.Rows
                dt.Rows(0)("IDPedido") = System.DBNull.Value
            Next
            BusinessHelper.UpdateTable(dt)
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarPedidosMultiEmpresaDelete(ByVal data As DataRow, ByVal services As ServiceProvider)
        '//control pedidos multiempresa
        Dim grp As New GRPPedidoVentaCompraLinea
        Dim control As DataTable = grp.TrazaPCPrincipal(data("IDPedido"))
        If Not control Is Nothing AndAlso control.Rows.Count > 0 Then

            If Length(control.Rows(0)("IDPVSecundaria")) > 0 Then
                Dim NPedidoVenta As String = control.Rows(0)("NPVSecundaria")
                Dim DescBaseDatos As String = New NegocioGeneral().GetDataBaseDescription(CType(control.Rows(0)("IDBDSecundaria"), Guid))
                Throw New Exception("No se puede eliminar el pedido de compra. Este pedido ha generado una línea en el pedido de venta Nº " & NPedidoVenta & " en la empresa " & Quoted(DescBaseDatos) & ".")
            End If

        End If
    End Sub

#End Region
#Region " BusinessRules "

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim services As New ServiceProvider
        Dim oBRL As New SynonymousBusinessRules
        oBRL.AddSynonymous("FechaPedido", "Fecha")
        'CCInmovilizado(current, New ServiceProvider)
        oBRL = ProcessServer.ExecuteTask(Of BusinessRules, BusinessRules)(AddressOf ProcesoCompra.DetailBusinessRulesCab, oBRL, services)

        oBRL("IDProveedor") = AddressOf CambioProveedor
        oBRL.Add("Fecha", AddressOf CambioFechaPedido)
        oBRL.Add("IDAlmacen", AddressOf ProcesoComunes.CambioAlmacen)
        oBRL.Add("IDTipoCompra", AddressOf CambioTipoCompra)
        oBRL.Add("IDCentroGestion", AddressOf ProcesoComunes.CambioCentroGestion)
        Return oBRL
    End Function
    <Task()> Public Shared Sub CambioProveedor(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)

        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ProcesoCompra.CambioProveedor, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ProcesoComunes.CambioMoneda, data, services)
        Dim dir As New DataDireccionProv(enumpdTipoDireccion.pdDireccionPedido, "IDDireccion", data.Current)
        ProcessServer.ExecuteTask(Of DataDireccionProv)(AddressOf ProcesoCompra.AsignarDireccionProveedor, dir, services)
        Dim obs As New DataObservaciones(GetType(PedidoCompraCabecera).Name, "Texto", data.Current)
        ProcessServer.ExecuteTask(Of DataObservaciones)(AddressOf ProcesoCompra.AsignarObservacionesProveedor, obs, services)

        If Length(data.Current("IDProveedor")) > 0 Then
            Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
            Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(data.Current("IDProveedor"))
            data.Current("Dto") = ProvInfo.DtoComercial
        Else
            data.Current("Dto") = 0
        End If
    End Sub
    <Task()> Public Shared Sub CambioTipoCompra(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Current("IDPedido")) > 0 Then
            Dim dtPCL As DataTable = New PedidoCompraLinea().Filter(New NumberFilterItem("IDPedido", data.Current("IDPedido")))
            If Not IsNothing(dtPCL) AndAlso dtPCL.Rows.Count > 0 Then
                ApplicationService.GenerateError("No es posible modificar el Tipo Compra, existen líneas de pedido.")
            End If
        End If
    End Sub
    <Task()> Public Shared Sub CambioFechaPedido(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)

        If data.ColumnName = "Fecha" Then
            '//Hay que ponerlo en los dos campos indicados en el Synonimous.
            data.Current(data.ColumnName) = data.Value
            data.Current("FechaPedido") = data.Value
        End If
        If Length(data.Current("FechaPedido")) > 0 Then
            ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ProcesoComunes.AsignarEjercicioContablePedido, data.Current, services)
            ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ProcesoComunes.CambioMoneda, data, services)
        Else
            data.Current("IDEjercicio") = DBNull.Value
        End If
        'TODO
        'CCInmovilizado(data.Current, New ServiceProvider)
    End Sub
    <Task()> Public Shared Sub CCInmovilizado(ByVal current As IPropertyAccessor, ByVal services As ServiceProvider)
        If Length(current("IDEjercicio")) > 0 AndAlso Length(current("IDPedido")) > 0 Then
            '// Comprobar si para las líneas del Pedido que existan, las C.Contables son de inmovilizado o 
            '// no según corresponda. 

            Dim objNegPCL As New PedidoCompraLinea
            Dim dtPCL As DataTable = objNegPCL.Filter(New NumberFilterItem("IDPedido", current("IDPedido")))
            If dtPCL.Columns.Contains("Inmovilizado") AndAlso dtPCL.Columns.Contains("CContable") Then
                For Each drRow As DataRow In dtPCL.Rows
                    Dim dataCCI As New ProcesoCompra.DataCContableInmovilizado
                    dataCCI.IDEjercicio = current("IDEjercicio") & String.Empty
                    dataCCI.CContable = drRow("CContable") & String.Empty
                    dataCCI.Inmovilizado = Nz(drRow("Inmovilizado"), False)
                    ProcessServer.ExecuteTask(Of ProcesoCompra.DataCContableInmovilizado)(AddressOf ProcesoCompra.ValidarCuentaInmovilizado, dataCCI, services)
                Next
            End If
        End If
    End Sub
    Private Function ExistenLineasPedido(ByVal intIDPedido As Integer) As Boolean
        Dim dtPCL As DataTable = New PedidoCompraLinea().Filter(New NumberFilterItem("IDPedido", intIDPedido))
        Return (Not IsNothing(dtPCL) AndAlso dtPCL.Rows.Count > 0)
    End Function
#End Region
#Region " Precio Optimo "

    <Task()> Public Shared Sub PrecioOptimo(ByVal IdPedido As Integer, ByVal services As ServiceProvider)
        Dim DocPed As DocumentoPedidoCompra = ProcessServer.ExecuteTask(Of Integer, DocumentoPedidoCompra)(AddressOf CrearDocumento, IdPedido, services)
        ProcessServer.ExecuteTask(Of DocumentoPedidoCompra)(AddressOf CalculoPrecioOptimo, DocPed, services)
        ProcessServer.ExecuteTask(Of DocumentoPedidoCompra)(AddressOf ProcesoPedidoCompra.CalcularAnalitica, DocPed, services)
        ProcessServer.ExecuteTask(Of DocumentoPedidoCompra)(AddressOf ProcesoPedidoCompra.CalcularBasesImponibles, DocPed, services)
        ProcessServer.ExecuteTask(Of DocumentoPedidoCompra)(AddressOf ProcesoComunes.TotalDocumento, DocPed, services)
        ProcessServer.ExecuteTask(Of DocumentoPedidoCompra)(AddressOf Business.General.Comunes.UpdateDocument, DocPed, services)
    End Sub

    <Task()> Public Shared Function CrearDocumento(ByVal IDPedido As Integer, ByVal services As ServiceProvider) As DocumentoPedidoCompra
        Return New DocumentoPedidoCompra(IDPedido)
    End Function

    <Task()> Public Shared Sub CalculoPrecioOptimo(ByVal Doc As DocumentoPedidoCompra, ByVal services As ServiceProvider)
        If Doc Is Nothing OrElse Doc.dtLineas Is Nothing OrElse Doc.dtLineas.Rows.Count = 0 Then Exit Sub

        'Recogemos los articulos que esten relacionados con ese Pedido.
        Dim dtArticulosPedido As DataTable = New BE.DataEngine().Filter("vNegPedidoCompraLineaArticulos", New StringFilterItem("IDPedido", Doc.HeaderRow("IDPedido")))
        Dim f As New Filter
        For Each drArticuloPedido As DataRow In dtArticulosPedido.Select
            f.Clear()
            f.Add("IDArticulo", drArticuloPedido("IDArticulo"))
            f.Add("Estado", FilterOperator.NotEqual, enumpvlEstado.pvlServido)
            'Recogemos las lineas del pedido que tengan el articulo de este momento
            Dim QPedida As Double = Nz(Doc.dtLineas.Compute("SUM(QPedida)", f.Compose(New AdoFilterComposer)), 0)

            Dim dataTarifa As New DataCalculoTarifaCompra
            dataTarifa.IDArticulo = drArticuloPedido("IDArticulo")
            dataTarifa.IDProveedor = Doc.IDProveedor
            dataTarifa.Cantidad = QPedida
            dataTarifa.Fecha = Doc.Fecha
            dataTarifa.IDMoneda = Doc.IDMoneda
            ProcessServer.ExecuteTask(Of DataCalculoTarifaCompra)(AddressOf ProcesoCompra.TarifaCompra, dataTarifa, services)
            If Not dataTarifa.DatosTarifa Is Nothing Then

                Dim context As New BusinessData(Doc.HeaderRow)
                Dim PCL As New PedidoCompraLinea
                Dim WhereArticuloNoServido As String = f.Compose(New AdoFilterComposer)
                For Each drPedidoLineaArticulo As DataRow In Doc.dtLineas.Select(WhereArticuloNoServido)
                    PCL.ApplyBusinessRule("Precio", dataTarifa.DatosTarifa.Precio, drPedidoLineaArticulo, context)
                    PCL.ApplyBusinessRule("Dto1", dataTarifa.DatosTarifa.Dto1, drPedidoLineaArticulo, context)
                    PCL.ApplyBusinessRule("Dto2", dataTarifa.DatosTarifa.Dto2, drPedidoLineaArticulo, context)
                    PCL.ApplyBusinessRule("Dto3", dataTarifa.DatosTarifa.Dto3, drPedidoLineaArticulo, context)
                    PCL.ApplyBusinessRule("UDValoracion", dataTarifa.DatosTarifa.UDValoracion, drPedidoLineaArticulo, context)

                    If Length(dataTarifa.DatosTarifa.SeguimientoTarifa) > 0 Then
                        drPedidoLineaArticulo("SeguimientoTarifa") = dataTarifa.DatosTarifa.SeguimientoTarifa
                    End If
                Next
            End If
            QPedida = 0
        Next
    End Sub

#End Region
#Region " Gestión MultiEmpresa "

    Private Sub ActualizarPedidosMultiEmpresa(ByVal dr As DataRow, Optional ByVal blnDelete As Boolean = False)
        Dim grp As New GRPPedidoVentaCompraLinea
        Dim control As DataTable = grp.TrazaPCPrincipal(dr("IDPedido"))
        If Not control Is Nothing AndAlso control.Rows.Count > 0 Then
            Dim General As New NegocioGeneral
            If blnDelete Then
                If Length(control.Rows(0)("IDPVSecundaria")) > 0 Then
                    Dim NPedidoVenta As String = control.Rows(0)("NPVSecundaria")
                    Dim DescBaseDatos As String = General.GetDataBaseDescription(CType(control.Rows(0)("IDBDSecundaria"), Guid))
                    Throw New Exception("No se puede eliminar el pedido de compra. Este pedido ha generado una línea en el pedido de venta Nº " & NPedidoVenta & " en la empresa " & Quoted(DescBaseDatos) & ".")
                Else
                    For Each drControl As DataRow In control.Rows
                        grp.Delete(drControl("IDPVLinea"))
                    Next
                    BusinessHelper.UpdateTable(control)
                End If
            Else
                If Length(control.Rows(0)("IDPVPrincipal")) > 0 Then
                    Dim pedidoVenta As DataRow = New PedidoVentaCabecera().GetItemRow(control.Rows(0)("IDPVPrincipal"))
                    '//Mantener esta igualdad
                    dr("EntregaProveedor") = pedidoVenta("EntregaProveedor")
                Else
                    If AreDifferents(dr("IDProveedor"), dr("IDProveedor", DataRowVersion.Original)) _
                    Or AreDifferents(dr("EntregaProveedor"), dr("EntregaProveedor", DataRowVersion.Original)) Then
                        '//Nota: el campo EmpresaGrupo solo se puede modificar si cambia el proveedor
                        If Length(control.Rows(0)("IDPVSecundaria")) > 0 Then
                            Dim descBaseDatos As String = General.GetDataBaseDescription(control.Rows(0)("IDBDSecundaria"))
                            Throw New Exception("Este pedido está relacionado con el pediodo de venta Nº " & control.Rows(0)("NPVSecundaria") _
                            & " de la base de datos " & Quoted(descBaseDatos) & "." & ControlChars.NewLine _
                            & "No se permite modificar el proveedor o datos que dependan del mismo.")
                        Else
                            '//Actualizar registros de la tabla tbGRPPedidoVentaCompraLinea
                            For Each traza As DataRow In control.Rows
                                traza("EntregaProveedor") = dr("EntregaProveedor")
                            Next
                            BusinessHelper.UpdateTable(control)
                        End If
                    End If
                End If

                If Length(control.Rows(0)("IDPVSecundaria")) > 0 Then
                    '//grabar el mismo texto en el pedido de venta de la bbdd secundaria
                    If AreDifferents(dr("Texto"), dr("Texto", DataRowVersion.Original)) Then
                        Dim databaseBak As String = AdminData.GetSessionDataBase()
                        Try
                            AdminData.SetSessionDataBase(control.Rows(0)("IDBDSecundaria"))
                            Dim pedidoVenta As DataTable = New PedidoVentaCabecera().SelOnPrimaryKey(control.Rows(0)("IDPVSecundaria"))
                            If pedidoVenta.Rows.Count > 0 Then
                                pedidoVenta.Rows(0)("Texto") = dr("Texto")
                                BusinessHelper.UpdateTable(pedidoVenta)
                            End If
                        Catch ex As Exception
                            AdminData.RollBackTx()
                            Throw ex
                        Finally
                            AdminData.SetSessionDataBase(databaseBak)
                        End Try
                    End If
                End If
            End If
        End If
    End Sub

#End Region
#Region " Seguimiento "
    Public Function SeguimientoPedidoVenta(ByVal IDPedido As Integer) As DataTable
        Return New BE.DataEngine().Filter("vConsultaPedidosCompraVenta", New NumberFilterItem("IDPCPrincipal", IDPedido))
    End Function
#End Region
End Class

