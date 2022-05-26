Imports Solmicro.Expertis.Business.Negocio.NegocioGeneral

Public Class _PedidoVentaLinea
    Public Const IDLineaPedido As String = "IDLineaPedido"
    Public Const IDPedido As String = "IDPedido"
    Public Const IDArticulo As String = "IDArticulo"
    Public Const DescArticulo As String = "DescArticulo"
    Public Const RefCliente As String = "RefCliente"
    Public Const DescRefCliente As String = "DescRefCliente"
    Public Const IDAlmacen As String = "IDAlmacen"
    Public Const IDTipoIVA As String = "IDTipoIVA"
    Public Const FechaEntrega As String = "FechaEntrega"
    Public Const PedidoCliente As String = "PedidoCliente"
    Public Const IDUdMedida As String = "IDUdMedida"
    Public Const IDUdInterna As String = "IDUdInterna"
    Public Const CContable As String = "CContable"
    Public Const QPedida As String = "QPedida"
    Public Const QServida As String = "QServida"
    Public Const QAlbaran As String = "QAlbaran"
    Public Const QDisponible As String = "QDisponible"
    Public Const Precio As String = "Precio"
    Public Const PrecioA As String = "PrecioA"
    Public Const PrecioB As String = "PrecioB"
    Public Const UdValoracion As String = "UdValoracion"
    Public Const Importe As String = "Importe"
    Public Const ImporteA As String = "ImporteA"
    Public Const ImporteB As String = "ImporteB"
    Public Const Estado As String = "Estado"
    Public Const Confirmado As String = "Confirmado"
    Public Const PreparadoExp As String = "PreparadoExp"
    Public Const Dto1 As String = "Dto1"
    Public Const Dto2 As String = "Dto2"
    Public Const Dto3 As String = "Dto3"
    Public Const Dto As String = "Dto"
    Public Const DtoProntoPago As String = "DtoProntoPago"
    Public Const Texto As String = "Texto"
    Public Const IdPrograma As String = "IdPrograma"
    Public Const IdLineaPrograma As String = "IdLineaPrograma"
    Public Const Regalo As String = "Regalo"
    Public Const Prioridad As String = "Prioridad"
    Public Const IDPromocionLinea As String = "IDPromocionLinea"
    Public Const IdOrdenLinea As String = "IdOrdenLinea"
    Public Const IdLineaPedidoCompra As String = "IdLineaPedidoCompra"
    Public Const EspecificacionesArticulo As String = "EspecificacionesArticulo"
    Public Const Factor As String = "Factor"
    Public Const QInterna As String = "QInterna"
    Public Const Muelle As String = "Muelle"
    Public Const PuntoDescarga As String = "PuntoDescarga"
    Public Const Revision As String = "Revision"
    Public Const IDLineaOfertaDetalle As String = "IDLineaOfertaDetalle"
    Public Const Eliminar As String = "Eliminar"
    Public Const Deposito As String = "Deposito"
    Public Const QFacturada As String = "QFacturada"
    Public Const SeguimientoTarifa As String = "SeguimientoTarifa"
    Public Const FechaCreacionAudi As String = "FechaCreacionAudi"
    Public Const FechaModificacionAudi As String = "FechaModificacionAudi"
    Public Const UsuarioAudi As String = "UsuarioAudi"
    Public Const PedidoVentaOrigen As String = "PedidoVentaOrigen"
    Public Const IDCentroGestion As String = "IDCentroGestion"
    Public Const FechaEntregaModificado As String = "FechaEntregaModificado"
    Public Const PedidoClienteDestino As String = "PedidoClienteDestino"
    Public Const PrecioCosteA As String = "PrecioCosteA"
    Public Const PrecioCosteB As String = "PrecioCosteB"
    Public Const FechaPreparacion As String = "FechaPreparacion"
    Public Const IDTipoLinea As String = "IDTipoLinea"
End Class

Public Class PedidoVentaLinea
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Private Const cnEntidad As String = "tbPedidoVentaLinea"

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub


#Region " Validate "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ProcesoComercial.DetailCommonUpdateRules)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarFechaEntregaObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarAlmacenObligatorio)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarAlmacenBloqueado)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarCantidadLineaPedido)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarArticuloBloqueado)
    End Sub

    <Task()> Public Shared Sub ValidarArticuloBloqueado(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDArticulo")) <> 0 Then
            Dim Cabecera As DataTable = New PedidoVentaCabecera().SelOnPrimaryKey(data("IDPedido"))
            If Not Cabecera Is Nothing AndAlso Cabecera.Rows.Count > 0 Then
                Dim StDatos As New Cliente.DataBloqArtClie
                StDatos.IDArticulo = data("IDArticulo") : StDatos.IDCliente = Cabecera.Rows(0)("IDCliente")
                If ProcessServer.ExecuteTask(Of Cliente.DataBloqArtClie, Boolean)(AddressOf Cliente.ComprobarBloqueoArticuloCliente, StDatos, services) Then
                    ApplicationService.GenerateError("El Artículo está bloqueado para este Cliente.")
                End If
            End If
        End If
    End Sub
#End Region
#Region " Delete "

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf Business.General.Comunes.BeginTransaction)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ValidarEstadoLineaDel)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ValidarOrdenFabricacionOrigen)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ActualizarOfertaComercialDetalle)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ActualizarLineaProgramaDel)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ActualizarPedidosMultiEmpresa)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf Business.General.Comunes.DeleteEntityRow)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf Business.General.Comunes.MarcarComoEliminado)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ActualizarLineaPromocion)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ActualizarCantidadPromocionada)
    End Sub
    <Task()> Public Shared Sub ValidarEstadoLineaDel(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data("Estado") <> enumpvlEstado.pvlPedido Then
            ApplicationService.GenerateError("No se puede eliminar la línea porque está Servida, Parcialmente Servida o Cerrada.")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarOrdenFabricacionOrigen(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim ClsOF As BusinessHelper = BusinessHelper.CreateBusinessObject("OrdenFabricacionOrigen")
        Dim FilOF As New Filter
        FilOF.Add("IDOrigen", FilterOperator.Equal, data("IDLineaPedido"))
        FilOF.Add("TipoOrigen", FilterOperator.Equal, CInt(enumofoTipoOrigen.ofoPedidoVenta))
        Dim DtOF As DataTable = ClsOF.Filter(FilOF)
        If Not DtOF Is Nothing AndAlso DtOF.Rows.Count > 0 Then
            ApplicationService.GenerateError("No se puede eliminar la línea porque tiene relacionada una Orden de Fabricación.")
        End If
    End Sub

    <Task()> Public Shared Function NoHaSidoEliminada(ByVal data As DataRow, ByVal services As ServiceProvider) As Boolean
        Dim listaEliminados As LineasPedidoEliminadas = services.GetService(Of LineasPedidoEliminadas)()
        Dim haSidoEliminado As Boolean = listaEliminados.IDLineas.Contains(data("IDLineaPedido"))
        If haSidoEliminado Then services.GetService(Of DeleteProcessContext).Deleted = haSidoEliminado
        Return Not haSidoEliminado
    End Function

    '<Task()> Public Shared Sub EliminarLineasPromocion(ByVal data As DataRow, ByVal services As ServiceProvider)
    '    '//Si es un artículo de promoción, pero no es un regalo. O bien, es un regalo que no sea de una promoción.
    '    If (Not data("Regalo") AndAlso Length("IDPromocionLinea") > 0) OrElse (data("Regalo") AndAlso Length("IDPromocionLinea") = 0) Then

    '    End If
    '    ProcessServer.ExecuteTask(Of DataRow)(AddressOf Business.General.Comunes.DeleteEntityRow, data, services)
    '    ProcessServer.ExecuteTask(Of DataRow)(AddressOf Business.General.Comunes.MarcarComoEliminado, data, services)


    '    ProcessServer.ExecuteTask(Of DataRow)(AddressOf ActualizarLineaPromocion, data, services)
    '    ProcessServer.ExecuteTask(Of DataRow)(AddressOf ActualizarCantidadPromocionada, data, services)
    'End Sub
    <Task()> Public Shared Sub ActualizarLineaPromocion(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDPromocionLinea")) > 0 And data("Regalo") = 0 Then
            Dim StDatos As New PromocionLinea.DatosActuaLinPromoDr(data.Table, True)
            ProcessServer.ExecuteTask(Of PromocionLinea.DatosActuaLinPromoDr)(AddressOf PromocionLinea.ActualizarLineaPromocion, StDatos, services)
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarOfertaComercialDetalle(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDLineaOfertaDetalle")) > 0 Then
            Dim oferta As BusinessHelper = BusinessHelper.CreateBusinessObject("OfertaComercialDetalle")
            Dim detalle As DataTable = oferta.SelOnPrimaryKey(data("IDLineaOfertaDetalle"))
            If Not detalle Is Nothing AndAlso detalle.Rows.Count > 0 Then
                detalle.Rows(0)("EstadoVenta") = False
                BusinessHelper.UpdateTable(detalle)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarLineaProgramaDel(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Not data Is Nothing Then
            If Length(data("IDLineaPrograma")) > 0 Then
                Dim Programa As DataTable = New ProgramaLinea().SelOnPrimaryKey(data("IDLineaPrograma"))
                If Not Programa Is Nothing AndAlso Programa.Rows.Count > 0 Then
                    Dim DblConfir As Double = Programa.Rows(0)("QConfirmada") - data("QPedida")
                    Programa.Rows(0)("QConfirmada") = DblConfir
                    Programa.Rows(0)("Confirmada") = IIf(DblConfir > 0, True, False)
                    Programa.Rows(0)("FechaConfirmacion") = IIf(DblConfir > 0, Programa.Rows(0)("FechaConfirmacion"), DBNull.Value)
                    BusinessHelper.UpdateTable(Programa)
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarPedidosMultiEmpresa(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim NPedidoCompra, NPedidoVenta, DescBaseDatos As String
        AdminData.BeginTx()
        Dim grp As New GRPPedidoVentaCompraLinea
        Dim control As DataTable = grp.TrazaPVLPrincipal(data("IDLineaPedido"))
        If Not control Is Nothing AndAlso control.Rows.Count > 0 Then
            If Length(control.Rows(0)("IDPVSecundaria")) > 0 Then
                NPedidoVenta = control.Rows(0)("NPVSecundaria")
                DescBaseDatos = New NegocioGeneral().GetDataBaseDescription(CType(control.Rows(0)("IDBDSecundaria"), Guid))
                Throw New Exception("No se puede eliminar el pedido de venta. Este pedido ha generado una línea en el pedido de venta Nº " & NPedidoVenta & " en la empresa " & Quoted(DescBaseDatos) & ".")
            ElseIf Length(control.Rows(0)("IDPCPrincipal")) > 0 Then
                NPedidoCompra = control.Rows(0)("NPCPrincipal")
                Throw New Exception("No se puede eliminar el pedido de venta. Este pedido ha generado una línea en el pedido de compra Nº " & NPedidoCompra & ".")
            Else
                grp.Delete(control.Rows(0)("IDPVLinea"))
            End If
        Else
            Dim BaseDatos As DataTable = AdminData.GetData("vConsultaBaseDatosPrincipal", New NumberFilterItem("IDLineaPedido", data("IDLineaPedido")))
            If Not BaseDatos Is Nothing AndAlso BaseDatos.Rows.Count > 0 Then
                Dim databaseBak As String = AdminData.GetSessionDataBase()
                Dim BDInfo As New DataBasesDatosMultiempresa(AdminData.GetConnectionInfo.IDDataBase, BaseDatos.Rows(0)("BaseDatos"))

                Try
                    AdminData.CommitTx(True)
                    If Length(BDInfo.IDBaseDatosSecundaria) > 0 Then AdminData.SetCurrentConnection(BDInfo.IDBaseDatosSecundaria)
                    Dim grp2 As New GRPPedidoVentaCompraLinea
                    Dim control2 As DataTable = grp2.TrazaPVLSecundaria(data("IDLineaPedido"))
                    If Not control2 Is Nothing AndAlso control2.Rows.Count > 0 Then
                        If Length(control2.Rows(0)("IDPVSecundaria")) > 0 Then
                            control2.Rows(0)("IDPVSecundaria") = DBNull.Value
                            control2.Rows(0)("NPVSecundaria") = DBNull.Value
                            control2.Rows(0)("IDLineaPVSecundaria") = DBNull.Value
                            control2.Rows(0)("IDBDSecundaria") = DBNull.Value
                            AdminData.SetData(control2)
                        End If
                    End If
                Catch ex As Exception
                    AdminData.RollBackTx(True)

                    If Length(BDInfo.IDBaseDatosPrincipal) > 0 Then AdminData.SetCurrentConnection(BDInfo.IDBaseDatosPrincipal)

                    Throw ex
                Finally
                    AdminData.CommitTx(True)
                    AdminData.SetCurrentConnection(BDInfo.IDBaseDatosPrincipal)
                End Try
            End If

        End If
    End Sub

    <Task()> Public Shared Sub ActualizarCantidadPromocionada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDPromocionLinea")) > 0 AndAlso data("Regalo") = 0 Then
            ProcessServer.ExecuteTask(Of Integer)(AddressOf PromocionLinea.ActualizarLineaPromocionQ, data("IDPromocionLinea"), services)
        End If
    End Sub


#End Region
#Region " Business Rules "

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim services As New ServiceProvider
        Dim oBRL As New SynonymousBusinessRules
        oBRL.AddSynonymous("QPedida", "Cantidad")
        '//BusinessRules  -  Comercial
        oBRL = ProcessServer.ExecuteTask(Of BusinessRules, BusinessRules)(AddressOf ProcesoComercial.DetailBusinessRulesLin, oBRL, services)

        '//BusinessRules - Específicas PVL
        oBRL.Add("IDAlmacen", AddressOf CambioAlmacen)
        Return oBRL
    End Function

    <Task()> Public Shared Sub CambioAlmacen(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDAlmacen")) > 0 Then
            Dim Almacenes As EntityInfoCache(Of AlmacenInfo) = services.GetService(Of EntityInfoCache(Of AlmacenInfo))()
            Dim AlmInfo As AlmacenInfo = Almacenes.GetEntity(data.Current("IDAlmacen"))
            If AlmInfo.Bloqueado Then
                ApplicationService.GenerateError("El Almacén {0} está bloqueado.", data.Current("IDAlmacen"))
            Else
                If Length(data.Current("IDArticulo")) > 0 AndAlso Length(data.Current("IDAlmacen")) > 0 Then
                    Dim DtArtAlm As DataTable = New ArticuloAlmacen().SelOnPrimaryKey(data.Current("IDArticulo"), data.Current("IDAlmacen"))
                    If Not DtArtAlm Is Nothing AndAlso DtArtAlm.Rows.Count > 0 Then
                        data.Current("StockFisico") = Nz(DtArtAlm.Rows(0)("StockFisico"), 0)
                    Else : data.Current("StockFisico") = 0
                    End If
                Else : data.Current("StockFisico") = 0
                End If
            End If
        End If
    End Sub

#End Region
#Region "Seguimiento"
    Public Function SeguimientoPedidoCompra(ByVal IDLineaPedido As Integer) As DataTable
        Dim dt As DataTable = New BE.DataEngine().Filter("vConsultaPedidosCompraVentaLinea", New NumberFilterItem("IDLineaPVPrincipal", IDLineaPedido))
        If dt.Rows.Count = 0 Then
            dt = New BE.DataEngine().Filter("vConsultaPedidosCompraVentaLinea", New NumberFilterItem("IDLineaPVSecundaria", IDLineaPedido))
        End If
        dt.Columns.Add("EmpresaGrupo", GetType(Boolean))
        dt.Columns.Add("EntregaProveedor", GetType(Boolean))
        dt.Columns.Add("DescBaseDatosPrincipal", GetType(String))
        dt.Columns.Add("DescBaseDatosSecundaria", GetType(String))
        If dt.Rows.Count > 0 Then

            If Not dt.Rows(0).IsNull("IDLineaPVPrincipal") AndAlso dt.Rows(0)("IDLineaPVPrincipal") = IDLineaPedido Then
                If Length(dt.Rows(0)("IDPVPrincipal")) Then
                    Dim pedido As DataRow = New PedidoVentaCabecera().GetItemRow(dt.Rows(0)("IDPVPrincipal"))
                    dt.Rows(0)("EmpresaGrupo") = pedido("EmpresaGrupo")
                    dt.Rows(0)("EntregaProveedor") = pedido("EntregaProveedor")
                End If
            ElseIf dt.Rows(0)("IDLineaPVSecundaria") = IDLineaPedido Then
                Dim pedido As DataRow = New PedidoVentaCabecera().GetItemRow(dt.Rows(0)("IDPVSecundaria"))
                dt.Rows(0)("EmpresaGrupo") = pedido("EmpresaGrupo")
                dt.Rows(0)("EntregaProveedor") = pedido("EntregaProveedor")
            End If

            If Not dt.Rows(0).IsNull("IDBDPrincipal") Then
                Dim db As DataTable = New BE.DataEngine().Filter("xDataBase", New GuidFilterItem("IDBaseDatos", CType(dt.Rows(0)("IDBDPrincipal"), Guid)), , , , True)
                If db.Rows.Count > 0 Then
                    dt.Rows(0)("DescBaseDatosPrincipal") = db.Rows(0)("DescBaseDatos")
                End If
            End If
            If Not dt.Rows(0).IsNull("IDBDSecundaria") Then
                Dim db As DataTable = New BE.DataEngine().Filter("xDataBase", New GuidFilterItem("IDBaseDatos", CType(dt.Rows(0)("IDBDSecundaria"), Guid)), , , , True)
                If db.Rows.Count > 0 Then
                    dt.Rows(0)("DescBaseDatosSecundaria") = db.Rows(0)("DescBaseDatos")
                End If
            End If
        End If
        Return dt
    End Function
#End Region

    Public Sub CambiarDatosDisponible(ByVal intIDLineaPedido As Integer, ByVal dtmNuevaFecha As Date, _
                                      ByVal dblNuevaCantidad As Double, ByRef dblNuevaCantidadPendiente As Double)

        Dim dtpvl As DataTable = Me.SelOnPrimaryKey(intIDLineaPedido)
        If Not IsNothing(dtpvl) AndAlso dtpvl.Rows.Count > 0 Then
            dtpvl.Rows(0)("FechaEntrega") = dtmNuevaFecha
            dtpvl.Rows(0)("QPedida") = dblNuevaCantidad
            dblNuevaCantidadPendiente = dblNuevaCantidad - dtpvl.Rows(0)("QServida")

            Update(dtpvl)
            CalcularDisponible(dtpvl.Rows(0))
        End If
    End Sub
#Region " CalcularDisponible "

    Public Sub CalcularDisponible(ByVal dt As DataTable)
        For Each dr As DataRow In dt.Rows
            CalcularDisponible(dr)
        Next
    End Sub

    Public Sub CalcularDisponible(ByVal dr As DataRow)
        If dr.Table.Columns.Contains("IDArticulo") AndAlso dr.Table.Columns.Contains("IDAlmacen") Then
            CalcularDisponible(dr("IDArticulo") & String.Empty, dr("IDAlmacen") & String.Empty)
        End If
    End Sub

    ' Esta funcion recalcula el disponible de cada articulo
    Public Sub CalcularDisponible(ByVal strIDArticulo As String, ByVal strIDAlmacen As String)
        If Len(strIDArticulo) > 0 And Len(strIDAlmacen) > 0 Then
            Dim strWhere As String = "IdArticulo='" & strIDArticulo & "' AND IdAlmacen='" & strIDAlmacen & "'"
            Dim dblDisponible As Double

            'Se calcula el StockFisico
            Dim aa = New ArticuloAlmacen
            Dim dtaa As DataTable = aa.Filter("StockFisico", strWhere)
            If Not IsNothing(dtaa) AndAlso dtaa.Rows.Count > 0 Then
                dblDisponible = dtaa.Rows(0)("StockFisico")
            End If

            'Se recalcula el nuevo disponible
            strWhere = strWhere & " AND Estado <> " & enumpvlEstado.pvlServido & " AND Estado <> " & enumpvlEstado.pvlCerrado
            Dim dtPVL As DataTable = Filter(, strWhere, "FechaEntrega ASC")
            If Not IsNothing(dtPVL) AndAlso dtPVL.Rows.Count Then
                For Each dr As DataRow In dtPVL.Rows
                    dr("QDisponible") = dblDisponible
                    dblDisponible = dblDisponible - dr("QInterna") + (dr("QServida") * dr("Factor"))
                Next
                Update(dtPVL)
            End If
        End If
    End Sub

#End Region
    Public Sub CambiarEstado(ByVal dtLineasPedido As DataTable, ByVal intEstado As Integer)
        If Not IsNothing(dtLineasPedido) AndAlso dtLineasPedido.Rows.Count > 0 Then
            For Each dr As DataRow In dtLineasPedido.Rows
                dr("Estado") = intEstado
            Next
            BusinessHelper.UpdateTable(dtLineasPedido)
        Else
            ApplicationService.GenerateError("No hay líneas para actualizar.")
        End If
    End Sub
    Public Sub ConfirmarExpedicion(ByVal dtLineasPedido As DataTable)
        If Not IsNothing(dtLineasPedido) AndAlso dtLineasPedido.Rows.Count > 0 Then
            For Each dr As DataRow In dtLineasPedido.Rows
                dr("QAlbaran") = dr("CantidadMarca1")
                dr("Confirmado") = 1
                If IsDate(dr("CantidadMarcaFecha2")) Then
                    dr("FechaEntregaModificado") = dr("CantidadMarcaFecha2")
                Else
                    If IsDate(dr("FechaEntrega")) Then
                        dr("FechaEntregaModificado") = dr("FechaEntrega")
                    Else
                        dr("FechaEntregaModificado") = Date.Today
                    End If
                End If
            Next
            BusinessHelper.UpdateTable(dtLineasPedido)
        Else
            ApplicationService.GenerateError("No hay líneas para confirmar la expedición.")
        End If
    End Sub
    Public Sub AnularConfirmacion(ByVal dtLineasPedido As DataTable)
        If Not IsNothing(dtLineasPedido) AndAlso dtLineasPedido.Rows.Count > 0 Then
            For Each dr As DataRow In dtLineasPedido.Rows
                dr("QAlbaran") = 0
                dr("Confirmado") = 0
                dr("FechaEntregaModificado") = System.DBNull.Value
            Next
            BusinessHelper.UpdateTable(dtLineasPedido)
        Else
            ApplicationService.GenerateError("No hay líneas para anular la confirmación.")
        End If
    End Sub
    Public Function ObtenerCantidadPendiente(ByVal f As Filter) As DataTable
        Return New BE.DataEngine().Filter("SELECT SUM(QPedida - QServida) AS QPendiente FROM " & Me.Table, f)
    End Function

End Class

#Region " Código que no se utiliza ¿? "
#Region " Update -  Revisar donde poner lo que queda "

'Private Sub TratarFechaEntrega(ByVal dr As DataRow)
'    Dim drCab As DataRow = New PedidoVentaCabecera().GetItemRow(dr("IDPedido"))
'    If Length(drCab("FechaEntrega")) = 0 Then
'        drCab("FechaEntrega") = dr("FechaEntrega")
'        BusinessHelper.UpdateTable(drCab.Table)
'    End If
'End Sub

'<Task()> public Shared Sub TratarPromocion(ByVal dr As DataRow, ByVal services As ServiceProvider)
'    Dim pl As New PromocionLinea
'    If dr.RowState = DataRowState.Added Then
'        Dim dtPromLineaOLD As DataTable = pl.SelOnPrimaryKey(dr("IDPromocionLinea"))
'        If Not IsNothing(dtPromLineaOLD) AndAlso dtPromLineaOLD.Rows.Count > 0 Then
'            If dr("QPedida") >= dtPromLineaOLD.Rows(0)("QMinPedido") Then
'                If dr("QPedida") > dtPromLineaOLD.Rows(0)("QMaxPedido") Then
'                    dr("QPedida") = dtPromLineaOLD.Rows(0)("QMaxPedido")
'                End If
'                Dim StDatos As New PromocionLinea.DatosActuaLinPromoDr
'                StDatos.Dr = dr
'                StDatos.Delete = False
'                ProcessServer.ExecuteTask(Of PromocionLinea.DatosActuaLinPromoDr)(AddressOf PromocionLinea.ActualizarLineaPromocion, StDatos, services)
'                ADDLineaRegaloPedido(dr)
'            End If
'        End If
'    ElseIf dr.RowState = DataRowState.Modified Then
'        Dim intIDPromocionLineaOLD As Integer
'        Dim strIDPromocionOLD As String
'        Dim dblQPedidaOLD As Double
'        If Nz(dr("QPedida", DataRowVersion.Original), 0) <> Nz(dr("QPedida"), 0) Then
'            intIDPromocionLineaOLD = Nz(dr("IDPromocionLinea", DataRowVersion.Original), 0)
'            strIDPromocionOLD = dr("IDPromocion", DataRowVersion.Original) & String.Empty
'            dblQPedidaOLD = dr("QPedida", DataRowVersion.Original)
'        End If

'        If intIDPromocionLineaOLD = 0 Then
'            intIDPromocionLineaOLD = Nz(dr("IDPromocionLinea"), 0)
'        End If
'        'Se actualiza la cantidad promocionada
'        Dim dtPromLineaOLD As DataTable = pl.SelOnPrimaryKey(intIDPromocionLineaOLD)
'        If Not IsNothing(dtPromLineaOLD) AndAlso dtPromLineaOLD.Rows.Count > 0 Then
'            If dr("QPedida") < dtPromLineaOLD.Rows(0)("QMinPedido") Then
'                dr("IDPromocionLinea") = intIDPromocionLineaOLD
'                dr("IDPromocion") = strIDPromocionOLD
'                dr("QPedida") = dr("QPedida") - dblQPedidaOLD
'                Dim StDatos As New PromocionLinea.DatosActuaLinPromoDr
'                StDatos.Dr = dr
'                StDatos.Delete = False
'                ProcessServer.ExecuteTask(Of PromocionLinea.DatosActuaLinPromoDr)(AddressOf PromocionLinea.ActualizarLineaPromocion, StDatos, services)
'            Else
'                Dim dblQ As Double = dr("QPedida")
'                If dr("QPedida") > dtPromLineaOLD.Rows(0)("QMaxPedido") Then
'                    dr("IDPromocionLinea") = intIDPromocionLineaOLD
'                    dr("IDPromocion") = strIDPromocionOLD
'                    dblQ = dtPromLineaOLD.Rows(0)("QMaxPedido")
'                End If
'                dr("QPedida") = dr("QPedida") - dblQPedidaOLD
'                Dim StDatos As New PromocionLinea.DatosActuaLinPromoDr
'                StDatos.Dr = dr
'                StDatos.Delete = False
'                ProcessServer.ExecuteTask(Of PromocionLinea.DatosActuaLinPromoDr)(AddressOf PromocionLinea.ActualizarLineaPromocion, StDatos, services)

'                dr("QPedida") = dblQ
'                ADDLineaRegaloPedido(dr, False)
'            End If
'        End If

'        If Length(dr("IDPromocionLinea")) > 0 AndAlso dr("Estado") = enumpvlEstado.pvlCerrado Then
'            'Si ha cambiado el estado a cerrado hay que comprobar si
'            'existe alguna diferencia entre la cantidad pedida y la cantidad servida.
'            If dr("Estado", DataRowVersion.Original) & String.Empty <> dr("Estado") Then
'                If dr("QPedida") <> dr("QServida") Then
'                    'Hay que actualizar la cantidad promocionada
'                    Dim dtPromLinea As DataTable = pl.SelOnPrimaryKey(dr("IDPromocionLinea"))
'                    If Not IsNothing(dtPromLinea) AndAlso dtPromLinea.Rows.Count > 0 Then
'                        dtPromLinea.Rows(0)("QPromocionada") = dtPromLinea.Rows(0)("QPromocionada") - dr("QPedida") + dr("QServida")
'                        BusinessHelper.UpdateTable(dtPromLinea)
'                    End If
'                End If
'            End If
'        End If
'    End If
'End Sub

'#Region " REGALOS "

'    Private Sub ADDLineaRegaloPedido(ByVal drLinea As DataRow, _
'                                     Optional ByVal blnActulizarPromo As Boolean = True)

'        If Not IsNothing(drLinea) AndAlso Length(drLinea("IDPromocionLinea")) > 0 Then
'            Dim dblQPrev As Double = drLinea("QPedida")

'            Dim f As New Filter
'            f.Add(New NumberFilterItem("IDPromocionLinea", drLinea("IDPromocionLinea")))
'            f.Add(New NumberFilterItem("IDPedido", drLinea("IDPedido")))

'            Dim dtArticuloRegalo As DataTable = New BE.DataEngine().Filter("vNegPromocionArticulosRegaloPedido", f)
'            If Not IsNothing(dtArticuloRegalo) AndAlso dtArticuloRegalo.Rows.Count > 0 Then
'                Dim strAlmacenPred As String = New Parametro().AlmacenPredeterminado()
'                Dim strTipoLinea As String = ProcessServer.ExecuteTask(Of Object, String)(AddressOf TipoLinea.TipoLineaRegalo, Nothing, New ServiceProvider)

'                Dim intOrden As Integer
'                Dim dtPVL As DataTable = Me.Filter(f, "IDOrdenLinea DESC")
'                If Not IsNothing(dtPVL) AndAlso dtPVL.Rows.Count > 0 Then
'                    intOrden = dtPVL.Rows.Count
'                End If

'                Dim PVC As New PedidoVentaCabecera
'                Dim drPVC As DataRow = PVC.GetItemRow(drLinea("IDPedido"))

'                Dim context As New BusinessData
'                context("IDCliente") = drPVC("IDCliente")

'                For Each drArticuloRegalo As DataRow In dtArticuloRegalo.Rows
'                    'Nuevo registro
'                    Dim drPVL As DataRow = dtPVL.NewRow
'                    drPVL("IDLineaPedido") = AdminData.GetAutoNumeric
'                    drPVL("Estado") = enumpvlEstado.pvlPedido
'                    drPVL("IDTipoLinea") = strTipoLinea
'                    drPVL("IDPedido") = drPVC("IDPedido")
'                    drPVL("IDCentroGestion") = drPVC("IDCentroGestion")

'                    drPVL = ApplyBusinessRule("IDArticulo", drArticuloRegalo("IDArticuloRegalo"), drPVL, context)

'                    drPVL("FechaEntrega") = IIf(Length(drPVC("FechaEntrega")) > 0, drPVC("FechaEntrega"), drPVC("FechaPedido"))
'                    drPVL("Regalo") = True

'                    'En el campo Cantidad guardamos la Cantidad indicada con el ArticuloRegalo
'                    drPVL("QPedida") = Fix((dblQPrev / drArticuloRegalo("QPedida"))) * drArticuloRegalo("QRegalo")
'                    If drPVL("QPedida") = 0 Then
'                        drPVL("QPedida") = drArticuloRegalo("QRegalo")
'                    End If

'                    'Se incrementa el IDOrden para cada linea de regalo generada
'                    intOrden = intOrden + 1
'                    drPVL("IDOrdenLinea") = intOrden

'                    drPVL = ApplyBusinessRule("QPedida", drPVL("QPedida"), drPVL, context)
'                    drPVL("IDPromocion") = drLinea("IDPromocion")

'                    dtPVL.Rows.Add(drPVL)

'                    BusinessHelper.UpdateTable(dtPVL)
'                Next

'                If blnActulizarPromo Then
'                    'Actualización QPromocionada
'                    Dim PL As New PromocionLinea
'                    Dim drPromocionLinea As DataRow = PL.GetItemRow(drLinea("IDPromocionLinea"))
'                    drPromocionLinea("QPromocionada") = drPromocionLinea("QPromocionada") + dblQPrev
'                    BusinessHelper.UpdateTable(drPromocionLinea.Table)
'                End If
'            End If
'        End If
'    End Sub

'#End Region

#End Region

#Region " ActualizarPrograma "
'Public Sub ActualizarPrograma(ByVal lineasPedido As DataTable, Optional ByVal blnDelete As Boolean = False, Optional ByVal dtConfirmaciones As DataTable = Nothing)
'    If Not IsNothing(lineasPedido) AndAlso lineasPedido.Rows.Count Then

'        Dim ofLineaPrograma As New Filter
'        Dim dFechaConfirmacion As Date = cnMinDate
'        For Each lineaPedido As DataRow In lineasPedido.Rows
'            If Not IsNothing(dtConfirmaciones) AndAlso dtConfirmaciones.Rows.Count > 0 Then
'                '//Buscamos la fecha de confirmación en las líneas de programa a confirmar.
'                ofLineaPrograma.Clear()
'                ofLineaPrograma.Add(New NumberFilterItem("IDLineaPrograma", lineaPedido("IDLineaPrograma")))
'                Dim adrConfirmaciones() As DataRow = dtConfirmaciones.Select(ofLineaPrograma.Compose(New AdoFilterComposer))
'                If Not IsNothing(adrConfirmaciones) AndAlso Length(adrConfirmaciones(0)("FechaConfirmacionNew")) > 0 Then
'                    dFechaConfirmacion = adrConfirmaciones(0)("FechaConfirmacionNew")
'                Else
'                    dFechaConfirmacion = cnMinDate
'                End If
'            End If

'            ActualizarPrograma(lineaPedido, blnDelete, dFechaConfirmacion)
'        Next
'    End If
'End Sub

'Public Sub ActualizarPrograma(ByVal lineaPedido As DataRow, Optional ByVal blndelete As Boolean = False, Optional ByVal dFechaConfirmacion As Date = cnMinDate)
'    If Not lineaPedido Is Nothing Then
'        If Length(lineaPedido("IDLineaPrograma")) > 0 Then
'            Dim PL As New ProgramaLinea
'            Dim Programa As DataTable = PL.SelOnPrimaryKey(lineaPedido("IDLineaPrograma"))
'            If Not Programa Is Nothing AndAlso Programa.Rows.Count > 0 Then
'                If blndelete Then
'                    Dim DblConfir As Double = Programa.Rows(0)("QConfirmada") - lineaPedido("QPedida")
'                    Programa.Rows(0)("QConfirmada") = IIf(DblConfir > 0, True, False)
'                    Programa.Rows(0)("Confirmada") = False
'                    Programa.Rows(0)("FechaConfirmacion") = DBNull.Value
'                Else
'                    Dim dblQModificada As Integer
'                    If lineaPedido.RowState = DataRowState.Modified Then
'                        dblQModificada = lineaPedido("QPedida", DataRowVersion.Original)
'                    End If
'                    Programa.Rows(0)("QConfirmada") = Nz(Programa.Rows(0)("QConfirmada"), 0) + (lineaPedido("QPedida") - dblQModificada)
'                    Programa.Rows(0)("Confirmada") = True
'                    If dFechaConfirmacion <> cnMinDate Then
'                        Programa.Rows(0)("FechaConfirmacion") = dFechaConfirmacion
'                    Else
'                        Programa.Rows(0)("FechaConfirmacion") = Today
'                    End If

'                    Programa.Rows(0)("FechaEntrega") = lineaPedido("FechaEntrega")
'                End If
'                AdminData.SetData(Programa)
'            End If
'        End If
'    End If
'End Sub

#End Region

#End Region