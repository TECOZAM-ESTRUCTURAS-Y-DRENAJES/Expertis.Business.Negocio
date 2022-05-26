Imports Solmicro.Expertis.Business.Negocio.NegocioGeneral

Public Class _PedidoCompraLinea
    Public Const IDLineaPedido As String = "IDLineaPedido"
    Public Const IDPedido As String = "IDPedido"
    Public Const IDArticulo As String = "IDArticulo"
    Public Const DescArticulo As String = "DescArticulo"
    Public Const NSerie As String = "NSerie"
    Public Const RefProveedor As String = "RefProveedor"
    Public Const DescRefProveedor As String = "DescRefProveedor"
    Public Const FechaEntrega As String = "FechaEntrega"
    Public Const IDAlmacen As String = "IDAlmacen"
    Public Const IdOperario As String = "IdOperario"
    Public Const IDTipoIva As String = "IDTipoIva"
    Public Const CContable As String = "CContable"
    Public Const QPedida As String = "QPedida"
    Public Const QServida As String = "QServida"
    Public Const Precio As String = "Precio"
    Public Const PrecioA As String = "PrecioA"
    Public Const PrecioB As String = "PrecioB"
    Public Const UdValoracion As String = "UdValoracion"
    Public Const IDUdMedida As String = "IDUdMedida"
    Public Const IDUdInterna As String = "IDUdInterna"
    Public Const Estado As String = "Estado"
    Public Const Dto1 As String = "Dto1"
    Public Const Dto2 As String = "Dto2"
    Public Const Dto3 As String = "Dto3"
    Public Const Dto As String = "Dto"
    Public Const DtoProntoPago As String = "DtoProntoPago"
    Public Const Importe As String = "Importe"
    Public Const ImporteA As String = "ImporteA"
    Public Const ImporteB As String = "ImporteB"
    Public Const Marca As String = "Marca"
    Public Const IDTrabajo As String = "IDTrabajo"
    Public Const IDObra As String = "IDObra"
    Public Const IdOrdenLinea As String = "IdOrdenLinea"
    Public Const IdOferta As String = "IdOferta"
    Public Const IdLineaOferta As String = "IdLineaOferta"
    Public Const IDSolicitud As String = "IDSolicitud"
    Public Const IdLineaSolicitud As String = "IdLineaSolicitud"
    Public Const IdContrato As String = "IdContrato"
    Public Const IdLineaContrato As String = "IdLineaContrato"
    Public Const Factor As String = "Factor"
    Public Const QInterna As String = "QInterna"
    Public Const IDLineaMaterial As String = "IDLineaMaterial"
    Public Const IDLineaOfertaDetalle As String = "IDLineaOfertaDetalle"
    Public Const Texto As String = "Texto"
    Public Const TipoLineaCompra As String = "TipoLineaCompra"
    Public Const IDOrdenRuta As String = "IDOrdenRuta"
    Public Const IDLineaPadre As String = "IDLineaPadre"
    Public Const IDMntoOTPrev As String = "IDMntoOTPrev"
    Public Const IdPrograma As String = "IdPrograma"
    Public Const IdLineaPrograma As String = "IdLineaPrograma"
    Public Const FechaCreacionAudi As String = "FechaCreacionAudi"
    Public Const FechaModificacionAudi As String = "FechaModificacionAudi"
    Public Const UsuarioAudi As String = "UsuarioAudi"
    Public Const PedidoVentaOrigen As String = "PedidoVentaOrigen"
    Public Const IdLineaContratoSub As String = "IdLineaContratoSub"
    Public Const IDCentroGestion As String = "IDCentroGestion"
    Public Const IDActivoAImputar As String = "IDActivoAImputar"
    Public Const FechaEntregaModificado As String = "FechaEntregaModificado"
    Public Const QTiempo As String = "QTiempo"
    Public Const Inmovilizado As String = "Inmovilizado"
    Public Const SeguimientoTarifa As String = "SeguimientoTarifa"

End Class

Public Class LineasPedidoCompraEliminadas
    Public IDLineas As Hashtable

    Public Sub New()
        IDLineas = New Hashtable
    End Sub
End Class

Public Class PedidoCompraLinea
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper
    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private _PCC As _PedidoCompraCabecera
    Private _PCL As _PedidoCompraLinea

    Private Const cnEntidad As String = "tbPedidoCompraLinea"


#Region " Validate "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ProcesoCompra.DetailCommonUpdateRules)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarFechaEntregaObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarAlmacenObligatorio)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarUnidadMedida)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarCantidadLineaPedido)
    End Sub

#End Region
#Region " Delete"
    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf General.Comunes.BeginTransaction)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ValidarEstadoLineaDel)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf DeleteComponentes)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ActualizarSolicitudDelete)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ActualizarProgramaDelete)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ActualizarLineasObraDelete)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ActualizarOfertaDelete)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ActualizarOfertaComercialDetalle)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ActualizarOrdenRutaDelete)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ActualizarLineaOTDelete)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ActualizarPedidosMultiEmpresaDelete)
    End Sub
    <Task()> Public Shared Function NoHaSidoEliminada(ByVal data As DataRow, ByVal services As ServiceProvider) As Boolean
        Dim listaEliminados As LineasPedidoCompraEliminadas = services.GetService(Of LineasPedidoCompraEliminadas)()
        Dim haSidoEliminado As Boolean = listaEliminados.IDLineas.Contains(data("IDLineaPedido"))
        If haSidoEliminado Then services.GetService(Of DeleteProcessContext).Deleted = haSidoEliminado
        Return Not haSidoEliminado
    End Function


    <Task()> Public Shared Sub DeleteComponentes(ByVal data As DataRow, ByVal services As ServiceProvider)
        Select Case data("TipoLineaCompra")
            Case enumaclTipoLineaAlbaran.aclSubcontratacion
                Dim f As New Filter
                f.Add(New NumberFilterItem("IDLineaPadre", data("IDLineaPedido")))
                Dim PCL As New PedidoCompraLinea
                Dim dtComponentes As DataTable = PCL.Filter(f)
                For Each dr As DataRow In dtComponentes.Rows
                    PCL.DeleteRowCascade(dr, services)

                    Dim listaEliminados As LineasPedidoCompraEliminadas = services.GetService(Of LineasPedidoCompraEliminadas)()
                    listaEliminados.IDLineas.Add(dr("IDLineaPedido"), dr("IDLineaPedido"))
                Next
            Case enumaclTipoLineaAlbaran.aclComponente
                If Length(data("IDOrdenRuta")) > 0 Then
                    '//Si es de subcontratación le decimos que ya está borrado, para que el motor no vuelva a intentar a borrarla.
                    Dim listaEliminados As LineasPedidoCompraEliminadas = services.GetService(Of LineasPedidoCompraEliminadas)()
                    listaEliminados.IDLineas.Add(data("IDLineaPedido"), data("IDLineaPedido"))
                Else
                    If Length(data("IDLineaPadre")) > 0 Then
                        Dim dtExistePadre As DataTable = New PedidoCompraLinea().SelOnPrimaryKey(data("IDLineaPadre"))
                        If Not dtExistePadre Is Nothing AndAlso dtExistePadre.Rows.Count > 0 Then
                            ApplicationService.GenerateError("No se permite eliminar líneas de tipo Componente.")
                        End If
                    End If
                End If
        End Select
    End Sub
    <Task()> Public Shared Sub ValidarEstadoLineaDel(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data("Estado") <> enumpclEstado.pclpedido Then
            ApplicationService.GenerateError("No se puede eliminar la linea porque esta Servida, Parcialmente Servida o Cerrada.")
        End If

    End Sub
    <Task()> Public Shared Sub ActualizarSolicitudDelete(ByVal data As DataRow, ByVal services As ServiceProvider)
        'Se cambia el estado del albarán (tanto de la cabecera como de las líneas).
        If Length(data("IDLineaSolicitud")) > 0 Then
            data("QPedida") = 0
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoPedidoCompra.ActualizarSolicitudLinea, data, services)
        End If
    End Sub
    <Task()> Public Shared Sub ActualizarProgramaDelete(ByVal data As DataRow, ByVal services As ServiceProvider)
        'Se cambia el estado del albarán (tanto de la cabecera como de las líneas).
        If Length(data("IDLineaPrograma")) > 0 Then
            Dim datosActProg As New ProcesoPedidoCompra.DataActualizarProgramaLinea(data, True)
            ProcessServer.ExecuteTask(Of ProcesoPedidoCompra.DataActualizarProgramaLinea)(AddressOf ProcesoPedidoCompra.ActualizarProgramaLinea, datosActProg, services)
        End If
    End Sub
    <Task()> Public Shared Sub ActualizarLineasObraDelete(ByVal data As DataRow, ByVal services As ServiceProvider)
        'Se cambia el estado del albarán (tanto de la cabecera como de las líneas).
        If Length(data("IDLineaMaterial")) > 0 OrElse Length(data("IDTrabajo")) > 0 Then
            data("QInterna") = 0
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoPedidoCompra.ActualizarLineasObra, data, services)
        End If
    End Sub
    <Task()> Public Shared Sub ActualizarOfertaDelete(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDOferta")) > 0 Then
            data("QPedida") = 0
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoPedidoCompra.ActualizarOfertaCompraLinea, data, services)
        End If
    End Sub
    <Task()> Public Shared Sub ActualizarOfertaComercialDetalle(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDLineaOfertaDetalle")) > 0 Then
            Dim oferta As BusinessHelper = BusinessHelper.CreateBusinessObject("OfertaComercialDetalle")
            Dim detalle As DataTable = oferta.SelOnPrimaryKey(data("IDLineaOfertaDetalle"))
            If Not detalle Is Nothing AndAlso detalle.Rows.Count > 0 Then
                detalle.Rows(0)("EstadoCompra") = enumocdEstadoCompraVenta.ecvPendiente
                BusinessHelper.UpdateTable(detalle)
            End If
        End If
    End Sub
    <Task()> Public Shared Sub ActualizarOrdenRutaDelete(ByVal data As DataRow, ByVal services As ServiceProvider)
        'Se cambia el estado del albarán (tanto de la cabecera como de las líneas).
        If Length(data("IDOrdenRuta")) > 0 Then
            data("QPedida") = 0
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoPedidoCompra.ActualizarQEnviadaOrdenRuta, data, services)
        End If
    End Sub
    <Task()> Public Shared Sub ActualizarLineaOTDelete(ByVal data As DataRow, ByVal services As ServiceProvider)
        'Se cambia el estado del albarán (tanto de la cabecera como de las líneas).
        If Length(data("IDMntoOTPrev")) > 0 Then
            data("QPedida") = 0
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoPedidoCompra.ActualizarLineaMantenimiento, data, services)
        End If
    End Sub
    <Task()> Public Shared Sub ActualizarPedidosMultiEmpresaDelete(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim grp As New GRPPedidoVentaCompraLinea
        data("QPedida") = 0
        Dim control As DataTable = grp.TrazaPCLPrincipal(data("IDLineaPedido"))
        If Not control Is Nothing AndAlso control.Rows.Count > 0 Then
            If Length(control.Rows(0)("IDPVSecundaria")) > 0 Then
                Dim NPedidoVenta As String = control.Rows(0)("NPVSecundaria")
                Dim DescBaseDatos As String = New NegocioGeneral().GetDataBaseDescription(CType(control.Rows(0)("IDBDSecundaria"), Guid))
                Throw New Exception("No se puede eliminar el pedido de compra. Este pedido ha generado una línea en el pedido de venta Nº " & NPedidoVenta & " en la empresa " & Quoted(DescBaseDatos) & ".")
            Else
                Dim PVL As New PedidoVentaLinea
                For Each dr As DataRow In control.Rows
                    If Length(dr("IDLineaPvPrincipal")) > 0 AndAlso Length(dr("IDPVLinea")) > 0 Then
                        Dim dtPVL As DataTable = PVL.SelOnPrimaryKey(dr("IDLineaPvPrincipal"))
                        If Not dtPVL Is Nothing AndAlso dtPVL.Rows.Count > 0 Then dr.Delete() 'grp.Delete(dr("IDPVLinea"))
                    End If
                Next
                BusinessHelper.UpdateTable(control)
            End If
        End If
    End Sub
#End Region
#Region " Business Rules "

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim services As New ServiceProvider
        Dim oBRL As New SynonymousBusinessRules
        oBRL.AddSynonymous("QPedida", "Cantidad")

        '//BusinessRules - Genéricas del circuito de compra
        oBRL = ProcessServer.ExecuteTask(Of BusinessRules, BusinessRules)(AddressOf ProcesoCompra.DetailBusinessRulesLin, oBRL, services)

        ''//BusinessRules - Específicas PCL  
        oBRL("NSerie") = AddressOf CambioNSerie  '
        oBRL("Estado") = AddressOf CambioEstado
        oBRL("FechaEntrega") = AddressOf CambioFechaEntrega
        Return oBRL
    End Function
    <Task()> Public Shared Sub CambioNSerie(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)

        If Length(data.Value) > 0 Then
            Dim ns As DataTable = New ArticuloNSerie().Filter(New StringFilterItem("NSerie", data.Value))
            If ns.Rows.Count > 0 Then
                data.Current("IDArticulo") = ns.Rows(0)("IDArticulo")
                data.Current("QPedida") = 1
            Else
                ApplicationService.GenerateError("El número de serie no existe.")
            End If
        End If

    End Sub
    <Task()> Public Shared Sub CambioEstado(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)

        'Control de estados: excepto si el original es servido, se puede cambiar a cualquier estado (incluido pasar de cerrado a otro)
        If data.Current(data.ColumnName) = enumpclEstado.pclservido Then 'Or rs.Fields("Estado").OriginalValue = pclCerrado Then
            data.Value = data.Current(data.ColumnName)
        ElseIf data.Current(data.ColumnName) = enumpclEstado.pclservido And Nz(data.Current("QServida"), 0) < Nz(data.Current("QPedida"), 0) Then
            data.Value = data.Current(data.ColumnName)
        ElseIf data.Current(data.ColumnName) = enumpclEstado.pclparcservido And Nz(data.Current("QServida"), 0) = 0 Then
            data.Value = data.Current(data.ColumnName)
        End If

    End Sub

    <Task()> Public Shared Sub CambioFechaEntrega(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then data.Current("FechaEntregaModificado") = data.Value
    End Sub

#End Region
#Region " Seguimiento"
    <Task()> Public Shared Function SeguimientoPedidoVenta(ByVal IDLineaPedido As Integer, ByVal services As ServiceProvider) As DataTable
        Dim dt As DataTable = New BE.DataEngine().Filter("vConsultaPedidosCompraVentaLinea", New NumberFilterItem("IDLineaPCPrincipal", IDLineaPedido))
        dt.Columns.Add("EmpresaGrupo", GetType(Boolean))
        dt.Columns.Add("EntregaProveedor", GetType(Boolean))
        dt.Columns.Add("DescBaseDatosPrincipal", GetType(String))
        dt.Columns.Add("DescBaseDatosSecundaria", GetType(String))
        If dt.Rows.Count > 0 Then
            Dim pedido As DataRow = New PedidoCompraCabecera().GetItemRow(dt.Rows(0)("IDPCPrincipal"))
            dt.Rows(0)("EmpresaGrupo") = pedido("EmpresaGrupo")
            dt.Rows(0)("EntregaProveedor") = pedido("EntregaProveedor")

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
#Region " Actualizaciones "

    <Task()> Public Shared Sub ActualizarPedidosMultiEmpresa(ByVal IDPedido As Integer, ByVal services As ServiceProvider)
        Dim grp As New GRPPedidoVentaCompraLinea
        Dim control As DataTable = grp.TrazaPCPrincipal(IDPedido)
        If Not control Is Nothing AndAlso control.Rows.Count > 0 Then
            If Length(control.Rows(0)("IDPVSecundaria")) > 0 Then
                Dim NPedidoVenta As String = control.Rows(0)("NPVSecundaria")
                Dim DescBaseDatos As String = New NegocioGeneral().GetDataBaseDescription(CType(control.Rows(0)("IDBDSecundaria"), Guid))
                Throw New Exception("No se puede eliminar el pedido de compra. Este pedido ha generado una línea en el pedido de venta Nº " & NPedidoVenta & " en la empresa " & Quoted(DescBaseDatos) & ".")
            Else
                For Each dr As DataRow In control.Rows
                    grp.Delete(dr("IDPVLinea"))
                Next
                BusinessHelper.UpdateTable(control)
            End If
        End If
    End Sub

#End Region

 
    Private Function GrabarEspecificaciones(ByVal strIdOferta As String, ByVal intIDPedido As Integer) As Integer
        ' Comprobar Si esta grabada ya la especificacion
        Dim PCE As New PedidoCompraEspecificacion
        Dim dtPCEspec As DataTable = PCE.Filter(New NumberFilterItem("IDLineaPedido", intIDPedido))
        If dtPCEspec.Rows.Count = 0 Then
            ' Busco las especificaciones de la oferta
            Dim fwOfEspec As BusinessHelper = New BusinessHelper("OfertaEspecificacion")
            Dim dtOfEspec As DataTable = fwOfEspec.Filter(New StringFilterItem("IDOferta", strIdOferta))

            If Not dtOfEspec Is Nothing AndAlso dtOfEspec.Rows.Count > 0 Then
                ' Grabar las especificaciones del articulo
                dtPCEspec = PCE.AddNew
                For Each drOFEspec As DataRow In dtOfEspec.Rows
                    Dim drPCEspec As DataRow = dtPCEspec.NewRow()

                    drPCEspec("IDLineaPedido").Value = intIDPedido
                    drPCEspec("DescEspecificacion").Value = drOFEspec("DescEspecificacion")
                    drPCEspec("Valor").Value = drOFEspec("Valor")
                    dtPCEspec.ImportRow(drPCEspec)
                Next

                PCE.Update(dtPCEspec)
            End If
        End If
    End Function
    Public Sub CerrarLineas(ByVal dtPCL As DataTable)
        If Not dtPCL Is Nothing AndAlso dtPCL.Rows.Count > 0 Then
            Dim StrPed(dtPCL.Rows.Count - 1) As String
            Dim i As Integer = 0
            For Each Dr As DataRow In dtPCL.Select
                StrPed(i) = Dr("IDLineaPedido")
                i += 1
            Next
            Dim DtUpdate As DataTable = Me.Filter(New InListFilterItem("IDLineaPedido", StrPed, FilterType.Numeric))
            For Each DrUpd As DataRow In DtUpdate.Select
                DrUpd("Estado") = enumpclEstado.pclCerrado
            Next
            Me.Update(DtUpdate)
        End If
    End Sub

End Class

#Region " Código que ya no se utiliza ¿? "

'TODO: Este código de actulizaciones, parece que ya no se utilizan
'Private Sub ActualizarLineaMaterial(ByVal intIDLineaMaterial As Integer, ByVal QPedida As Double)
'    If intIDLineaMaterial > 0 Then

'        Dim dt As DataTable
'        Dim ovl As BusinessHelper
'        ovl = BusinessHelper.CreateBusinessObject("ObraMaterial")
'        dt = ovl.Filter(, "IDLineaMaterial= " & intIDLineaMaterial)

'        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
'            dt.Rows(0)("QPedida") = dt.Rows(0)("QPedida") - QPedida
'            BusinessHelper.UpdateTable(dt)
'        End If
'    End If
'End Sub

'Private Sub ActualizarLineaTrabajo(ByVal intIDTrabajo As Integer, ByVal QPedida As Double)
'    If intIDTrabajo > 0 Then

'        Dim dt As DataTable
'        Dim ovl As BusinessHelper
'        ovl = BusinessHelper.CreateBusinessObject("ObraTrabajo")
'        dt = ovl.Filter(, "IDTrabajo= " & intIDTrabajo)

'        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
'            dt.Rows(0)("QPedida") = dt.Rows(0)("QPedida") - QPedida

'            BusinessHelper.UpdateTable(dt)
'        End If
'    End If
'End Sub

'#Region " ActualizarLineaOT "

'    Public Sub ActualizarLineaOT(ByVal lineasPedido As DataTable, Optional ByVal blnDelete As Boolean = False)
'        If Not IsNothing(lineasPedido) AndAlso lineasPedido.Rows.Count Then
'            For Each lineaPedido As DataRow In lineasPedido.Rows
'                ActualizarLineaOT(lineaPedido, blnDelete)
'            Next
'        End If
'    End Sub

'    Private Sub ActualizarLineaOT(ByVal lineaPedido As DataRow, Optional ByVal blnDelete As Boolean = False)
'        If Not IsNothing(lineaPedido) AndAlso Length(lineaPedido("IDMntoOTPrev")) > 0 Then
'            Dim OTPrev As BusinessHelper
'            OTPrev = BusinessHelper.CreateBusinessObject("MntoOTPrevLinea")
'            Dim dtOT As DataTable = OTPrev.SelOnPrimaryKey(lineaPedido("IDMntoOTPrev"))
'            If Not dtOT Is Nothing AndAlso dtOT.Rows.Count > 0 Then
'                If blnDelete Then
'                    dtOT.Rows(0)("QPedida") = dtOT.Rows(0)("QPedida") - lineaPedido("QPedida")
'                Else
'                    Dim dblQModificada As Integer
'                    If lineaPedido.RowState = DataRowState.Modified Then
'                        dblQModificada = lineaPedido("QPedida", DataRowVersion.Original)
'                    End If
'                    dtOT.Rows(0)("QPedida") = Nz(dtOT.Rows(0)("QPedida"), 0) + (lineaPedido("QPedida") - dblQModificada)
'                End If
'                BusinessHelper.UpdateTable(dtOT)
'            End If
'        End If
'    End Sub

'#End Region

'    Private Sub ActualizarOfertaCompra(ByVal strIdOferta As String)
'        If Len(strIdOferta) > 0 Then
'            Dim OC As BusinessHelper = BusinessHelper.CreateBusinessObject("OfertaCabecera")
'            Dim dt As DataTable = OC.SelOnPrimaryKey(strIdOferta)
'            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
'                dt.Rows(0)("Estado") = enumOfertaCabecera.ocAdjudicada
'                dt.Rows(0)("IDPedido") = System.DBNull.Value
'                BusinessHelper.UpdateTable(dt)
'            End If
'        End If
'    End Sub

'    'Private Sub ActualizarOfertaComercialDetalle(ByVal intIDLineaOfertaDetalle As Integer)
'    '    If intIDLineaOfertaDetalle > 0 Then
'    '        Dim OC As BusinessHelper = BusinessHelper.CreateBusinessObject("OfertaComercialDetalle")
'    '        Dim dt As DataTable = OC.SelOnPrimaryKey(intIDLineaOfertaDetalle)
'    '        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
'    '            dt.Rows(0)("EstadoCompra") = False

'    '            AdminData.SetData(dt)
'    '        End If
'    '    End If
'    'End Sub

'    Private Function ActualizarOrdenRuta(ByVal lineaPedido As DataRow)
'        If IsNumeric(lineaPedido(_PCL.IDOrdenRuta)) Then
'            Dim IDOrdenRuta As Integer = lineaPedido(_PCL.IDOrdenRuta)
'            Dim QPedida As Double = lineaPedido(_PCL.QPedida)
'            Dim QPedidaOriginal As Double = lineaPedido(_PCL.QPedida, DataRowVersion.Original)
'            Dim QServida As Double = lineaPedido(_PCL.QServida)
'            Dim QServidaOriginal As Double = lineaPedido(_PCL.QServida, DataRowVersion.Original)
'            Dim Factor As Double = lineaPedido(_PCL.Factor)

'            If Factor <= 0 Then Factor = 1

'            Dim OC As BusinessHelper

'            OC = BusinessHelper.CreateBusinessObject(AdminData.GetEntityInfo("OrdenRuta"))
'            Dim dt As DataTable = OC.Filter(, "IDOrdenRuta= " & IDOrdenRuta)

'            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
'                Dim IncQPedida As Double = QPedida - QPedidaOriginal
'                Dim IncQServida As Double = QServida - QServidaOriginal

'                dt.Rows(0)("QEnviada") = dt.Rows(0)("QEnviada") + (Factor * (IncQPedida - IncQServida))
'                If dt.Rows(0)("QEnviada") < 0 Then dt.Rows(0)("QEnviada") = 0

'                BusinessHelper.UpdateTable(dt)
'            End If
'        End If
'    End Function

'#Region " ActualizarPrograma "

'    Public Sub ActualizarPrograma(ByVal lineaPedido As DataRow, Optional ByVal blnDelete As Boolean = False, Optional ByVal dFechaConfirmacion As Date = cnMinDate)
'        If Not IsNothing(lineaPedido) Then
'            If IsNumeric(lineaPedido("IDLineaPrograma")) Then
'                Dim pl As New ProgramaCompraLinea
'                Dim Programa As DataTable = pl.SelOnPrimaryKey(lineaPedido("IDLineaPrograma"))
'                If Not IsNothing(Programa) AndAlso Programa.Rows.Count Then
'                    If blnDelete Then
'                        Programa.Rows(0)("Confirmada") = False
'                        Programa.Rows(0)("QConfirmada") = Programa.Rows(0)("QConfirmada") - lineaPedido("QPedida")
'                        Programa.Rows(0)("FechaConfirmacion") = DBNull.Value
'                    Else
'                        Dim dblQModificada As Integer
'                        If lineaPedido.RowState = DataRowState.Modified Then
'                            dblQModificada = lineaPedido("QPedida", DataRowVersion.Original)
'                        End If
'                        Programa.Rows(0)("QConfirmada") = Nz(Programa.Rows(0)("QConfirmada"), 0) + (lineaPedido("QPedida") - dblQModificada)
'                        Programa.Rows(0)("Confirmada") = CBool(enumplEstadoLinea.plConfirmada)
'                        If dFechaConfirmacion <> cnMinDate Then
'                            Programa.Rows(0)("FechaConfirmacion") = dFechaConfirmacion
'                        Else
'                            Programa.Rows(0)("FechaConfirmacion") = Today
'                        End If
'                        Programa.Rows(0)("FechaEntrega") = lineaPedido("FechaEntrega")
'                    End If
'                    BusinessHelper.UpdateTable(Programa)
'                End If
'            End If
'        End If
'    End Sub

'    Public Sub ActualizarPrograma(ByVal lineasPedido As DataTable, Optional ByVal blnDelete As Boolean = False, Optional ByVal dtConfirmaciones As DataTable = Nothing)
'        If Not IsNothing(lineasPedido) AndAlso lineasPedido.Rows.Count Then
'            Dim ofLineaPrograma As New Filter
'            Dim dFechaConfirmacion As Date = cnMinDate
'            For Each lineaPedido As DataRow In lineasPedido.Rows
'                If Not IsNothing(dtConfirmaciones) AndAlso dtConfirmaciones.Rows.Count > 0 Then
'                    '//Buscamos la fecha de confirmación en las líneas de programa a confirmar.
'                    ofLineaPrograma.Clear()
'                    ofLineaPrograma.Add(New NumberFilterItem("IDLineaPrograma", lineaPedido("IDLineaPrograma")))
'                    Dim WhereLineaPrograma As String = ofLineaPrograma.Compose(New AdoFilterComposer)
'                    Dim adrConfirmaciones() As DataRow = dtConfirmaciones.Select(WhereLineaPrograma)
'                    If Not IsNothing(adrConfirmaciones) AndAlso Length(adrConfirmaciones(0)("FechaConfirmacionNew")) > 0 Then
'                        dFechaConfirmacion = adrConfirmaciones(0)("FechaConfirmacionNew")
'                    Else
'                        dFechaConfirmacion = cnMinDate
'                    End If
'                End If

'                ActualizarPrograma(lineaPedido, blnDelete, dFechaConfirmacion)
'            Next
'        End If
'    End Sub

'#End Region

'#Region " ActualizarOfertaCompra "

'    Public Sub ActualizarOfertaCompra(ByVal lineaPedido As DataRow, Optional ByVal blnDelete As Boolean = False)

'        If Not IsNothing(lineaPedido) Then
'            If Length(lineaPedido("IDOferta")) > 0 Then
'                Dim of1 As BusinessHelper
'                of1 = BusinessHelper.CreateBusinessObject("OfertaCabecera")
'                Dim dtOferta As DataTable = of1.SelOnPrimaryKey(lineaPedido("IDOferta"))
'                If Not IsNothing(dtOferta) AndAlso dtOferta.Rows.Count Then
'                    If blnDelete Then
'                        dtOferta.Rows(0)("Estado") = enumOfertaCabecera.ocAdjudicada
'                        dtOferta.Rows(0)("IDPedido") = System.DBNull.Value
'                    Else
'                        dtOferta.Rows(0)("Estado") = enumOfertaCabecera.ocCerrada
'                        dtOferta.Rows(0)("IDPedido") = lineaPedido("IDPedido")
'                    End If
'                    BusinessHelper.UpdateTable(dtOferta)
'                End If
'            End If
'        End If
'    End Sub

'    Public Sub ActualizarOfertaCompra(ByVal lineasPedido As DataTable)
'        If Not IsNothing(lineasPedido) AndAlso lineasPedido.Rows.Count Then
'            For Each lineaPedido As DataRow In lineasPedido.Rows
'                ActualizarOfertaCompra(lineaPedido)
'                If Length(lineaPedido("IDLineaSolicitud")) Then
'                    ActualizarSolicitudes(lineaPedido, False)
'                End If
'            Next
'        End If
'    End Sub

'#End Region

'#Region " ActualizarSolicitudes "

'    Public Sub ActualizarSolicitudes(ByVal lineaPedido As DataRow, Optional ByVal blnDelete As Boolean = False)
'        If Not IsNothing(lineaPedido) Then
'            If IsNumeric(lineaPedido("IDLineaSolicitud")) Then
'                Dim sl As BusinessHelper
'                sl = BusinessHelper.CreateBusinessObject("SolicitudCompraLinea")
'                Dim dtSolicitud As DataTable = sl.SelOnPrimaryKey(lineaPedido("IDLineaSolicitud"))
'                If Not IsNothing(dtSolicitud) AndAlso dtSolicitud.Rows.Count Then
'                    If blnDelete And Length(lineaPedido("IDOferta")) = 0 Then
'                        dtSolicitud.Rows(0)("FechaEstado") = Date.Today
'                        dtSolicitud.Rows(0)("QTramitada") = dtSolicitud.Rows(0)("QTramitada") - lineaPedido("QPedida")
'                        If dtSolicitud.Rows(0)("QTramitada") = 0 Then
'                            dtSolicitud.Rows(0)("Estado") = enumscEstado.scSolicitado
'                        End If
'                    ElseIf blnDelete And Length(lineaPedido("IDOferta")) <> 0 Then
'                        dtSolicitud.Rows(0)("Estado") = enumscEstado.scSolicitaOferta
'                    Else
'                        dtSolicitud.Rows(0)("FechaEstado") = Date.Today
'                        If Length(lineaPedido("IDOferta")) = 0 Then
'                            Dim dblQModificada As Integer
'                            If lineaPedido.RowState = DataRowState.Modified Then
'                                dblQModificada = lineaPedido("QPedida", DataRowVersion.Original)
'                            End If
'                            dtSolicitud.Rows(0)("QTramitada") = dtSolicitud.Rows(0)("QTramitada") + (lineaPedido("QPedida") - dblQModificada)
'                        End If
'                        dtSolicitud.Rows(0)("Estado") = enumscEstado.scPedido
'                    End If
'                    BusinessHelper.UpdateTable(dtSolicitud)
'                End If
'            End If
'        End If
'    End Sub

'    Public Sub ActualizarSolicitudes(ByVal lineasPedido As DataTable, Optional ByVal blnDelete As Boolean = False)
'        If Not IsNothing(lineasPedido) AndAlso lineasPedido.Rows.Count Then
'            For Each lineaPedido As DataRow In lineasPedido.Rows
'                ActualizarSolicitudes(lineaPedido, blnDelete)
'            Next
'        End If
'    End Sub

'#End Region

'Public Function Componentes(ByVal lineaPedido As DataRow) As DataTable
'    Dim newData As DataTable = Me.AddNew
'    Dim data As DataTable
'    Dim subcontratacionManual As Boolean

'    If IsNumeric(lineaPedido(_PCL.IDOrdenRuta)) Then
'        Dim f As New Filter
'        f.Add(New NumberFilterItem(_PCL.IDOrdenRuta, FilterOperator.Equal, lineaPedido(_PCL.IDOrdenRuta)))
'        Dim ordenRuta As DataTable
'        ordenRuta = New BE.DataEngine().Filter("tbOrdenRuta", f, "IDOrden,Secuencia")
'        If Not ordenRuta Is Nothing AndAlso ordenRuta.Rows.Count > 0 Then
'            If IsNumeric(ordenRuta.Rows(0)("Secuencia")) AndAlso ordenRuta.Rows(0)("Secuencia") <> 0 Then
'                'TODO. REvisar si hace falta esto
'                'data = ComponentesDeSubcontratacion(ordenRuta.Rows(0)("IDOrden"), lineaPedido(_PCL.IDArticulo), ordenRuta.Rows(0)("Secuencia"))
'            End If
'        End If
'    Else
'        subcontratacionManual = True
'        Dim f As New Filter
'        f.Add(New StringFilterItem(_PCL.IDArticulo, FilterOperator.Equal, lineaPedido(_PCL.IDArticulo)))
'        data = New BE.DataEngine().Filter("vNegArticuloCompPrimerNivelSubcontratacion", f)
'    End If

'    If Not data Is Nothing AndAlso data.Rows.Count > 0 Then
'        Dim UDS As New ArticuloUnidadAB
'        For Each componente As DataRow In data.Rows
'            Dim newrow As DataRow = newData.NewRow
'            '//copiar previamente los valores originales de la linea pedido padre
'            newrow.ItemArray = lineaPedido.ItemArray
'            newrow(_PCL.IDLineaPedido) = AdminData.GetAutoNumeric
'            newrow(_PCL.IDArticulo) = componente("IDComponente")
'            If Length(componente("DescComponente")) > 0 Then
'                newrow(_PCL.DescArticulo) = componente("DescComponente")
'            End If
'            If Length(componente("IDAlmacen")) > 0 Then
'                newrow(_PCL.IDAlmacen) = componente("IDAlmacen")
'            End If
'            If Length(componente("IdUdCompra")) > 0 Then
'                newrow(_PCL.IDUdMedida) = componente("IdUdCompra")
'            End If
'            newrow(_PCL.IDUdInterna) = componente("IDUDInterna")
'            newrow(_PCL.UdValoracion) = componente("UdValoracion")
'            newrow(_PCL.QInterna) = (Nz(lineaPedido(_PCL.QInterna), 0) * Nz(componente("Cantidad"), 0)) * (1 + (Nz(componente("Merma"), 0) / 100))
'            If subcontratacionManual Then
'                Dim StDatos As New ArticuloUnidadAB.DatosFactorConversion
'                StDatos.IDArticulo = newrow(_PCL.IDArticulo)
'                StDatos.IDUdMedidaA = newrow(_PCL.IDUdMedida)
'                StDatos.IDUdMedidaB = newrow(_PCL.IDUdInterna)
'                newrow(_PCL.Factor) = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, StDatos, New ServiceProvider)
'            Else
'                newrow(_PCL.Factor) = Nz((componente("FactorProduccion")), 0)
'                If Length(componente("IDCContable")) > 0 Then
'                    newrow("CContable") = componente("IDCContable")
'                End If
'            End If
'            If newrow(_PCL.Factor) = 0 Then
'                newrow(_PCL.Factor) = 1
'            End If
'            newrow(_PCL.QPedida) = newrow(_PCL.QInterna) / newrow(_PCL.Factor)
'            newrow(_PCL.Precio) = 0
'            newrow(_PCL.PrecioA) = 0
'            newrow(_PCL.PrecioB) = 0
'            newrow(_PCL.Importe) = 0
'            newrow(_PCL.ImporteA) = 0
'            newrow(_PCL.ImporteB) = 0
'            newrow(_PCL.TipoLineaCompra) = enumaclTipoLineaAlbaran.aclComponente
'            newrow(_PCL.IDLineaPadre) = lineaPedido(_PCL.IDLineaPedido)
'            newData.Rows.Add(newrow)
'        Next
'    End If

'    Return newData
'End Function
'Private Function ActualizarComponentes(ByVal lineaPedido As DataRow) As DataTable
'    Dim f As New Filter
'    f.Add(New NumberFilterItem(_PCL.IDPedido, FilterOperator.Equal, lineaPedido(_PCL.IDPedido)))
'    f.Add(New NumberFilterItem(_PCL.IDLineaPadre, FilterOperator.Equal, lineaPedido(_PCL.IDLineaPedido)))
'    f.Add(New NumberFilterItem(_PCL.TipoLineaCompra, FilterOperator.Equal, enumaclTipoLineaAlbaran.aclComponente))

'    Dim Componentes As DataTable = Me.Filter(f)
'    If Not Componentes Is Nothing AndAlso Componentes.Rows.Count > 0 Then
'        If Nz(lineaPedido(_PCL.QInterna, DataRowVersion.Original), 0) <> 0 Then
'            Dim factorVariacion As Double = lineaPedido(_PCL.QInterna) / lineaPedido(_PCL.QInterna, DataRowVersion.Original)
'            For Each componente As DataRow In Componentes.Rows
'                componente(_PCL.QPedida) = componente(_PCL.QPedida) * factorVariacion
'                componente(_PCL.QInterna) = componente(_PCL.QInterna) * factorVariacion
'            Next
'        End If
'    End If

'    BusinessHelper.UpdateTable(Componentes)
'    Return Componentes
'End Function

'Protected Overridable Function NuevaAnalitica(ByVal dr As DataRow) As DataTable
'    '//Provisional: En el proceso de generacion de pedidos de compra en la gestion de multiempresa
'    '//la llamada al metodo update de PedidoCompraLinea crea una analitica nueva partiendo de las 
'    '//lineas recien creadas. Se llama al update para aprovechar todo lo que hace, y eso en principio 
'    '//habria que mantenerlo.
'    '//En la gestion multiempresa la analitica del pedido de compra se crea a partir de las lineas
'    '//de pedido venta.
'    '//Para modificar este comportamiento se ha puesto overridable la parte que genera la analitica.

'    '//Esto es lo que hacia por defecto. 
'    'TODO: Pendiente de revisar esto
'    ' Return NegocioGeneral.NuevaAnalitica(dr)

'    '//En la clase de multiempresa se tendra una clase que hereda de PedidoCompraLinea y que 
'    '//sobreescribe esta funcion.
'End Function
#End Region
