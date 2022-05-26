Imports Solmicro.Expertis.Engine.BE.ApplicationService
Imports _ACC = Solmicro.Expertis.Business.Negocio._AlbaranCompraCabecera
Imports _ACL = Solmicro.Expertis.Business.Negocio._AlbaranCompraLinea
Imports _PCL = Solmicro.Expertis.Business.Negocio._PedidoCompraLinea
Imports _ACP = Solmicro.Expertis.Business.Negocio._AlbaranCompraPrecio
Imports _ACLT = Solmicro.Expertis.Business.Negocio._AlbaranCompraLote

Public Class _AlbaranCompraLinea
    Public Const IDLineaAlbaran As String = "IDLineaAlbaran"
    Public Const IDAlbaran As String = "IDAlbaran"
    Public Const IDLineaPedido As String = "IDLineaPedido"
    Public Const IDPedido As String = "IDPedido"
    Public Const IDArticulo As String = "IDArticulo"
    Public Const DescArticulo As String = "DescArticulo"
    Public Const RefProveedor As String = "RefProveedor"
    Public Const DescRefProveedor As String = "DescRefProveedor"
    Public Const IDAlmacen As String = "IDAlmacen"
    Public Const IDCondicionPago As String = "IDCondicionPago"
    Public Const IDFormaPago As String = "IDFormaPago"
    Public Const IDTipoIva As String = "IDTipoIva"
    Public Const Lote As String = "Lote"
    Public Const Ubicacion As String = "Ubicacion"
    Public Const UdValoracion As String = "UdValoracion"
    Public Const IDUdMedida As String = "IDUdMedida"
    Public Const IDUdInterna As String = "IDUdInterna"
    Public Const CContable As String = "CContable"
    Public Const QServida As String = "QServida"
    Public Const Precio As String = "Precio"
    Public Const PrecioA As String = "PrecioA"
    Public Const PrecioB As String = "PrecioB"
    Public Const Dto As String = "Dto"
    Public Const Dto1 As String = "Dto1"
    Public Const Dto2 As String = "Dto2"
    Public Const Dto3 As String = "Dto3"
    Public Const DtoProntoPago As String = "DtoProntoPago"
    Public Const Importe As String = "Importe"
    Public Const ImporteA As String = "ImporteA"
    Public Const ImporteB As String = "ImporteB"
    Public Const IDMovimiento As String = "IDMovimiento"
    Public Const EstadoFactura As String = "EstadoFactura"
    Public Const EstadoStock As String = "EstadoStock"
    Public Const IDTrabajo As String = "IDTrabajo"
    Public Const IDObra As String = "IDObra"
    Public Const TipoGastoObra As String = "TipoGastoObra"
    Public Const IDLineaPadre As String = "IDLineaPadre"
    Public Const IDConcepto As String = "IDConcepto"
    Public Const TipoLineaAlbaran As String = "TipoLineaAlbaran"
    Public Const IdOrdenLinea As String = "IdOrdenLinea"
    Public Const IdContrato As String = "IdContrato"
    Public Const IdLineaContrato As String = "IdLineaContrato"
    Public Const IDLineaOfertaDetalle As String = "IDLineaOfertaDetalle"
    Public Const Factor As String = "Factor"
    Public Const QInterna As String = "QInterna"
    Public Const Texto As String = "Texto"
    Public Const IDOrdenRuta As String = "IDOrdenRuta"
    Public Const IDOFControl As String = "IDOFControl"
    Public Const ControlCalidad As String = "ControlCalidad"
    Public Const IDRecepcion As String = "IDRecepcion"
    Public Const IDMntoOTPrev As String = "IDMntoOTPrev"
    Public Const FechaCreacionAudi As String = "FechaCreacionAudi"
    Public Const FechaModificacionAudi As String = "FechaModificacionAudi"
    Public Const UsuarioAudi As String = "UsuarioAudi"
    Public Const IdLineaContratoSub As String = "IdLineaContratoSub"
    Public Const MotivoRecepcionCalidad As String = "MotivoRecepcionCalidad"
    Public Const IDOperario As String = "IDOperario"
    Public Const IDEstadoActivo As String = "IDEstadoActivo"
    Public Const IdCentroGestion As String = "IdCentroGestion"
    Public Const IDActivoAImputar As String = "IDActivoAImputar"
    Public Const QFacturada As String = "QFacturada"
    Public Const FechaEntregaModificado As String = "FechaEntregaModificado"
    Public Const IDLineaAlbaranMltEmprs As String = "IDLineaAlbaranMltEmprs"
    Public Const QTiempo As String = "QTiempo"
    Public Const Inmovilizado As String = "Inmovilizado"
    Public Const SeguimientoTarifa As String = "SeguimientoTarifa"
End Class

'Public Interface IObraMaterialControl
'    Sub EliminarLineaControlMaterialDesdeAlbaranCompra(ByVal IDLineaAlbaran As Integer, ByVal services As ServiceProvider)
'End Interface

Public Class AlbaranCompraLineaInfo
    Inherits ClassEntityInfo

    Public IDAlbaran As Integer
    Public IDLineaAlbaran As Integer
    Public IDAlmacen As String
    Public GeneradoControl As Boolean

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal data As DataRow)
        MyBase.New(data)
    End Sub

    Public Overloads Overrides Sub Fill(ByVal ParamArray PrimaryKey() As Object)
        Dim dt As DataTable
        If Not IsNothing(PrimaryKey) AndAlso PrimaryKey.Length > 0 AndAlso Length(PrimaryKey(0)) > 0 Then
            dt = New AlbaranCompraLinea().SelOnPrimaryKey(PrimaryKey(0))
        End If

        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
            Me.Fill(dt.Rows(0))
        End If
    End Sub

End Class

Public Class LineasAlbaranCompraEliminadas
    Public IDLineas As Hashtable

    Public Sub New()
        IDLineas = New Hashtable
    End Sub
End Class

Public Class AlbaranCompraLinea
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Private Const cnEntidad As String = "tbAlbaranCompraLinea"

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

#Region " AddNewForm "

    'Public Overrides Function AddNewForm() As System.Data.DataTable
    '    AddNewForm = MyBase.AddNewForm()
    '    AddNewForm.Rows(0)(_ACL.IDLineaAlbaran) = AdminData.GetAutoNumeric
    'End Function

#End Region

#Region " Delete "

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf General.Comunes.BeginTransaction)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ValidarDelLineaAlbaranFacturada)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ValidarDelLineaAlbaranImputadoGasto)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ValidarDelNumerosSerieDisponibles)
        'deleteProcess.AddTask(Of DataRow)(AddressOf ValidarDelLineaAlbaranComponente)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ActualizarLineaPedido)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ActualizarLineaPrograma)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf EliminarComponentes)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf PrepararArticuloUltimaCompra)
        'deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ProcesoAlbaranCompra.PrepararArticuloUltimaCompra)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ActualizarProduccion)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ActualizarMovimientosStock)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ActualizarAlbaranesMultiEmpresaLinea)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf EliminarControlObras)
    End Sub

    <Task()> Public Shared Function NoHaSidoEliminada(ByVal data As DataRow, ByVal services As ServiceProvider) As Boolean
        Dim listaEliminados As LineasAlbaranCompraEliminadas = services.GetService(Of LineasAlbaranCompraEliminadas)()
        Dim haSidoEliminado As Boolean = listaEliminados.IDLineas.Contains(data("IDLineaAlbaran"))
        If haSidoEliminado Then services.GetService(Of DeleteProcessContext).Deleted = haSidoEliminado
        Return Not haSidoEliminado
    End Function

    <Task()> Public Shared Sub ValidarDelLineaAlbaranImputadoGasto(ByVal Linea As DataRow, ByVal services As ServiceProvider)
        Dim f As New Filter(FilterUnionOperator.Or)
        f.Add(New NumberFilterItem("IDLineaAlbaran", Linea("IDLineaAlbaran")))
        f.Add(New NumberFilterItem("IDLineaAlbaranHija", Linea("IDLineaAlbaran")))
        Dim dtACPrecio As DataTable = New AlbaranCompraPrecio().Filter(f)
        If dtACPrecio.Rows.Count > 0 Then
            ApplicationService.GenerateError("No se puede borrar la línea de Albarán. Está tiene líneas de gasto asociadas.")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarDelNumerosSerieDisponibles(ByVal Linea As DataRow, ByVal services As ServiceProvider)
        If Length(Linea("Lote")) > 0 Then
            Dim f As New Filter
            ''vCtlCiDisponibilidadNSerie
            f.Add(New StringFilterItem("NSerie", Linea("Lote")))
            f.Add(New StringFilterItem("IDArticulo", Linea("IDArticulo")))
            f.Add(New BooleanFilterItem("Disponible", False))
            Dim dtNserie As DataTable = New BE.DataEngine().Filter("vCtlCiDisponibilidadNSerie", f)
            If dtNserie.Rows.Count > 0 Then
                ApplicationService.GenerateError("No se puede borrar la línea de Albarán. El Número de Serie: | ya no está disponible.", Linea("Lote"))
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarDelLineaAlbaranFacturada(ByVal Linea As DataRow, ByVal services As ServiceProvider)
        If Linea("EstadoFactura") = enumaclEstadoFactura.aclParcFacturado OrElse _
           Linea("EstadoFactura") = enumaclEstadoFactura.aclFacturado Then
            ApplicationService.GenerateError("No se puede borrar la línea de Albarán. Está Facturada o Parcialmente Facturada.")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarDelLineaAlbaranComponente(ByVal Linea As DataRow, ByVal services As ServiceProvider)
        If IsNumeric(Linea(_ACL.IDLineaPadre)) And Linea(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclComponente And Length(Linea(_ACL.IDOrdenRuta)) <> 0 Then
            ApplicationService.GenerateError("No se permite eliminar líneas de tipo Componente.")
        End If
    End Sub

    <Task()> Public Shared Sub EliminarComponentes(ByVal Linea As System.Data.DataRow, ByVal services As ServiceProvider)
        'KITS Y SUBCONTR. que no vienen de producción
        If (Linea(_ACL.EstadoStock) = enumaclEstadoStock.aclSinGestion And Linea(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclKit) _
                     Or (Linea(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclSubcontratacion And IsDBNull(Linea(_ACL.IDOrdenRuta))) Then
            Dim ACL As New AlbaranCompraLinea
            Dim componentes As DataTable = ACL.Filter(New NumberFilterItem("IDLineaPadre", Linea("IDLineaAlbaran")))
            For Each drv As DataRow In componentes.Rows
                Dim drACL As DataRow = ACL.GetItemRow(drv("IDLineaALbaran"))
                ProcessServer.ExecuteTask(Of DataRow)(AddressOf EliminarMovimientoStock, drACL, services)
                If Length(drACL("IDLineaPedido")) > 0 Then
                    ProcessServer.ExecuteTask(Of DataRow)(AddressOf ActualizarLineaPedido, drACL, services)
                End If
                Dim listaEliminados As LineasAlbaranCompraEliminadas = services.GetService(Of LineasAlbaranCompraEliminadas)()
                ACL.DeleteRowCascade(drv, services)
                listaEliminados.IDLineas.Add(drv("IDLineaAlbaran"), drv("IDLineaAlbaran"))
            Next
        Else
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf DeleteComponentes, Linea, services)
        End If
    End Sub

    <Task()> Public Shared Sub EliminarMovimientoStock(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim updateData As StockUpdateData = ProcessServer.ExecuteTask(Of DataRow, StockUpdateData)(AddressOf EliminarMovimientoLineaAlbaran, data, services)
        If Not updateData Is Nothing Then
            If updateData.Estado <> EstadoStock.Actualizado Then
                If Not updateData.Log Is Nothing AndAlso Length(updateData.Log) > 0 Then Throw New Exception(updateData.Log)
            End If
        End If
    End Sub

    <Task()> Public Shared Function EliminarMovimientoLineaAlbaran(ByVal lineaAlbaran As DataRow, ByVal services As ServiceProvider) As StockUpdateData
        Dim updateData As StockUpdateData
        ProcessServer.ExecuteTask(Of Object)(AddressOf Comunes.BeginTransaction, Nothing, services)

        If IsNumeric(lineaAlbaran(_ACL.IDMovimiento)) Then
            '//Correccion movimiento de entrada
            Dim datElimMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Eliminar, lineaAlbaran(_ACL.IDMovimiento))
            updateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento, StockUpdateData)(AddressOf ProcesoStocks.ActualizarMovimiento, datElimMovto, services)
            If updateData Is Nothing OrElse updateData.Estado <> EstadoStock.Actualizado Then
                ProcessServer.ExecuteTask(Of Boolean)(AddressOf Comunes.RollbackTransaction, True, services)

                Return updateData
            End If
        End If

        ProcessServer.ExecuteTask(Of Boolean)(AddressOf Comunes.CommitTransaction, True, services)

        Return updateData
    End Function

    <Task()> Public Shared Sub DeleteComponentes(ByVal data As DataRow, ByVal services As ServiceProvider)
        Select Case data("TipoLineaAlbaran")
            Case enumaclTipoLineaAlbaran.aclSubcontratacion
                Dim f As New Filter
                f.Add(New NumberFilterItem("IDLineaPadre", data("IDLineaAlbaran")))
                Dim ACL As New AlbaranCompraLinea
                Dim dtComponentes As DataTable = ACL.Filter(f)
                For Each dr As DataRow In dtComponentes.Rows
                    ACL.DeleteRowCascade(dr, services)
                    Dim listaEliminados As LineasAlbaranCompraEliminadas = services.GetService(Of LineasAlbaranCompraEliminadas)()
                    listaEliminados.IDLineas.Add(dr("IDLineaAlbaran"), dr("IDLineaAlbaran"))
                Next
            Case enumaclTipoLineaAlbaran.aclComponente
                If Length(data("IDOrdenRuta")) > 0 Then
                    '//Si es de subcontratación le decimos que ya está borrado, para que el motor no vuelva a intentar a borrarla.
                    Dim dpc As DeleteProcessContext = services.GetService(Of DeleteProcessContext)()
                    dpc.Deleted = True
                Else
                    If Length(data("IDLineaPadre")) > 0 Then
                        Dim dtExistePadre As DataTable = New AlbaranCompraLinea().SelOnPrimaryKey(data("IDLineaPadre"))
                        If Not dtExistePadre Is Nothing AndAlso dtExistePadre.Rows.Count > 0 Then
                            ApplicationService.GenerateError("No se permite eliminar líneas de tipo Componente.")
                        End If
                    End If
                End If
        End Select
    End Sub


    <Task()> Public Shared Sub ActualizarLineaPedido(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDLineaPedido")) > 0 Then
            data("QServida") = 0
            data("QRechazada") = 0
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoAlbaranCompra.ActualizarLineaPedido, data, services)
            ProcessServer.ExecuteTask(Of Object)(AddressOf ProcesoAlbaranCompra.GrabarPedidos, Nothing, services)
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarLineaPrograma(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDLineaPedido")) > 0 Then
            Dim pl As New PedidoCompraLinea
            Dim DtPedido As DataTable = pl.SelOnPrimaryKey(data("IDLineaPedido"))
            For Each lineaPedido As DataRow In DtPedido.Select
                If Length(lineaPedido("IDPrograma")) > 0 AndAlso Length(lineaPedido("IDLineaPrograma")) > 0 Then
                    Dim datosActProg As New ProcesoAlbaranCompra.DataActualizarProgramaLinea(lineaPedido("IDLineaPrograma"), data, True)
                    ProcessServer.ExecuteTask(Of ProcesoAlbaranCompra.DataActualizarProgramaLinea)(AddressOf ProcesoAlbaranCompra.ActualizarProgramaLinea, datosActProg, services)
                End If
            Next
        End If
    End Sub

    <Task()> Public Shared Sub PrepararArticuloUltimaCompra(ByVal LineaAlbaran As DataRow, ByVal services As ServiceProvider)
        Dim rwArt As DataRow = New Articulo().GetItemRow(LineaAlbaran("IDArticulo"))
        Dim FilUltCompra As New Filter
        FilUltCompra.Add("IDArticulo", FilterOperator.Equal, rwArt("IDArticulo"))
        FilUltCompra.Add("IDAlbaran", FilterOperator.NotEqual, LineaAlbaran("IDAlbaran"))
        Dim dtAlbaUltimaFecha As DataTable = New BE.DataEngine().Filter("vAlbaranCompraFecha", FilUltCompra, "TOP 1 IDAlbaran,FechaAlbaran,TipoLineaAlbaran,IDOrdenRuta,IDProveedor,QInterna,ImporteA,ImporteB", "FechaAlbaran DESC, IDAlbaran DESC")
        If Not IsNothing(dtAlbaUltimaFecha) AndAlso dtAlbaUltimaFecha.Rows.Count Then
            If dtAlbaUltimaFecha.Rows(0)("TipoLineaAlbaran") = enumaclTipoLineaAlbaran.aclNormal OrElse _
              (dtAlbaUltimaFecha.Rows(0)("TipoLineaAlbaran") = enumaclTipoLineaAlbaran.aclSubcontratacion AndAlso _
               Length(dtAlbaUltimaFecha.Rows(0)("IDOrdenRuta")) > 0) OrElse _
               dtAlbaUltimaFecha.Rows(0)("TipoLineaAlbaran") = enumaclTipoLineaAlbaran.aclSubcontratacion OrElse _
               dtAlbaUltimaFecha.Rows(0)("TipoLineaAlbaran") = enumaclTipoLineaAlbaran.aclKit Then
                If dtAlbaUltimaFecha.Rows(0)("QInterna") > 0 Then
                    rwArt("FechaUltimaCompra") = dtAlbaUltimaFecha.Rows(0)("FechaAlbaran")
                    rwArt("IdProveedorUltimaCompra") = dtAlbaUltimaFecha.Rows(0)("IDProveedor")
                    rwArt("PrecioUltimaCompraA") = dtAlbaUltimaFecha.Rows(0)("ImporteA") / dtAlbaUltimaFecha.Rows(0)("QInterna")
                    rwArt("PrecioUltimaCompraB") = dtAlbaUltimaFecha.Rows(0)("ImporteB") / dtAlbaUltimaFecha.Rows(0)("QInterna")
                    BE.BusinessHelper.UpdateTable(rwArt.Table)
                End If
            End If
        Else
            rwArt("FechaUltimaCompra") = DBNull.Value
            rwArt("IdProveedorUltimaCompra") = String.Empty
            rwArt("PrecioUltimaCompraA") = 0
            rwArt("PrecioUltimaCompraB") = 0
            BE.BusinessHelper.UpdateTable(rwArt.Table)
        End If
    End Sub

    <Task()> Public Shared Sub EliminarControlObras(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDObra")) > 0 Then
            Dim GeneradoControl As Boolean = ProcessServer.ExecuteTask(Of DataRow, Boolean)(AddressOf ActualizacionControlObras.AlbaranGeneradoControl, data, services)
            If GeneradoControl Then
                Dim dataDelete As New ActualizacionControlObras.dataDeleteControlObras(data, ActualizacionControlObras.enumOrigen.Albaran)
                Select Case CType(data("TipoGastoObra"), enumfclTipoGastoObra)
                    Case enumfclTipoGastoObra.enumfclMaterial
                        ProcessServer.ExecuteTask(Of ActualizacionControlObras.dataDeleteControlObras)(AddressOf ActualizacionControlObras.DeleteObraMaterialControl, dataDelete, services)
                    Case enumfclTipoGastoObra.enumfclGastos
                        ProcessServer.ExecuteTask(Of ActualizacionControlObras.dataDeleteControlObras)(AddressOf ActualizacionControlObras.DeleteObraGastoControl, dataDelete, services)
                    Case enumfclTipoGastoObra.enumfclVarios
                        ProcessServer.ExecuteTask(Of ActualizacionControlObras.dataDeleteControlObras)(AddressOf ActualizacionControlObras.DeleteObraVariosControl, dataDelete, services)
                End Select
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarProduccion(ByVal linea As DataRow, ByVal services As ServiceProvider)
        If linea(_ACL.EstadoStock) = enumaclEstadoStock.aclActualizado Then
            'Si se trata de una subcontratacion de produccion
            If linea(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclSubcontratacion And IsNumeric(linea(_ACL.IDOrdenRuta)) Then
                ProcessServer.ExecuteTask(Of DataRow)(AddressOf EliminarComponentesProduccion, linea, services)
                'Else
                '    Dim updateData As StockUpdateData = ProcessServer.ExecuteTask(Of DataRow, StockUpdateData)(AddressOf EliminarMovimiento, linea, services)
                '    If Not updateData Is Nothing Then
                '        If updateData.Estado <> EstadoStock.Actualizado AndAlso Not IsNothing(updateData.Log) Then
                '            'Se deja porque bajaUpdateData.Detalle ya viene traduccido
                '            Throw New Exception(updateData.Log)
                '        End If
                '    End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarMovimientosStock(ByVal linea As DataRow, ByVal services As ServiceProvider)
        If linea(_ACL.EstadoStock) = enumaclEstadoStock.aclActualizado Then
            'Si se trata de una subcontratacion de produccion no hay movimiento
            If (linea(_ACL.TipoLineaAlbaran) <> enumaclTipoLineaAlbaran.aclSubcontratacion) Or _
            (linea(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclSubcontratacion And Nz(linea(_ACL.IDOrdenRuta), 0) = 0) Then
                Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
                If AppParams.GestionInventarioPermanente Then
                    If linea("Contabilizado") <> CInt(enumContabilizado.NoContabilizado) Then
                        ApplicationService.GenerateError("Debe descontabilizar primero la/s línea/s para poder eliminar el Movimiento.")
                    End If
                End If
                Dim updateData As StockUpdateData = ProcessServer.ExecuteTask(Of DataRow, StockUpdateData)(AddressOf EliminarMovimiento, linea, services)
                If Not updateData Is Nothing Then
                    If updateData.Estado <> EstadoStock.Actualizado AndAlso Not IsNothing(updateData.Log) Then
                        'Se deja porque bajaUpdateData.Detalle ya viene traduccido
                        Throw New Exception(updateData.Log)
                    End If
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub EliminarComponentesProduccion(ByVal dtrSource As System.Data.DataRow, ByVal services As ServiceProvider)
        'SUBCONTR. que  vienen de producción
        If Nz(dtrSource(_ACL.IDOFControl), 0) <> 0 Then
            Dim ofc As BusinessHelper = BusinessHelper.CreateBusinessObject("OFControl")
            CType(ofc, IControlProduccion).EliminarOFControl(dtrSource(_ACL.IDOFControl))
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarAlbaranesMultiEmpresaLinea(ByVal linea As DataRow, ByVal services As ServiceProvider)
        '//Control albaranes multiempresa
        Dim grp As New GRPAlbaranVentaCompraLinea
        Dim control As DataTable = grp.TrazaAVLPrincipal(linea("IDLineaALbaran"))
        If Not control Is Nothing AndAlso control.Rows.Count > 0 Then
            control.Rows(0)("IDAVPrincipal") = DBNull.Value
            control.Rows(0)("NAVPrincipal") = DBNull.Value
            control.Rows(0)("IDLineaAVPrincipal") = DBNull.Value
            BusinessHelper.UpdateTable(control)
        Else
            control = grp.TrazaAVLSecundaria(linea("IDLineaALbaran"))
            If Not control Is Nothing AndAlso control.Rows.Count > 0 Then
                Dim msg As String = "Este albarán está relacionado con un pedido entre empresas del grupo."
                msg = String.Concat(msg, ControlChars.NewLine, "Deberá eliminar el albarán completo.")
                Throw New Exception(msg)
            End If
        End If
    End Sub

#End Region

#Region " Validate "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ProcesoCompra.DetailCommonUpdateRules)    'Validaciones Generales Compras 
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarAlbaranFacturado)
        'validateProcess.AddTask(Of DataRow)(AddressOf ProcesoAlbaranVenta.ValidarEstadoLinea)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarCantidadLineaAlbaran)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarAlmacenObligatorio)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarFormaPagoObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarCondicionPagoObligatoria)
    End Sub

    <Task()> Public Shared Sub ValidarAlbaranFacturado(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDAlbaran")) <> 0 Then
            Dim Cabecera As DataTable = New AlbaranCompraCabecera().SelOnPrimaryKey(data("IDAlbaran"))
            If Not Cabecera Is Nothing AndAlso Cabecera.Rows.Count > 0 Then
                If Cabecera.Rows(0)("Estado") = enumavcEstadoFactura.avcFacturado Then
                    ApplicationService.GenerateError("El Albarán está Facturado.")
                End If
            End If
        End If
    End Sub

#End Region

#Region " Business Rules "

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim services As New ServiceProvider
        'TODO: Para qué sirve esto?
        '//?
        'If Not Context Is Nothing Then
        '    If Not Context.Contains("NSerieObligatorio") Then
        '        If Length(current("IDArticulo")) > 0 Then
        '            Dim articulo As DataRow = New Articulo().GetItemRow(current("IDArticulo"))
        '            Context("NSerieObligatorio") = articulo("NSerieObligatorio")
        '        Else
        '            Context("NSerieObligatorio") = False
        '        End If
        '    End If
        'End If

        'If Context("NSerieObligatorio") Then
        '    If (ColumnName = "QServida" Or ColumnName = "QInterna" Or ColumnName = "Factor") And Length(current("Lote")) > 0 Then
        '        If (Nz(Value, 0) <> 1 AndAlso Nz(Value, 0) <> -1) Then
        '            ApplicationService.GenerateError("La cantidad interna debe ser la unidad para un artículo con número de serie.")
        '        End If
        '    End If
        'End If
        '//?

        Dim oBRL As New SynonymousBusinessRules
        oBRL.AddSynonymous("QServida", "Cantidad")

        '//BusinessRules - Genéricas del circuito de comercial
        oBRL = ProcessServer.ExecuteTask(Of BusinessRules, BusinessRules)(AddressOf ProcesoCompra.DetailBusinessRulesLin, oBRL, services)


        ''//BusinessRules - Específicas AVL  
        oBRL("IDArticulo") = AddressOf CambioArticuloAlbaran  'Específica ACL
        oBRL.Add("TipoGastoObra", AddressOf CambioTipoGastoObra)
        oBRL.Add("IDConcepto", AddressOf CambioConcepto)
        oBRL.Add("IdActivoAImputar", AddressOf CambioActivoAImputar)
        oBRL.Add("Lote", AddressOf CambioLote)
        oBRL.Add("IDCondicionPago", AddressOf ProcesoComunes.CambioCondicionPagoLineas)

        '   oBRL.Add("NObra", AddressOf CambioNObra)
        Return oBRL
    End Function

    <Task()> Public Shared Sub CambioArticuloAlbaran(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ProcesoCompra.CambioArticulo, data, services)
        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
        Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.Current("IDArticulo"))
        If ArtInfo.GestionStock Then
            data.Current("EstadoStock") = CInt(enumaclEstadoStock.aclNoActualizado)
        Else
            data.Current("EstadoStock") = CInt(enumaclEstadoStock.aclSinGestion)
        End If
    End Sub

    <Task()> Public Shared Sub CambioTipoGastoObra(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("TipoGastoObra")) = 0 Then data.Current("TipoGastoObra") = enumfclTipoGastoObra.enumfclMaterial
        data.Current("IDLineaPadre") = System.DBNull.Value
        If data.Current("TipoGastoObra") <> enumfclTipoGastoObra.enumfclMaterial Then
            data.Current("IDConcepto") = System.DBNull.Value
        Else
            data.Current("IDConcepto") = data.Current("IdArticulo")
        End If
    End Sub

    <Task()> Public Shared Sub CambioConcepto(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDConcepto")) > 0 Then
            Dim f As New Filter
            Dim strFrom, strEntidad As String
            If data.Current("TipoGastoObra") = enumfclTipoGastoObra.enumfclGastos Then
                f.Add(New StringFilterItem("IDGasto", data.Current("IDConcepto")))
                strFrom = "tbMaestroGasto"
                strEntidad = "Gastos"
            ElseIf data.Current("TipoGastoObra") = enumfclTipoGastoObra.enumfclVarios Then
                f.Add(New StringFilterItem("IDVarios", data.Current("IDConcepto")))
                strFrom = "tbMaestroVarios"
                strEntidad = "Varios"
            End If
            If Len(strFrom) > 0 Then
                Dim dt As DataTable = New BE.DataEngine().Filter(strFrom, f)
                If Not dt Is Nothing AndAlso dt.Rows.Count = 0 Then
                    ApplicationService.GenerateError("El Cod. Concepto introducido no existe en |.", Quoted(strEntidad))
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioActivoAImputar(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDActivoAImputar")) > 0 Then
            Dim dt As DataTable = New Activo().SelOnPrimaryKey(data.Current("IDActivoAImputar"))
            If Not dt Is Nothing AndAlso dt.Rows.Count = 0 Then
                ApplicationService.GenerateError("El Activo no existe.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioLote(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("Lote")) > 0 And data.Context("NSerieObligatorio") Then
            data.Current("IDOperario") = ProcessServer.ExecuteTask(Of Object, String)(AddressOf Operario.ObtenerIDOperarioUsuario, Nothing, services)
            data.Current("IDEstadoActivo") = NegocioGeneral.ESTADOACTIVO_DISPONIBLE
            Dim serie As DataTable = New ArticuloNSerie().Filter(New StringFilterItem("NSerie", data.Current("Lote")))
            If serie.Rows.Count > 0 Then
                Dim p As New Parametro
                Dim blnActivo As Boolean = p.GestionNumeroSerieConActivos()
                'Si llevo gestion de activos correlativa a numeros de serie, el numero de serie no se puede repetir para diferentes artículos.
                If blnActivo AndAlso AreDifferents(data.Current("IDArticulo"), serie.Rows(0)("IDArticulo")) Then
                    'Me.ApplyBusinessRule("IDArticulo", serie.Rows(0)("IDArticulo"), data.Current, data.Context)
                    Dim brd As New BusinessRuleData("IDArticulo", serie.Rows(0)("IDArticulo"), data.Current, data.Context)
                    CambioArticuloAlbaran(brd, services)
                End If
            End If
            data.Current("Factor") = 1
        End If
    End Sub

    <Task()> Public Shared Sub CambioNObra(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) = 0 Then
            data.Current("IDObra") = DBNull.Value
        End If
    End Sub

#End Region

    'TODO: STOCKS
#Region " Gestión de stocks "

    'Public Function CorregirMovimiento(ByVal LineasAlbaran As DataTable) As StockUpdateData()
    '    Dim updateData(-1) As StockUpdateData
    '    Me.BeginTx()
    '    For Each lineaAlbaran As DataRow In LineasAlbaran.Rows
    '        Dim data As StockUpdateData = Me.CorregirMovimiento(lineaAlbaran)
    '        If Not data Is Nothing Then
    '            ReDim Preserve updateData(UBound(updateData) + 1)
    '            updateData(UBound(updateData)) = data
    '        End If
    '    Next
    '    Me.CommitTx()
    '    Return updateData
    'End Function

    'Public Function CorregirMovimiento(ByVal lineaAlbaran As DataRow) As StockUpdateData
    '    '//Lineas de albaran de tipo subcontratacion se actualizan desde el control de la produccion.
    '    If Nz(lineaAlbaran(_ACL.IDOrdenRuta), 0) = 0 Then
    '        Dim updateData As StockUpdateData
    '        Dim IDLineaMovimiento As Integer
    '        Dim Cantidad As Double
    '        Dim PrecioA As Double
    '        Dim PrecioB As Double

    '        '//Importes extras
    '        Dim ImporteExtraA As Double
    '        Dim ImporteExtraB As Double
    '        Dim Importes As DataTable = New AlbaranCompraPrecio().Filter(New NumberFilterItem(_ACP.IDLineaAlbaran, lineaAlbaran(_ACL.IDLineaAlbaran)))
    '        For Each importe As DataRow In Importes.Rows
    '            ImporteExtraA += importe(_ACP.ImporteA)
    '            ImporteExtraB += importe(_ACP.ImporteB)
    '        Next

    '        Cantidad = lineaAlbaran(_ACL.QInterna)
    '        If Cantidad <> 0 Then
    '            Dim m As New Moneda
    '            Dim monedaA As MonedaInfo = m.MonedaA
    '            Dim monedaB As MonedaInfo = m.MonedaB
    '            PrecioA = xRound(ImporteExtraA + (lineaAlbaran(_ACL.PrecioA) / lineaAlbaran(_ACL.Factor) / lineaAlbaran(_ACL.UdValoracion) * (1 - lineaAlbaran(_ACL.Dto1) / 100) * (1 - lineaAlbaran(_ACL.Dto2) / 100) * (1 - lineaAlbaran(_ACL.Dto3) / 100) * (1 - lineaAlbaran(_ACL.Dto) / 100) * (1 - lineaAlbaran(_ACL.DtoProntoPago) / 100)), monedaA.NDecimalesPrecio)
    '            PrecioB = xRound(ImporteExtraB + (lineaAlbaran(_ACL.PrecioB) / lineaAlbaran(_ACL.Factor) / lineaAlbaran(_ACL.UdValoracion) * (1 - lineaAlbaran(_ACL.Dto1) / 100) * (1 - lineaAlbaran(_ACL.Dto2) / 100) * (1 - lineaAlbaran(_ACL.Dto3) / 100) * (1 - lineaAlbaran(_ACL.Dto) / 100) * (1 - lineaAlbaran(_ACL.DtoProntoPago) / 100)), monedaB.NDecimalesPrecio)
    '        End If

    '        Me.BeginTx()
    '        Dim stockobj As New Stock
    '        Dim lote As DataTable
    '        Dim f As New Filter
    '        f.Add(New NumberFilterItem(_ACLT.IDLineaAlbaran, lineaAlbaran(_ACL.IDLineaAlbaran)))
    '        lote = New AlbaranCompraLote().Filter(f)
    '        If lote.Rows.Count > 0 Then
    '            For Each dr As DataRow In lote.Rows
    '                If Not dr.IsNull(_ACLT.IDMovimientoEntrada) Then
    '                    '//Correccion movimiento de entrada
    '                    IDLineaMovimiento = dr(_ACLT.IDMovimientoEntrada)
    '                    updateData = stockobj.CorregirMovimiento(IDLineaMovimiento, PrecioA, PrecioB)
    '                    If updateData Is Nothing OrElse updateData.Estado <> EstadoStock.Actualizado Then
    '                        Me.RollbackTx()
    '                        Return updateData
    '                    End If
    '                End If
    '            Next
    '        Else
    '            If Not lineaAlbaran.IsNull(_ACL.IDMovimiento) Then
    '                IDLineaMovimiento = lineaAlbaran(_ACL.IDMovimiento)
    '                updateData = stockobj.CorregirMovimiento(IDLineaMovimiento, Cantidad, PrecioA, PrecioB)
    '                If updateData Is Nothing OrElse updateData.Estado <> EstadoStock.Actualizado Then
    '                    Me.RollbackTx()
    '                    Return updateData
    '                End If
    '            End If
    '        End If
    '        Me.CommitTx()

    '        Return updateData
    '    End If
    'End Function

    '<Task()> Public Shared Function EliminarMovimiento(ByVal LineasAlbaran As DataTable, ByVal services As ServiceProvider) As StockUpdateData()
    '    Dim updateData(-1) As StockUpdateData
    '    Me.BeginTx()
    '    For Each lineaAlbaran As DataRow In LineasAlbaran.Rows
    '        Dim data As StockUpdateData = EliminarMovimiento(lineaAlbaran)
    '        If Not data Is Nothing Then
    '            ReDim Preserve updateData(UBound(updateData) + 1)
    '            updateData(UBound(updateData)) = data
    '        End If
    '    Next
    '    Me.CommitTx()
    '    Return updateData
    'End Function

    <Task()> Public Shared Function EliminarMovimiento(ByVal lineaAlbaran As DataRow, ByVal services As ServiceProvider) As StockUpdateData
        If IsNumeric(lineaAlbaran(_ACL.IDMovimiento)) Then
            AdminData.BeginTx()
            '//Correccion movimiento de entrada
            Dim datElimMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Eliminar, lineaAlbaran(_ACL.IDMovimiento))
            Dim updateData As StockUpdateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento, StockUpdateData)(AddressOf ProcesoStocks.ActualizarMovimiento, datElimMovto, services)
            If updateData Is Nothing OrElse updateData.Estado <> EstadoStock.Actualizado Then
                AdminData.RollBackTx(True)
                Return updateData
            End If

            AdminData.CommitTx(True)
            Return updateData
        End If
    End Function

    'Public Function ActualizarStock(ByVal LineasAlbaran As DataTable) As StockUpdateData()
    '    Dim lineas(-1) As DataRow
    '    For Each dr As DataRow In LineasAlbaran.Rows
    '        ArrayManager.Copy(dr, lineas)
    '    Next
    '    Return Me.ActualizarStock(lineas)
    'End Function

    'Public Function ActualizarStock(ByVal LineasAlbaran As DataRow()) As StockUpdateData()
    '    Dim updateDataArray(-1) As StockUpdateData
    '    Dim acc As New AlbaranCompraCabecera
    '    Dim aclt As New AlbaranCompraLote
    '    Dim ofc As BusinessHelper = BusinessHelper.CreateBusinessObject("OFControl")
    '    Dim [or] As BusinessHelper = BusinessHelper.CreateBusinessObject("OrdenRuta")
    '    Dim oe As BusinessHelper = BusinessHelper.CreateBusinessObject("OrdenEstructura")

    '    Dim OperarioGenerico As String = New Parametro().OperarioGenerico()

    '    For Each dr As DataRow In LineasAlbaran
    '        Me.BeginTx()
    '        Dim Linea As DataTable = Me.SelOnPrimaryKey(dr(_ACL.IDLineaAlbaran))
    '        If Linea.Rows.Count > 0 Then
    '            Dim lineaAlbaran As DataRow = Linea.Rows(0)
    '            If lineaAlbaran(_ACL.EstadoStock) = enumaclEstadoStock.aclNoActualizado Then
    '                If lineaAlbaran(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclNormal _
    '                Or lineaAlbaran(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclKit _
    '                Or lineaAlbaran(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclRealquiler _
    '                Or (lineaAlbaran(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclSubcontratacion And IsDBNull(lineaAlbaran(_ACL.IDOrdenRuta))) _
    '                Or (lineaAlbaran(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclComponente And IsDBNull(lineaAlbaran(_ACL.IDOrdenRuta))) Then
    '                    '//Linea NORMAL, o de tipo KIT, o SUBCONTRATACION MANUAL(que NO proviene de una OF)

    '                    Dim Lotes(-1) As DataTable
    '                    Dim Cabecera As DataTable = acc.SelOnPrimaryKey(lineaAlbaran(_ACL.IDAlbaran))
    '                    If Cabecera.Rows.Count > 0 Then
    '                        Dim updateData() As StockUpdateData
    '                        Dim lote As DataTable
    '                        Dim f As New Filter
    '                        f.Add(New NumberFilterItem(_ACLT.IDLineaAlbaran, lineaAlbaran(_ACL.IDLineaAlbaran)))
    '                        lote = aclt.Filter(f)
    '                        If lote.Rows.Count > 0 Then
    '                            updateData = Me.ActualizarStock(Cabecera.Rows(0), lineaAlbaran, lote)
    '                            ArrayManager.Copy(lote, Lotes)
    '                        Else
    '                            updateData = Me.ActualizarStock(Cabecera.Rows(0), lineaAlbaran)
    '                        End If
    '                        ArrayManager.Copy(updateData, updateDataArray)

    '                        AdminData.SetData(Cabecera)
    '                        AdminData.SetData(Linea)
    '                        AdminData.SetData(Lotes)
    '                    End If

    '                ElseIf lineaAlbaran(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclSubcontratacion And Nz(lineaAlbaran(_ACL.IDOrdenRuta), 0) <> 0 Then
    '                    '//Lineas de SUBCONTRATACION QUE PROVIENEN DE UNA ORDEN DE FABRICACION.
    '                    '//Las lineas de albaran se actualizaran automaticamente, independientemente
    '                    '//de que la actualizacion del stock se haga correctamente. Los movimientos 
    '                    '//pendientes de actualizar, si existen, se gestionaran desde el programa
    '                    '//'Movimientos de Stock asociados a la Orden' (tbOFControlEstructura).

    '                    Dim ParteTrabajo As DataTable
    '                    Dim produccionLog As ControlProduccionUpdateData

    '                    '//parte de trabajo (registro de tbOFControl)
    '                    If IsNumeric(lineaAlbaran("IDOFControl")) Then
    '                        ParteTrabajo = ofc.SelOnPrimaryKey(lineaAlbaran("IDOFControl"))
    '                    Else
    '                        Dim Cabecera As DataTable = acc.SelOnPrimaryKey(lineaAlbaran(_ACL.IDAlbaran))
    '                        If Cabecera.Rows.Count > 0 Then
    '                            Dim FechaActualizacion As Date = Cabecera.Rows(0)("FechaAlbaran")

    '                            Dim operacion As DataRow
    '                            If IsNumeric(lineaAlbaran("IDOrdenRuta")) Then
    '                                operacion = [or].GetItemRow(lineaAlbaran("IDOrdenRuta"))
    '                                ParteTrabajo = ofc.AddNewForm()
    '                                Dim parte As DataRow = ParteTrabajo.Rows(0)
    '                                parte = ofc.ApplyBusinessRule("IDOrden", operacion("IDOrden"), parte)
    '                                parte("FechaInicio") = FechaActualizacion
    '                                parte("FechaFin") = FechaActualizacion
    '                                parte("IDOperario") = OperarioGenerico
    '                                parte("IDOrdenRuta") = operacion("IDOrdenRuta")
    '                                parte("Secuencia") = operacion("Secuencia")
    '                                parte = ofc.ApplyBusinessRule("Secuencia", parte("Secuencia"), parte)
    '                                parte("QBuenaUdProduccion") = lineaAlbaran("QServida")
    '                                parte("QRechazadaUdProduccion") = 0
    '                                parte("QDudosaUdProduccion") = 0
    '                                parte = ofc.ApplyBusinessRule("QBuenaUdProduccion", parte("QBuenaUdProduccion"), parte)
    '                                parte = ofc.ApplyBusinessRule("QRechazadaUdProduccion", parte("QRechazadaUdProduccion"), parte)
    '                                parte = ofc.ApplyBusinessRule("QDudosaUdProduccion", parte("QDudosaUdProduccion"), parte)

    '                                lineaAlbaran("IDOFControl") = parte("IDOFControl")
    '                                lineaAlbaran("EstadoStock") = enumaclEstadoStock.aclActualizado
    '                            End If
    '                        End If
    '                    End If

    '                    '//Obtener las lineas componentes de la linea de subcontratacion actual
    '                    Dim f As New Filter
    '                    f.Add(New NumberFilterItem("IDLineaPadre", lineaAlbaran(_ACL.IDLineaAlbaran)))
    '                    Dim componentes As DataTable = Me.Filter(f)
    '                    If componentes.Rows.Count > 0 Then
    '                        For Each componente As DataRow In componentes.Rows
    '                            componente("IDOFControl") = lineaAlbaran("IDOFControl")
    '                            componente("EstadoStock") = enumaclEstadoStock.aclActualizado
    '                        Next
    '                    End If

    '                    produccionLog = CType(ofc, IControlProduccion).ControlProduccion(ParteTrabajo)
    '                    If Not produccionLog Is Nothing Then
    '                        ArrayManager.Copy(produccionLog.Entradas, updateDataArray)
    '                        ArrayManager.Copy(produccionLog.Salidas, updateDataArray)
    '                    End If

    '                    AdminData.SetData(Linea)
    '                    AdminData.SetData(componentes)
    '                End If
    '            End If
    '        End If
    '        Me.CommitTx()
    '    Next

    '    Return updateDataArray
    'End Function

    'Friend Function ActualizarStock(ByVal cabeceraAlbaran As DataRow, ByVal lineaAlbaran As DataRow) As StockUpdateData()
    '    Dim updateData As StockUpdateData
    '    updateData = ActualizarStockTx(cabeceraAlbaran, lineaAlbaran)
    '    If Not updateData Is Nothing Then
    '        If updateData.Estado = EstadoStock.Actualizado Then
    '            lineaAlbaran(_ACL.IDMovimiento) = updateData.IDLineaMovimiento
    '            lineaAlbaran(_ACL.EstadoStock) = enumaclEstadoStock.aclActualizado
    '        Else
    '            lineaAlbaran(_ACL.IDMovimiento) = DBNull.Value
    '            If updateData.Estado = EstadoStock.NoActualizado Then
    '                lineaAlbaran(_ACL.EstadoStock) = enumaclEstadoStock.aclNoActualizado
    '            ElseIf updateData.Estado = EstadoStock.SinGestion Then
    '                lineaAlbaran(_ACL.EstadoStock) = enumaclEstadoStock.aclSinGestion
    '            End If
    '        End If
    '    End If

    '    'Para que quede igual que en ventas esta funcion devuelve un array
    '    Dim updateDataArray(0) As StockUpdateData
    '    updateDataArray(0) = updateData
    '    Return updateDataArray
    'End Function

    'Friend Function ActualizarStock(ByVal cabeceraAlbaran As DataRow, ByVal lineaAlbaran As DataRow, ByVal AlbaranLote As DataTable) As StockUpdateData()
    '    Dim updateDataArray(-1) As StockUpdateData
    '    If Not AlbaranLote Is Nothing AndAlso AlbaranLote.Rows.Count > 0 Then
    '        For Each lote As DataRow In AlbaranLote.Rows
    '            Dim updateData As StockUpdateData
    '            updateData = ActualizarStockTx(cabeceraAlbaran, lineaAlbaran, lote)
    '            If Not updateData Is Nothing Then
    '                If updateData.Estado = EstadoStock.Actualizado Then
    '                    lote(_ACLT.IDMovimientoEntrada) = updateData.IDLineaMovimiento
    '                Else
    '                    lote(_ACLT.IDMovimientoEntrada) = DBNull.Value
    '                End If

    '                ReDim Preserve updateDataArray(UBound(updateDataArray) + 1)
    '                updateDataArray(UBound(updateDataArray)) = updateData
    '            End If
    '        Next

    '        For Each updateItem As StockUpdateData In updateDataArray
    '            If Not updateItem Is Nothing Then
    '                If updateItem.Estado = EstadoStock.Actualizado Then
    '                    lineaAlbaran(_ACL.EstadoStock) = enumaclEstadoStock.aclActualizado
    '                ElseIf updateItem.Estado = EstadoStock.NoActualizado Then
    '                    lineaAlbaran(_ACL.EstadoStock) = enumaclEstadoStock.aclNoActualizado
    '                    Exit For
    '                ElseIf updateItem.Estado = EstadoStock.SinGestion Then
    '                    lineaAlbaran(_ACL.EstadoStock) = enumaclEstadoStock.aclSinGestion
    '                    Exit For
    '                End If
    '            End If
    '        Next
    '    Else
    '        ReDim Preserve updateDataArray(UBound(updateDataArray) + 1)
    '        updateDataArray(UBound(updateDataArray)) = Me.LogActualizarStock("El lote es obligatorio.", lineaAlbaran(_ACL.IDArticulo), lineaAlbaran(_ACL.IDAlmacen))
    '        Return updateDataArray
    '    End If

    '    Return updateDataArray
    'End Function

    'Private Function ActualizarStockTx(ByVal cabeceraAlbaran As DataRow, ByVal lineaAlbaran As DataRow, Optional ByVal lineaAlbaranLote As DataRow = Nothing) As StockUpdateData
    '    Dim stockobj As New Stock
    '    Dim NumeroMovimiento As Long
    '    If IsNumeric(cabeceraAlbaran(_ACC.NMovimiento)) AndAlso Not (cabeceraAlbaran(_ACC.NMovimiento) = 0) Then
    '        NumeroMovimiento = cabeceraAlbaran(_ACC.NMovimiento)
    '    Else
    '        NumeroMovimiento = stockobj.NumeroMovimiento
    '        cabeceraAlbaran(_ACC.NMovimiento) = NumeroMovimiento
    '    End If

    '    '//Volcar los datos en un objeto stockData sin importar si la gestion es normal, por lotes o numeros de serie.
    '    '//Este control ya se hace en las funciones de actualizacion el stock.
    '    Dim data As New StockData
    '    data.Articulo = lineaAlbaran(_ACL.IDArticulo)
    '    data.Almacen = lineaAlbaran(_ACL.IDAlmacen)
    '    If Length(lineaAlbaran(_ACL.Lote)) > 0 Then
    '        data.Lote = lineaAlbaran(_ACL.Lote)
    '        data.NSerie = lineaAlbaran(_ACL.Lote)
    '    End If
    '    If Length(lineaAlbaran(_ACL.Ubicacion)) > 0 Then
    '        data.Ubicacion = lineaAlbaran(_ACL.Ubicacion)
    '    End If
    '    data.IDDocumento = cabeceraAlbaran(_ACC.IDAlbaran)
    '    data.Documento = cabeceraAlbaran(_ACC.NAlbaran)
    '    data.FechaDocumento = cabeceraAlbaran(_ACC.FechaAlbaran)
    '    If Not lineaAlbaranLote Is Nothing Then
    '        data.Lote = lineaAlbaranLote(_ACLT.Lote) & String.Empty
    '        data.Ubicacion = lineaAlbaranLote(_ACLT.Ubicacion) & String.Empty
    '        data.Cantidad = lineaAlbaranLote(_ACLT.QInterna)
    '    Else
    '        data.Cantidad = lineaAlbaran(_ACL.QInterna)
    '    End If
    '    If IsNumeric(lineaAlbaran(_ACL.IDObra)) Then
    '        data.Obra = lineaAlbaran(_ACL.IDObra)
    '    End If
    '    If Length(lineaAlbaran(_ACL.IDEstadoActivo)) > 0 Then
    '        data.EstadoNSerie = lineaAlbaran(_ACL.IDEstadoActivo)
    '    End If
    '    If Length(lineaAlbaran(_ACL.IDOperario)) > 0 Then
    '        data.Operario = lineaAlbaran(_ACL.IDOperario)
    '    Else
    '        Dim StrOperario As String = New Operario().ObtenerIDOperarioUsuario
    '        If Len(StrOperario) > 0 Then data.Operario = StrOperario
    '    End If
    '    '//Importes extras
    '    Dim ImporteExtraA As Double
    '    Dim ImporteExtraB As Double
    '    Dim ACP As New AlbaranCompraPrecio
    '    Dim Importes As DataTable = ACP.Filter(New NumberFilterItem(_ACP.IDLineaAlbaran, FilterOperator.Equal, lineaAlbaran(_ACL.IDLineaAlbaran)))
    '    If Not Importes Is Nothing Then
    '        For Each importe As DataRow In Importes.Rows
    '            ImporteExtraA = ImporteExtraA + importe(_ACP.ImporteA)
    '            ImporteExtraB = ImporteExtraB + importe(_ACP.ImporteB)
    '        Next
    '    End If
    '    If lineaAlbaran(_ACL.Factor) <> 0 And lineaAlbaran(_ACL.UdValoracion) <> 0 Then
    '        Dim m As New Moneda
    '        Dim monedaA As MonedaInfo = m.MonedaA
    '        Dim monedaB As MonedaInfo = m.MonedaB
    '        data.PrecioA = xRound(ImporteExtraA + (lineaAlbaran(_ACL.PrecioA) / lineaAlbaran(_ACL.Factor) / lineaAlbaran(_ACL.UdValoracion) * (1 - lineaAlbaran(_ACL.Dto1) / 100) * (1 - lineaAlbaran(_ACL.Dto2) / 100) * (1 - lineaAlbaran(_ACL.Dto3) / 100) * (1 - lineaAlbaran(_ACL.Dto) / 100) * (1 - lineaAlbaran(_ACL.DtoProntoPago) / 100)), monedaA.NDecimalesPrecio)
    '        data.PrecioB = xRound(ImporteExtraB + (lineaAlbaran(_ACL.PrecioB) / lineaAlbaran(_ACL.Factor) / lineaAlbaran(_ACL.UdValoracion) * (1 - lineaAlbaran(_ACL.Dto1) / 100) * (1 - lineaAlbaran(_ACL.Dto2) / 100) * (1 - lineaAlbaran(_ACL.Dto3) / 100) * (1 - lineaAlbaran(_ACL.Dto) / 100) * (1 - lineaAlbaran(_ACL.DtoProntoPago) / 100)), monedaB.NDecimalesPrecio)
    '    End If


    '    '//Determinar el tipo de movimiento que por defecto es entrada de albaran
    '    If lineaAlbaran(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclNormal _
    '    Or lineaAlbaran(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclKit Then
    '        data.TipoMovimiento = enumTipoMovimiento.tmEntAlbaranCompra
    '    ElseIf lineaAlbaran(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclSubcontratacion Then
    '        data.TipoMovimiento = enumTipoMovimiento.tmEntSubcontratacion
    '    ElseIf lineaAlbaran(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclComponente Then
    '        'si la linea es componente hay que determinar si pertenece a un kit o es un componente de subcontratacion manual
    '        Dim lineaPadre As DataTable = Me.SelOnPrimaryKey(lineaAlbaran(_ACL.IDLineaPadre))
    '        If Not lineaPadre Is Nothing AndAlso lineaPadre.Rows.Count > 0 Then
    '            If lineaPadre.Rows(0)(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclKit Then
    '                data.TipoMovimiento = enumTipoMovimiento.tmEntAlbaranCompra
    '            ElseIf lineaPadre.Rows(0)(_ACL.TipoLineaAlbaran) = enumaclTipoLineaAlbaran.aclSubcontratacion Then
    '                data.Cantidad = data.Cantidad * -1
    '                data.TipoMovimiento = enumTipoMovimiento.tmSalSubcontratacion
    '            End If
    '        End If
    '    End If
    '    '/// Campos específicos de GAM ///
    '    If lineaAlbaran.Table.Columns.Contains("CodProveedorAlquiler") Then data.CodProveedorAlquiler = Nz(lineaAlbaran("CodProveedorAlquiler"), Nothing)
    '    If lineaAlbaran.Table.Columns.Contains("MaquinaRealquilada") Then data.MaquinaRealquilada = Nz(lineaAlbaran("MaquinaRealquilada"), False)
    '    If lineaAlbaran.Table.Columns.Contains("MaquinaRealquiladaGrupo") Then data.MaquinaRealquiladaGrupo = Nz(lineaAlbaran("MaquinaRealquiladaGrupo"), False)
    '    '////////////////////////////////
    '    Dim updateData As StockUpdateData
    '    If data.TipoMovimiento = enumTipoMovimiento.tmSalSubcontratacion Then
    '        updateData = stockobj.Salida(NumeroMovimiento, data)
    '    Else
    '        updateData = stockobj.Entrada(NumeroMovimiento, data)
    '    End If

    '    If Not updateData Is Nothing Then
    '        If updateData.Estado = EstadoStock.NoActualizado Then
    '            Me.RollbackTx()
    '        Else
    '            PrepararActivoUltimaCompra(lineaAlbaran)
    '        End If
    '    End If
    '    Return updateData
    'End Function

    'Private Function LogActualizarStock(ByVal log As String, Optional ByVal Articulo As String = Nothing, Optional ByVal Almacen As String = Nothing, Optional ByVal Lote As String = Nothing, Optional ByVal Ubicacion As String = Nothing) As StockUpdateData
    '    Dim auxData As New StockData
    '    auxData.Articulo = Articulo
    '    auxData.Almacen = Almacen
    '    auxData.Lote = Lote
    '    auxData.Ubicacion = Ubicacion
    '    Dim auxUpdateData As New StockUpdateData
    '    auxUpdateData.StockData = auxData
    '    auxUpdateData.Log = log
    '    Return auxUpdateData
    'End Function

#End Region

#Region " CALIDAD "

    <Task()> Public Shared Function DesmarcarControlCalidad(ByVal IdLineaAlbaran As Integer, ByVal services As ServiceProvider) As Boolean
        Dim ACL As New AlbaranCompraLinea()
        Dim dtACL As DataTable = ACL.SelOnPrimaryKey(IdLineaAlbaran)
        If Not dtACL Is Nothing AndAlso dtACL.Rows.Count > 0 Then
            dtACL.Rows(0)("ControlCalidad") = False
            ACL.Update(dtACL)
            Return True
        Else
            Return False
        End If
    End Function

#End Region

    'TODO: PENDIENTE
#Region " Actualizar Activo "

    Private Sub PrepararActivoUltimaCompra(ByVal IntIDAlbaran As Integer)
        Dim DtACL As DataTable = New AlbaranCompraLinea().Filter(New FilterItem("IDAlbaran", FilterOperator.Equal, IntIDAlbaran))
        If Not DtACL Is Nothing AndAlso DtACL.Rows.Count > 0 Then
            PrepararActivoUltimaCompra(DtACL)
        End If
    End Sub

    Private Sub PrepararActivoUltimaCompra(ByVal dt As DataTable)
        Dim strIN As String
        Dim strIdArticulo As String

        If Not IsNothing(dt) And dt.Rows.Count > 0 Then
            For Each dr As DataRow In dt.Rows
                If InStr(strIN, dr("IDArticulo")) = 0 Then
                    strIdArticulo = dr("IDArticulo")
                    If Len(strIN) Then
                        strIN = strIN & ","
                    End If
                    strIN = strIN & "'" & strIdArticulo & "'"
                    PrepararActivoUltimaCompra(dr)
                End If
            Next
        End If
    End Sub

    Private Sub PrepararActivoUltimaCompra(ByVal dr As DataRow)
        If Not dr Is Nothing Then
            Dim drArticulo As DataRow = New Articulo().GetItemRow(dr("IDArticulo"))
            If Not IsNothing(drArticulo) AndAlso Nz(drArticulo("NSerieObligatorio"), False) Then
                If Length(dr("Lote")) > 0 Then
                    Dim objFilter As New Filter
                    objFilter.Add(New StringFilterItem("IDArticulo", dr("IDArticulo")))
                    objFilter.Add(New StringFilterItem("NSerie", dr("Lote")))

                    Dim dtArtNSerie As DataTable = New BE.DataEngine().Filter("vFrmArticuloNSerie", objFilter)
                    If Not IsNothing(dtArtNSerie) AndAlso dtArtNSerie.Rows.Count > 0 Then
                        Dim dtActivo As DataTable = New Activo().SelOnPrimaryKey(dtArtNSerie.Rows(0)("IDActivo"))
                        If Not IsNothing(dtActivo) AndAlso dtActivo.Rows.Count > 0 Then
                            dtActivo.Rows(0)("IDProveedor") = drArticulo("IdProveedorUltimaCompra")
                            dtActivo.Rows(0)("PrecioUltimaCompra") = drArticulo("PrecioUltimaCompraA")
                            BE.BusinessHelper.UpdateTable(dtActivo)
                        End If
                    End If
                End If
            End If
        End If
    End Sub

#End Region

#Region "Otros"

    <Task()> Public Shared Function GetIDAlbaranLinea(ByVal data As Object, ByVal services As ServiceProvider) As Integer
        Return AdminData.GetAutoNumeric
    End Function

    <Serializable()> _
  Public Class DataBuscarArt
        Public IDArticulo As String
        Public NSerie As Boolean = False

        Public Sub New()
        End Sub
    End Class

    <Task()> Public Shared Function BuscarArticulo(ByVal data As String, ByVal services As ServiceProvider) As DataBuscarArt
        Dim DataReturn As New DataBuscarArt
        Dim DtArt As DataTable
        Dim ClsArt As New Articulo
        'Búsqueda por artículo principal
        DtArt = ClsArt.SelOnPrimaryKey(data)
        If DtArt Is Nothing OrElse DtArt.Rows.Count = 0 Then
            'Búsqueda por RefProveedor
            Dim DtRef As DataTable = New ArticuloProveedor().Filter(New FilterItem("RefProveedor", FilterOperator.Equal, data))
            If Not DtRef Is Nothing AndAlso DtRef.Rows.Count > 0 Then
                DataReturn.IDArticulo = DtRef.Rows(0)("IDArticulo")
            Else
                'Búsqueda por código barras Artículo
                DtArt = ClsArt.Filter(New FilterItem("CodigoBarras", FilterOperator.Equal, data))
                If Not DtArt Is Nothing AndAlso DtArt.Rows.Count > 0 Then
                    DataReturn.IDArticulo = DtArt.Rows(0)("IDArticulo")
                Else
                    'Búsqueda por múltiples códigos de barras de artículo
                    Dim DtCods As DataTable = New ArticuloCodigoBarras().Filter(New FilterItem("CodigoBarras", FilterOperator.Equal, data))
                    If Not DtCods Is Nothing AndAlso DtCods.Rows.Count > 0 Then
                        DataReturn.IDArticulo = DtCods.Rows(0)("IDArticulo")
                    Else
                        'Búsqueda por número de Serie
                        Dim DtNSerie As DataTable = New ArticuloNSerie().Filter(New FilterItem("NSerie", FilterOperator.Equal, data))
                        If Not DtNSerie Is Nothing AndAlso DtNSerie.Rows.Count > 0 Then
                            Dim FilEstado As New Filter
                            FilEstado.Add("Disponible", FilterOperator.Equal, 1)
                            FilEstado.Add("IDEstadoActivo", FilterOperator.Equal, DtNSerie.Rows(0)("IDEstadoActivo"))
                            Dim DtEstado As DataTable = New BE.DataEngine().Filter("tbMntoEstadoActivo", FilEstado)
                            If Not DtEstado Is Nothing AndAlso DtEstado.Rows.Count > 0 Then
                                DataReturn.IDArticulo = DtNSerie.Rows(0)("IDArticulo")
                                DataReturn.NSerie = True
                            Else
                                'Búsqueda por Lote
                                Dim DtLote As DataTable = New ArticuloAlmacenLote().Filter(New FilterItem("Lote", FilterOperator.Equal, data))
                                If Not DtLote Is Nothing AndAlso DtLote.Rows.Count > 0 Then
                                    DataReturn.IDArticulo = DtLote.Rows(0)("IDArticulo")
                                End If
                            End If
                        Else
                            'Búsqueda por Lote
                            Dim DtLote As DataTable = New ArticuloAlmacenLote().Filter(New FilterItem("Lote", FilterOperator.Equal, data))
                            If Not DtLote Is Nothing AndAlso DtLote.Rows.Count > 0 Then
                                DataReturn.IDArticulo = DtLote.Rows(0)("IDArticulo")
                            End If
                        End If
                    End If
                End If
            End If
        Else
            Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
            Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data)

            If Not ArtInfo.Activo Then
                ApplicationService.GenerateError("El artículo | no está activo.", Quoted(data))
            ElseIf Not ArtInfo.Compra Then
                ApplicationService.GenerateError("El artículo | no es de tipo compra.", Quoted(data))
            Else
                DataReturn.IDArticulo = data
            End If
        End If
        Return DataReturn
    End Function

#End Region

End Class

