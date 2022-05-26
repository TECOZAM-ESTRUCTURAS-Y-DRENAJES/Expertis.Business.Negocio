Imports Solmicro.Expertis.Engine.BE.ApplicationService
Imports _AVC = Solmicro.Expertis.Business.Negocio._AlbaranVentaCabecera
Imports _AVL = Solmicro.Expertis.Business.Negocio._AlbaranVentaLinea
Imports _PVL = Solmicro.Expertis.Business.Negocio._PedidoVentaLinea
Imports _AA = Solmicro.Expertis.Business.Negocio._ArticuloAlmacen
Imports _AVLT = Solmicro.Expertis.Business.Negocio._AlbaranVentaLote
Imports System.Collections.Generic

Public Class _AlbaranVentaLinea
    Public Const IDLineaAlbaran As String = "IDLineaAlbaran"
    Public Const IDAlbaran As String = "IDAlbaran"
    Public Const IDLineaPedido As String = "IDLineaPedido"
    Public Const IDOrdenLinea As String = "IDOrdenLinea"
    Public Const IDPedido As String = "IDPedido"
    Public Const PedidoCliente As String = "PedidoCliente"
    Public Const IDArticulo As String = "IDArticulo"
    Public Const DescArticulo As String = "DescArticulo"
    Public Const RefCliente As String = "RefCliente"
    Public Const DescRefCliente As String = "DescRefCliente"
    Public Const IDAlmacen As String = "IDAlmacen"
    Public Const IDFormaPago As String = "IDFormaPago"
    Public Const IDCondicionPago As String = "IDCondicionPago"
    Public Const IDTipoIva As String = "IDTipoIva"
    Public Const CContable As String = "CContable"
    Public Const EstadoStock As String = "EstadoStock"
    Public Const EstadoFactura As String = "EstadoFactura"
    Public Const IDMovimiento As String = "IDMovimiento"
    Public Const IDUdMedida As String = "IDUdMedida"
    Public Const IDUdInterna As String = "IDUdInterna"
    Public Const QServida As String = "QServida"
    Public Const QFacturada As String = "QFacturada"
    Public Const Precio As String = "Precio"
    Public Const PrecioA As String = "PrecioA"
    Public Const PrecioB As String = "PrecioB"
    Public Const Dto As String = "Dto"
    Public Const DtoProntoPago As String = "DtoProntoPago"
    Public Const Dto1 As String = "Dto1"
    Public Const Dto2 As String = "Dto2"
    Public Const Dto3 As String = "Dto3"
    Public Const Importe As String = "Importe"
    Public Const ImporteA As String = "ImporteA"
    Public Const ImporteB As String = "ImporteB"
    Public Const UdValoracion As String = "UdValoracion"
    Public Const QServidaUd As String = "QServidaUd"
    Public Const Texto As String = "Texto"
    Public Const PVP As String = "PVP"
    Public Const PVPA As String = "PVPA"
    Public Const PVPB As String = "PVPB"
    Public Const Lote As String = "Lote"
    Public Const Ubicacion As String = "Ubicacion"
    Public Const Regalo As String = "Regalo"
    Public Const PuntosUtilizados As String = "PuntosUtilizados"
    Public Const ImportePVP As String = "ImportePVP"
    Public Const ImportePVPA As String = "ImportePVPA"
    Public Const ImportePVPB As String = "ImportePVPB"
    Public Const IDPromocionLinea As String = "IDPromocionLinea"
    Public Const IDTrabajo As String = "IDTrabajo"
    Public Const IDObra As String = "IDObra"
    Public Const IDLineaMaterial As String = "IDLineaMaterial"
    Public Const IDLineaPadre As String = "IDLineaPadre"
    Public Const TipoLineaAlbaran As String = "TipoLineaAlbaran"
    Public Const Facturable As String = "Facturable"
    Public Const Factor As String = "Factor"
    Public Const QInterna As String = "QInterna"
    Public Const IDArticuloContenedor As String = "IDArticuloContenedor"
    Public Const QEtiContenedor As String = "QEtiContenedor"
    Public Const IDArticuloEmbalaje As String = "IDArticuloEmbalaje"
    Public Const QEtiEmbalaje As String = "QEtiEmbalaje"
    Public Const Revision As String = "Revision"
    Public Const IDLineaOfertaDetalle As String = "IDLineaOfertaDetalle"
    Public Const PrecioUltimaCompra As String = "PrecioUltimaCompra"
    Public Const IDMovimientoEntrada As String = "IDMovimientoEntrada"
    Public Const SeguimientoTarifa As String = "SeguimientoTarifa"
    Public Const FechaAlbaranEnvio As String = "FechaAlbaranEnvio"
    Public Const IDLineaAlbaranDeposito As String = "IDLineaAlbaranDeposito"
    Public Const FechaPrevistaRetorno As String = "FechaPrevistaRetorno"
    Public Const PuntosMarketing As String = "PuntosMarketing"
    Public Const IDSalidaContenedor As String = "IDSalidaContenedor"
    Public Const IDEntradaContenedor As String = "IDEntradaContenedor"
    Public Const FechaAlquiler As String = "FechaAlquiler"
    Public Const HoraAlquiler As String = "HoraAlquiler"
    Public Const IDAlbaranDeposito As String = "IDAlbaranDeposito"
    Public Const TipoFactAlquiler As String = "TipoFactAlquiler"
    Public Const IDOperario As String = "IDOperario"
    Public Const IDEstadoActivo As String = "IDEstadoActivo"
    Public Const IDCentroGestion As String = "IDCentroGestion"
    Public Const FechaEntregaModificado As String = "FechaEntregaModificado"
    Public Const FechaRetornoDiasMinimos As String = "FechaRetornoDiasMinimos"
    Public Const TextoContacto As String = "TextoContacto"
    Public Const ConsumoAlquiler As String = "ConsumoAlquiler"
    Public Const FechaCreacionAudi As String = "FechaCreacionAudi"
    Public Const FechaModificacionAudi As String = "FechaModificacionAudi"
    Public Const UsuarioAudi As String = "UsuarioAudi"
    Public Const IDMovimientoSalida As String = "IDMovimientoSalida"
    Public Const PrecioCosteA As String = "PrecioCosteA"
    Public Const PrecioCosteB As String = "PrecioCosteB"
    Public Const IDTipoLinea As String = "IDTipoLinea"
    Public Const IDDireccionFra As String = "IDDireccionFra"
    Public Const IDClienteBanco As String = "IDClienteBanco"
    Public Const QPendienteDevolverAInicio As String = "QPendienteDevolverAInicio"
    Public Const ARNAlbaranRecogida As String = "ARNAlbaranRecogida"
End Class

#Region "Control de la OTS"

Public Interface IControlOT
    Sub GenerarNuevaOTDEsdeRetornos(ByVal dtOT As DataTable, ByVal services As ServiceProvider)
End Interface

#End Region

Public Class AlbaranVentaLinea
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Private Const cnEntidad As String = "tbAlbaranVentaLinea"

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    <Serializable()> _
    Public Class udtComponentes
        Public lngTipoLinea As enumavlTipoLineaAlbaran
        Public DtComponentes As DataTable
    End Class

#Region " Delete "

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf Business.General.Comunes.BeginTransaction)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ValidarDelEstadoAVL)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ValidarDelTipoAlbaran)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ValidarDelEstadoAlbRetorno)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ValidarDelEstadoAlbVentaOrigen)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ValidarDelOTAsociada)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf EliminarContabilizacion)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ActualizarLineaPedido)
        'deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ActualizarLineaPromocionPorBorrado)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ProcesoAlbaranVentaObras.ActualizarObraMaterialLineaPorBorrado)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ProcesoAlbaranVentaAlquiler.ActualizarOrdenesServicioPorBorrado)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf EliminarMovimientoStock)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf EliminarComponentes)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ActualizarAlbaranesMultiEmpresa)
        'deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ActualizarPuntosMarketing)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf Business.General.Comunes.DeleteEntityRow)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf Business.General.Comunes.MarcarComoEliminado)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ActualizarLineaPromocionPorBorrado)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ActualizarCantidadPromocionada)
    End Sub

    <Task()> Public Shared Function NoHaSidoEliminada(ByVal data As DataRow, ByVal services As ServiceProvider) As Boolean
        Dim listaEliminados As LineasAlbaranEliminadas = services.GetService(Of LineasAlbaranEliminadas)()
        Dim haSidoEliminado As Boolean = listaEliminados.IDLineas.Contains(data("IDLineaAlbaran"))
        If haSidoEliminado Then services.GetService(Of DeleteProcessContext).Deleted = haSidoEliminado
        Return Not haSidoEliminado
    End Function

    <Task()> Public Shared Sub EliminarContabilizacion(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
        If Not AppParams.GestionInventarioPermanente Then Exit Sub

        If data("Contabilizado") <> enumContabilizado.NoContabilizado Then
           Dim IDLineas(0) As Object
            IDLineas(0) = data("IDLineaAlbaran")

            Dim datIStock As New ProcesoStocks.DataCreateIStockClass(InventariosPermanentes.ENSAMBLADO_INV_PERMANENTE_STOCKS, InventariosPermanentes.CLASE_INV_PERMANENTE_STOCKS)
            Dim IStockClass As IStockInventarioPermanente = ProcessServer.ExecuteTask(Of ProcesoStocks.DataCreateIStockClass, IStockInventarioPermanente)(AddressOf ProcesoStocks.CreateIStockInventarios, datIStock, services)
            If Not IStockClass Is Nothing Then
                Dim datGetDesconta As New ProcesoAlbaranVenta.DataGetLineasDescontabilizar(IDLineas)
                datGetDesconta = ProcessServer.ExecuteTask(Of ProcesoAlbaranVenta.DataGetLineasDescontabilizar, ProcesoAlbaranVenta.DataGetLineasDescontabilizar)(AddressOf ProcesoAlbaranVenta.GetLineasDescontabilizar, datGetDesconta, services)

                IStockClass.SincronizarDescontaAlbaranVenta(datGetDesconta.ApuntesAlbaran, services)

                '//Si eliminamos la conta, tenemos que actualizar el campo 'Contabilizado' para que el resto de tareas que lo consultan, sepan que ya se ha hecho la descontabilización
                Dim datValEstado As New Comunes.DataValidarEstado(data("IDLineaAlbaran"), enumDiarioTipoApunte.AlbaranVenta)
                data("Contabilizado") = ProcessServer.ExecuteTask(Of Comunes.DataValidarEstado, Integer)(AddressOf Comunes.ValidarEstado, datValEstado, services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarDelEstadoAVL(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data(_AVL.EstadoFactura) = enumavlEstadoFactura.avlParcFacturado Then
            ApplicationService.GenerateError("No se puede borrar el Albarán. Está parcialmente facturado.")
        End If
        If data(_AVL.EstadoFactura) = enumavlEstadoFactura.avlFacturado Then
            ApplicationService.GenerateError("No se puede borrar el Albarán. Está total o parcialmente facturado.")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarDelTipoAlbaran(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim strTipoAlbaran As String
        Dim dtAVC As DataTable = New AlbaranVentaCabecera().SelOnPrimaryKey(data(_AVL.IDAlbaran))
        If Not dtAVC Is Nothing AndAlso dtAVC.Rows.Count > 0 Then
            strTipoAlbaran = dtAVC.Rows(0)(_AVC.IDTipoAlbaran) & String.Empty
        End If

        Dim AppParamsAlb As ParametroAlbaranVenta = services.GetService(Of ParametroAlbaranVenta)()
        If Length(strTipoAlbaran) > 0 Then
            If strTipoAlbaran = AppParamsAlb.TipoAlbaranDeDeposito Then
                Dim AVL As New AlbaranVentaLinea  '//Hay que crear este instancia para poder hacer el Filter.
                Dim dtDeposito As DataTable = AVL.Filter(New NumberFilterItem("IDLineaAlbaranDeposito", data("IDLineaAlbaran")))
                If Not dtDeposito Is Nothing AndAlso dtDeposito.Rows.Count > 0 Then
                    ApplicationService.GenerateError("No se puede borrar el Albarán de Depósito, tiene un Albarán de Retorno asociado.")
                End If
            End If

            Dim dt As DataTable
            If strTipoAlbaran = AppParamsAlb.TipoAlbaranDeDeposito Then
                dt = AdminData.GetData("tbPreventivoContadorHist", New NumberFilterItem("IDLineaAlbaran", data("IDLineaAlbaran")), "IdHistoricoContador")
            ElseIf strTipoAlbaran = AppParamsAlb.TipoAlbaranRetornoAlquiler Then
                dt = AdminData.GetData("tbPreventivoContadorHist", New NumberFilterItem("IDLineaAlbaranRetorno", data("IDLineaAlbaran")), "IdHistoricoContador")
            End If
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("La linea tiene lineas de medidas de contadores asociadas.")
            End If

            If strTipoAlbaran = AppParamsAlb.TipoAlbaranDeDeposito AndAlso Length(data(_AVL.ARNAlbaranRecogida)) > 0 Then
                ApplicationService.GenerateError("La linea tiene Avisos de retornos activos.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarDelEstadoAlbRetorno(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim f As New Filter
        f.Add(New NumberFilterItem("IDLineaAlbaranRetorno", data("IDLineaAlbaran")))
        Dim dt As DataTable = AdminData.GetData("tbFacturaVentaLinea", f, "IdLineaAlbaranRetorno")
        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            ApplicationService.GenerateError("No se puede borrar el Albarán. Está parcialmente facturado.")
        End If

        dt = AdminData.GetData("tbObraTrabajoFacturacion", f, "IdLineaAlbaranRetorno")
        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            ApplicationService.GenerateError("No se puede borrar el Albarán. Está parcialmente facturado.")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarDelEstadoAlbVentaOrigen(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim f As New Filter
        f.Add(New NumberFilterItem("IDLineaAlbaranVentaOrigen", data("IDLineaAlbaran")))
        Dim dt As DataTable = AdminData.GetData("tbObraTrabajoFacturacion", f, "IdLineaAlbaranVentaOrigen")
        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            ApplicationService.GenerateError("No se puede borrar el Albarán. Está parcialmente facturado.")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarDelOTAsociada(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim f As New Filter
        f.Add(New NumberFilterItem("IDLineaAlbaranRetorno", data("IDLineaAlbaran")))
        Dim dt As DataTable = AdminData.GetData("tbMntoOT", f, "IDOT")
        If dt.Rows.Count > 0 Then
            f.Clear()
            f.Add(New NumberFilterItem("IDOT", dt.Rows(0)("IDOT")))
            dt = AdminData.GetData("tbMntoOTAccionOT", f, "IDOT")
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("No se ha podido borrar la línea porque no se ha podido borrar la OT asociada.")
            End If
            dt = AdminData.GetData("tbMntoOTCausaOT", f, "IDOT")
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("No se ha podido borrar la línea porque no se ha podido borrar la OT asociada.")
            End If
            dt = AdminData.GetData("tbMntoOTDefectoOT", f, "IDOT")
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("No se ha podido borrar la línea porque no se ha podido borrar la OT asociada.")
            End If
            dt = AdminData.GetData("tbPreventivoActivo", f, "IDOT")
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("No se ha podido borrar la línea porque no se ha podido borrar la OT asociada.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarLineaPedido(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data(_AVL.IDLineaPedido)) > 0 Then
            data(_AVL.QServida) = 0
            'Dim fLineasPedido As New Filter
            'fLineasPedido.Add(New IsNullFilterItem("IDAlbaran", False))
            'fLineasPedido.Add(New IsNullFilterItem("IDLineaAlbaran", False))
            'Dim strLineasPedido As String = fLineasPedido.Compose(New AdoFilterComposer)
            Dim Pedidos As New System.Collections.Generic.List(Of DataTable)
            Dim Doc As New DocumentoAlbaranVenta(data("IDAlbaran"))
            Dim DataActua As New ProcesoAlbaranVentaPedidos.DataActuaPedidos(Doc, data, Pedidos)
            ProcessServer.ExecuteTask(Of ProcesoAlbaranVentaPedidos.DataActuaPedidos)(AddressOf ProcesoAlbaranVentaPedidos.ActualizarLineasPedido, DataActua, services)
            If Not Pedidos Is Nothing AndAlso Pedidos.Count > 0 Then
                For Each Pedido As DataTable In Pedidos
                    BusinessHelper.UpdateTable(Pedido)
                Next
            End If
            'ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoAlbaranVentaPedidos.ActualizarLineaPedido, data, services)
            'ProcessServer.ExecuteTask(Of Object)(AddressOf ProcesoAlbaranVentaPedidos.GrabarPedidos, Nothing, services)
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarLineaPromocionPorBorrado(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data(_AVL.IDPromocionLinea)) > 0 And data("Regalo") = 0 Then
            Dim datosPromo As New PromocionLinea.DatosActuaLinPromoDr(data.Table, True)
            ProcessServer.ExecuteTask(Of PromocionLinea.DatosActuaLinPromoDr)(AddressOf PromocionLinea.ActualizarLineaPromocion, datosPromo, services)
        End If
    End Sub

    <Task()> Public Shared Sub EliminarComponentes(ByVal data As DataRow, ByVal services As ServiceProvider)
        If (data.IsNull(_AVL.IDLineaPadre) AndAlso data(_AVL.TipoLineaAlbaran) = enumavlTipoLineaAlbaran.avlKit) OrElse _
            (data(_AVL.TipoLineaAlbaran) = enumavlTipoLineaAlbaran.avlFantasma) Then
            Dim f As New Filter
            f.Add(New NumberFilterItem("IDLineaPadre", data("IDLineaAlbaran")))
            Dim AVL As New AlbaranVentaLinea
            Dim componentes As DataTable = AVL.Filter(f)
            If Not IsNothing(componentes) AndAlso componentes.Rows.Count > 0 Then
                Dim listaEliminados As LineasAlbaranEliminadas = services.GetService(Of LineasAlbaranEliminadas)()
                Dim dv As New DataView(componentes)
                For Each drv As DataRowView In dv

                    Dim fRecur As New Filter
                    fRecur.Add(New NumberFilterItem("IDLineaPadre", drv("IDLineaAlbaran")))
                    Dim componentesRecur As DataTable = AVL.Filter(fRecur)
                    If componentesRecur.Rows.Count > 0 Then
                        Dim dvRecur As New DataView(componentesRecur)
                        For Each drvRecur As DataRowView In dvRecur
                            Dim drAVLRecur As DataRow = AVL.GetItemRow(drvRecur("IDLineaAlbaran"))
                            ProcessServer.ExecuteTask(Of DataRow)(AddressOf EliminarMovimientoStock, drAVLRecur, services)
                            AVL.DeleteRowCascade(drAVLRecur, services)
                            listaEliminados.IDLineas.Add(drvRecur("IDLineaALbaran"), drvRecur("IDLineaALbaran"))
                        Next
                    End If


                    Dim drAVL As DataRow = AVL.GetItemRow(drv("IDLineaAlbaran"))
                    ProcessServer.ExecuteTask(Of DataRow)(AddressOf EliminarMovimientoStock, drAVL, services)
                    ' Dim listaEliminados As LineasAlbaranEliminadas = services.GetService(Of LineasAlbaranEliminadas)()
                    AVL.DeleteRowCascade(drAVL, services)
                    listaEliminados.IDLineas.Add(drv("IDLineaALbaran"), drv("IDLineaALbaran"))

                Next
            End If
        End If
    End Sub


    <Task()> Public Shared Sub EliminarMovimientoStock(ByVal data As DataRow, ByVal services As ServiceProvider)
        '//Si la linea tiene lotes la correccion de movimientos esta en el Delete de AlbaranVentaLote.
        '//En cualquier caso se tiene que llamar a EliminarMovimiento de AlbaranVentaLinea porque pueden 
        '//existir movimientos de contenedores, depositos, movimientos de numeros de serie, etc.

        Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
        If AppParams.GestionInventarioPermanente Then
            If data("Contabilizado") <> CInt(enumContabilizado.NoContabilizado) Then
                ApplicationService.GenerateError("Debe descontabilizar primero la/s línea/s para poder eliminar el Movimiento.")
            Else
                If Nz(data("TipoLineaAlbaran"), 0) = enumavlTipoLineaAlbaran.avlKit Then
                    Dim f As New Filter
                    f.Add(New NumberFilterItem("TipoLineaAlbaran", enumavlTipoLineaAlbaran.avlComponente))
                    f.Add(New NumberFilterItem("IDLineaPadre", data("IDLineaAlbaran")))
                    f.Add(New NumberFilterItem("Contabilizado", FilterOperator.NotEqual, enumContabilizado.NoContabilizado))
                    Dim dtComponentesContabilizados As DataTable = New AlbaranVentaLinea().Filter(f)
                    If dtComponentesContabilizados.Rows.Count > 0 Then
                        ApplicationService.GenerateError("No se puede eliminar el Kit, ya que alguno de sus componentes está contabilizado.")
                    End If
                End If
            End If
        End If

        Dim updateData As StockUpdateData = ProcessServer.ExecuteTask(Of DataRow, StockUpdateData)(AddressOf EliminarMovimientoLineaAlbaran, data, services)
        If Not updateData Is Nothing Then
            If updateData.Estado <> EstadoStock.Actualizado Then
                'Se deja porque bajaUpdateData.Detalle ya viene traduccido
                If Not updateData.Log Is Nothing AndAlso Length(updateData.Log) > 0 Then Throw New Exception(updateData.Log)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarAlbaranesMultiEmpresa(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim grp As New GRPAlbaranVentaCompraLinea
        Dim control As DataTable = grp.TrazaAVLPrincipal(data("IDLineaAlbaran"))
        If Not control Is Nothing AndAlso control.Rows.Count > 0 Then
            control.Rows(0)("IDAVPrincipal") = DBNull.Value
            control.Rows(0)("NAVPrincipal") = DBNull.Value
            control.Rows(0)("IDLineaAVPrincipal") = DBNull.Value
            BusinessHelper.UpdateTable(control)
        Else
            control = grp.TrazaAVLSecundaria(data("IDLineaAlbaran"))
            If Not control Is Nothing AndAlso control.Rows.Count > 0 Then
                Dim msg As String = "Este albarán está relacionado con un pedido entre empresas del grupo."
                msg = String.Concat(msg, ControlChars.NewLine, "Deberá eliminar el albarán completo.")
                Throw New Exception(msg)
            End If
        End If
    End Sub

    '<Task()> Public Shared Sub ActualizarPuntosMarketing(ByVal data As DataRow, ByVal services As ServiceProvider)
    '    Dim DtParam As DataTable = New Parametro().SelOnPrimaryKey("PUNTOS_IMP")
    '    If Not DtParam Is Nothing AndAlso DtParam.Rows.Count > 0 Then
    '        If DtParam.Rows(0)("Valor") > 0 Then
    '            ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoAlbaranVenta.BorradoPuntosMarketing, data, services)
    '        End If
    '    End If
    'End Sub

    <Task()> Public Shared Sub ActualizarCantidadPromocionada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDPromocionLinea")) > 0 AndAlso data("Regalo") = 0 Then
            ProcessServer.ExecuteTask(Of Integer)(AddressOf PromocionLinea.ActualizarLineaPromocionQ, data("IDPromocionLinea"), services)
        End If
    End Sub

#End Region

#Region " Validate "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ProcesoComercial.DetailCommonUpdateRules)    'Validaciones Generales Comercial 
        'validateProcess.AddTask(Of DataRow)(AddressOf ProcesoAlbaranVenta.ValidarEstadoLinea)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarAlbaranFacturado)
        validateProcess.AddTask(Of DataRow)(AddressOf ProcesoAlbaranVenta.ValidarArticuloBloqueado)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarCantidadLineaAlbaran)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarAlmacenObligatorio)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarFormaPagoObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarCondicionPagoObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf ProcesoAlbaranVenta.ValidarArticuloContenedor)
    End Sub

    <Task()> Public Shared Sub ValidarAlbaranFacturado(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDAlbaran")) <> 0 Then
            Dim Cabecera As DataTable = New AlbaranVentaCabecera().SelOnPrimaryKey(data("IDAlbaran"))
            If Not Cabecera Is Nothing AndAlso Cabecera.Rows.Count > 0 Then
                If Cabecera.Rows(0)("Estado") = enumavcEstadoFactura.avcFacturado Then
                    ApplicationService.GenerateError("El Albarán está Facturado.")
                End If
            End If
        End If
    End Sub

#End Region

#Region " Update "


    'Public Overloads Overrides Function Update(ByVal dttSource As System.Data.DataTable) As System.Data.DataTable
    '    If Not IsNothing(dttSource) AndAlso dttSource.Rows.Count > 0 Then
    '        Dim carrier(-1) As DataTable
    '        Dim com As New Comercial
    '        Dim AVC As New AlbaranVentaCabecera
    '        Dim services As New ServiceProvider
    '        Dim AppParams As ParametroAlbaranVenta = services.GetService(GetType(ParametroAlbaranVenta))
    '        Dim AppParamsConta As ParametroContabilidadVenta = services.GetService(GetType(ParametroContabilidadVenta))

    '        For Each dr As DataRow In dttSource.Rows
    '            If dr.RowState And (DataRowState.Added Or DataRowState.Modified) Then
    '                'TODO: Revisar esto
    '                'com.DetailCommonUpdateRules(dr, services)
    '                If Length(dr(_AVL.IDAlmacen)) = 0 Then ApplicationService.GenerateError("El Almacen es obligatorio.")
    '                If Length(dr(_AVL.IDFormaPago)) = 0 Then ApplicationService.GenerateError("La Forma Pago es obligatoria.")
    '                If Length(dr(_AVL.IDCondicionPago)) = 0 Then ApplicationService.GenerateError("La Condición de Pago es obligatoria.")

    '                If Not IsNumeric(dr(_AVL.QServida)) Then
    '                    ApplicationService.GenerateError("La cantidad no es válida.")
    '                ElseIf dr(_AVL.QServida) = 0 Then
    '                    ApplicationService.GenerateError("La cantidad no puede ser cero.")
    '                End If

    '                Dim Cabecera As DataRow = AVC.GetItemRow(dr(_AVL.IDAlbaran))
    '                If Cabecera("Estado") = enumavcEstadoFactura.avcFacturado Then
    '                    ApplicationService.GenerateError("El Albarán ya está Facturado.")
    '                End If
    '                Dim Analitica As DataTable
    '                Dim Representantes As DataTable
    '                Dim Componentes As DataTable

    '                dr = MantenimientoValoresAyB(dr, Cabecera(_AVC.IDMoneda), Cabecera(_AVC.CambioA), Cabecera(_AVC.CambioB))

    '                If dr.RowState = DataRowState.Added Then
    '                    dr(_AVL.IDDireccionFra) = Cabecera(_AVC.IDDireccionFra)
    '                    dr(_AVL.IDClienteBanco) = Cabecera(_AVC.IDClienteBanco)
    '                    dr(_AVL.EstadoFactura) = Cabecera(_AVC.Estado)
    '                    If IsDBNull(dr(_AVL.IDLineaAlbaran)) OrElse Nz(dr(_AVL.IDLineaAlbaran), 0) = 0 Then
    '                        dr(_AVL.IDLineaAlbaran) = AdminData.GetAutoNumeric
    '                    End If

    '                    If dr(_AVL.EstadoFactura) = enumavlEstadoFactura.avlNoFacturado Then
    '                        If IsDBNull(dr(_AVL.TipoLineaAlbaran)) Then
    '                            dr(_AVL.TipoLineaAlbaran) = enumavlTipoLineaAlbaran.avlNormal
    '                        End If
    '                        dr(_AVL.QPendienteDevolverAInicio) = dr(_AVL.QServida)
    '                        If AppParamsConta.Contabilidad AndAlso AppParamsConta.Analitica.AplicarAnalitica Then
    '                            Analitica = NuevaAnalitica(dr)
    '                        End If
    '                        Representantes = com.NuevoRepresentante(dr, New MonedaCache)

    '                        '//Comprobar si articulo es Kit
    '                        If Not (dr(_AVL.TipoLineaAlbaran) = enumavlTipoLineaAlbaran.avlComponente) Then
    '                            Componentes = Me.ComponentesDePrimerNivel(dr)
    '                            If Not Componentes Is Nothing AndAlso Componentes.Rows.Count > 0 Then
    '                                dr(_AVL.TipoLineaAlbaran) = enumavlTipoLineaAlbaran.avlKit
    '                            End If
    '                        End If

    '                        dr(_AVL.EstadoStock) = enumavlEstadoStock.avlNoActualizado
    '                        If AppParams.ActualizacionAutomaticaStock Then
    '                            Dim updateData() As StockUpdateData
    '                            updateData = ProcesoAlbaranVenta.ActualizarStock(Cabecera, dr)
    '                            If Not Componentes Is Nothing AndAlso Componentes.Rows.Count > 0 Then
    '                                For Each componente As DataRow In Componentes.Rows
    '                                    updateData = ProcesoAlbaranVenta.ActualizarStock(Cabecera, componente)
    '                                Next
    '                            End If
    '                            If Cabecera.RowState = DataRowState.Modified And Cabecera("NMovimiento") <> Cabecera("NMovimiento", DataRowVersion.Original) Then
    '                                AdminData.SetData(Cabecera.Table)
    '                            End If
    '                        End If

    '                        '///Lotes: Lineas de albaran generadas a partir de otras lineas de albaran, que a su vez provienen de un pedido
    '                        If Length(dr(_AVL.Lote)) > 0 And IsNumeric(dr(_AVL.IDLineaPedido)) Then
    '                            ActualizarLineaPedido(Cabecera(_AVC.IDTipoAlbaran) & String.Empty, dr)
    '                        End If
    '                    Else
    '                        If dr(_AVL.EstadoFactura) = enumavlEstadoFactura.avlParcFacturado Then ApplicationService.GenerateError("No se puede borrar el Albarán. Está parcialmente facturado.")
    '                        If dr(_AVL.EstadoFactura) = enumavlEstadoFactura.avlFacturado Then ApplicationService.GenerateError("No se puede borrar el Albarán. Está total o parcialmente facturado.")
    '                    End If
    '                ElseIf dr.RowState = DataRowState.Modified Then
    '                    If Not Nz(Cabecera(_AVC.Automatico), False) Then
    '                        If Not (dr(_AVL.EstadoFactura, DataRowVersion.Original) = enumaclEstadoFactura.aclNoFacturado) Then
    '                            If dr(_AVL.IDArticulo) <> dr(_AVL.IDArticulo, DataRowVersion.Original) Or dr(_AVL.QServida) <> dr(_AVL.QServida, DataRowVersion.Original) Then
    '                                If dr(_AVL.EstadoFactura, DataRowVersion.Original) = enumaclEstadoFactura.aclParcFacturado Then ApplicationService.GenerateError("No se puede modificar el Albaráan. Está parcialmente facturado.")
    '                                If dr(_AVL.EstadoFactura, DataRowVersion.Original) = enumaclEstadoFactura.aclFacturado Then ApplicationService.GenerateError("No se puede modificar el Albarán. Está facturado.")
    '                            End If
    '                        End If
    '                    End If
    '                    If Length(dr(_AVL.IDArticuloContenedor, DataRowVersion.Original)) > 0 Then
    '                        If (Length(dr(_AVL.IDArticuloContenedor)) = 0 AndAlso Length(dr(_AVL.IDArticuloContenedor, DataRowVersion.Original)) > 0) _
    '                        OrElse (dr(_AVL.IDArticuloContenedor) <> dr(_AVL.IDArticuloContenedor, DataRowVersion.Original) _
    '                        AndAlso Length(dr(_AVL.IDEntradaContenedor)) > 0 AndAlso Length(dr(_AVL.IDSalidaContenedor)) > 0) Then
    '                            ApplicationService.GenerateError("No se puede modificar el Articulo Contenedor, se ha actualizado ya el stock del contenedor.")
    '                        End If
    '                    End If

    '                    dr(_AVL.QPendienteDevolverAInicio) = dr(_AVL.QServida)
    '                    If AppParamsConta.Contabilidad AndAlso AppParamsConta.Analitica.AplicarAnalitica Then
    '                        Analitica = ActualizarAnalitica(dr)
    '                    End If
    '                    Representantes = com.ActualizarRepresentantes(dr)

    '                    If dr(_AVL.TipoLineaAlbaran) = enumavlTipoLineaAlbaran.avlKit Then
    '                        Componentes = ProcesoAlbaranVenta.ActualizarComponentes(dr)
    '                    End If

    '                    If dr(_AVL.QServida, DataRowVersion.Original) <> dr(_AVL.QServida) _
    '                    Or dr(_AVL.QInterna, DataRowVersion.Original) <> dr(_AVL.QInterna) _
    '                    Or dr(_AVL.ImporteA, DataRowVersion.Original) <> dr(_AVL.ImporteA) _
    '                    Or dr(_AVL.ImporteB, DataRowVersion.Original) <> dr(_AVL.ImporteB) _
    '                    Or dr(_AVL.Precio, DataRowVersion.Original) <> dr(_AVL.Precio) _
    '                    Or dr(_AVL.QEtiContenedor, DataRowVersion.Original) <> dr(_AVL.QEtiContenedor) Then
    '                        If dr(_AVL.Precio) <> dr(_AVL.Precio, DataRowVersion.Original) _
    '                        And dr(_AVL.QServida) = dr(_AVL.QServida, DataRowVersion.Original) _
    '                        And dr(_AVL.QInterna) = dr(_AVL.QInterna, DataRowVersion.Original) _
    '                        And dr(_AVL.QEtiContenedor) = dr(_AVL.QEtiContenedor, DataRowVersion.Original) Then
    '                            Dim DtArt As DataTable = New Articulo().SelOnPrimaryKey(dr(_AVL.IDArticulo))
    '                            If Not DtArt Is Nothing AndAlso DtArt.Rows.Count > 0 Then
    '                                If DtArt.Rows(0)("RecalcularValoracion") = CInt(enumtaValoracionSalidas.taMantenerPrecio) Then
    '                                    Dim updateData As StockUpdateData = Me.CorregirMovimiento(dr)
    '                                    If updateData Is Nothing Then
    '                                        If dr(_AVL.EstadoStock) = EstadoStock.Actualizado Or dr(_AVL.EstadoStock) = EstadoStock.SinGestion Then
    '                                            If Not Componentes Is Nothing AndAlso Componentes.Rows.Count > 0 Then
    '                                                Me.CorregirMovimiento(Componentes)
    '                                            End If
    '                                        End If
    '                                    ElseIf updateData.Estado = EstadoStock.NoActualizado Then
    '                                        Throw New Exception(updateData.Detalle)
    '                                    End If
    '                                End If
    '                            End If
    '                        Else
    '                            Dim updateData As StockUpdateData = Me.CorregirMovimiento(dr)
    '                            If updateData Is Nothing Then
    '                                If dr(_AVL.EstadoStock) = EstadoStock.Actualizado Or dr(_AVL.EstadoStock) = EstadoStock.SinGestion Then
    '                                    If Not Componentes Is Nothing AndAlso Componentes.Rows.Count > 0 Then
    '                                        Me.CorregirMovimiento(Componentes)
    '                                    End If
    '                                End If
    '                            ElseIf updateData.Estado = EstadoStock.NoActualizado Then
    '                                Throw New Exception(updateData.Detalle)
    '                            End If
    '                        End If
    '                    End If

    '                    If AppParams.ActualizacionAutomaticaStock Then
    '                        If (Length(dr("Lote", DataRowVersion.Original)) = 0 AndAlso Length(dr("Lote")) > 0) OrElse _
    '                        (Length(dr("Lote", DataRowVersion.Original)) > 0 AndAlso Length(dr("Lote")) > 0 AndAlso _
    '                        dr("Lote", DataRowVersion.Original) <> dr("Lote")) Then
    '                            Dim updateData() As StockUpdateData = ProcesoAlbaranVenta.ActualizarStock(Cabecera, dr)
    '                            If Not Componentes Is Nothing AndAlso Componentes.Rows.Count > 0 Then
    '                                For Each componente As DataRow In Componentes.Rows
    '                                    updateData = ProcesoAlbaranVenta.ActualizarStock(Cabecera, componente)
    '                                Next
    '                            End If
    '                        End If
    '                    End If

    '                    If Not IsDBNull(dr(_AVL.IDLineaPedido)) And Not (dr(_AVL.EstadoFactura) = enumavlEstadoFactura.avlFacturado) Then
    '                        If Nz(dr(_AVL.QServida), 0) <> Nz(dr(_AVL.QServida, DataRowVersion.Original), 0) Then
    '                            ActualizarLineaPedido(Cabecera(_AVC.IDTipoAlbaran) & String.Empty, dr)
    '                        End If
    '                    End If
    '                End If

    '                If Not IsNothing(Analitica) Then
    '                    ReDim Preserve carrier(UBound(carrier) + 1) : carrier(UBound(carrier)) = Analitica
    '                    If Not Analitica Is Nothing Then Analitica.Dispose()
    '                End If
    '                If Not IsNothing(Representantes) Then
    '                    ReDim Preserve carrier(UBound(carrier) + 1) : carrier(UBound(carrier)) = Representantes
    '                    If Not Representantes Is Nothing Then Representantes.Dispose()
    '                End If
    '                If Not IsNothing(Componentes) Then
    '                    ReDim Preserve carrier(UBound(carrier) + 1) : carrier(UBound(carrier)) = Componentes
    '                    If Not Componentes Is Nothing Then Componentes.Dispose()
    '                End If

    '            End If
    '        Next
    '        Dim dtAVL As DataTable = dttSource.Copy
    '        AdminData.SetData(dttSource)
    '        AdminData.SetData(carrier)
    '        Me.Updated(dtAVL)
    '    End If

    '    Return dttSource
    'End Function

    'Friend Function ActualizarComponentes(ByVal lineaAlbaran As DataRow) As DataTable
    '    Dim f As New Filter
    '    f.Add(New NumberFilterItem(_AVL.IDAlbaran, FilterOperator.Equal, lineaAlbaran(_AVL.IDAlbaran)))
    '    f.Add(New NumberFilterItem(_AVL.IDLineaPadre, FilterOperator.Equal, lineaAlbaran(_AVL.IDLineaAlbaran)))
    '    f.Add(New NumberFilterItem(_AVL.TipoLineaAlbaran, FilterOperator.Equal, enumavlTipoLineaAlbaran.avlComponente))

    '    Dim Componentes As DataTable = Me.Filter(f)
    '    If Not Componentes Is Nothing AndAlso Componentes.Rows.Count > 0 Then
    '        If Nz(lineaAlbaran(_AVL.QInterna, DataRowVersion.Original), 0) <> 0 Then
    '            Dim factorVariacion As Double = lineaAlbaran(_AVL.QInterna) / lineaAlbaran(_AVL.QInterna, DataRowVersion.Original)
    '            For Each componente As DataRow In Componentes.Rows
    '                componente(_AVL.QServida) = componente(_AVL.QServida) * factorVariacion
    '                componente(_AVL.QInterna) = componente(_AVL.QInterna) * factorVariacion
    '            Next
    '        End If
    '    End If

    '    AdminData.SetData(Componentes)
    '    Return Componentes
    'End Function


#End Region

#Region " BusinessRule "

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim services As New ServiceProvider
        Dim oBRL As New SynonymousBusinessRules
        oBRL.AddSynonymous("QServida", "Cantidad")

        '//BusinessRules - Genéricas del circuito de comercial
        oBRL = ProcessServer.ExecuteTask(Of BusinessRules, BusinessRules)(AddressOf ProcesoComercial.DetailBusinessRulesLin, oBRL, services)

        '//BusinessRules - Específicas AVL  
        oBRL("IDArticulo") = AddressOf CambioArticuloAlbaran     'Específica AVL
        oBRL("CodigoBarras") = AddressOf CambioArticuloAlbaran   'Específica AVL
        oBRL("RefCliente") = AddressOf CambioArticuloAlbaran     'Específica AVL
        oBRL("Cantidad") = AddressOf CambioCantidadAlbaran       'Específica AVL
        oBRL.Add("IDArticuloContenedor", AddressOf CambioArticuloContenedorEmbalaje)
        oBRL.Add("IDArticuloEmbalaje", AddressOf CambioArticuloContenedorEmbalaje)
        oBRL.Add("Lote", AddressOf CambioNSerie)
        oBRL.Add("IDCondicionPago", AddressOf ProcesoComunes.CambioCondicionPagoLineas)
        oBRL("NObra") = AddressOf CambioNObra
        Return oBRL
    End Function

    <Task()> Public Shared Sub CambioArticuloAlbaran(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ProcesoComercial.CambioArticulo, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ValidarArticuloCantidad, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf RecalcularEtiquetas, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf InicializarEstadoStock, data, services)
    End Sub

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
            'Búsqueda por RefCliente
            Dim DtRef As DataTable = New ArticuloCliente().Filter(New FilterItem("RefCliente", FilterOperator.Equal, data))
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
            ElseIf Not ArtInfo.Venta Then
                ApplicationService.GenerateError("El artículo | no es de tipo venta.", Quoted(data))
            Else
                DataReturn.IDArticulo = data
            End If
        End If
        Return DataReturn
    End Function

    <Task()> Public Shared Sub ValidarArticuloCantidad(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If data.ColumnName = "QInterna" Then data.Current(data.ColumnName) = data.Value
        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
        Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.Current("IDArticulo"))
        If ArtInfo.NSerieObligatorio AndAlso data.Current("QInterna") <> 1 And Length(data.Current("Lote")) > 0 Then
            ApplicationService.GenerateError("La cantidad interna debe ser la unidad para un artículo con número de serie.")
        End If
    End Sub
    <Task()> Public Shared Sub CambioCantidadAlbaran(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ProcesoComercial.CambioCantidad, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ValidarArticuloCantidad, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf RecalcularEtiquetas, data, services)
    End Sub

    <Task()> Public Shared Sub RecalcularEtiquetas(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If data.ColumnName = "IDArticulo" OrElse data.ColumnName = "Cantidad" Then data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDArticulo")) > 0 AndAlso Nz(data.Current("Cantidad"), 0) > 0 Then

            Dim ArtCltes As EntityInfoCache(Of ArticuloClienteInfo) = services.GetService(Of EntityInfoCache(Of ArticuloClienteInfo))()
            Dim ArtClteInfo As ArticuloClienteInfo = ArtCltes.GetEntity(data.Context("IDCliente"), data.Current("IDArticulo"))
            If Not ArtClteInfo Is Nothing AndAlso Length(ArtClteInfo.IDArticulo) > 0 AndAlso Length(ArtClteInfo.IDCliente) > 0 Then
                'Dim oRw As DataRow = dtArticuloCliente.Rows(0)
                If ArtClteInfo.QContenedor > 0 OrElse ArtClteInfo.QEmbalaje > 0 Then
                    data.Current("IDArticuloContenedor") = ArtClteInfo.IDArticuloContenedor
                    data.Current("IDArticuloEmbalaje") = ArtClteInfo.IDArticuloEmbalaje
                    If ArtClteInfo.QContenedor > 0 Then
                        data.Current("QEtiContenedor") = Math.Ceiling(Nz(data.Current("Cantidad"), 0) / ArtClteInfo.QContenedor)
                        'Dim intEntero As Integer = Decimal.Truncate(Nz(data.Current("Cantidad"), 0) / ArtClteInfo.QContenedor)
                        'Dim intResto As Integer = Nz(data.Current("Cantidad"), 0) Mod ArtClteInfo.QContenedor
                        'If intResto > 0 Then
                        '    data.Current("QEtiContenedor") = intEntero + 1
                        'Else
                        '    data.Current("QEtiContenedor") = intEntero
                        'End If
                    End If
                    If ArtClteInfo.QEmbalaje > 0 Then
                        data.Current("QEtiEmbalaje") = Math.Ceiling(Nz(data.Current("Cantidad"), 0) / ArtClteInfo.QEmbalaje)
                        'Dim intEntero As Integer = Decimal.Truncate(Nz(data.Current("Cantidad"), 0) / ArtClteInfo.QEmbalaje)
                        'Dim intResto As Integer = Nz(data.Current("Cantidad"), 0) Mod ArtClteInfo.QEmbalaje
                        'If intResto > 0 Then
                        '    data.Current("QEtiEmbalaje") = intEntero + 1
                        'Else
                        '    data.Current("QEtiEmbalaje") = intEntero
                        'End If
                    End If
                Else
                    Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(GetType(EntityInfoCache(Of ArticuloInfo)))
                    Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.Current("IDArticulo"))

                    data.Current("IDArticuloContenedor") = ArtInfo.IDArticuloContenedor
                    data.Current("IDArticuloEmbalaje") = ArtInfo.IDArticuloEmbalaje
                    If ArtInfo.QContenedor > 0 Then
                        data.Current("QEtiContenedor") = Math.Ceiling(Nz(data.Current("Cantidad"), 0) / ArtInfo.QContenedor)
                        'Dim intEntero As Integer = Decimal.Truncate(Nz(data.Current("Cantidad"), 0) / ArtInfo.QContenedor)
                        'Dim intResto As Integer = Nz(data.Current("Cantidad"), 0) Mod ArtInfo.QContenedor
                        'If intResto > 0 Then
                        '    data.Current("QEtiContenedor") = intEntero + 1
                        'Else
                        '    data.Current("QEtiContenedor") = intEntero
                        'End If
                    End If
                    If ArtInfo.QEmbalaje > 0 Then
                        data.Current("QEtiEmbalaje") = Math.Ceiling(Nz(data.Current("Cantidad"), 0) / ArtInfo.QEmbalaje)
                        'Dim intEntero As Integer = Decimal.Truncate(Nz(data.Current("Cantidad"), 0) / ArtInfo.QEmbalaje)
                        'Dim intResto As Integer = Nz(data.Current("Cantidad"), 0) Mod ArtInfo.QEmbalaje
                        'If intResto > 0 Then
                        '    data.Current("QEtiEmbalaje") = intEntero + 1
                        'Else
                        '    data.Current("QEtiEmbalaje") = intEntero
                        'End If
                    End If

                End If
            Else '//No hay un ArticuloCliente
                Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
                Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.Current("IDArticulo"))

                data.Current("IDArticuloContenedor") = ArtInfo.IDArticuloContenedor
                data.Current("IDArticuloEmbalaje") = ArtInfo.IDArticuloEmbalaje
                If ArtInfo.QContenedor > 0 Then
                    data.Current("QEtiContenedor") = Math.Ceiling(Nz(data.Current("Cantidad"), 0) / ArtInfo.QContenedor)
                    'Dim intEntero As Integer = Decimal.Truncate(Nz(data.Current("Cantidad"), 0) / ArtInfo.QContenedor)
                    'Dim intResto As Integer = Nz(data.Current("Cantidad"), 0) Mod ArtInfo.QContenedor
                    'If intResto > 0 Then
                    '    data.Current("QEtiContenedor") = intEntero + 1
                    'Else
                    '    data.Current("QEtiContenedor") = intEntero
                    'End If
                End If
                If ArtInfo.QEmbalaje > 0 Then
                    data.Current("QEtiEmbalaje") = Math.Ceiling(Nz(data.Current("Cantidad"), 0) / ArtInfo.QEmbalaje)
                    'Dim intEntero As Integer = Math.Ceiling(Nz(data.Current("Cantidad"), 0) / ArtInfo.QEmbalaje)
                    'Dim intResto As Double = Nz(data.Current("Cantidad"), 0) Mod ArtInfo.QEmbalaje
                    'If intResto > 0 Then
                    '    data.Current("QEtiEmbalaje") = intEntero + 1
                    'Else
                    '    data.Current("QEtiEmbalaje") = intEntero
                    'End If
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub InicializarEstadoStock(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
        Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.Current("IDArticulo"))
        If ArtInfo.GestionStock Then
            If ArtInfo.Fantasma Then
                data.Current("EstadoStock") = CInt(enumavlEstadoStock.avlSinGestion)
            Else
                data.Current("EstadoStock") = CInt(enumavlEstadoStock.avlNoActualizado)
            End If
        Else
            data.Current("EstadoStock") = CInt(enumavlEstadoStock.avlSinGestion)
        End If
        If Not data.Context Is Nothing AndAlso data.Context.ContainsKey("IDTipoAlbaran") AndAlso Length(data.Context("IDTipoAlbaran")) > 0 Then
            Dim TipoAlbInfo As ProcesoAlbaranVenta.TipoAlbaranInfo = ProcessServer.ExecuteTask(Of String, ProcesoAlbaranVenta.TipoAlbaranInfo)(AddressOf ProcesoAlbaranVenta.TipoDeAlbaran, data.Context("IDTipoAlbaran"), services)
            If TipoAlbInfo.Tipo = enumTipoAlbaran.ExpedDistribuidor OrElse TipoAlbInfo.Tipo = enumTipoAlbaran.AbonoDistribuidor Then
                data.Current("EstadoStock") = CInt(enumavlEstadoStock.avlSinGestion)
            End If
        End If
    End Sub


    <Task()> Public Shared Sub CambioArticuloContenedorEmbalaje(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current(data.ColumnName)) > 0 Then
            Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
            Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.Current(data.ColumnName))
            If Not ArtInfo.Embalaje Then
                ApplicationService.GenerateError("El artículo | no es de tipo embalaje.", Quoted(data.Current(data.ColumnName)))
            ElseIf Not ArtInfo.Activo Then
                ApplicationService.GenerateError("El artículo | no está activo.", Quoted(data.Current(data.ColumnName)))
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioNSerie(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("Lote")) > 0 AndAlso Not data.Context Is Nothing AndAlso Nz(data.Context("NSerieObligatorio"), False) Then
            Dim ArtNSeries As EntityInfoCache(Of ArticuloNSerieInfo) = services.GetService(Of EntityInfoCache(Of ArticuloNSerieInfo))()
            Dim ArtNSerieInfo As ArticuloNSerieInfo = ArtNSeries.GetEntity(data.Current("IDArticulo"), data.Current("Lote"))
            If Not ArtNSerieInfo Is Nothing AndAlso Length(ArtNSerieInfo.IDArticulo) > 0 AndAlso Length(ArtNSerieInfo.NSerie) > 0 Then
                data.Current("IDOperario") = ArtNSerieInfo.IDOperario
                data.Current("IDEstadoActivo") = ArtNSerieInfo.IDEstadoActivo
            Else
                ApplicationService.GenerateError("El número de serie indicado no existe o no está asignado a el artículo |.", Quoted(data.Current("IDArticulo")))
            End If

            data.Current("Factor") = 1
        End If
    End Sub

    <Task()> Public Shared Sub CambioNObra(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) = 0 Then data.Current("IDObra") = DBNull.Value
    End Sub

#End Region

#Region " Gestion de stocks "

    <Task()> Public Shared Function EliminarMovimiento(ByVal LineasAlbaran As DataTable, ByVal services As ServiceProvider) As StockUpdateData()
        Dim updateData(-1) As StockUpdateData

        AdminData.BeginTx()
        For Each lineaAlbaran As DataRow In LineasAlbaran.Rows
            Dim data As StockUpdateData = ProcessServer.ExecuteTask(Of DataRow, StockUpdateData)(AddressOf EliminarMovimientoLineaAlbaran, lineaAlbaran, services)
            If Not data Is Nothing Then
                ReDim Preserve updateData(UBound(updateData) + 1)
                updateData(UBound(updateData)) = data
            End If
        Next
        AdminData.CommitTx(True)

        Return updateData
    End Function

    <Task()> Public Shared Function EliminarMovimientoLineaAlbaran(ByVal lineaAlbaran As DataRow, ByVal services As ServiceProvider) As StockUpdateData
        Dim updateData As StockUpdateData
        AdminData.BeginTx()
        If IsNumeric(lineaAlbaran(_AVL.IDEntradaContenedor)) Then
            '//Correccion movimiento de entrada de contenedor
            Dim datElimMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Eliminar, lineaAlbaran(_AVL.IDEntradaContenedor))
            updateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento, StockUpdateData)(AddressOf ProcesoStocks.ActualizarMovimiento, datElimMovto, services)

            If updateData Is Nothing OrElse updateData.Estado <> EstadoStock.Actualizado Then
                AdminData.RollBackTx(True)
                Return updateData
            End If
        End If

        If IsNumeric(lineaAlbaran(_AVL.IDSalidaContenedor)) Then
            '//Correccion movimiento de salida de contenedor
            Dim datElimMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Eliminar, lineaAlbaran(_AVL.IDSalidaContenedor))
            updateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento, StockUpdateData)(AddressOf ProcesoStocks.ActualizarMovimiento, datElimMovto, services)

            If updateData Is Nothing OrElse updateData.Estado <> EstadoStock.Actualizado Then
                AdminData.RollBackTx(True)
                Return updateData
            End If
        End If

        If IsNumeric(lineaAlbaran(_AVL.IDMovimientoEntrada)) Then
            '//Correccion movimiento de entrada
            Dim datElimMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Eliminar, lineaAlbaran(_AVL.IDMovimientoEntrada))
            updateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento, StockUpdateData)(AddressOf ProcesoStocks.ActualizarMovimiento, datElimMovto, services)

            If updateData Is Nothing OrElse updateData.Estado <> EstadoStock.Actualizado Then
                AdminData.RollBackTx(True)
                Return updateData
            End If
        End If

        If IsNumeric(lineaAlbaran(_AVL.IDMovimiento)) Then
            '//Correccion movimiento de salida
            Dim datElimMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Eliminar, lineaAlbaran(_AVL.IDMovimiento))
            updateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento, StockUpdateData)(AddressOf ProcesoStocks.ActualizarMovimiento, datElimMovto, services)
            If updateData Is Nothing OrElse updateData.Estado <> EstadoStock.Actualizado Then
                AdminData.RollBackTx(True)

                Return updateData
            End If
        End If

        AdminData.CommitTx(True)
        Return updateData
    End Function

#End Region

#Region "Otros"

    <Task()> Public Shared Function GetIDAlbaranLinea(ByVal data As Object, ByVal services As ServiceProvider) As Integer
        Return AdminData.GetAutoNumeric
    End Function

#End Region

End Class
