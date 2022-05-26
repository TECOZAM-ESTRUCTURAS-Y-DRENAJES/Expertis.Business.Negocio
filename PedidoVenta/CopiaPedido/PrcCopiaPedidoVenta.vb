Public Class PrcCopiaPedidoVenta
    Inherits Process(Of Integer, CreateElement)

    'Lista de tareas a ejecutar
    Public Overrides Sub RegisterTasks()
        Me.AddTask(Of Integer, PedCabVentaCopia)(AddressOf DatosIniciales)
        Me.AddTask(Of PedCabVentaCopia, DocumentoPedidoVenta)(AddressOf ProcesoPedidoVenta.CrearDocumentoPedidoVenta)
        Me.AddTask(Of DocumentoPedidoVenta)(AddressOf CrearDocumentoCopia)
        Me.AddTask(Of DocumentoPedidoVenta)(AddressOf ProcesoPedidoVenta.CalcularImporteLineasPedido)
        Me.AddTask(Of DocumentoPedidoVenta)(AddressOf ProcesoPedidoVenta.CalcularRepresentantes)
        Me.AddTask(Of DocumentoPedidoVenta)(AddressOf ProcesoPedidoVenta.CalcularAnalitica)
        Me.AddTask(Of DocumentoPedidoVenta)(AddressOf ProcesoPedidoVenta.CalcularBasesImponibles)
        Me.AddTask(Of DocumentoPedidoVenta)(AddressOf ProcesoComunes.TotalDocumento)
        Me.AddTask(Of DocumentoPedidoVenta)(AddressOf Business.General.Comunes.UpdateDocument)
        Me.AddTask(Of DocumentoPedidoVenta, CreateElement)(AddressOf Resultado)
    End Sub

    <Task()> Public Shared Function DatosIniciales(ByVal data As Integer, ByVal services As ServiceProvider) As PedCabVentaCopia
        Dim dr As DataRow = New PedidoVentaCabecera().GetItemRow(data)
        Dim PVC As New PedCabVentaCopia(dr)
        Return PVC
    End Function

    <Task()> Public Shared Sub CrearDocumentoCopia(ByVal data As DocumentoPedidoVenta, ByVal services As ServiceProvider)
        Dim dtCab As DataTable = New PedidoVentaCabecera().SelOnPrimaryKey(CType(data.Cabecera, PedCabVentaCopia).IDPedido)
        Dim dtLineas As DataTable = New PedidoVentaLinea().Filter(New FilterItem("IDPedido", FilterOperator.Equal, CType(data.Cabecera, PedCabVentaCopia).IDPedido))
        Dim AppParamsConta As ParametroContabilidadVenta = services.GetService(GetType(ParametroContabilidadVenta))

        Dim drOrigenCabecera As DataRow = DtCab.Rows(0)
        For Each dc As DataColumn In DtCab.Columns
            If dc.ColumnName <> "IDPedido" AndAlso dc.ColumnName <> "NPedido" AndAlso _
               dc.ColumnName <> "IDEstadoPedido" AndAlso dc.ColumnName <> "PedidoCliente" AndAlso _
               dc.ColumnName <> "GastosEnvio" AndAlso dc.ColumnName <> "ImpPedido" AndAlso _
               dc.ColumnName <> "ImpPedidoA" AndAlso dc.ColumnName <> "ImpPedidoB" AndAlso _
               dc.ColumnName <> "IDEstadoPedido" AndAlso dc.ColumnName <> "Prioridad" AndAlso _
               dc.ColumnName <> "IdCentroSolicitante" AndAlso dc.ColumnName <> "PedidoInterno" AndAlso _
               dc.ColumnName <> "ImpTotal" AndAlso dc.ColumnName <> "ImpTotalA" AndAlso _
               dc.ColumnName <> "ImpTotalB" AndAlso dc.ColumnName <> "ImpIva" AndAlso _
               dc.ColumnName <> "ImpIvaA" AndAlso dc.ColumnName <> "ImpIvaB" AndAlso _
               dc.ColumnName <> "ImpRE" AndAlso dc.ColumnName <> "ImpREA" AndAlso _
               dc.ColumnName <> "ImpREB" AndAlso dc.ColumnName <> "ImpDto" AndAlso _
               dc.ColumnName <> "ImpDtoA" AndAlso dc.ColumnName <> "ImpDtoB" AndAlso _
               dc.ColumnName <> "PedidoClienteDestino" AndAlso dc.ColumnName <> "FechaAviso" AndAlso _
               dc.ColumnName <> "FechaPreparacion" AndAlso dc.ColumnName <> "PedidoInterno" AndAlso _
               dc.ColumnName <> "ImpDpp" AndAlso dc.ColumnName <> "ImpDppA" AndAlso _
               dc.ColumnName <> "ImpDppB" AndAlso dc.ColumnName <> "ImpRecFinan" AndAlso _
               dc.ColumnName <> "ImpRecFinanA" AndAlso dc.ColumnName <> "ImpRecFinanB" AndAlso _
               dc.ColumnName <> "BaseImponible" AndAlso dc.ColumnName <> "BaseImponibleA" AndAlso _
               dc.ColumnName <> "BaseImponibleB" AndAlso dc.ColumnName <> "ImportePorte" Then
                data.HeaderRow(dc.ColumnName) = drOrigenCabecera(dc)
            End If
        Next

        data.HeaderRow("IDPedido") = AdminData.GetAutoNumeric

        If Length(data.HeaderRow("IDContador")) = 0 Then
            ProcessServer.ExecuteTask(Of DocumentoPedidoVenta)(AddressOf SetContadorPredeterminadoEntidad, data, services)
        End If

        If Length(data.HeaderRow("IDContador")) > 0 Then
            Dim StDatos As New Contador.DatosCounterValue(data.HeaderRow("IDContador"), New PedidoVentaCabecera, "NPedido", "FechaPedido", data.HeaderRow("FechaPedido"))
            StDatos.IDEjercicio = data.HeaderRow("IDEjercicio") & String.Empty
            data.HeaderRow("NPedido") = ProcessServer.ExecuteTask(Of Contador.DatosCounterValue, String)(AddressOf Contador.CounterValue, StDatos, services)
        End If
        data.HeaderRow("FechaPedido") = Date.Today
        data.HeaderRow("FechaEntrega") = Date.Today
        data.HeaderRow("Estado") = enumpvcEstado.pvcPedido


        If AppParamsConta.Contabilidad Then
            Dim DataEjer As New DataEjercicio(New DataRowPropertyAccessor(data.HeaderRow), Today.Date)
            ProcessServer.ExecuteTask(Of DataEjercicio)(AddressOf NegocioGeneral.AsignarEjercicioContable, DataEjer, services)
        End If

        'Copia Líneas
        For Each drOrigenLinea As DataRow In dtLineas.Select
            Dim drDestinoLinea As DataRow = data.dtLineas.NewRow
            For Each dc As DataColumn In dtLineas.Columns
                If dc.ColumnName <> "IDLineaPedido" And dc.ColumnName <> "IDPedido" And _
                   dc.ColumnName <> "IDAlbaran" And dc.ColumnName <> "IDLineaAlbaran" And _
                   dc.ColumnName <> "FechaEntrega" And dc.ColumnName <> "PedidoCliente" And _
                   dc.ColumnName <> "IDPrograma" And dc.ColumnName <> "IDLineaPrograma" And _
                   dc.ColumnName <> "IDPromocionLinea" And dc.ColumnName <> "IdOrdenLinea" And _
                   dc.ColumnName <> "IdLineaPedidoCompra" And dc.ColumnName <> "QFacturada" And _
                   dc.ColumnName <> "PedidoVentaOrigen" And dc.ColumnName <> "FechaPreparacion" And _
                   dc.ColumnName <> "PedidoClienteDestino" And dc.ColumnName <> "IDLineaOfertaDetalle" And _
                    dc.ColumnName <> "IDPromocion" And dc.ColumnName <> "IDCertificacion" Then
                    drDestinoLinea(dc.ColumnName) = drOrigenLinea(dc)
                End If
            Next

            drDestinoLinea("FechaEntrega") = data.HeaderRow("FechaEntrega")
            drDestinoLinea("Estado") = enumpvlEstado.pvlPedido
            drDestinoLinea("QServida") = 0
            drDestinoLinea("QAlbaran") = 0
            drDestinoLinea("QDisponible") = 0
            drDestinoLinea("QFacturada") = 0
            drDestinoLinea("QTramitada") = 0

            drDestinoLinea("Confirmado") = False
            drDestinoLinea("PreparadoExp") = False


            drDestinoLinea("IDPedido") = data.HeaderRow("IDPedido")
            drDestinoLinea("IDLineaPedido") = AdminData.GetAutoNumeric

            data.dtLineas.Rows.Add(drDestinoLinea)
        Next
    End Sub

    <Task()> Public Shared Sub SetContadorPredeterminadoEntidad(ByVal Doc As DocumentoPedidoVenta, ByVal services As ServiceProvider)
        Dim dtCount As DataTable = New Contador().SelOnPrimaryKey(Doc.HeaderRow("IDContador"))
        If dtCount.Rows.Count = 0 Then
            Dim fContPred As New Filter
            fContPred.Add(New StringFilterItem("Entidad", Doc.EntidadCabecera))
            fContPred.Add(New BooleanFilterItem("Predeterminado", True))
            Dim dtContPred As DataTable = New EntidadContador().Filter(fContPred)
            If dtContPred.Rows.Count > 0 Then
                Doc.HeaderRow("IDContador") = dtContPred.Rows(0)("IDContador")
            End If
        End If
    End Sub

    <Task()> Public Shared Function Resultado(ByVal Doc As DocumentoPedidoVenta, ByVal services As ServiceProvider) As CreateElement
        Dim result As New CreateElement
        result.IDElement = Doc.HeaderRow("IDPedido")
        result.NElement = Doc.HeaderRow("NPedido")
        Return result
    End Function


End Class
