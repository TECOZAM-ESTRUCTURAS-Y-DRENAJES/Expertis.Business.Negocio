Public Class PrcCopiaPedidoCompra
    Inherits Process(Of Integer, CreateElement)

    Public Overrides Sub RegisterTasks()
        Me.AddTask(Of Integer, PedCabCompraCopia)(AddressOf DatosIniciales)
        Me.AddTask(Of PedCabCompraCopia, DocumentoPedidoCompra)(AddressOf ProcesoPedidoCompra.CrearDocumentoPedidoCompra)
        Me.AddTask(Of DocumentoPedidoCompra)(AddressOf CrearDocumentoCopia)
        Me.AddTask(Of DocumentoPedidoCompra)(AddressOf ProcesoPedidoCompra.CalcularAnalitica)
        Me.AddTask(Of DocumentoPedidoCompra)(AddressOf ProcesoPedidoCompra.CalcularBasesImponibles)
        Me.AddTask(Of DocumentoPedidoCompra)(AddressOf Business.General.Comunes.UpdateDocument)
        Me.AddTask(Of DocumentoPedidoCompra)(AddressOf CopiarGastos)
        Me.AddTask(Of DocumentoPedidoCompra, CreateElement)(AddressOf Resultado)
    End Sub

    <Task()> Public Shared Function DatosIniciales(ByVal data As Integer, ByVal services As ServiceProvider) As PedCabCompraCopia
        Dim Dr As DataRow = New PedidoCompraCabecera().GetItemRow(data)
        Dim PC As New PedCabCompraCopia(Dr)
        PC.Origen = enumOrigenPedidoCompra.PedidoCompra
        Return PC
    End Function

    <Task()> Public Shared Sub CrearDocumentoCopia(ByVal data As DocumentoPedidoCompra, ByVal services As ServiceProvider)
        Dim DtCab As DataTable = New PedidoCompraCabecera().SelOnPrimaryKey(CType(data.Cabecera, PedCabCompraCopia).IDPedido)
        Dim DtLineas As DataTable = New PedidoCompraLinea().Filter(New FilterItem("IDPedido", FilterOperator.Equal, CType(data.Cabecera, PedCabCompraCopia).IDPedido))
        Dim AppParamsConta As ParametroContabilidadCompra = services.GetService(GetType(ParametroContabilidadCompra))

        Dim drOrigenCabecera As DataRow = DtCab.Rows(0)
        For Each dc As DataColumn In DtCab.Columns
            If dc.ColumnName <> "IDPedido" And dc.ColumnName <> "NPedido" Then
                data.HeaderRow(dc.ColumnName) = drOrigenCabecera(dc)
            End If
        Next

        data.HeaderRow("IDPedido") = AdminData.GetAutoNumeric
        If Length(data.HeaderRow("IDContador")) > 0 Then
            ProcessServer.ExecuteTask(Of DocumentoPedidoCompra)(AddressOf SetContadorPredeterminadoEntidad, data, services)
            Dim StDatos As New Contador.DatosCounterValue(data.HeaderRow("IDContador"), New PedidoCompraCabecera, "NPedido", "FechaPedido", data.HeaderRow("FechaPedido"))
            StDatos.IDEjercicio = data.HeaderRow("IDEjercicio") & String.Empty
            data.HeaderRow("NPedido") = ProcessServer.ExecuteTask(Of Contador.DatosCounterValue, String)(AddressOf Contador.CounterValue, StDatos, services)
        End If
        data.HeaderRow("FechaPedido") = Date.Today
        data.HeaderRow("FechaEntrega") = Date.Today

        If AppParamsConta.Contabilidad Then
            Dim DataEjer As New DataEjercicio(New DataRowPropertyAccessor(data.HeaderRow), Today.Date)
            ProcessServer.ExecuteTask(Of DataEjercicio)(AddressOf NegocioGeneral.AsignarEjercicioContable, DataEjer, services)
        End If

        'Copia Líneas
        Dim StDataRelLineas As New List(Of DataRelLineas)
        For Each drOrigenLinea As DataRow In DtLineas.Select
            Dim drDestinoLinea As DataRow = data.dtLineas.NewRow
            drDestinoLinea("Estado") = enumpclEstado.pclpedido
            drDestinoLinea("IDPedido") = data.HeaderRow("IDPedido")
            drDestinoLinea("IDLineaPedido") = AdminData.GetAutoNumeric
            StDataRelLineas.Add(New DataRelLineas(drOrigenLinea("IDLineaPedido"), drDestinoLinea("IDLineaPedido")))
            For Each dc As DataColumn In DtLineas.Columns
                If dc.ColumnName <> "IDLineaPedido" And dc.ColumnName <> "IDPedido" And _
                   dc.ColumnName <> "Estado" And dc.ColumnName <> "IDLineaOferta" And _
                   dc.ColumnName <> "IDObra" And dc.ColumnName <> "IDTrabajo" And _
                   dc.ColumnName <> "IDLineaMaterial" And dc.ColumnName <> "IDMntoOTPrev" And _
                   dc.ColumnName <> "IDLineaContrato" And dc.ColumnName <> "IDLineaPrograma" And _
                   dc.ColumnName <> "IDOferta" And dc.ColumnName <> "IDOrdenRuta" And _
                   dc.ColumnName <> "IDPrograma" And dc.ColumnName <> "IDSolicitud" And _
                   dc.ColumnName <> "IDLineaOferta" And dc.ColumnName <> "IDContrato" And _
                   dc.ColumnName <> "IDLineaContrato" And dc.ColumnName <> "IDLineaOfertaDetalle" And _
                   dc.ColumnName <> "QServida" Then
                    drDestinoLinea(dc.ColumnName) = drOrigenLinea(dc)
                End If
            Next
            data.dtLineas.Rows.Add(drDestinoLinea)
        Next
        If StDataRelLineas.Count > 0 Then
            Dim StData As New DataPrcCopInfo(StDataRelLineas, DtLineas)
            services.RegisterService(StData)
        End If
    End Sub

    <Task()> Public Shared Sub CopiarGastos(ByVal data As DocumentoPedidoCompra, ByVal services As ServiceProvider)
        Dim StDataLineas As DataPrcCopInfo = services.GetService(Of DataPrcCopInfo)()
        If Not StDataLineas Is Nothing AndAlso StDataLineas.DataLineas.Count > 0 AndAlso StDataLineas.DtLineasOrigen.Rows.Count > 0 Then
            Dim ClsPedCompraPrecio As New PedidoCompraPrecio
            Dim DtPreciosDestino As DataTable = ClsPedCompraPrecio.AddNew
            For Each DrLinea As DataRow In StDataLineas.DtLineasOrigen.Select
                Dim DtPreciosOrigen As DataTable = ClsPedCompraPrecio.Filter(New FilterItem("IDLineaPedido", FilterOperator.Equal, DrLinea("IDLineaPedido")))
                If Not DtPreciosOrigen Is Nothing AndAlso DtPreciosOrigen.Rows.Count > 0 Then
                    For Each DrLineaPrecio As DataRow In DtPreciosOrigen.Select
                        Dim DrNewLineaDestino As DataRow = DtPreciosDestino.NewRow
                        DrNewLineaDestino.ItemArray = DrLineaPrecio.ItemArray
                        Dim StLinFind As List(Of Integer) = (From StData As DataRelLineas In StDataLineas.DataLineas Where StData.IDLineaOrigen = CInt(DrLineaPrecio("IDLineaPedido")) Select StData.IDLineaNuevo).ToList
                        If StLinFind.Count > 0 Then DrNewLineaDestino("IDLineaPedido") = StLinFind(0)
                        Dim StLinHijaFind As List(Of Integer) = (From StData As DataRelLineas In StDataLineas.DataLineas Where StData.IDLineaOrigen = CInt(Nz(DrLineaPrecio("IDLineaPedidoHija"), 0)) Select StData.IDLineaNuevo).ToList
                        If StLinHijaFind.Count > 0 Then DrNewLineaDestino("IDLineaPedidoHija") = StLinHijaFind(0)
                        DtPreciosDestino.Rows.Add(DrNewLineaDestino)
                    Next
                End If
            Next
            If Not DtPreciosDestino Is Nothing AndAlso DtPreciosDestino.Rows.Count > 0 Then
                ClsPedCompraPrecio.Update(DtPreciosDestino)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub SetContadorPredeterminadoEntidad(ByVal Doc As DocumentoPedidoCompra, ByVal services As ServiceProvider)
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

    <Task()> Public Shared Function Resultado(ByVal Doc As DocumentoPedidoCompra, ByVal services As ServiceProvider) As CreateElement
        Dim result As New CreateElement
        result.IDElement = Doc.HeaderRow("IDPedido")
        result.NElement = Doc.HeaderRow("NPedido")
        Return result
    End Function

End Class

<Serializable()> _
Public Class DataRelLineas
    Public IDLineaOrigen As Integer
    Public IDLineaNuevo As Integer

    Public Sub New()
    End Sub
    Public Sub New(ByVal IDLineaOrigen As Integer, ByVal IDLineaNuevo As Integer)
        Me.IDLineaOrigen = IDLineaOrigen
        Me.IDLineaNuevo = IDLineaNuevo
    End Sub
End Class

<Serializable()> _
Public Class DataPrcCopInfo
    Public DataLineas As List(Of DataRelLineas)
    Public DtLineasOrigen As DataTable

    Public Sub New()
    End Sub
    Public Sub New(ByVal Datalineas As List(Of DataRelLineas), ByVal DtLineasOrigen As DataTable)
        Me.DataLineas = Datalineas
        Me.DtLineasOrigen = DtLineasOrigen
    End Sub
End Class