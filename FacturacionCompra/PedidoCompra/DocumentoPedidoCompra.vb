Public Class DocumentoPedidoCompra
    Inherits DocumentoCompra

    Public Sub New(ByVal Cabecera As PedCabCompra, ByVal services As ServiceProvider)
        MyBase.New(Cabecera, services)
    End Sub

    Public Sub New(ByVal UpdtCtx As Engine.BE.UpdatePackage)
        MyBase.New(UpdtCtx)
    End Sub

    Public Sub New(ByVal ParamArray PrimaryKey() As Object)
        MyBase.New(PrimaryKey)
    End Sub

    Public Overrides Function EntidadAnalitica() As String
        Return GetType(PedidoCompraAnalitica).Name
    End Function

    Public Overrides Function EntidadCabecera() As String
        Return GetType(PedidoCompraCabecera).Name
    End Function

    Public Overrides Function EntidadLineas() As String
        Return GetType(PedidoCompraLinea).Name
    End Function


    Public Overrides Property Fecha(Optional ByVal OriginalValue As DataRowVersion = DataRowVersion.Current) As Date
        Get
            Return Me.HeaderRow("FechaPedido", OriginalValue)
        End Get
        Set(ByVal value As Date)
            Me.HeaderRow("FechaPedido") = value
        End Set
    End Property

    Public Overrides Function PrimaryKeyCab() As String()
        Return New String() {"IDPedido"}
    End Function

    Public Overrides Function PrimaryKeyLin() As String()
        Return New String() {"IDLineaPedido"}
    End Function

    Public Function KeyLinkTrazabilidad() As String()
        Return New String() {"IDPCPrincipal"}
    End Function


    Public Property Cabecera() As PedCabCompra
        Get
            Return MyBase.Cabecera
        End Get
        Set(ByVal value As PedCabCompra)
            MyBase.Cabecera = value
        End Set
    End Property

    Public ReadOnly Property dtTrazabilidad() As DataTable
        Get
            Return MyBase.Item(GetType(GRPPedidoVentaCompraLinea).Name)
        End Get
    End Property

    Protected Overrides Sub LoadEntitiesChilds(ByVal AddNew As Boolean, Optional ByVal UpdtCtx As Engine.BE.UpdatePackage = Nothing)
        MyBase.LoadEntitiesChilds(AddNew, UpdtCtx)

        '/////////////////////////////////////////// Trazabilidad  ///////////////////////////////////////////
        Dim dtPVCL As DataTable
        Dim GRP_PCL As New GRPPedidoVentaCompraLinea
        If AddNew Then  '//New de Procesos
            dtPVCL = GRP_PCL.AddNew
        ElseIf UpdtCtx Is Nothing Then  '//New de PrimaryKey
            Dim PKCabecera() As String = PrimaryKeyCab()
            dtPVCL = GRP_PCL.Filter(New FilterItem(KeyLinkTrazabilidad(0), HeaderRow(PKCabecera(0))))
        Else  '//New del formulario
            '// al crearse la bola desde el UpdatePackage, se debieran eliminar los conjuntos de datos que se van a tratar
            '// y dejar que el motor trate el resto  (MergeData)
            dtPVCL = MergeData(UpdtCtx, GetType(GRPPedidoVentaCompraLinea).Name, PrimaryKeyCab, KeyLinkTrazabilidad, True)
        End If
        Me.Add(GRP_PCL.GetType.Name, dtPVCL)
    End Sub

    ''// Esta variable, se utilizará sólo para PC desde Planificación, ya que no se pueden recuperar los datos desde negocio
    ''// se volcará aquí el datatable que viene desde presentación.
    'Friend DatosOrigen As DataTable

#Region " Métodos de integración con Albaranes "

    Public Sub SetQServida(ByVal IDLineaPedido As Integer, ByVal DiferenciaQServida As Double, ByVal services As ServiceProvider)
        If Me.dtLineas Is Nothing Then Exit Sub
        Dim DEL As New ProcesoPedidoCompra.DataEstadoLinea
        DEL.IDProveedor = Me.IDProveedor
        Dim fLineaPedido As New Filter
        fLineaPedido.Add(New NumberFilterItem("IDLineaPedido", IDLineaPedido))
        Dim WhereLineaPedido As String = fLineaPedido.Compose(New AdoFilterComposer)
        For Each lineaPedido As DataRow In Me.dtLineas.Select(WhereLineaPedido)
            lineaPedido("QServida") = Nz(lineaPedido("QServida"), 0) + DiferenciaQServida
            DEL.Linea = lineaPedido
            ProcessServer.ExecuteTask(Of ProcesoPedidoCompra.DataEstadoLinea)(AddressOf ProcesoPedidoCompra.AsignarEstadoLinea, DEL, services)
        Next
    End Sub

    'Friend Sub SetQAlbaran(ByVal IDLineaPedido As Integer, ByVal QAlbaran As Double, ByVal services As ServiceProvider)
    '    If Me.dtLineas Is Nothing Then Exit Sub
    '    Dim fLineaPedido As New Filter
    '    fLineaPedido.Add(New NumberFilterItem("IDLineaPedido", IDLineaPedido))
    '    For Each lineaPedido As DataRow In Me.dtLineas.Select(fLineaPedido.Compose(New AdoFilterComposer))
    '        lineaPedido("QAlbaran") = QAlbaran
    '    Next
    'End Sub

    'Friend Sub SetConfirmado(ByVal IDLineaPedido As Integer, ByVal Confirmado As Boolean, ByVal services As ServiceProvider)
    '    If Me.dtLineas Is Nothing Then Exit Sub
    '    Dim fLineaPedido As New Filter
    '    fLineaPedido.Add(New NumberFilterItem("IDLineaPedido", IDLineaPedido))
    '    For Each lineaPedido As DataRow In Me.dtLineas.Select(fLineaPedido.Compose(New AdoFilterComposer))
    '        lineaPedido("Confirmado") = Confirmado
    '    Next
    'End Sub


#End Region

End Class

