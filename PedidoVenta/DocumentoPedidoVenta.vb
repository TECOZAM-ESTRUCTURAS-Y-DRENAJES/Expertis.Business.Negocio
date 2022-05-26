Public Class DocumentoPedidoVenta
    Inherits DocumentoComercial

    Public Sub New(ByVal Cabecera As PedCab, ByVal services As ServiceProvider)
        MyBase.New(Cabecera, services)
    End Sub

    Public Sub New(ByVal UpdtCtx As Engine.BE.UpdatePackage)
        MyBase.New(UpdtCtx)
    End Sub

    Public Sub New(ByVal ParamArray PrimaryKey() As Object)
        MyBase.New(PrimaryKey)
    End Sub

    Public Overrides Function EntidadAnalitica() As String
        Return GetType(PedidoVentaAnalitica).Name
    End Function

    Public Overrides Function EntidadCabecera() As String
        Return GetType(PedidoVentaCabecera).Name
    End Function

    Public Overrides Function EntidadLineas() As String
        Return GetType(PedidoVentaLinea).Name
    End Function

    Public Overrides Function EntidadRepresentantes() As String
        Return GetType(PedidoVentaRepresentante).Name
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

    Public Property Cabecera() As PedCab
        Get
            Return MyBase.Cabecera
        End Get
        Set(ByVal value As PedCab)
            MyBase.Cabecera = value
        End Set
    End Property

    Protected Overrides Sub LoadEntityHeader(ByVal AddNew As Boolean, Optional ByVal Cab As Cabecera = Nothing, Optional ByVal UpdtCtx As Engine.BE.UpdatePackage = Nothing, Optional ByVal PrimaryKey() As Object = Nothing)
        MyBase.LoadEntityHeader(AddNew, CType(Cab, PedCab), UpdtCtx, PrimaryKey)
    End Sub

#Region " Métodos de integración con Albaranes "

    Public Overridable Sub SetQServida(ByVal IDLineaPedido As Integer, ByVal DiferenciaQServida As Double, ByVal services As ServiceProvider)
        If Me.dtLineas Is Nothing Then Exit Sub
        Dim fLineaPedido As New Filter
        fLineaPedido.Add(New NumberFilterItem("IDLineaPedido", IDLineaPedido))
        Dim WhereLineaPedido As String = fLineaPedido.Compose(New AdoFilterComposer)
        For Each lineaPedido As DataRow In Me.dtLineas.Select(WhereLineaPedido)
            lineaPedido("QServida") = Nz(lineaPedido("QServida"), 0) + DiferenciaQServida

            ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoPedidoVenta.AsignarEstadoLinea, lineaPedido, services)
        Next
    End Sub

    Public Overridable Sub SetQAlbaran(ByVal IDLineaPedido As Integer, ByVal QAlbaran As Double, ByVal services As ServiceProvider)
        If Me.dtLineas Is Nothing Then Exit Sub
        Dim fLineaPedido As New Filter
        fLineaPedido.Add(New NumberFilterItem("IDLineaPedido", IDLineaPedido))
        Dim WhereLineaPedido As String = fLineaPedido.Compose(New AdoFilterComposer)
        For Each lineaPedido As DataRow In Me.dtLineas.Select(WhereLineaPedido)
            lineaPedido("QAlbaran") = QAlbaran
        Next
    End Sub

    Public Overridable Sub SetConfirmado(ByVal IDLineaPedido As Integer, ByVal Confirmado As Boolean, ByVal services As ServiceProvider)
        If Me.dtLineas Is Nothing Then Exit Sub
        Dim fLineaPedido As New Filter
        fLineaPedido.Add(New NumberFilterItem("IDLineaPedido", IDLineaPedido))
        Dim WhereLineaPedido As String = fLineaPedido.Compose(New AdoFilterComposer)
        For Each lineaPedido As DataRow In Me.dtLineas.Select(WhereLineaPedido)
            lineaPedido("Confirmado") = Confirmado
        Next
    End Sub

    Public Overridable Sub SetDeposito(ByVal IDLineaPedido As Integer, ByVal EsDeposito As Boolean, ByVal services As ServiceProvider)
        If Me.dtLineas Is Nothing Then Exit Sub
        Dim fLineaPedido As New Filter
        fLineaPedido.Add(New NumberFilterItem("IDLineaPedido", IDLineaPedido))
        Dim WhereLineaPedido As String = fLineaPedido.Compose(New AdoFilterComposer)
        For Each lineaPedido As DataRow In Me.dtLineas.Select(WhereLineaPedido)
            lineaPedido("Deposito") = EsDeposito
        Next
    End Sub

    Public Overridable Sub SetQFacturada(ByVal IDLineaPedido As Integer, ByVal DiferenciaQFacturada As Double, ByVal services As ServiceProvider)
        If Me.dtLineas Is Nothing Then Exit Sub
        Dim fLineaPedido As New Filter
        fLineaPedido.Add(New NumberFilterItem("IDLineaPedido", IDLineaPedido))
        Dim WhereLineaPedido As String = fLineaPedido.Compose(New AdoFilterComposer)
        For Each lineaPedido As DataRow In Me.dtLineas.Select(WhereLineaPedido)
            lineaPedido("QFacturada") = Nz(lineaPedido("QFacturada"), 0) + DiferenciaQFacturada
        Next
    End Sub

#End Region

End Class
