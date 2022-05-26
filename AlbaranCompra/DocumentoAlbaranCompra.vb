Public Class DocumentoAlbaranCompra
    Inherits DocumentoCompra

    Public Sub New(ByVal Cabecera As AlbCabCompra, ByVal services As ServiceProvider)
        MyBase.New(Cabecera, services)
    End Sub

    Public Sub New(ByVal UpdtCtx As Engine.BE.UpdatePackage)
        MyBase.New(UpdtCtx)
    End Sub

    Public Sub New(ByVal ParamArray PrimaryKey() As Object)
        MyBase.New(PrimaryKey)
    End Sub

    Protected Overrides Sub LoadEntityHeader(ByVal AddNew As Boolean, Optional ByVal Cab As Cabecera = Nothing, Optional ByVal UpdtCtx As Engine.BE.UpdatePackage = Nothing, Optional ByVal PrimaryKey() As Object = Nothing)
        MyBase.LoadEntityHeader(AddNew, CType(Cab, AlbCabCompra), UpdtCtx, PrimaryKey)

        If Not Me.Cabecera Is Nothing Then
            Me.HeaderRow("Dto") = Me.Cabecera.Dto
            If Me.Cabecera.IDDireccion <> 0 Then Me.HeaderRow("IdDireccion") = Me.Cabecera.IDDireccion
            Me.HeaderRow("IDFormaEnvio") = Me.Cabecera.IDFormaEnvio
            Me.HeaderRow("IDCondicionEnvio") = Me.Cabecera.IDCondicionEnvio
            ' If Me.Cabecera.IDDireccionFra <> 0 Then Me.HeaderRow("IDDireccionFra") = Me.Cabecera.IDDireccionFra
            Me.HeaderRow("IDAlmacen") = Me.Cabecera.IDAlmacen
            Me.HeaderRow("Automatico") = True
            Me.HeaderRow("IDtipoCompra") = Me.Cabecera.IDTipoCompra
        End If
    End Sub

    Protected Overrides Sub LoadEntitiesGrandChilds(ByVal AddNew As Boolean)
        MyBase.LoadEntitiesGrandChilds(AddNew)

        Dim PKLineas() As String = PrimaryKeyLin()
        Dim ids(dtLineas.Rows.Count - 1) As Object
        For i As Integer = 0 To dtLineas.Rows.Count - 1
            ids(i) = dtLineas.Rows(i)(PKLineas(0))
        Next
        If ids.Length = 0 Then ids = New Object() {0}

        Dim f As New Filter
        f.Add(New InListFilterItem(PKLineas(0), ids, FilterType.Numeric))


        LoadEntityGrandChild(GetType(AlbaranCompraLote).Name, f, AddNew)
        LoadEntityGrandChild(GetType(AlbaranCompraPrecio).Name, f, AddNew)
        LoadEntityGrandChild("AlbaranCompraValoracion", f, AddNew)
    End Sub

    Public Property Cabecera() As AlbCabCompra
        Get
            Return MyBase.Cabecera
        End Get
        Set(ByVal value As AlbCabCompra)
            MyBase.Cabecera = value
        End Set
    End Property

    Public Overrides Function EntidadAnalitica() As String
        Return GetType(AlbaranCompraAnalitica).Name
    End Function

    Public Overrides Function EntidadCabecera() As String
        Return GetType(AlbaranCompraCabecera).Name
    End Function

    Public Overrides Function PrimaryKeyCab() As String()
        Return New String() {"IDAlbaran"}
    End Function

    Public Overrides Function PrimaryKeyLin() As String()
        Return New String() {"IDLineaAlbaran"}
    End Function

    Public Overrides Function EntidadLineas() As String
        Return GetType(AlbaranCompraLinea).Name
    End Function

    Public Overrides Property Fecha(Optional ByVal OriginalValue As DataRowVersion = DataRowVersion.Current) As Date
        Get
            Return Me.HeaderRow("FechaAlbaran", OriginalValue)
        End Get
        Set(ByVal value As Date)
            Me.HeaderRow("FechaAlbaran") = value

            Dim Monedas As New MonedaCache
            Me.MonedaA = Monedas.MonedaA
            Me.MonedaB = Monedas.MonedaB
            Me.Moneda = Monedas.GetMoneda(IDMoneda, Fecha)
            
            Me.IDMoneda = HeaderRow("IDMoneda")
            Me.HeaderRow("CambioA") = Me.Moneda.CambioA
            Me.HeaderRow("CambioB") = Me.Moneda.CambioB

            Me.CambioA = Me.Moneda.CambioA
            Me.CambioB = Me.Moneda.CambioB
        End Set
    End Property

    Public ReadOnly Property dtLote() As DataTable
        Get
            Return MyBase.Item(GetType(AlbaranCompraLote).Name)
        End Get
    End Property

    Public ReadOnly Property dtPrecios() As DataTable
        Get
            Return MyBase.Item(GetType(AlbaranCompraPrecio).Name)
        End Get
    End Property

    Public ReadOnly Property dtValoracion() As DataTable
        Get
            Return MyBase.Item("AlbaranCompraValoracion")
        End Get
    End Property

#Region " Métodos de integración con Facturas "

    Public Sub SetQFacturada(ByVal IDLineaAlbaran As Integer, ByVal DiferenciaQFacturada As Double, ByVal services As ServiceProvider)
        If Me.dtLineas Is Nothing Then Exit Sub
        Dim fLineaAlbaran As New Filter
        fLineaAlbaran.Add(New NumberFilterItem("IDLineaAlbaran", IDLineaAlbaran))
        Dim WhereLineaAlbaran As String = fLineaAlbaran.Compose(New AdoFilterComposer)
        For Each lineaAlbaran As DataRow In Me.dtLineas.Select(WhereLineaAlbaran)
            lineaAlbaran("QFacturada") = Nz(lineaAlbaran("QFacturada"), 0) + DiferenciaQFacturada

            ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoAlbaranCompra.AsignarEstadoLinea, lineaAlbaran, services)
        Next
        ProcessServer.ExecuteTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.ActualizarEstadoAlbaran, Me, services)
    End Sub

#End Region

End Class

