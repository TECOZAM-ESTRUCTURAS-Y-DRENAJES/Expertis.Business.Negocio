Public Class DocumentoAlbaranVenta
    Inherits DocumentoComercial

    Public Sub New(ByVal Cabecera As AlbCabVenta, ByVal services As ServiceProvider)
        MyBase.New(Cabecera, services)
    End Sub

    Public Sub New(ByVal UpdtCtx As Engine.BE.UpdatePackage)
        MyBase.New(UpdtCtx)
    End Sub

    Public Sub New(ByVal ParamArray PrimaryKey() As Object)
        MyBase.New(PrimaryKey)
    End Sub

    Protected Overrides Sub LoadEntityHeader(ByVal AddNew As Boolean, Optional ByVal Cab As Cabecera = Nothing, Optional ByVal UpdtCtx As Engine.BE.UpdatePackage = Nothing, Optional ByVal PrimaryKey() As Object = Nothing)
        MyBase.LoadEntityHeader(AddNew, CType(Cab, AlbCabVenta), UpdtCtx, PrimaryKey)

        If Not Me.Cabecera Is Nothing Then
            Me.HeaderRow("DtoAlbaran") = Me.Cabecera.Dto
            If Me.Cabecera.IdDireccion <> 0 Then Me.HeaderRow("IdDireccion") = Me.Cabecera.IdDireccion
            Me.HeaderRow("IDFormaEnvio") = Me.Cabecera.IDFormaEnvio
            Me.HeaderRow("IDCondicionEnvio") = Me.Cabecera.IDCondicionEnvio
            If TypeOf Me.Cabecera Is AlbCabVentaPedido Then
                Me.HeaderRow("IDModoTransporte") = CType(Me.Cabecera, AlbCabVentaPedido).IDModoTransporte
                Me.HeaderRow("PedidoCliente") = CType(Me.Cabecera, AlbCabVentaPedido).PedidoCliente
                Me.HeaderRow("ResponsableExpedicion") = CType(Me.Cabecera, AlbCabVentaPedido).Responsable
            End If
            If Me.Cabecera.IDDireccionFra <> 0 Then Me.HeaderRow("IDDireccionFra") = Me.Cabecera.IDDireccionFra
            Me.HeaderRow("IDAlmacen") = Me.Cabecera.IDAlmacen
            Me.HeaderRow("Automatico") = True
            If Length(Me.HeaderRow("ResponsableExpedicion")) = 0 Then
                Dim strIDOper As String = ProcessServer.ExecuteTask(Of Object, String)(AddressOf Operario.ObtenerIDOperarioUsuario, Nothing, New ServiceProvider)
                If Len(strIDOper) > 0 Then
                    Me.HeaderRow("ResponsableExpedicion") = strIDOper
                End If
            End If
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

        LoadEntityGrandChild(GetType(AlbaranVentaLote).Name, f, AddNew)

        Dim idsLotes(dtLote.Rows.Count - 1) As Object
        For i As Integer = 0 To dtLote.Rows.Count - 1
            idsLotes(i) = dtLote.Rows(i)("IDLineaLote")
        Next

        Dim FSeg As New Filter
        If idsLotes.Length > 0 Then
            FSeg.Add(New InListFilterItem("IDLineaLote", idsLotes, FilterType.Numeric))
        Else : FSeg.Add(New NoRowsFilterItem())
        End If
        LoadEntityGrandChild(GetType(AlbaranVentaSeguimiento).Name, FSeg, AddNew)
    End Sub

    Public Property Cabecera() As AlbCabVenta
        Get
            Return MyBase.Cabecera
        End Get
        Set(ByVal value As AlbCabVenta)
            MyBase.Cabecera = value
        End Set
    End Property

    Public Overrides Function EntidadAnalitica() As String
        Return GetType(AlbaranVentaAnalitica).Name
    End Function

    Public Overrides Function EntidadCabecera() As String
        Return GetType(AlbaranVentaCabecera).Name
    End Function

    Public Overrides Function PrimaryKeyCab() As String()
        Return New String() {"IDAlbaran"}
    End Function

    Public Overrides Function PrimaryKeyLin() As String()
        Return New String() {"IDLineaAlbaran"}
    End Function

    Public Overrides Function EntidadLineas() As String
        Return GetType(AlbaranVentaLinea).Name
    End Function

    Public Overrides Function EntidadRepresentantes() As String
        Return GetType(AlbaranVentaRepresentante).Name
    End Function

    Public Function EntidadSeguimiento() As String
        Return GetType(AlbaranVentaSeguimiento).Name
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
            Return MyBase.Item(GetType(AlbaranVentaLote).Name)
        End Get
    End Property

    Public ReadOnly Property dtSeguimiento() As DataTable
        Get
            Return MyBase.Item(GetType(AlbaranVentaSeguimiento).Name)
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

            ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoAlbaranVenta.AsignarEstadoLinea, lineaAlbaran, services)
        Next
        ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.ActualizarEstadoAlbaran, Me, services)
    End Sub

    Public Sub SetTipoDocumentoFactura()
        If Me.HeaderRow("Ticket") Then Me.HeaderRow("TipoDocumento") = enumTipoDocumento.tdFactura
    End Sub

#End Region

End Class