Public MustInherit Class DocumentoComercial
    Inherits DocumentCabLin

    Public Cliente As ClienteInfo
    Public MustOverride Function EntidadRepresentantes() As String

#Region " Creación de instancias "

    '//New a utilizar en los procesos de creación de elementos de tipo cabecera/lineas
    Public Sub New(ByVal Cabecera As ComercialCab, ByVal services As ServiceProvider)
        MyBase.New(Cabecera, services)
    End Sub

    '//New a utilizar desde presentación (utilizado por el motor para realizar las actualizaciones de los elementos cabecera/lineas)
    Public Sub New(ByVal UpdtCtx As UpdatePackage)
        '// al crearse la bola desde el UpdatePackage, se debieran eliminar los conjuntos de datos que se van a tratar
        '// y dejar que el motor trate el resto

        MyBase.New(UpdtCtx)
    End Sub

    '//New utilizado para obtener un Documento alamacenado en la BBDD.
    Public Sub New(ByVal ParamArray PrimaryKey() As Object)
        MyBase.New(PrimaryKey)
    End Sub

    Protected Overrides Sub LoadEntityHeader(ByVal AddNew As Boolean, Optional ByVal Cab As Cabecera = Nothing, Optional ByVal UpdtCtx As UpdatePackage = Nothing, Optional ByVal PrimaryKey() As Object = Nothing)
        MyBase.LoadEntityHeader(AddNew, CType(Cab, ComercialCab), UpdtCtx, PrimaryKey)
        If AddNew Then '//New de Procesos
            Me.HeaderRow("IDCliente") = Me.Cabecera.IDCliente
            Me.HeaderRow("Edi") = Me.Cabecera.Edi
        End If
    End Sub
    Protected Overrides Sub LoadEntitiesChilds(ByVal AddNew As Boolean, Optional ByVal UpdtCtx As UpdatePackage = Nothing)
        MyBase.LoadEntitiesChilds(AddNew, UpdtCtx)
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

        LoadEntityGrandChild(EntidadRepresentantes, f, AddNew)

    End Sub


    Protected Overrides Sub Inicializar(ByVal services As ServiceProvider)
        MyBase.Inicializar(services)

        Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
        If Cliente Is Nothing Then
            Me.Cliente = Clientes.GetEntity(IDCliente)
        End If

        Me.IDCliente = HeaderRow("IDCliente")
    End Sub

#End Region

    Public Property Cabecera() As ComercialCab
        Get
            Return MyBase.Cabecera
        End Get
        Set(ByVal value As ComercialCab)
            MyBase.Cabecera = value
        End Set
    End Property

    Public Property IDCliente() As String
        Get
            Return Me.HeaderRow("IDCliente") & String.Empty
        End Get
        Set(ByVal value As String)
            Me.HeaderRow("IDCliente") = value
        End Set
    End Property

  
    Public ReadOnly Property dtVentaRepresentante() As DataTable
        Get
            Return MyBase.Item(EntidadRepresentantes)
        End Get
    End Property

End Class
