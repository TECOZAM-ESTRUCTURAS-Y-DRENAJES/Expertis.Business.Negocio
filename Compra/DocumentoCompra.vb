Public MustInherit Class DocumentoCompra
    Inherits DocumentCabLin

    Public Proveedor As ProveedorInfo

#Region " Creación de instancias "

    '//New a utilizar en los procesos de creación de elementos de tipo cabecera/lineas
    Public Sub New(ByVal Cabecera As CompraCab, ByVal services As ServiceProvider)
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
        MyBase.LoadEntityHeader(AddNew, CType(Cab, CompraCab), UpdtCtx, PrimaryKey)

        If AddNew Then '//New de Procesos
            Me.HeaderRow("IDProveedor") = Me.Cabecera.IDProveedor
        End If
    End Sub
    Protected Overrides Sub LoadEntitiesChilds(ByVal AddNew As Boolean, Optional ByVal UpdtCtx As UpdatePackage = Nothing)
        MyBase.LoadEntitiesChilds(AddNew, UpdtCtx)
    End Sub
    Protected Overrides Sub LoadEntitiesGrandChilds(ByVal AddNew As Boolean)
        MyBase.LoadEntitiesGrandChilds(AddNew)
    End Sub

    Protected Overrides Sub Inicializar(ByVal services As ServiceProvider)
        MyBase.Inicializar(services)

        Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
        If Proveedor Is Nothing Then
            Me.Proveedor = Proveedores.GetEntity(IDProveedor)
        End If
        If Me.HeaderRow.Table.Columns.Contains("Dto") AndAlso (Me.HeaderRow.RowState = DataRowState.Added OrElse (Me.HeaderRow.RowState = DataRowState.Modified AndAlso Me.HeaderRow("IDProveedor") <> Me.HeaderRow("IDProveedor", DataRowVersion.Original))) Then
            Me.HeaderRow("Dto") = Me.Proveedor.DtoComercial
        End If
        Me.IDProveedor = HeaderRow("IDProveedor")
    End Sub

#End Region


    Public Property IDProveedor() As String
        Get
            Return Me.HeaderRow("IDProveedor") & String.Empty
        End Get
        Set(ByVal value As String)
            Me.HeaderRow("IDProveedor") = value
        End Set
    End Property

    Public Property Cabecera() As CompraCab
        Get
            Return MyBase.Cabecera
        End Get
        Set(ByVal value As CompraCab)
            MyBase.Cabecera = value
        End Set
    End Property


End Class
