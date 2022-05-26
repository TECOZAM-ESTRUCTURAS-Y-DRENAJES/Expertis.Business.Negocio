Public Class DocumentoFacturaVenta
    Inherits DocumentoComercial


    '//New a utilizar en los procesos de creación de elementos de tipo cabecera/lineas
    Public Sub New(ByVal Cabecera As FraCab, ByVal services As ServiceProvider)
        MyBase.New(Cabecera, services)
    End Sub
    '//New a utilizar desde presentación (utilizado por el motor para realizar las actualizaciones de los elementos cabecera/lineas)
    Public Sub New(ByVal UpdtCtx As Engine.BE.UpdatePackage)
        MyBase.New(UpdtCtx)
    End Sub
    '//New utilizado para obtener un Documento almacenado en la BBDD.
    Public Sub New(ByVal ParamArray PrimaryKey() As Object)
        MyBase.New(PrimaryKey)
    End Sub


    Protected Overrides Sub LoadEntityHeader(ByVal AddNew As Boolean, Optional ByVal Cab As Cabecera = Nothing, Optional ByVal UpdtCtx As Engine.BE.UpdatePackage = Nothing, Optional ByVal PrimaryKey() As Object = Nothing)
        MyBase.LoadEntityHeader(AddNew, CType(Cab, FraCab), UpdtCtx, PrimaryKey)
        'Me.Cabecera = CType(Cabecera, FraCab)

        If Not Me.Cabecera Is Nothing Then
            ' Me.HeaderRow("DtoFactura") = Me.Cabecera.Dto
            If Me.Cabecera.IDClienteBanco <> 0 Then Me.HeaderRow("IDClienteBanco") = Me.Cabecera.IDClienteBanco
            If Me.Cabecera.IDDireccion <> 0 Then Me.HeaderRow("IDDireccion") = Me.Cabecera.IDDireccion
            'If Me.Cabecera.IDDireccionFra <> 0 Then Me.HeaderRow("IDDireccionFra") = Me.Cabecera.IDDireccionFra
            If Me.Cabecera.IDObra <> 0 Then Me.HeaderRow("IDObra") = Me.Cabecera.IDObra
            If Cabecera.GetType Is GetType(FraCabAlbaran) Then
                If Length(CType(Cabecera, FraCabAlbaran).IDTPV) > 0 Then Me.HeaderRow("IDTPV") = CType(Cabecera, FraCabAlbaran).IDTPV
            End If
        End If
    End Sub

    Protected Overrides Sub LoadEntitiesChilds(ByVal AddNew As Boolean, Optional ByVal UpdtCtx As Engine.BE.UpdatePackage = Nothing)
        MyBase.LoadEntitiesChilds(AddNew, UpdtCtx)

        Dim oFBI As New FacturaVentaBaseImponible : Dim oCBR As New Cobro
        Dim dtFVBI As DataTable : Dim dtCobro As DataTable
        Dim oFVI As New FacturaVentaImpuesto : Dim dtFVI As DataTable

        If AddNew Then  '//New de Procesos
            dtFVBI = oFBI.AddNew
            dtCobro = oCBR.AddNew
            dtFVI = oFVI.AddNew
        ElseIf UpdtCtx Is Nothing Then  '//New de PrimaryKey
            Dim PKCabecera() As String = PrimaryKeyCab()
            dtFVBI = New FacturaVentaBaseImponible().Filter(New FilterItem(PKCabecera(0), HeaderRow(PKCabecera(0))))
            dtCobro = New Cobro().Filter(New FilterItem(PKCabecera(0), HeaderRow(PKCabecera(0))))
            dtFVI = oFVI.Filter(New FilterItem(PKCabecera(0), HeaderRow(PKCabecera(0))))
        Else  '//New del formulario
            '// al crearse la bola desde el UpdatePackage, se debieran eliminar los conjuntos de datos que se van a tratar
            '// y dejar que el motor trate el resto  (MergeData)
            dtFVBI = MergeData(UpdtCtx, GetType(FacturaVentaBaseImponible).Name, New String() {"IDFactura"}, New String() {"IDFactura"}, True)
            dtCobro = MergeData(UpdtCtx, GetType(Cobro).Name, New String() {"IDFactura"}, New String() {"IDFactura"}, True)
            dtFVI = MergeData(UpdtCtx, oFVI.GetType.Name, New String() {"IDFactura"}, New String() {"IDFactura"}, True)
        End If
        Me.Add(oFBI.GetType.Name, dtFVBI)
        Me.Add(oCBR.GetType.Name, dtCobro)
        Me.Add(oFVI.GetType.Name, dtFVI)

    End Sub



    Public Property Cabecera() As FraCab
        Get
            Return MyBase.Cabecera
        End Get
        Set(ByVal value As FraCab)
            MyBase.Cabecera = value
        End Set
    End Property

    

    Public Overrides Function EntidadCabecera() As String
        Return GetType(FacturaVentaCabecera).Name
    End Function

    Public Overrides Function EntidadLineas() As String
        Return GetType(FacturaVentaLinea).Name
    End Function


    Public Overrides Function EntidadRepresentantes() As String
        Return GetType(FacturaVentaRepresentante).Name
    End Function

    Public Overrides Function EntidadAnalitica() As String
        Return GetType(FacturaVentaAnalitica).Name
    End Function

    Public Overrides Function PrimaryKeyCab() As String()
        Return New String() {"IDFactura"}
    End Function

    Public Overrides Function PrimaryKeyLin() As String()
        Return New String() {"IDLineaFactura"}
    End Function


    Public Overrides Property Fecha(Optional ByVal OriginalValue As DataRowVersion = DataRowVersion.Current) As Date
        Get
            Return Me.HeaderRow("FechaFactura", OriginalValue)
        End Get
        Set(ByVal value As Date)
            HeaderRow("FechaFactura") = value
        End Set
    End Property

    Public ReadOnly Property dtFVBI() As DataTable
        Get
            Return MyBase.Item(GetType(FacturaVentaBaseImponible).Name)
        End Get
    End Property

    Public ReadOnly Property dtCobros() As DataTable
        Get
            Return MyBase.Item(GetType(Cobro).Name)
        End Get
    End Property

    Public ReadOnly Property dtImpuestos() As DataTable
        Get
            Return Me.Item("FacturaVentaImpuesto")
        End Get
    End Property
End Class