Public Class DocumentoFacturaCompra
    Inherits DocumentoCompra

    ' Public AntesImpuestos As Boolean

    '//New a utilizar en los procesos de creación de elementos de tipo cabecera/lineas
    Public Sub New(ByVal Cabecera As FraCabCompra, ByVal services As ServiceProvider)
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
        MyBase.LoadEntityHeader(AddNew, CType(Cab, FraCabCompra), UpdtCtx, PrimaryKey)

        If Not Cab Is Nothing Then
            Me.HeaderRow("DtoFactura") = CType(Cab, FraCabCompra).Dto
            '    If CType(Cabecera, FraCabCompra).IDProveedorBanco <> 0 Then Me.HeaderRow("IDClienteBanco") = CType(Cabecera, FraCabCompra).IDProveedorBanco
            If CType(Cab, FraCabCompra).IDDireccion <> 0 Then Me.HeaderRow("IDDireccion") = CType(Cab, FraCabCompra).IDDireccion
            If CType(Cab, FraCabCompra).IDObra <> 0 Then Me.HeaderRow("IDObra") = CType(Cab, FraCabCompra).IDObra
        End If
    End Sub

    Protected Overrides Sub LoadEntitiesChilds(ByVal AddNew As Boolean, Optional ByVal UpdtCtx As Engine.BE.UpdatePackage = Nothing)
        MyBase.LoadEntitiesChilds(AddNew, UpdtCtx)

        Dim oFBI As New FacturaCompraBaseImponible : Dim oPAG As New Pago
        Dim dtFCBI As DataTable : Dim dtPago As DataTable
        Dim oFCI As New FacturaCompraImpuesto : Dim dtFCI As DataTable

        If AddNew Then  '//New de Procesos
            dtFCBI = oFBI.AddNew
            dtPago = oPAG.AddNew
            dtFCI = oFCI.AddNew
        ElseIf UpdtCtx Is Nothing Then  '//New de PrimaryKey
            Dim PKCabecera() As String = PrimaryKeyCab()
            dtFCBI = New FacturaCompraBaseImponible().Filter(New FilterItem(PKCabecera(0), HeaderRow(PKCabecera(0))))
            dtPago = New Pago().Filter(New FilterItem(PKCabecera(0), HeaderRow(PKCabecera(0))))
            dtFCI = oFCI.Filter(New FilterItem(PKCabecera(0), HeaderRow(PKCabecera(0))))
        Else  '//New del formulario
            '// al crearse la bola desde el UpdatePackage, se debieran eliminar los conjuntos de datos que se van a tratar
            '// y dejar que el motor trate el resto  (MergeData)
            dtFCBI = MergeData(UpdtCtx, GetType(FacturaCompraBaseImponible).Name, New String() {"IDFactura"}, New String() {"IDFactura"}, True)
            dtPago = MergeData(UpdtCtx, GetType(Pago).Name, New String() {"IDFactura"}, New String() {"IDFactura"}, True)
            dtFCI = MergeData(UpdtCtx, oFCI.GetType.Name, New String() {"IDFactura"}, New String() {"IDFactura"}, True)
        End If
        Me.Add(oFBI.GetType.Name, dtFCBI)
        Me.Add(oPAG.GetType.Name, dtPago)
        Me.Add(oFCI.GetType.Name, dtFCI)

    End Sub



    Public Property Cabecera() As FraCabCompra
        Get
            Return MyBase.Cabecera
        End Get
        Set(ByVal value As FraCabCompra)
            MyBase.Cabecera = value
        End Set
    End Property

    Public Overrides Function EntidadAnalitica() As String
        Return GetType(FacturaCompraAnalitica).Name
    End Function

    Public Overrides Function EntidadCabecera() As String
        Return GetType(FacturaCompraCabecera).Name
    End Function

    Public Overrides Function PrimaryKeyCab() As String()
        Return New String() {"IDFactura"}
    End Function

    Public Overrides Function PrimaryKeyLin() As String()
        Return New String() {"IDLineaFactura"}
    End Function

    Public Overrides Function EntidadLineas() As String
        Return GetType(FacturaCompraLinea).Name
    End Function

   
    Public Overrides Property Fecha(Optional ByVal OriginalValue As DataRowVersion = DataRowVersion.Current) As Date
        Get
            Return Me.HeaderRow("FechaFactura", OriginalValue)
        End Get
        Set(ByVal value As Date)
            HeaderRow("FechaFactura") = value
        End Set
    End Property

    Public ReadOnly Property dtFCBI() As DataTable
        Get
            Return MyBase.Item(GetType(FacturaCompraBaseImponible).Name)
        End Get
    End Property

    Public ReadOnly Property dtPagos() As DataTable
        Get
            Return MyBase.Item(GetType(Pago).Name)
        End Get
    End Property

    Public ReadOnly Property dtImpuestos() As DataTable
        Get
            Return MyBase.Item("FacturaCompraImpuesto")
        End Get
    End Property

End Class

 
