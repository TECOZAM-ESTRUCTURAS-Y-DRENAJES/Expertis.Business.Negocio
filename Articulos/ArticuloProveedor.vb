Public Class ArticuloProveedorInfo
    Inherits ClassEntityInfo

    Public IDProveedor As String
    Public IDArticulo As String
    Public RefProveedor As String
    Public DescRefProveedor As String
    Public IDUDCompra As String
    Public UdValoracion As Integer
    'Public Revision As String
    'Public IDLineaOfertaDetalle As Integer
    Public Precio As Double
    Public Dto1 As Double
    Public Dto2 As Double
    Public Dto3 As Double
    'Public IDArticuloContenedor As String
    'Public QContenedor As Double
    'Public IDArticuloEmbalaje As String
    'Public QEmbalaje As Double

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal data As DataRow)
        MyBase.New(data)
    End Sub

    Public Overloads Overrides Sub Fill(ByVal ParamArray PrimaryKey() As Object)
        Dim dtArtProv As DataTable = New ArticuloProveedor().SelOnPrimaryKey(PrimaryKey)
        If dtArtProv.Rows.Count > 0 Then Me.Fill(dtArtProv.Rows(0))
    End Sub

End Class

Public Class ArticuloProveedor

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbArticuloProveedor"

#End Region

#Region "Eventos RegisterDeleteTasks"

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf ActualizarPrincipal)
    End Sub

    <Task()> Public Shared Sub ActualizarPrincipal(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data("Principal") Then
            Dim dt As DataTable = New ArticuloProveedor().Filter(New FilterItem("IDArticulo", FilterOperator.Equal, data("IDArticulo")))
            If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
                dt.Rows(0)("Principal") = True
                BusinessHelper.UpdateTable(dt)
            End If
        End If
    End Sub

#End Region

#Region "Eventos GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBrl As New BusinessRules
        oBrl.Add("UDValoracion", AddressOf ComprobarValor)
        oBrl.Add("Dto1", AddressOf ComprobarValor)
        oBrl.Add("Dto2", AddressOf ComprobarValor)
        oBrl.Add("Dto3", AddressOf ComprobarValor)
        oBrl.Add("IDArticulo", AddressOf BRCambioArticulo)
        Return oBrl
    End Function

    <Task()> Public Shared Sub ComprobarValor(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Not IsNumeric(data.Value) Then ApplicationService.GenerateError("Campo no numérico.")
    End Sub

    <Task()> Public Shared Sub BRCambioArticulo(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            data.Current(data.ColumnName) = data.Value
            ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf CambioArticulo, data.Current, services)
        End If
    End Sub

    <Task()> Public Shared Sub CambioArticuloDR(ByVal data As DataRow, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf CambioArticulo, New DataRowPropertyAccessor(data), services)
    End Sub

    <Task()> Public Shared Sub CambioArticulo(ByVal data As IPropertyAccessor, ByVal services As ServiceProvider)
        If Not IsDBNull(data("IDArticulo")) Then
            Dim dt As DataTable = ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf Articulo.ValidaExisteArticulo, data("IDArticulo"), services)
            data("DescArticulo") = dt.Rows(0)("DescArticulo")
            data("IDUDCompra") = dt.Rows(0)("IDUDCompra")
            data("UDValoracion") = dt.Rows(0)("UDValoracion")
        End If
    End Sub

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDProveedor")) = 0 Then ApplicationService.GenerateError("El Proveedor es un dato obligatorio.")
        If Length(data("IDArticulo")) = 0 Then ApplicationService.GenerateError("El Artículo es obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dtAP As DataTable = New ArticuloProveedor().SelOnPrimaryKey(data("IDProveedor"), data("IDArticulo"))
            If Not dtAP Is Nothing AndAlso dtAP.Rows.Count > 0 Then
                ApplicationService.GenerateError("El Artículo | ya tiene asociado el Proveedor |.", Quoted(data("IDArticulo")), Quoted(data("IDProveedor")))
            End If
        End If
    End Sub

#End Region

#Region "Eventos RegisterUpdateTask"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AplicarDecimales)
        'updateProcess.AddTask(Of DataRow)(AddressOf Business.General.Comunes.UpdateEntityRow)
        'updateProcess.AddTask(Of DataRow)(AddressOf Business.General.Comunes.MarcarComoActualizado)
        'updateProcess.AddTask(Of DataRow)(AddressOf TratarPrincipal)
    End Sub

    '<Task()> Public Shared Sub TratarPrincipal(ByVal data As DataRow, ByVal services As ServiceProvider)
    '    Dim objFilter As New Filter
    '    objFilter.Add(New StringFilterItem("IDArticulo", data("IDArticulo")))
    '    objFilter.Add(New BooleanFilterItem("Principal", True))
    '    objFilter.Add(New StringFilterItem("IDProveedor", FilterOperator.NotEqual, data("IDProveedor")))

    '    Dim dtPrincipal As DataTable = New ArticuloProveedor().Filter(objFilter)
    '    If IsNothing(dtPrincipal) OrElse dtPrincipal.Rows.Count = 0 Then
    '        'data("Principal") = True
    '    Else
    '        If Nz(data("Principal"), False) Then
    '            If data("IDProveedor") <> dtPrincipal.Rows(0)("IDProveedor") Then
    '                dtPrincipal.Rows(0)("Principal") = False
    '                BusinessHelper.UpdateTable(dtPrincipal)
    '            End If
    '        End If
    '    End If

    'End Sub

    <Task()> Public Shared Sub AplicarDecimales(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDProveedor")) > 0 Then
            Dim Proveedors As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
            Dim ProInfo As ProveedorInfo = Proveedors.GetEntity(data("IDProveedor"))
            Dim IDMoneda As String
            If Length(ProInfo.IDMoneda) > 0 Then
                IDMoneda = ProInfo.IDMoneda
            Else
                Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
                IDMoneda = Monedas.MonedaA.ID
            End If
            Dim datosDec As New DataAplicarDecimalesMoneda(IDMoneda, data)
            ProcessServer.ExecuteTask(Of DataAplicarDecimalesMoneda)(AddressOf NegocioGeneral.AplicarDecimalesMoneda, datosDec, services)
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Serializable()> _
    Public Class DatosArtRef
        Public IDProveedor As String
        Public Referencia As String
    End Class

    <Task()> Public Shared Function ObtenerArticuloRef(ByVal data As DatosArtRef, ByVal services As ServiceProvider) As String
        Dim strIDArticulo As String
        If Len(data.IDProveedor) > 0 And Len(data.Referencia) > 0 Then
            Dim objFilter As New Filter
            objFilter.Add(New StringFilterItem("IDProveedor", data.IDProveedor))
            objFilter.Add(New StringFilterItem("RefProveedor", data.Referencia))

            Dim dtReferencia As DataTable = New ArticuloProveedor().Filter(objFilter)

            If Not IsNothing(dtReferencia) AndAlso dtReferencia.Rows.Count > 0 Then
                strIDArticulo = dtReferencia.Rows(0)("IDArticulo") & String.Empty
            End If
        End If

        Return strIDArticulo
    End Function

    <Task()> Public Shared Function ProveedorPredeterminadoArticuloDt(ByVal data As DataTable, ByVal services As ServiceProvider) As String
        Dim strArticulosSinProv As String
        If Not IsNothing(data) AndAlso data.Rows.Count > 0 Then
            Dim strIDProveedor As String
            For Each dr As DataRow In data.Select
                strIDProveedor = ProcessServer.ExecuteTask(Of String, String)(AddressOf ProveedorPredeterminadoArticulo, dr("IDArticulo"), services)
                If Length(strIDProveedor) > 0 Then
                    dr("IDProveedor") = strIDProveedor
                Else
                    If Length(strArticulosSinProv) > 0 Then strArticulosSinProv = strArticulosSinProv & ","
                    strArticulosSinProv = strArticulosSinProv & dr("IDArticulo")
                End If
            Next
            BusinessHelper.UpdateTable(data)
        End If
        Return strArticulosSinProv
    End Function

    <Task()> Public Shared Function ProveedorPredeterminadoArticulo(ByVal data As String, ByVal services As ServiceProvider) As String
        Dim strIDProveedor As String
        If Len(data) > 0 Then
            Dim objFilter As New Filter
            objFilter.Add(New StringFilterItem("IDArticulo", data))
            objFilter.Add(New BooleanFilterItem("Principal", True))
            Dim dt As DataTable = New ArticuloProveedor().Filter(objFilter)
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                strIDProveedor = dt.Rows(0)("IDProveedor") & String.Empty
            End If
        End If
        Return strIDProveedor
    End Function

    <Serializable()> _
    Public Class DatosProvPredetArtCant
        Public IDArticulo As String
        Public QCantidad As Double
    End Class

    <Task()> Public Shared Function ProveedorPredeterminadoArticuloCant(ByVal data As DatosProvPredetArtCant, ByVal services As ServiceProvider) As DataTable
        Dim DtProvPred As New DataTable
        With DtProvPred
            .Columns.Add("Precio", GetType(Double))
            .Columns.Add("Dto1", GetType(Double))
            .Columns.Add("Dto2", GetType(Double))
            .Columns.Add("Dto3", GetType(Double))
            .Columns.Add("UDValoracion", GetType(Integer))
        End With
        Dim drNew As DataRow = DtProvPred.NewRow
        Dim f As New Filter
        f.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
        f.Add(New BooleanFilterItem("Principal", True))

        Dim dtProv As DataTable = New ArticuloProveedor().Filter(f)
        If Not dtProv Is Nothing AndAlso dtProv.Rows.Count > 0 Then
            Dim drProv As DataRow = dtProv.Rows(0)
            If Length(drProv("UDValoracion")) Then drNew("UDValoracion") = drProv("UDValoracion")

            f.Clear()
            f.Add("IDProveedor", FilterOperator.Equal, drProv("IDProveedor"), FilterType.String)
            f.Add("IDArticulo", FilterOperator.Equal, data.IDArticulo, FilterType.String)
            f.Add("QDesde", FilterOperator.LessThanOrEqual, data.QCantidad, FilterType.Numeric)
            Dim dtPrecio As DataTable = New ArticuloProveedorLinea().Filter(f, "QDesde DESC")
            If Not dtPrecio Is Nothing AndAlso dtPrecio.Rows.Count > 0 Then
                drNew("Precio") = dtPrecio.Rows(0)("Precio")
                drNew("Dto1") = dtPrecio.Rows(0)("Dto1")
                drNew("Dto2") = dtPrecio.Rows(0)("Dto2")
                drNew("Dto3") = dtPrecio.Rows(0)("Dto3")
            Else
                drNew("Precio") = 0 : drNew("Dto1") = 0 : drNew("Dto2") = 0 : drNew("Dto3") = 0
            End If
            DtProvPred.Rows.Add(drNew.ItemArray)
            Return DtProvPred
        End If
    End Function

#End Region

End Class