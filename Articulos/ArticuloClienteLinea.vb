Public Class ArticuloClienteLinea
#Region "Constructor"
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbArticuloClienteLinea"
#End Region
#Region "Eventos GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBrl As New BusinessRules
        oBrl.Add("QDesde", AddressOf CambioDatos)
        oBrl.Add("Dto1", AddressOf CambioDatos)
        oBrl.Add("Dto2", AddressOf CambioDatos)
        oBrl.Add("Dto3", AddressOf CambioDatos)
        oBrl.Add("Precio", AddressOf CambioPrecios)
        oBrl.Add("PVP", AddressOf CambioPrecios)
        Return oBrl
    End Function

    <Task()> Public Shared Sub CambioDatos(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Not IsNumeric(data.Value) Then ApplicationService.GenerateError("Campo no numérico.")
    End Sub
    <Task()> Public Shared Sub CambioPrecios(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Not IsNumeric(data.Value) Then ApplicationService.GenerateError("Campo no numérico.")
        data.Current(data.ColumnName) = data.Value
        Dim strIDTipoIVA As String
        If Not IsNothing(data.Context) AndAlso data.Context.Contains("IdTipoIVA") Then
            strIDTipoIVA = data.Context("IdTipoIVA")
        ElseIf Not IsNothing(data.Current) AndAlso data.Current.Contains("IDTipoIVA") Then
            strIDTipoIVA = data.Current("IdTipoIVA")
        Else
            Dim DtArt As DataTable = New Articulo().SelOnPrimaryKey(data.Current("IDArticulo"))
            If Not DtArt Is Nothing AndAlso DtArt.Rows.Count > 0 Then strIDTipoIVA = Nz(DtArt.Rows(0)("IDTipoIVA"), String.Empty)
        End If
        Dim dblFactor As Double
        If Length(strIDTipoIVA) > 0 Then
            Dim oIVA As New TipoIva
            Dim dtIVA As DataTable = oIVA.SelOnPrimaryKey(strIDTipoIVA)
            If Not IsNothing(dtIVA) AndAlso dtIVA.Rows.Count > 0 Then
                dblFactor = dtIVA.Rows(0)("Factor") / 100
            End If
        End If

        Dim clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
        Dim CliInfo As ClienteInfo = clientes.GetEntity(data.Context("IDCliente"))
        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
        Dim m As MonedaInfo = Monedas.GetMoneda(CliInfo.Moneda, Today())

        If data.ColumnName = "Precio" Then
            ' Hallamos el precio de venta al público...
            data.Current(data.ColumnName) = xRound(data.Value, m.NDecimalesPrecio)
            data.Current("PVP") = xRound(data.Current("Precio") * (1 + dblFactor), m.NDecimalesImporte)
        Else
            ' Si modifica el precio de venta al público hallamos el precio...
            data.Current(data.ColumnName) = xRound(data.Value, m.NDecimalesImporte)
            data.Current("Precio") = xRound(data.Current("PVP") / (1 + dblFactor), m.NDecimalesPrecio)
        End If

    End Sub
#End Region
#Region "RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDCliente")) = 0 Then ApplicationService.GenerateError("El Cliente es un dato obligatorio.")
        If Length(data("IDArticulo")) = 0 Then ApplicationService.GenerateError("El Artículo es obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dtACL As DataTable = New ArticuloClienteLinea().SelOnPrimaryKey(data("IDCliente"), data("IDArticulo"), data("QDesde"))
            If Not dtACL Is Nothing AndAlso dtACL.Rows.Count > 0 Then
                ApplicationService.GenerateError("Esa Cantidad ya existe para ese Artículo-Cliente.")
            End If
        End If
    End Sub

#End Region
    
    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf Negocio.ArticuloCliente.AplicarDecimales)
    End Sub

    'Public Function ObtenerFactorIVA(ByVal strIDTipoIVA As String) As Double
    '    Dim dblFactor As Double
    '    If Len(strIDTipoIVA) > 0 Then
    '        Dim oIVA As New TipoIva
    '        Dim dtIVA As DataTable = oIVA.SelOnPrimaryKey(strIDTipoIVA)
    '        If Not dtIVA Is Nothing AndAlso dtIVA.Rows.Count > 0 Then
    '            dblFactor = dtIVA.Rows(0)("Factor")
    '        End If
    '    End If
    '    Return dblFactor
    'End Function



End Class