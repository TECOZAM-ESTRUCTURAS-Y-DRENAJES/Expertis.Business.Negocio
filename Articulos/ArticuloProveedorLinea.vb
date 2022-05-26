Public Class ArticuloProveedorLinea

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbArticuloProveedorLinea"

#End Region

#Region "Eventos GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBrl As New BusinessRules
        oBrl.Add("QDesde", AddressOf CambioDatos)
        oBrl.Add("Precio", AddressOf CambioDatos)
        oBrl.Add("Dto1", AddressOf CambioDatos)
        oBrl.Add("Dto2", AddressOf CambioDatos)
        oBrl.Add("Dto3", AddressOf CambioDatos)
        Return oBrl
    End Function

    <Task()> Public Shared Sub CambioDatos(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Not IsNumeric(data.Value) Then ApplicationService.GenerateError("Campo no numérico.")
    End Sub

#End Region

#Region "RegisterValidateTasks"

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
            Dim dtAPL As DataTable = New ArticuloProveedorLinea().SelOnPrimaryKey(data("IDProveedor"), data("IDArticulo"), data("QDesde"))
            If Not dtAPL Is Nothing AndAlso dtAPL.Rows.Count > 0 Then
                ApplicationService.GenerateError("Esa Cantidad ya existe para ese Artículo-Proveedor.")
            End If
        End If
    End Sub

#End Region

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf Negocio.ArticuloProveedor.AplicarDecimales)
    End Sub

End Class