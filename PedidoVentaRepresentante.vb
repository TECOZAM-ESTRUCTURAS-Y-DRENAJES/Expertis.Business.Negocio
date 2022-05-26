Public Class PedidoVentaRepresentante
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper
    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbPedidoVentaRepresentante"

#Region " Business Rules "

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim services As New ServiceProvider
        Dim oBRL As New BusinessRules
        oBRL = ProcessServer.ExecuteTask(Of BusinessRules, BusinessRules)(AddressOf ProcesoComercial.RepresentantesCommonBusinessRules, oBRL, services)
        Return oBRL
    End Function

#End Region


#Region " Validate "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarRegistroExistente)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarRepresentanteExistente)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarImporteValido)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarComisionValida)
    End Sub

    <Task()> Public Shared Sub ValidarRegistroExistente(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim DtPVR As DataTable = New PedidoVentaRepresentante().SelOnPrimaryKey(data("IDLineaPedido"), data("IDRepresentante"))
            If Not DtPVR Is Nothing AndAlso DtPVR.Rows.Count > 0 Then
                ApplicationService.GenerateError("Ya existe un registro con ese ID Linea y ese Representante.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarRepresentanteExistente(ByVal data As DataRow, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ProcesoComercial.ValidarRepresentanteExistente, New DataRowPropertyAccessor(data), services)
    End Sub

    <Task()> Public Shared Sub ValidarImporteValido(ByVal data As DataRow, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ProcesoComercial.ValidarImporteValido, New DataRowPropertyAccessor(data), services)
    End Sub

    <Task()> Public Shared Sub ValidarComisionValida(ByVal data As DataRow, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ProcesoComercial.ValidarComisionValida, New DataRowPropertyAccessor(data), services)
    End Sub

#End Region


End Class