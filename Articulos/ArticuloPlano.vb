Public Class ArticuloPlano

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbArticuloPlano"

#End Region

#Region "Eventos GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBrl As New BusinessRules
        oBrl.Add("FechaEdicion", AddressOf CambioFecha)
        oBrl.Add("FechaVigor", AddressOf CambioFecha)
        Return oBrl
    End Function

    <Task()> Public Shared Sub CambioFecha(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Not IsDate(data.Value) Then ApplicationService.GenerateError("Campo no numérico.")
    End Sub

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("CodigoPlano")) = 0 Then ApplicationService.GenerateError("Código plano es un dato obligatorio.") '
        If Length(data("DescPlano")) = 0 Then ApplicationService.GenerateError("La descripción del plano es un dato obligatorio.") ' '
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClave)
    End Sub

    <Task()> Public Shared Sub AsignarClave(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added AndAlso Length(data("IDArticuloPlano")) = 0 Then
            data("IDArticuloPlano") = AdminData.GetAutoNumeric
        End If
    End Sub

#End Region

End Class