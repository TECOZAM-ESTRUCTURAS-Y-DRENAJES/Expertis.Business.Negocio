Public Class AvalProveedor

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbAvalProveedor"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDProveedor")) <= 0 Then ApplicationService.GenerateError("El Número de Proveedor es un dato Obligatorio.")
        If Length(data("IDObra")) <= 0 Then ApplicationService.GenerateError("El Número de Obra es un dato Obligatorio.")
        If Length(data("Importe")) <= 0 Then ApplicationService.GenerateError("El Importe del Aval del Proveedor es un Dato Obligatorio.")
        If Length(data("FechaVencimiento")) <= 0 Then ApplicationService.GenerateError("La Fecha del Vencimiento es un dato Obligatorio.")
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClave)
    End Sub

    <Task()> Public Shared Sub AsignarClave(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDAvalProveedor")) = 0 Then data("IDAvalProveedor") = AdminData.GetAutoNumeric()
        End If
    End Sub

#End Region

End Class