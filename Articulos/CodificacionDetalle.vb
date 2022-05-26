Public Class CodificacionDetalle

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbCodificacionDetalle"

#End Region

#Region "Eventos ValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarTablaCaracteristica)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("Orden")) = 0 Then ApplicationService.GenerateError("El campo Orden es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarTablaCaracteristica(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDCampo")) = 0 AndAlso Length(data("IDCaracteristica")) = 0 Then
            ApplicationService.GenerateError("Debe especificar o Campo de tabla o Característica para configurar los datos.", vbNewLine)
        End If
        If Length(data("IDCampo")) > 0 AndAlso Length(data("IDCaracteristica")) > 0 Then
            ApplicationService.GenerateError("El campo Tabla y Característica no pueden tener valores simultaneamente.|Por favor, revise los datos.", vbNewLine)
        End If
    End Sub

#End Region

#Region "Eventos UpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClavePrimaria)
    End Sub

    <Task()> Public Shared Sub AsignarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDCodigoDetalle")) = 0 Then data("IDCodigoDetalle") = AdminData.GetAutoNumeric
        End If
    End Sub

#End Region

End Class