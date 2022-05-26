Public Class Modelo

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroModelo"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("DescModelo")) = 0 Then ApplicationService.GenerateError("La Descripción es obligatoria.")
        If Length(data("IDModelo")) = 0 Then ApplicationService.GenerateError("El modelo es obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dtModelo As DataTable = New Modelo().SelOnPrimaryKey(data("IDModelo"))
            If Not dtModelo Is Nothing AndAlso dtModelo.Rows.Count > 0 Then ApplicationService.GenerateError("Código de modelo duplicado")
        End If
    End Sub

#End Region

End Class