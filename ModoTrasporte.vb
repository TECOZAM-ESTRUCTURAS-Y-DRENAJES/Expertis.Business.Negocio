Public Class ModoTrasporte

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroModoTrasporte"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDModoTransporte")) = 0 Then ApplicationService.GenerateError("El código del Modo de Transporte es obligatorio.")
        If Length(data("DescModoTransporte")) = 0 Then ApplicationService.GenerateError("La descripción del Modo de Transporte es obligatoria.")
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dtModo As DataTable = New ModoTrasporte().SelOnPrimaryKey(data("IDModoTransporte"))
            If Not dtModo Is Nothing AndAlso dtModo.Rows.Count > 0 Then
                ApplicationService.GenerateError("El Código del Modo de Transporte está duplicado.")
            End If
        End If
    End Sub

#End Region

End Class