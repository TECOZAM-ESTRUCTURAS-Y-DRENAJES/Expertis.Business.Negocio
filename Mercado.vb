Public Class Mercado

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroMercado"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDMercado")) = 0 Then ApplicationService.GenerateError("El Identificativo es obligatorio.")
        If Length(data("DescMercado")) = 0 Then ApplicationService.GenerateError("La Descripción es obligatoria.")
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = New Mercado().SelOnPrimaryKey(data("IDMercado"))
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then ApplicationService.GenerateError("El Código de Mercado está duplicado.")
        End If
    End Sub

#End Region

End Class