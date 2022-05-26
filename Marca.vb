Public Class Marca

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroMarca"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("DescMarca")) = 0 Then ApplicationService.GenerateError("La Descripci�n es obligatoria.")
        If Length(data("IDMarca")) = 0 Then ApplicationService.GenerateError("La marca es obligatoria.")
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = New Marca().SelOnPrimaryKey(data("IDMarca"))
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then ApplicationService.GenerateError("Marca duplicada")
        End If
    End Sub

#End Region

End Class