Public Class ClaseAval

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper
    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroAvalClase"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDAvalClase")) = 0 Then ApplicationService.GenerateError("La Clase de Aval es un Dato Obligatorio")
        If Length(data("DescAvalClase")) = 0 Then ApplicationService.GenerateError("La Descripción es un Dato Obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = New ClaseAval().SelOnPrimaryKey(data("IDAvalClase"))
            If Not IsNothing(dt) AndAlso Not dt.Rows.Count.Equals(0) Then
                ApplicationService.GenerateError("La Clase de Aval introducida ya Existe.")
            End If
        End If
    End Sub

#End Region

End Class