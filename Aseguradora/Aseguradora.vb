Public Class Aseguradora

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroAseguradora"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("DescAseguradora")) = 0 Then ApplicationService.GenerateError("Introduzca la descripción de la aseguradora")
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDAseguradora")) > 0 Then
                Dim DtTemp As DataTable = New Aseguradora().SelOnPrimaryKey(data("IDAseguradora"))
                If Not DtTemp Is Nothing AndAlso DtTemp.Rows.Count > 0 Then
                    ApplicationService.GenerateError("Ya existe una Aseguradora con esa clave.")
                End If
            Else : ApplicationService.GenerateError("Introduzca el código de la Asegurador")
            End If
        End If
    End Sub

#End Region

End Class