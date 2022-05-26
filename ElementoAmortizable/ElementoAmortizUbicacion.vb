Public Class ElementoAmortizUbicacion

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroElementoAmortizUbicacion"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDCentroGestion")) = 0 Then ApplicationService.GenerateError("Introduzca el centro de gestión")
        If Length(data("IDOperario")) = 0 Then ApplicationService.GenerateError("Introduzca el operario")
        If Length(data("Cantidad")) = 0 Then ApplicationService.GenerateError("La cantidad del artículo debe ser positiva")
    End Sub

#End Region

End Class