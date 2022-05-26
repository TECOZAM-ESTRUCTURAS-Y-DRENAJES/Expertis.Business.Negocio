Public Class HistoricoTipoIVA

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbHistoricoTipoIVA"

#End Region

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("FechaDesde")) = 0 Then ApplicationService.GenerateError("La Fecha Desde es un dato obligatorio.")
        If Length(data("FechaHasta")) = 0 Then ApplicationService.GenerateError("La Fecha Hasta es un dato obligatorio.")
        If Length(data("Factor")) = 0 Then ApplicationService.GenerateError("El Factor es un dato obligatorio.")
        If Length(data("IvaRE")) = 0 Then ApplicationService.GenerateError("El IVA R.E. es un dato obligatorio.")
        If Length(data("IVAIntrastat")) = 0 Then ApplicationService.GenerateError("El IVA Intrastat es un dato obligatorio.")
        If Length(data("IVASinRepercutir")) = 0 Then ApplicationService.GenerateError("El IVA Sin Repercutir es un dato obligatorio.")
    End Sub

End Class