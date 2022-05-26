Public Class PedidoVentaAnalitica
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper
    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbPedidoVentaAnalitica"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBRL As New BusinessRules
        oBRL.Add("Importe", AddressOf NegocioGeneral.AnaliticaCommonBusinessRules)
        oBRL.Add("Porcentaje", AddressOf NegocioGeneral.AnaliticaCommonBusinessRules)
        Return oBRL
    End Function

    'Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
    '    MyBase.RegisterValidateTasks(validateProcess)
    '    validateProcess.AddTask(Of DataTable)(AddressOf NegocioGeneral.AnaliticaCommonValidateRules)
    'End Sub

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.AnaliticaCommonValidateRules)
    End Sub

End Class