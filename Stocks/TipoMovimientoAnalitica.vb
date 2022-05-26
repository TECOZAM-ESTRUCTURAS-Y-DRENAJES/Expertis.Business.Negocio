Public Class TipoMovimientoAnalitica
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroTipoMovimientoAnalitica"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBRL As New BusinessRules
        oBRL.Add("Porcentaje", AddressOf CambioPorcentaje)
        Return oBRL
    End Function

    <Task()> Public Shared Sub CambioPorcentaje(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Not IsNumeric(data.Current("Porcentaje")) OrElse data.Current("Porcentaje") < 0 Then
            ApplicationService.GenerateError("El {0} no es válido.", Quoted(data.Current("Porcentaje")))
        ElseIf data.Current("Porcentaje") > 100 Then
            ApplicationService.GenerateError("El porcentaje no puede ser superior al 100%.")
        Else
            data.Current("Porcentaje") = xRound(data.Current("Porcentaje"), 2)
        End If
    End Sub

End Class
