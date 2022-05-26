Public Class FacturaCompraBaseImponible
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper
    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbFacturaCompraBaseImponible"
#Region " Gestión cambio base imponible para ajustar facturas "
    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBRL As New BusinessRules
        oBRL.Add("BaseImponible", AddressOf ProcesoComunes.CalcularIVA)
        Return oBRL
    End Function
#End Region

End Class

