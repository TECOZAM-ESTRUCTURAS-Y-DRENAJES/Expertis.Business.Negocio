<Serializable()> _
Public Class DataBaseImponible
    Public IDTipoIva As String
    'Public BaseImponibleNormal As Double
    'Public BaseImponibleEspecial As Double
    Public BaseImponible As Double
    'Public BaseImponibleNormalA As Double
    'Public BaseImponibleEspecialA As Double
    Public BaseImponibleA As Double
    'Public BaseImponibleNormalB As Double
    'Public BaseImponibleEspecialB As Double
    Public BaseImponibleB As Double
    Public ImporteIVA As Double
    Public ImporteIVAA As Double
    Public ImporteIVAB As Double
    Public ImporteIVANoDeducible As Double
    Public ImporteIVANoDeducibleA As Double
    Public ImporteIVANoDeducibleB As Double
    Public PorcenIVANoDeducible As Double
    Public Sub New(ByVal TipoIva As String)
        IDTipoIva = TipoIva
    End Sub
End Class
