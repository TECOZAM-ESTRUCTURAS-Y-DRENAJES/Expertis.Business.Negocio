<Serializable()> _
Public Class DataPrcFacturacionGeneral

    Public IDAlbaranes() As Integer
    Public IDContador As String
    Public DteFechaFactura As Date
    Public ConPropuesta As Boolean = True

    Public Sub New(ByVal IDAlbaranes() As Integer, ByVal IDContador As String, ByVal DteFechaFactura As Date, Optional ByVal ConPropuesta As Boolean = True)
        Me.IDAlbaranes = IDAlbaranes
        Me.IDContador = IDContador
        Me.DteFechaFactura = DteFechaFactura
        Me.ConPropuesta = ConPropuesta
    End Sub
End Class
