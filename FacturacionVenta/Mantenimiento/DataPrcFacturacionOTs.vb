<Serializable()> _
Public Class DataPrcFacturacionOTs

    Public IDMntoOTControl() As Integer
    Public IDContador As String
    Public FechaFactura As Date?

    Public Sub New(ByVal IDMntoOTControl() As Integer, Optional ByVal IDContador As String = Nothing, Optional ByVal FechaFactura As Date = Nothing)
        Me.IDMntoOTControl = IDMntoOTControl
        If Length(IDContador) > 0 Then Me.IDContador = IDContador
        If Length(FechaFactura) > 0 AndAlso FechaFactura <> cnMinDate Then Me.FechaFactura = FechaFactura
    End Sub

End Class
