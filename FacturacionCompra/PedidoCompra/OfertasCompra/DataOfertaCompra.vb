<Serializable()> _
Public Class DataOfertaCompra
    Public IDLineaOferta As Integer
    Public QOferta As Double

    Public Sub New(ByVal IDLineaOferta As Integer, ByVal QOferta As Double)
        Me.IDLineaOferta = IDLineaOferta
        Me.QOferta = QOferta
    End Sub
End Class
