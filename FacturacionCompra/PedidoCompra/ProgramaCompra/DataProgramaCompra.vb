<Serializable()> _
Public Class DataProgramaCompra
    Public IDLineaPrograma As Integer
    Public Cantidad As Double     '?
    Public QConfirmada As Double   '?
    Public FechaConfirmacion As Date

    Public Sub New(ByVal IDLineaPrograma As Integer, ByVal Cantidad As Double, ByVal QConfirmada As Double, Optional ByVal FechaConfirmacion As Date = cnMinDate)
        Me.IDLineaPrograma = IDLineaPrograma
        Me.Cantidad = Cantidad
        Me.QConfirmada = QConfirmada
        If FechaConfirmacion <> cnMinDate Then Me.FechaConfirmacion = FechaConfirmacion
    End Sub
End Class