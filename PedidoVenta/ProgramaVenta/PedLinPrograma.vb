Public Class PedLinPrograma
    Public IDPrograma As String
    Public IDLineaPrograma As Integer
    Public Cantidad As Double
    Public FechaConfirmacion As Date?
    Public QConfirmada As Double

    Public Sub New(ByVal oRow As DataRow)
        IDPrograma = oRow("IDPrograma")
        IDLineaPrograma = oRow("IDLineaPrograma")
        Cantidad = Double.NaN
        QConfirmada = Double.NaN
    End Sub

End Class
