Public Class FraLinObra
    Public IDObra As Integer
    Public NObra As String
    Public NumeroPedido As String
    Public IDCentroGestion As String

    Public Sub New(ByVal oRow As DataRow)
        IDObra = oRow("IDObra")
        NObra = oRow("NObra")
        If Length(oRow("NumeroPedido")) > 0 Then NumeroPedido = oRow("NumeroPedido")
        If Length(oRow("IDCentroGestion")) > 0 Then IDCentroGestion = oRow("IDCentroGestion")
    End Sub

End Class
