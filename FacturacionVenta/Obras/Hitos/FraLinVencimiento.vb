Public Class FraLinVencimiento
    Inherits FraLinObra

    Public IdVencimiento As Integer
    Public DescVencimiento As String
    'Public IdArticulo As String

    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(oRow)

        IdVencimiento = oRow("IDVencimiento")
        DescVencimiento = oRow("DescVencimiento") & String.Empty
        'IdArticulo = oRow("DescVencimiento") & String.Empty
    End Sub

End Class
