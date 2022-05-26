Public Class FraLinEntregaCta
    Public IDEntrega As Integer
    Public IDArticulo As String
    Public CContable As String
    Public Cantidad As Double
    Public Precio As Double
    'Public IDObra As Integer   '?SE coge de la cabecera?

    Public Sub New(ByVal oRow As DataRow)
        Me.IDEntrega = oRow("IDEntrega")
        Me.IDArticulo = oRow("IDArticulo")
        Me.CContable = oRow("CCArticulo")
        Me.Cantidad = 1
        Me.Precio = oRow("Importe")
        '   Me.IDObra = oRow("IDObra")
    End Sub
End Class