Public Class FraCabMnto
    Inherits FraCab

    Public IDOT As Integer
    Public NROT As String

    Public IDDiaPago As String

    Public Lineas(-1) As FraLinMnto

    Public Sub New(ByVal oRow As IPropertyAccessor)
        MyBase.New(oRow)
        IDOT = oRow("IDOT")
        NROT = oRow("NROT")
        Fecha = oRow("Fecha")
        If Length(oRow("IDDiaPago")) > 0 Then IDDiaPago = oRow("IDDiaPago")
    End Sub

    Public Sub Add(ByVal lin As FraLinMnto)
        ReDim Preserve Lineas(Lineas.Length)
        Lineas(Lineas.Length - 1) = lin
    End Sub
End Class
