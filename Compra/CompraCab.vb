Public Class CompraCab
    Inherits Cabecera

    Public IDProveedor As String

    Public Sub New(ByVal oRow As IPropertyAccessor)
        MyBase.New(oRow)
        IDProveedor = oRow("IDProveedor")
    End Sub

End Class
