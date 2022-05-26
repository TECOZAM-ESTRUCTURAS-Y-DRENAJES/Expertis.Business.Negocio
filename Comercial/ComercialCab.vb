Public Class ComercialCab
    Inherits Cabecera

    Public IDFormaEnvio As String
    Public IDCondicionEnvio As String
    Public IDModoTransporte As String

    Public ObsComerciales As String
    Public IDCliente As String
    Public Edi As Boolean

    Public Sub New(ByVal oRow As IPropertyAccessor)
        MyBase.New(oRow)
        IDCliente = oRow("IDCliente")
        Edi = Nz(oRow("EDI"), False)

        If oRow.ContainsKey("IDFormaEnvio") AndAlso Length(oRow("IDFormaEnvio")) > 0 Then IDFormaEnvio = oRow("IDFormaEnvio")
        If oRow.ContainsKey("IDCondicionEnvio") AndAlso Length(oRow("IDCondicionEnvio")) > 0 Then IDCondicionEnvio = oRow("IDCondicionEnvio")
        If oRow.ContainsKey("IDModoTransporte") AndAlso Length(oRow("IDModoTransporte")) > 0 Then IDModoTransporte = oRow("IDModoTransporte")
    End Sub

End Class

