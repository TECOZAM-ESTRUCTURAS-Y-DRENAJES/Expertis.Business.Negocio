
Public Class FraCabEntregaCta
    Inherits FraCab

    Public IDEntrega As Integer
    Public CContableCliente As String
    Public IDTipoCobro As String
    Public TipoEntrega As Integer

    Public Lineas(-1) As FraLinEntregaCta

    Public Sub New(ByVal oRow As IPropertyAccessor)
        MyBase.New(oRow)
        Me.IDEntrega = oRow("IDEntrega")
        Me.IDCliente = oRow("IDCliente")
        Me.Fecha = oRow("FechaEntrega")
        If Length(oRow("IDObra")) > 0 Then Me.IDObra = oRow("IDObra")
        If Length(oRow("IDBancoPropio")) > 0 Then Me.IDBancoPropio = oRow("IDBancoPropio")
        If Length(oRow("IDMoneda")) > 0 Then Me.IDMoneda = oRow("IDMoneda")
        Me.CContableCliente = oRow("CCClienteProveedor")
        Me.TipoEntrega = oRow("TipoEntrega")
        Me.IDTipoCobro = oRow("IDTipoCobroPago")
    End Sub

    Public Sub Add(ByVal lin As FraLinEntregaCta)
        ReDim Preserve Lineas(Lineas.Length)
        Lineas(Lineas.Length - 1) = lin
    End Sub
End Class