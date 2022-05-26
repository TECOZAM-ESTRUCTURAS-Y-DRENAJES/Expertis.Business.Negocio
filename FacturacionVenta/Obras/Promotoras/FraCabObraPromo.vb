Public Class FraCabObraPromo
    Inherits FraCabVencimiento

    Public DescObra As String
    Public NRegistro As Integer
    Public TipoFactura As Integer
    Public CCPrestamoHip As String
    Public CCCompCliente As String
    Public CCAnticipo As String
    Public IDLocal As Integer
    Public DescLocal As String
    Public CobroGenerado As Boolean

    Public Lineas(-1) As FraLinObraPromo

    Public Sub New(ByVal oRow As DataRow, ByVal TipoFactura As Integer) 'IPropertyAccessor)
        MyBase.New(oRow)

        Me.TipoFactura = TipoFactura
        Me.CCPrestamoHip = oRow("CCPrestamoHip") & String.Empty
        Me.CCCompCliente = oRow("CCCompCliente") & String.Empty
        Me.CCAnticipo = oRow("CCAnticipo") & String.Empty
        Me.DescObra = oRow("DescObra") & String.Empty
        Me.IDLocal = oRow("IDLocal")
        Me.DescLocal = oRow("DescLocal") & String.Empty
        Me.CobroGenerado = Nz(oRow("CobroGenerado"), False)
        Me.Fecha = Nz(oRow("FechaVencimiento"), Date.Today)
    End Sub

    Public Sub Add(ByVal lin As FraLinObraPromo)
        ReDim Preserve Lineas(Lineas.Length)
        Lineas(Lineas.Length - 1) = lin
    End Sub

End Class