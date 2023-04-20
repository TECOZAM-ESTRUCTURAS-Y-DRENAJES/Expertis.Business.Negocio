<Serializable()> _
    Public Class DataPrcCrearPedidoCompraPrograma
    Public Programas() As DataProgramaCompra
    Public IDContador As String

    Public Sub New(ByVal Programas() As DataProgramaCompra, ByVal IDContador As String)
        Me.Programas = Programas
        Me.IDContador = IDContador
    End Sub

End Class
