
<Serializable()> _
Public Class DataPrcCrearPedidoVentaPrograma
    Public Programas() As DataPrograma
    Public IDContador As String

    Public Sub New(ByVal Programas() As DataPrograma, ByVal IDContador As String)
        Me.Programas = Programas
        Me.IDContador = IDContador
    End Sub

End Class
