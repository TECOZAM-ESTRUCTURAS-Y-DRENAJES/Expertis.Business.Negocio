Public Class PedidoCompraEspecificacion
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper
    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbPedidoCompraEspecificacion"

End Class