Public Class PedCabCompraObra
    Inherits PedCabCompra

    Public PorMateriales As Boolean
    Public PorTrabajos As Boolean

    Public Sub New(ByVal oRow As DataRow, ByVal PorMateriales As Boolean, ByVal PorTrabajos As Boolean)
        MyBase.New(oRow)
        Me.PorMateriales = PorMateriales
        Me.PorTrabajos = PorTrabajos
        If Me.PorMateriales Then
            MyBase.ViewName = "vFrmMntoObraGeneraCompra"
        Else
            MyBase.ViewName = "vFrmMntoObraGeneraCompraTrabajo"
        End If

        MyBase.Origen = enumOrigenPedidoCompra.Obras
    End Sub

    Public Overrides Function PrimaryKeyCabOrigen() As String
        Return "IDObra"
    End Function

    Public Overrides Function FieldNOrigen() As String
        Return "NObra"
    End Function
End Class
