Public Class DataDireccionProv
    Public Datos As IPropertyAccessor
    Public Field As String
    Public TipoDireccion As enumpdTipoDireccion

    Public Sub New(ByVal TipoDireccion As enumpdTipoDireccion, ByVal Field As String, ByVal datos As IPropertyAccessor)
        Me.TipoDireccion = TipoDireccion
        Me.Datos = datos
        Me.Field = Field
    End Sub
End Class
