Public Class DataDireccionClte
    Public Datos As IPropertyAccessor
    Public Field As String
    Public TipoDireccion As enumcdTipoDireccion

    Public Sub New(ByVal TipoDireccion As enumcdTipoDireccion, ByVal Field As String, ByVal datos As IPropertyAccessor)
        Me.TipoDireccion = TipoDireccion
        Me.Datos = datos
        Me.Field = Field
    End Sub
End Class
