Public Class DataObservaciones
    Public Entity As String
    Public Datos As IPropertyAccessor
    Public Field As String

    Public Sub New(ByVal Entity As String, ByVal Field As String, ByVal datos As IPropertyAccessor)
        Me.Entity = Entity
        Me.Datos = datos
        Me.Field = Field
    End Sub
End Class
