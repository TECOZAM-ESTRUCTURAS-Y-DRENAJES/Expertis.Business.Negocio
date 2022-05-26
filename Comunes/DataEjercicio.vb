Public Class DataEjercicio
    Public Datos As IPropertyAccessor
    Public Fecha As Date

    'Public Sub New()
    '    MyBase.New()
    'End Sub

    Public Sub New(ByVal Datos As IPropertyAccessor, ByVal Fecha As Date)
        Me.Datos = Datos
        Me.Fecha = Fecha
    End Sub

End Class