Public Class DataCambioMoneda
    Public Row As IPropertyAccessor
    Public IDMonedaOld As String
    Public IDMonedaNew As String
    Public Fecha As Date

    Public Sub New(ByVal Row As IPropertyAccessor, ByVal IDMonedaOld As String, ByVal IDMonedaNew As String, Optional ByVal Fecha As Date = cnMinDate)
        Me.Row = Row
        Me.IDMonedaOld = IDMonedaOld
        Me.IDMonedaNew = IDMonedaNew
        Me.Fecha = Fecha
    End Sub
End Class
