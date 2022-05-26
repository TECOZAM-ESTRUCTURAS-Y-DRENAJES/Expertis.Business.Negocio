Public Class Cabecera

    Public IDMoneda As String
    Public IDFormaPago As String
    Public IDCondicionPago As String
    ' Public IDBancoPropio As String
    Public IDCentroGestion As String
    Public Fecha As Date
    
    Public Sub New(ByVal oRow As IPropertyAccessor)
        IDMoneda = oRow("IDMoneda") & String.Empty
        IDFormaPago = oRow("IDFormaPago") & String.Empty
        IDCondicionPago = oRow("IDCondicionPago") & String.Empty
        'If Not oRow.IsNull("IDBancoPropio") Then IDBancoPropio = oRow("IDBancoPropio")
        IDCentroGestion = oRow("IDCentroGestion") & String.Empty
    End Sub

End Class
