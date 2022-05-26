Public Class DataCalculoRepresentantes
    Public IDCliente As String
    Public IDArticulo As String
    Public IDObra As Integer?
    Public Cantidad As Double
    'Friend Orden As Integer

    Public Representantes As DataTable          '//DataTable en el que se irán añadiendo los distintos representantes con sus importes/comisiones

    Public Sub New(ByVal IDCliente As String, ByVal IDArticulo As String, ByVal Cantidad As Double)
        Me.IDCliente = IDCliente
        Me.IDArticulo = IDArticulo
        Me.Cantidad = Cantidad
    End Sub

    Public Sub New(ByVal IDCliente As String, ByVal IDArticulo As String, ByVal Cantidad As Double, ByVal IDObra As Integer)
        Me.IDCliente = IDCliente
        Me.IDArticulo = IDArticulo
        Me.Cantidad = Cantidad
        Me.IDObra = IDObra
    End Sub

    Public Sub AddRepresentante(ByVal IDRepresentante As String, ByVal Porcentaje As Boolean, ByVal Comision As Double) ', ByVal Importe As Double) ', ByVal Orden As Integer)
        If Representantes Is Nothing Then CrearDTRepresentante()
        Dim f As New Filter
        f.Add(New StringFilterItem("IDRepresentante", IDRepresentante))
        Dim WhereRepresentante As String = f.Compose(New AdoFilterComposer)
        Dim ExisteRepresentante() As DataRow = Representantes.Select(WhereRepresentante)
        If ExisteRepresentante Is Nothing OrElse ExisteRepresentante.Length = 0 Then
            Dim dr As DataRow = Representantes.NewRow
            dr("IDRepresentante") = IDRepresentante
            dr("Comision") = Comision
            dr("Porcentaje") = Porcentaje
            ' dr("orden") = Orden
            Representantes.Rows.Add(dr)
        End If
    End Sub
    Public Sub AddRepresentantes(ByVal dtRepresentantes As DataTable)
        If dtRepresentantes Is Nothing Then Exit Sub
        For Each drRepresentante As DataRow In dtRepresentantes.Rows
            AddRepresentante(drRepresentante("IDRepresentante"), drRepresentante("Porcentaje"), drRepresentante("Comision"))
        Next
    End Sub

    Public Sub CrearDTRepresentante()
        Representantes = New DataTable
        Representantes.RemotingFormat = SerializationFormat.Binary
        Representantes.Columns.Add("IDRepresentante", GetType(String))
        Representantes.Columns.Add("Comision", GetType(Double))
        Representantes.Columns.Add("Importe", GetType(Double))
        Representantes.Columns.Add("Porcentaje", GetType(Boolean))
        'Representantes.Columns.Add("Orden", GetType(Integer))
    End Sub

End Class
