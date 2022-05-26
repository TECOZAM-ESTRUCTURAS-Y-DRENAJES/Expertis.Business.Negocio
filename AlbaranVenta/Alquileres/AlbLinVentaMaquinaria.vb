Public Class AlbLinVentaMaquinaria
    Inherits AlbLinVentaObras

    Public IDActivo As String
    Public IDObra As Integer
    Public IDTrabajo As Integer
    Public TipoFactAlquiler As Integer
    Public Precio As Double
    Public Dto1 As Double
    Public row As DataRow

    Public OrigenDatos As enumOrigenDatosLineaVentaMaquinaria

    Public Enum enumOrigenDatosLineaVentaMaquinaria
        ObraMaterial
        Activo
    End Enum

    Public Overrides Function PrimaryKeyLinOrigen() As String
        Return New String("") '("IDActivo")
    End Function

    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(oRow)
        Me.row = oRow

        If oRow.Table.Columns.Contains("IDArticulo") Then
            Me.OrigenDatos = enumOrigenDatosLineaVentaMaquinaria.Activo
        ElseIf oRow.Table.Columns.Contains("IDMaterial") Then
            Me.OrigenDatos = enumOrigenDatosLineaVentaMaquinaria.ObraMaterial
        End If
    End Sub

End Class
