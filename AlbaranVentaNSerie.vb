Public Class AlbaranVentaNSerie
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub
    Private Const cnEntidad As String = "tbAlbaranVentaNSerie"

    Public Overloads Function Update(ByVal dttSource As System.Data.DataTable) As System.Data.DataTable
        If Not dttSource Is Nothing AndAlso dttSource.Rows.Count > 0 Then

            Me.BeginTx()
            For Each dr As DataRow In dttSource.Rows
                If dr.RowState = DataRowState.Added Then
                    dr("AlbaranVentaNSerie") = AdminData.GetAutoNumeric
                    If Length(dr("IDArticulo")) = 0 Then ApplicationService.GenerateError("El Articulo es un dato obligatorio.")
                    If Length(dr("NSerie")) = 0 Then ApplicationService.GenerateError("La Referencia es un dato obligatorio.")

                End If

            Next
            MyBase.Update(dttSource)
        End If

        Return dttSource
    End Function
End Class


