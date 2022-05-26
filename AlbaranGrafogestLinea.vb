Public Class AlbaranGrafogestLinea
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Private Const cnEntidad As String = "tbAlbaranGrafogestLinea"

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Public Overloads Sub Delete(ByVal intIDAlbaranGrafogestLinea As Integer)


    End Sub

    Public Overloads Function Update(ByVal dttSource As System.Data.DataTable) As System.Data.DataTable
        If Not IsNothing(dttSource) AndAlso dttSource.Rows.Count > 0 Then
            Me.BeginTx()
            For Each dr As DataRow In dttSource.Rows
                If dr.RowState = DataRowState.Added Or dr.RowState = DataRowState.Modified Then
                    If Length(dr("IDLineaAlbaranGrafogest")) = 0 Then
                        dr("IDLineaAlbaranGrafogest") = AdminData.GetAutoNumeric
                    End If
                End If
            Next
            AdminData.SetData(dttSource)
        End If
        Return dttSource
    End Function

End Class
