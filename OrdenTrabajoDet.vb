Option Strict Off
Option Explicit On
Option Compare Text

Public Class OrdenTrabajodet
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbOrdenTrabajodet"

    Public Overloads Sub Delete(ByVal strIDOrdenTrabajoDet As DataRow)

    End Sub

    Public Overloads Function Update(ByVal dttSource As System.Data.DataTable) As System.Data.DataTable
        If Not dttSource Is Nothing AndAlso dttSource.Rows.Count > 0 Then
            Me.BeginTx()
            For Each dr As DataRow In dttSource.Rows
                'If Length(dr("DescPrevision")) = 0 Then ApplicationService.GenerateError("La Descripción de la Previsión es obligatoria")

                Dim dtTarifa As DataTable

                If dr.RowState = DataRowState.Added Then

                    If Not IsDBNull(dr("idOrdentrabajo")) Then
                        dr("idOrdentrabajodet") = AdminData.GetAutoNumeric
                    End If

                    ''Comprobación de la existencia de la Prevision
                    'dtTarifa = SelOnPrimaryKey(dr("IDPrevision"))
                    'If dtTarifa.Rows.Count <> 0 Then GenerateMessage("La Previsión ya existe", Me.GetType.Name & ".Update")

                End If
            Next
            AdminData.SetData(dttSource)
        End If
        Return dttSource
    End Function

    Public Overrides Function AddNewForm() As DataTable
        Dim dt As DataTable = MyBase.AddNewForm

        dt.Rows(0)("idOrdentrabajodet") = AdminData.GetAutoNumeric
        Return dt

    End Function

End Class
