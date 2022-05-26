Option Strict Off
Option Explicit On
Option Compare Text

Imports Solmicro.Expertis.Engine.UI

Public Class PrevisionTecoCabecera
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbPrevisionTecoCabecera"

    'Public Overloads Sub Delete(ByVal strIDPrevision As String)
    '    If Not MyBase.Delete(strIDPrevision) Then
    '        ApplicationService.GenerateError(DELETECONSTRAINTMESSAGE)
    '    Else

    '    End If
    'End Sub

    Public Overloads Function Update(ByVal dttSource As System.Data.DataTable) As System.Data.DataTable
        If Not dttSource Is Nothing AndAlso dttSource.Rows.Count > 0 Then
            Me.BeginTx()
            For Each dr As DataRow In dttSource.Rows
                If Length(dr("DescPrevision")) = 0 Then ApplicationService.GenerateError("La Descripción de la Previsión es obligatoria")

                Dim dtTarifa As DataTable

                If dr.RowState = DataRowState.Added Then

                    If Not IsDBNull(dr("IDPrevision")) Then
                        dr("IDPrevision") = AdminData.GetAutoNumeric
                    End If

                    'Comprobación de la existencia de la Prevision
                    dtTarifa = SelOnPrimaryKey(dr("IDPrevision"))
                    If dtTarifa.Rows.Count <> 0 Then MsgBox("La Previsión ya existe")

                End If
            Next
            UpdateTable(dttSource)
        End If
        Return dttSource
    End Function

    Public Overrides Function AddNewForm() As DataTable
        Dim dt As DataTable = MyBase.AddNewForm

        Try
            Dim cLineas As New PrevisionTecoLinea
            Dim dtCabecera As DataTable = Nothing
            If comprobarCabecera(dtCabecera) < 0 Then

            End If
            If dtCabecera.Rows.Count <= 0 OrElse IsDBNull(dtCabecera.Rows(0)("IDPrevision")) Then
                ' Generar por ejercicio las diferentes situaciones
                '' Si llega hasta aqui generar la cabecera
                dt.Rows(0)("IDPrevision") = AdminData.GetAutoNumeric
                dt.Rows(0)("DescPrevision") = "PREVISIÓN FINANCIERA TECOZAM"

                ' Grabar cabecera
                UpdateTable(dt)
                dt.AcceptChanges()
            Else
                dt.Rows(0).ItemArray = dtCabecera.Rows(0).ItemArray

            End If
            ' Crear los detalles
            Dim dtDetalles As New DataTable
            If cLineas.CrearActLineas(dt.Rows(0)("IDPrevision")) < 0 Then
                Return Nothing
            End If
            ' Dejar el estado para q no cree otra
            Return dt
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error al crear gestión de cobros.")
        End Try

    End Function

    Private Function comprobarCabecera(ByRef dt As DataTable) As Short
        dt = AdminData.GetData("SELECT * FROM xMaestros5.dbo.tbPrevisionTecoCabecera")
        ' Bien
        Return 1
    End Function
End Class
