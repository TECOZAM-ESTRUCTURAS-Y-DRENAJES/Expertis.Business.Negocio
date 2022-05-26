Option Strict Off
Option Explicit On 
Option Compare Text

Public Class GestionCobrosLineas
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbGestionCobrosLineas"

    Public Overloads Sub Delete(ByVal data As DataRow, ByVal strIDGestion As String)
        If Not MyBase.Delete(data(strIDGestion)) Then
            ApplicationService.GenerateError(DELETECONSTRAINTMESSAGE)
        End If
    End Sub

    Public Overloads Function Update(ByVal dttSource As System.Data.DataTable) As System.Data.DataTable
        If Not dttSource Is Nothing AndAlso dttSource.Rows.Count > 0 Then
            Me.BeginTx()
            For Each dr As DataRow In dttSource.Rows

                Dim dtTarifa As DataTable

                If dr.RowState = DataRowState.Added Then

                    If Not IsDBNull(dr("idGestionCobrosLin")) Then
                        dr("idGestionCobrosLin") = AdminData.GetAutoNumeric
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

        dt.Rows(0)("idGestionCobrosLin") = AdminData.GetAutoNumeric
        Return dt

    End Function

    ' Función para obtener las diferentes situaciones de cobro
    Private Function TraerTiposCobro() As DataTable
        Dim dtTiposCobro As New DataTable
        Try
            dtTiposCobro = AdminData.GetData("SELECT IDEstado, DescEstado FROM tbMaestroEstadoPago")
            ' Control de filas
            If dtTiposCobro.Rows.Count <= 0 Then
                MsgBox("No se han obtenido estados de cobro.", MsgBoxStyle.Exclamation, "Sin estados de cobro")
                Return Nothing
            End If
            ' Bien
            Return dtTiposCobro
        Catch ex As Exception
            MsgBox("Se produjo un error al obtener los estados de cobro." & ex.Message, MsgBoxStyle.Exclamation, "Sin estados de cobro")
            Return Nothing
        End Try
    End Function

    'Función para crear las lineas por cada cabecera
    Public Function CrearLineas(ByVal iIdCabecera As Integer) As DataTable
        Try
            Dim dt As DataTable = MyBase.AddNewForm
            Dim dtCobros As New DataTable
            Dim shFila As Short = 0
            dtCobros = TraerTiposCobro()
            ' Por cada tipo de cobro
            If Not IsNothing(dtCobros) Then
                For shcont As Short = 0 To dtCobros.Rows.Count - 1
                    ' Por las 12 mensualidades del año
                    ' Crear lineas
                    For shcontMeses As Short = 0 To 11
                        dt.Rows(shFila)("idGestionCobrosLin") = AdminData.GetAutoNumeric
                        dt.Rows(shFila)("idGestionCobros") = iIdCabecera
                        dt.Rows(shFila)("mes") = shcontMeses + 1
                        dt.Rows(shFila)("situacion") = dtCobros.Rows(shcont)("IDEstado")
                        'dt.Rows(shFila)("DescEstado") = dtCobros.Rows(shcont)("DescEstado")
                        dt.Rows(shFila)("impcobros") = 0

                        'IBIS. David. 18/10/2010. Cambiado el For de 5 a 8, porqué añadimos 2 nuevas filas.
                        For shPagos As Short = 1 To 4
                            dt.Rows(shFila)("imppagos" & shPagos.ToString) = 0
                        Next

                        ' Control de última línea de cada estado
                        Dim dfila As DataRow
                        dfila = dt.NewRow
                        dt.Rows.Add(dfila)
                        shFila += 1
                    Next
                Next
                ' Borrar última linea generada de más
                dt.Rows(dt.Rows.Count - 1).Delete()
            Else
                MsgBox("No se han cargado los tipos de pago.", MsgBoxStyle.Exclamation, "Error al obtener datos")
            End If
            ' De todas Retorna un dt Válido
            'Ibis 10-11-2011 Nuevo Proceso para traer los cobros
            Dim DtSCobros As DataTable
            DtSCobros = CrearLineasCobros(iIdCabecera)

            If Not IsNothing(DtSCobros) AndAlso DtSCobros.Rows.Count > 0 Then
                For Each dr As DataRow In DtSCobros.Rows
                    dt.ImportRow(dr)
                Next

            End If

            Return dt
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error")
            Return Nothing
        End Try
    End Function

#Region " Ibis Computer"
    Public Function CrearLineasCobros(ByVal iIdCabecera As Integer) As DataTable
        Try
            Dim dt As DataTable = MyBase.AddNewForm
            Dim dtCobros As New DataTable
            Dim shFila As Short = 0
            dtCobros = AdminData.GetData("SELECT IDEstado, DescEstado FROM tbMaestroEstadoCobro where IDEstado in (0,1,2,99)")
            ' Por cada tipo de cobro
            If Not IsNothing(dtCobros) Then
                For shcont As Short = 0 To dtCobros.Rows.Count - 1
                    ' Por las 12 mensualidades del año
                    ' Crear lineas
                    For shcontMeses As Short = 0 To 11
                        dt.Rows(shFila)("idGestionCobrosLin") = AdminData.GetAutoNumeric
                        dt.Rows(shFila)("idGestionCobros") = iIdCabecera
                        dt.Rows(shFila)("mes") = shcontMeses + 1
                        dt.Rows(shFila)("situacion") = dtCobros.Rows(shcont)("IDEstado")
                        'dt.Rows(shFila)("DescEstado") = dtCobros.Rows(shcont)("DescEstado")
                        dt.Rows(shFila)("impcobros") = 0
                        dt.Rows(shFila)("Tipo") = "Cobro"


                        For shPagos As Short = 1 To 4
                            dt.Rows(shFila)("imppagos" & shPagos.ToString) = 0
                        Next


                        Dim dfila As DataRow
                        dfila = dt.NewRow
                        dt.Rows.Add(dfila)
                        shFila += 1
                    Next
                Next
                ' Borrar última linea generada de más
                dt.Rows(dt.Rows.Count - 1).Delete()
            End If

            Return dt
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error")
            Return Nothing
        End Try
    End Function
#End Region
End Class
