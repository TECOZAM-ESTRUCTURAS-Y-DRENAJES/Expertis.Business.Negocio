Option Strict Off
Option Explicit On 
Option Compare Text

Public Class GestionCobrosCabecera

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Property iEjer() As Integer
        Get
            Return iEjercicio
        End Get
        Set(ByVal Value As Integer)
            iEjercicio = Value
        End Set
    End Property

    Private iEjercicio As Integer = 0

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbGestionCobrosCabecera"

    Public Overridable Overloads Sub Delete(ByVal data As DataRow, ByVal strIDGestion As String)
        ' Borrar tambien las lineas
        Try
            If Not MyBase.Delete(data(strIDGestion)) Then
                ApplicationService.GenerateError(DELETECONSTRAINTMESSAGE)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Public Overloads Function Update(ByVal dttSource As System.Data.DataTable) As System.Data.DataTable
        If Not dttSource Is Nothing AndAlso dttSource.Rows.Count > 0 Then
            Me.BeginTx()
            For Each dr As DataRow In dttSource.Rows

                Dim dtTarifa As DataTable

                If dr.RowState = DataRowState.Added Then

                    If Not IsDBNull(dr("idGestionCobros")) Then
                        dr("idGestionCobros") = AdminData.GetAutoNumeric
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
        Try
            Dim cLineas As New GestionCobrosLineas
            Dim dt As DataTable = ComprobarEjercicio(iEjer)
            If dt.Rows.Count <= 0 OrElse IsDBNull(dt.Rows(0)("idgestioncobros")) Then
                ' Generar por ejercicio las diferentes situaciones
                '' Si llega hasta aqui generar la cabecera
                dt = MyBase.AddNewForm()
                dt.Rows(0)("idgestioncobros") = AdminData.GetAutoNumeric
                dt.Rows(0)("idejercicio") = iEjer
                dt.Rows(0)("comentario") = "CONTROL DE PAGOS " & iEjer & "."
                ' Grabar cabecera
                'AdminData.SetData(dt)
                dt.AcceptChanges()
                ' Crear los detalles
                Dim dtDetalles As New DataTable
                dtDetalles = cLineas.CrearLineas(dt.Rows(0)("idgestioncobros"))
                ' Si no es nula grabar lineas
                If Not IsNothing(dtDetalles) Then
                    '' Obtener los pagos

                    'IBIS. David. 18/10/2010. En el For cambiamos el 5 por un 8, porqué hemos añadido 2 empresas más 
                    For shEmp As Short = 1 To 8

                        'IBIS. David. 18/10/2010. Nos saltamos el 6, porqué es la columna 'Salario'
                        If shEmp <> 6 Then
                            ObtenerPagos(dtDetalles, shEmp, dt.Rows(0)("idejercicio"))
                            ObtenerCobros(dtDetalles, shEmp, dt.Rows(0)("idejercicio"))
                        End If

                    Next
                    AdminData.SetData(dtDetalles)
                End If
                dtDetalles = Nothing
            End If
            ' Dejar el estado para q no cree otra
            Return dt
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error al crear gestión de cobros.")
        End Try

    End Function

    Private Function ComprobarEjercicio(ByVal strEjercicio As Integer) As DataTable
        Dim dt As New DataTable
        dt = AdminData.GetData("SELECT * FROM tbGestionCobrosCabecera WHERE idejercicio = " & strEjercicio)

        ' De todas Retorna un dt Válido
        Return dt
    End Function

    Public Function ActualizarCobros(ByVal strIdcobros As String, ByVal shEjercicio As String) As Short
        '' Con el id de la cabecera actualizar las lineas
        ' Crear los detalles
        Try
            Dim dtDetalles As New DataTable
            Dim dttmp As New DataTable
            Dim cLineas As New GestionCobrosLineas

            ' Resetear pagos a 0
            'IBIS. David. 18/10/2010. Se añaden 2 nuevas columnas, impppagos7 e imppagos 8.
            Dim strUpdate As String = "UPDATE tbGestionCobrosLineas SET imppagos1 = 0,imppagos2 = 0, imppagos3 = 0, imppagos4 = 0, imppagos5 = 0 WHERE idGestionCobros = " & strIdcobros
            'AdminData.ExecuteNonQuery(strUpdate, False)
            AdminData.Execute(strUpdate, ExecuteCommand.ExecuteNonQuery, False)


            ' Cargar la estructura
            dtDetalles = cLineas.AddNewForm()
            dtDetalles.Rows(0).Delete()

            'Coger todos los valores menos los de cobros
            dttmp = AdminData.GetData("SELECT * FROM tbGestionCobrosLineas WHERE idGestionCobros = " & strIdcobros)

            Dim dfila As DataRow
            For shcont As Short = 0 To dttmp.Rows.Count - 1
                dfila = dtDetalles.NewRow
                dfila.ItemArray = dttmp.Rows(shcont).ItemArray
                dtDetalles.Rows.Add(dfila)
            Next

            dtDetalles.AcceptChanges()

            ' Control de filas
            If dtDetalles.Rows.Count <= 0 Then
                MsgBox("No se han obtenido detalles del registro.", MsgBoxStyle.Exclamation, "Sin detalles")
                Return -1
            End If

            ' Bien
            ' Si no es nula grabar lineas
            If Not IsNothing(dtDetalles) Then

                '' Obtener los pagos
                For shEmp As Short = 1 To 5
                    'IBIS. David. 18/10/2010. Cambiamos en el for de 5 a 8, y quitamos el 6 por ser la col. salario
                    If shEmp <> 6 Then
                        ObtenerPagos(dtDetalles, shEmp, shEjercicio)
                        ObtenerCobros(dtDetalles, shEmp, shEjercicio)
                    End If
                Next

                'Acumular los saldos desde enero a diciembre,descartado 20/05/2009 mostrar por meses
                '''''''''AcumularSaldos(dtDetalles)
                'Actualizar utilizando la clase
                cLineas.Update(dtDetalles)
            End If

            ' Liberar memoria
            dttmp.Dispose()
            dttmp = Nothing
            cLineas = Nothing
            dtDetalles.Dispose()
            dtDetalles = Nothing
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error al actualizar lineas.")
        End Try
        '' Bien
        Return 1
    End Function
    Public Function ObtieneTabla(ByVal strIdcobros As String, ByVal shEjercicio As String) As DataTable
        '' Con el id de la cabecera actualizar las lineas
        ' Crear los detalles
        Try
            Dim dtDetalles As New DataTable
            Dim dttmp As New DataTable
            Dim cLineas As New GestionCobrosLineas

            ' Resetear pagos a 0
            'IBIS. David. 18/10/2010. Se añaden 2 nuevas columnas, impppagos7 e imppagos 8.
            Dim strUpdate As String = "UPDATE tbGestionCobrosLineas SET imppagos1 = 0,imppagos2 = 0, imppagos3 = 0, imppagos4 = 0, imppagos5 = 0 WHERE idGestionCobros = " & strIdcobros
            'AdminData.ExecuteNonQuery(strUpdate, False)
            AdminData.Execute(strUpdate, ExecuteCommand.ExecuteNonQuery, False)


            ' Cargar la estructura
            dtDetalles = cLineas.AddNewForm()
            dtDetalles.Rows(0).Delete()

            'Coger todos los valores menos los de cobros
            dttmp = AdminData.GetData("SELECT * FROM tbGestionCobrosLineas WHERE idGestionCobros = " & strIdcobros)

            Dim dfila As DataRow
            For shcont As Short = 0 To dttmp.Rows.Count - 1
                dfila = dtDetalles.NewRow
                dfila.ItemArray = dttmp.Rows(shcont).ItemArray
                dtDetalles.Rows.Add(dfila)
            Next

            dtDetalles.AcceptChanges()

            ' Control de filas
            If dtDetalles.Rows.Count <= 0 Then
                MsgBox("No se han obtenido detalles del registro.", MsgBoxStyle.Exclamation, "Sin detalles")
            End If

            ' Bien
            ' Si no es nula grabar lineas
            If Not IsNothing(dtDetalles) Then

                '' Obtener los pagos
                For shEmp As Short = 1 To 5
                    'IBIS. David. 18/10/2010. Cambiamos en el for de 5 a 8, y quitamos el 6 por ser la col. salario
                    If shEmp <> 6 Then
                        ObtenerPagos(dtDetalles, shEmp, shEjercicio)
                        ObtenerCobros(dtDetalles, shEmp, shEjercicio)
                    End If
                Next

                'Acumular los saldos desde enero a diciembre,descartado 20/05/2009 mostrar por meses
                '''''''''AcumularSaldos(dtDetalles)
                'Actualizar utilizando la clase
                'dttmp.Dispose()
                'dttmp = Nothing
                'cLineas = Nothing
                'dtDetalles.Dispose()
                'dtDetalles = Nothing
                Return dtDetalles
            End If



        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error al actualizar lineas.")
        End Try
        '' Bien

    End Function

    ' Función que obtiene los diferentes cobros por cada entidad
    Public Sub ObtenerPagos(ByRef dtDetalles As DataTable, ByVal shNumEmpresa As Short, ByVal iEjercicio As Integer)
        Try
            Dim sEmpresa, sConsulta As String
            Select Case shNumEmpresa
                Case 1
                    sEmpresa = "xTecozam50R2"
                Case 2
                    'sEmpresa = "xDrenajes50R2"
                    sEmpresa = "xFerrallas50R2"
                Case 3
                    'sEmpresa = "xFerrallas50R2"
                    sEmpresa = "xSecozam50R2"
                Case 4
                    'sEmpresa = "xSecozam50R2"
                    sEmpresa = "xDrenajesPortugal50R2"
                Case 5
                    ' Solicitado con la creación de la nueva sociedad
                    'sEmpresa = "xTecozamUnitedKingdomR2"
                Case 6
                    ' Columna de salarios

                    'Case 7
                    '    'IBIS. David. 18/10/2010. Añadida nueva columna 
                    '    sEmpresa = "xTecozamPortugal50R2"
                    'Case 8
                    '    'IBIS. David. 18/10/2010. Añadida nueva columna 
                    '    sEmpresa = "xDrenajesPortugal50R2"

                    'Case Else
                    '    'sEmpresa = "xMaestros5"
            End Select
            sConsulta = "SELECT SUM(ImpVencimientoA) AS impSituacion, Situacion,MONTH(FechaVencimiento) As Mes " & _
                        "FROM " & sEmpresa & " .dbo.frmPagos " & _
                        "WHERE (FechaVencimiento >= '01-01-" & iEjercicio & "') " & _
                        " AND   (FechaVencimiento < '01-01-" & iEjercicio + 1 & "') " & _
                        " GROUP BY Situacion,MONTH(FechaVencimiento)"

            ' Obtener cobros por empresa
            Dim dtCobros As New DataTable
            dtCobros = AdminData.GetData(sConsulta)
            ' Si ha encontrado filas
            If dtCobros.Rows.Count > 0 Then
                Dim dfilas() As DataRow

                '' Recorrer todos los cobros encontrados y asignar a la fila y columna correcta
                For shcont As Short = 0 To dtCobros.Rows.Count - 1
                    dfilas = dtDetalles.Select("Tipo = 'Pago' and mes = " & dtCobros.Rows(shcont)("mes") & " and situacion = " & dtCobros.Rows(shcont)("situacion"))

                    'Control de fila encontrada en los detalles del control de cobros
                    If Not IsNothing(dfilas) And dfilas.Length > 0 Then
                        dfilas(0)("imppagos" & shNumEmpresa) = dtCobros.Rows(shcont)("impSituacion")
                        dfilas(0)("fechaModificacionAudi") = Now
                    End If

                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error al obtener los pagos.")
        End Try
    End Sub


    ' 15/05/2009 acumular saldos por cada tipo en diciembre por cada estado el importe de todo el año
    ' Descartada 20/05/2009. Solo mostrar lo del mes
    Private Sub AcumularSaldos(ByRef dtDetalles As DataTable)
        ' Coger todos los estados
        Dim shEstados As New ArrayList
        ' Recorrer las filas para los estados
        For shcont As Short = 0 To dtDetalles.Rows.Count - 1
            If shEstados.Contains(dtDetalles.Rows(shcont)("situacion")) = False Then
                shEstados.Add(dtDetalles.Rows(shcont)("situacion"))
            End If
        Next
        ' Tengo los diferentes puestos , obtener las filas ordenadas por mes para ir añadiendo valores de pagos
        For shcont As Short = 0 To shEstados.Count - 1
            Dim dfEstado As DataRow() = dtDetalles.Select("situacion=" & CStr(shEstados(shcont)), "mes")
            For ShFilas As Short = 1 To dfEstado.Length - 1
                dfEstado(ShFilas)("imppagos1") += dfEstado(ShFilas - 1)("imppagos1")
                dfEstado(ShFilas)("imppagos2") += dfEstado(ShFilas - 1)("imppagos2")
                dfEstado(ShFilas)("imppagos3") += dfEstado(ShFilas - 1)("imppagos3")
                dfEstado(ShFilas)("imppagos4") += dfEstado(ShFilas - 1)("imppagos4")
                dfEstado(ShFilas)("imppagos5") += dfEstado(ShFilas - 1)("imppagos5")
                'IBIS. David. 18/10/2010. Añadidas las columnas 7 y 8
                'dfEstado(ShFilas)("imppagos7") += dfEstado(ShFilas - 1)("imppagos7")
                'dfEstado(ShFilas)("imppagos8") += dfEstado(ShFilas - 1)("imppagos8")
            Next
        Next
    End Sub

#Region "Ibis Desarrollo"
    ' IBIS PATRICIA: Opción para crear cobros de cada entidad 10-11-2011
    Public Sub ObtenerCobros(ByRef dtDetalles As DataTable, ByVal shNumEmpresa As Short, ByVal iEjercicio As Integer)
        Try
            Dim sEmpresa, sConsulta As String
            Select Case shNumEmpresa
                Case 1
                    sEmpresa = "xTecozam50R2"
                Case 2
                    'sEmpresa = "xDrenajes50R2"
                    sEmpresa = "xFerrallas50R2"
                Case 3
                    'sEmpresa = "xFerrallas50R2"
                    sEmpresa = "xSecozam50R2"
                Case 4
                    'sEmpresa = "xSecozam50R2"
                    sEmpresa = "xDrenajesPortugal50R2"
                Case 5
                    ' Solicitado con la creación de la nueva sociedad
                    'sEmpresa = "xTecozamUnitedKingdomR2"
                Case 6
                    ' Columna de salarios

                    'Case 7
                    '    'IBIS. David. 18/10/2010. Añadida nueva columna 
                    '    sEmpresa = "xTecozamPortugal50R2"
                    'Case 8
                    '    'IBIS. David. 18/10/2010. Añadida nueva columna 
                    '    sEmpresa = "xDrenajesPortugal50R2"

                    'Case Else
                    '    'sEmpresa = "xMaestros5"
            End Select
            sConsulta = "SELECT SUM(ImpVencimientoA) AS impSituacion, Situacion,MONTH(FechaVencimiento) As Mes " & _
                        "FROM " & sEmpresa & " .dbo.frmCobros " & _
                        "WHERE (FechaVencimiento >= '01-01-" & iEjercicio & "') " & _
                        " AND   (FechaVencimiento < '01-01-" & iEjercicio + 1 & "') " & _
                        " GROUP BY Situacion,MONTH(FechaVencimiento)"

            ' Obtener cobros por empresa
            Dim dtCobros As New DataTable
            dtCobros = AdminData.GetData(sConsulta)

            ' Si ha encontrado filas
            If dtCobros.Rows.Count > 0 Then
                Dim dfilas() As DataRow

                '' Recorrer todos los cobros encontrados y asignar a la fila y columna correcta
                For shcont As Short = 0 To dtCobros.Rows.Count - 1
                    dfilas = dtDetalles.Select("Tipo = 'Cobro' and mes = " & dtCobros.Rows(shcont)("mes") & " and situacion = " & dtCobros.Rows(shcont)("situacion"))

                    'Control de fila encontrada en los detalles del control de cobros
                    If Not IsNothing(dfilas) And dfilas.Length > 0 Then
                        dfilas(0)("imppagos" & shNumEmpresa) = dtCobros.Rows(shcont)("impSituacion")
                        dfilas(0)("fechaModificacionAudi") = Now
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error al obtener los pagos.")
        End Try
    End Sub
#End Region
End Class

