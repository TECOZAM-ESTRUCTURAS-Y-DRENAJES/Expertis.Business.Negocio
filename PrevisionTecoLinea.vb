Option Strict Off
Option Explicit On
Option Compare Text

Public Class PrevisionTecoLinea
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbPrevisionTecoLinea"

    'Public Overloads Sub Delete(ByVal strIDPrevision As String)
    '    Dim DtDelete As DataTable = MyBase.SelOnPrimaryKey(strIDPrevision)
    '    If Not MyBase.Delete(DtDelete) Then
    '        ApplicationService.GenerateError(DELETECONSTRAINTMESSAGE)
    '    Else

    '    End If
    'End Sub

    Public Overloads Function Update(ByVal dttSource As System.Data.DataTable) As System.Data.DataTable
        If Not dttSource Is Nothing AndAlso dttSource.Rows.Count > 0 Then
            Me.BeginTx()
            For Each dr As DataRow In dttSource.Rows
                'If Lenght(dr("DescPrevision")) = 0 Then ApplicationService.GenerateError("La Descripción de la Previsión es obligatoria")

                'Dim dtTarifa As DataTable

                If dr.RowState = DataRowState.Added Then

                    If Not IsDBNull(dr("IDPrevision")) Then
                        dr("IDPrevisionLinea") = AdminData.GetAutoNumeric
                    End If

                    ''Comprobación de la existencia de la Prevision
                    'dtTarifa = SelOnPrimaryKey(dr("IDPrevision"))
                    'If dtTarifa.Rows.Count <> 0 Then GenerateMessage("La Previsión ya existe", Me.GetType.Name & ".Update")

                End If
            Next
            UpdateTable(dttSource)
        End If
        Return dttSource
    End Function

    Public Overrides Function AddNewForm() As DataTable
        Dim dt As DataTable = MyBase.AddNewForm

        dt.Rows(0)("IDPrevisionLinea") = AdminData.GetAutoNumeric
        Return dt

    End Function

    'Función para Crear/Actualiza/Borra las lineas por cada cabecera ajustandose a los bancos en cada sociedad
    Public Function CrearActLineas(ByVal iIdCabecera As Integer) As Short
        Try
            ' Cargar las lineas q pueda tener previamente
            Dim dtOriginal As DataTable
            Dim dtCobros As DataTable = Nothing
            Dim dtPagos As DataTable = Nothing
            Dim dtBpropios, dtBPropiosTotal As DataTable
            Dim sselect As String
            Dim row, rowcopy, rows(), rowsAux() As DataRow
            dtOriginal = AdminData.GetData("SELECT * FROM tbprevisiontecolinea WHERE  idprevision = " & iIdCabecera)
            '' Coger los bp de todas las entidades
            'Armozam, ojo con espacios para q entren todas las entidades
            sselect = "SELECT 'FERRALLAS' AS EMPRESA,idBancoPropio, DescBancoPropio,riesgoPagare FROM xFerrallas50R2_COPIA.dbo.tbMaestroBancoPropio " & _
                      " WHERE (snbaja <> 1 OR snbaja IS NULL ) " & _
                      "  AND (DescBancoPropio LIKE '1%')"
            dtBpropios = AdminData.GetData(sselect)
            dtBPropiosTotal = dtBpropios.Clone

            'Obtener cobros
            If obtenerCobros(dtCobros, "xFerrallas50R2_COPIA", "FERRALLAS") < 0 Then
                Return Nothing
            End If
            'Obtener pagos,ojo dejo espacio para los nombres de empresa
            If obtenerPagos(dtPagos, "xFerrallas50R2_COPIA", "FERRALLAS") < 0 Then
                Return Nothing
            End If
            For Each row In dtBpropios.Rows
                rowcopy = dtBPropiosTotal.NewRow
                rowcopy.ItemArray = row.ItemArray
                dtBPropiosTotal.Rows.Add(rowcopy)
            Next
            'Dyezam
            sselect = sselect.Replace("xFerrallas50R2_COPIA", "xDrenajes50R2")
            sselect = sselect.Replace("FERRALLAS", "DYEZAM")
            dtBpropios = AdminData.GetData(sselect)
            'Obtener cobros
            If obtenerCobros(dtCobros, "xDrenajes50R2", "DYEZAM") < 0 Then
                Return Nothing
            End If
            'Obtener pagos
            If obtenerPagos(dtPagos, "xDrenajes50R2", "DYEZAM") < 0 Then
                Return Nothing
            End If
            ' Añadir las filas al final
            For Each row In dtBpropios.Rows
                rowcopy = dtBPropiosTotal.NewRow
                rowcopy.ItemArray = row.ItemArray
                dtBPropiosTotal.Rows.Add(rowcopy)
            Next
            'Armozam
            sselect = sselect.Replace("xDrenajes4", "xArmadores4")
            sselect = sselect.Replace("DYEZAM", "ARMOZAM")
            dtBpropios = AdminData.GetData(sselect)
            'Obtener cobros
            If obtenerCobros(dtCobros, "xArmadores4", "ARMOZAM") < 0 Then
                Return Nothing
            End If
            'Obtener pagos
            If obtenerPagos(dtPagos, "xArmadores4", "ARMOZAM") < 0 Then
                Return Nothing
            End If
            ' Añadir las filas al final
            For Each row In dtBpropios.Rows
                rowcopy = dtBPropiosTotal.NewRow
                rowcopy.ItemArray = row.ItemArray
                dtBPropiosTotal.Rows.Add(rowcopy)
            Next
            'Tecozam
            sselect = sselect.Replace("xArmadores4", "xTecozam50R2Demo2")
            sselect = sselect.Replace("ARMOZAM", "TECOZAM")
            dtBpropios = AdminData.GetData(sselect)
            'Obtener cobros
            If obtenerCobros(dtCobros, "xTecozam50R2Demo2", "TECOZAM") < 0 Then
                Return Nothing
            End If
            'Obtener pagos
            If obtenerPagos(dtPagos, "xTecozam50R2Demo2", "TECOZAM") < 0 Then
                Return Nothing
            End If
            ' Añadir las filas al final
            For Each row In dtBpropios.Rows
                rowcopy = dtBPropiosTotal.NewRow
                rowcopy.ItemArray = row.ItemArray
                dtBPropiosTotal.Rows.Add(rowcopy)
            Next
            ''''''''''''Secozam
            '''''''''''sselect = sselect.Replace("xTecozam4", "xSecozam4")
            '''''''''''sselect = sselect.Replace("TECOZAM", "SECOZAM")
            '''''''''''dtBpropios = AdminData.GetData(sselect)
            ''''''''''''Obtener cobros
            '''''''''''If obtenerCobros(dtCobros, "xSecozam4", "SECOZAM") < 0 Then
            '''''''''''    Return Nothing
            '''''''''''End If
            ''''''''''''Obtener pagos
            '''''''''''If obtenerPagos(dtPagos, "xSecozam4", "SECOZAM") < 0 Then
            '''''''''''    Return Nothing
            '''''''''''End If
            '''''''''''' Añadir las filas al final
            '''''''''''For Each row In dtBpropios.Rows
            '''''''''''    rowcopy = dtBPropiosTotal.NewRow
            '''''''''''    rowcopy.ItemArray = row.ItemArray
            '''''''''''    dtBPropiosTotal.Rows.Add(rowcopy)
            '''''''''''Next
            ' Recorro todas las filas y las q YA no estén ya las elimino
            If dtOriginal.Rows.Count > 0 Then
                For Each row In dtOriginal.Rows
                    rows = dtBPropiosTotal.Select("empresa = '" & row("descripEmpresa") & "' AND idBancoPropio = '" & row("idBancoPropio") & "'")
                    If rows.Length <= 0 Then
                        row.Delete()
                    End If
                Next
            End If
            ' Ya tengo todos los B.P. de todas las entidades crear/actualizar las lineas
            For Each dr As DataRow In dtBPropiosTotal.Rows
                Dim dt As New DataTable
                Dim f As New Filter
                f.Add("DescripEmpresa", FilterOperator.Equal, dr("EMPRESA"))
                f.Add("IDBancoPropio", FilterOperator.Equal, dr("IDBancoPropio"))
                dt = Me.Filter(f)
                If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                    ' Fila encontrada, actualizar
                    dt.Rows(0)("fechaprev") = Now
                    ' Columnas de importes
                    dt.Rows(0)("impriesgo") = dr("riesgoPagare")
                    ' Coger los datos de cobros
                    If Not IsNothing(dtCobros) Then
                        rowsAux = dtCobros.Select("EMPRESA='" & dr("EMPRESA") & "' AND IDBANCOPROPIO = '" & dr("idBancoPropio") & "'")
                        If rowsAux.Length > 0 Then
                            dt.Rows(0)("IMPDISPUESTO") = rowsAux(0)("impVencimiento")
                        End If
                    Else
                        dt.Rows(0)("IMPDISPUESTO") = 0
                    End If
                    ' Coger los datos de pagos
                    If Not IsNothing(dtPagos) Then
                        rowsAux = dtPagos.Select("EMPRESA='" & dr("EMPRESA") & "' AND IDBANCOPROPIO = '" & dr("idBancoPropio") & "'")
                        If rowsAux.Length > 0 Then
                            dt.Rows(0)("IMPVENCIMIENTO") = rowsAux(0)("impVencimiento")
                        End If
                    Else
                        dt.Rows(0)("IMPVENCIMIENTO") = 0
                    End If
                    CerosFila(dt.Rows(0))
                    Me.Update(dt)


                Else
                    ' Nueva fila
                    Dim dtNew As New DataTable
                    dtNew = Me.AddNewForm
                    dtNew.Rows(0)("IDPrevision") = iIdCabecera
                    dtNew.Rows(0)("fechaprev") = Now
                    dtNew.Rows(0)("descripempresa") = dr("EMPRESA")
                    dtNew.Rows(0)("IDBancoPropio") = dr("idBancoPropio")
                    dtNew.Rows(0)("DescripBancoPropio") = dr("DescBancoPropio")
                    dtNew.Rows(0)("usuariocrea") = "INFOR"
                    ' Columnas de importes
                    dtNew.Rows(0)("impriesgo") = dr("riesgoPagare")
                    'USUARIO expertisapp.username
                    ' Coger los datos de cobros
                    If Not IsNothing(dtCobros) Then
                        rowsAux = dtCobros.Select("EMPRESA='" & dr("EMPRESA") & "' AND IDBANCOPROPIO = '" & dr("idBancoPropio").ToString & "'")
                        If rowsAux.Length > 0 Then
                            dtNew.Rows(0)("IMPDISPUESTO") = rowsAux(0)("impVencimiento")
                        End If
                    Else
                        dtNew.Rows(0)("IMPDISPUESTO") = 0
                    End If
                    ' Coger los datos de pagos, falla idbanco alfanumérico
                    If Not IsNothing(dtPagos) Then
                        rowsAux = dtPagos.Select("EMPRESA='" & dr("EMPRESA") & "' AND IDBANCOPROPIO = '" & dr("idBancoPropio").ToString & "'")
                        If rowsAux.Length > 0 Then
                            dtNew.Rows(0)("IMPVENCIMIENTO") = rowsAux(0)("impVencimiento")
                        End If
                    Else
                        dtNew.Rows(0)("IMPDISPUESTO") = 0
                    End If
                    ' Añadir al repositorio final
                    CerosFila(dtNew.Rows(0))
                    Me.Update(dtNew)
                End If
            Next

            'For shcont As Short = 0 To dtBPropiosTotal.Rows.Count - 1

            '    rows = dtOriginal.Select("descripEmpresa = '" & dtBPropiosTotal.Rows(shcont)("EMPRESA") & "' AND idBancoPropio = '" & dtBPropiosTotal.Rows(shcont)("idBancoPropio") & "'")
            '    'Control de no encontrada
            '    If rows.Length <= 0 Then
            '        'row = dtOriginal.NewRow
            '        'row("IDPrevisionLinea") = AdminData.GetAutoNumeric
            '        'row("IDPrevision") = iIdCabecera
            '        'row("fechaprev") = Now
            '        'row("descripempresa") = dtBPropiosTotal.Rows(shcont)("EMPRESA")
            '        'row("IDBancoPropio") = dtBPropiosTotal.Rows(shcont)("idBancoPropio")
            '        'row("DescripBancoPropio") = dtBPropiosTotal.Rows(shcont)("DescBancoPropio")
            '        'row("usuariocrea") = "INFORMÁTICA"
            '        '' Columnas de importes
            '        'row("impriesgo") = dtBPropiosTotal.Rows(shcont)("riesgoPagare")
            '        '' Coger los datos de cobros
            '        'If Not IsNothing(dtCobros) Then
            '        '    rowsAux = dtCobros.Select("EMPRESA='" & dtBPropiosTotal.Rows(shcont)("EMPRESA") & "' AND IDBANCOPROPIO = '" & dtBPropiosTotal.Rows(shcont)("idBancoPropio").ToString & "'")
            '        '    If rowsAux.Length > 0 Then
            '        '        row("IMPDISPUESTO") = rowsAux(0)("impVencimiento")
            '        '    End If
            '        'Else
            '        '    row("IMPDISPUESTO") = 0
            '        'End If
            '        '' Coger los datos de pagos, falla idbanco alfanumérico
            '        'If Not IsNothing(dtPagos) Then
            '        '    rowsAux = dtPagos.Select("EMPRESA='" & dtBPropiosTotal.Rows(shcont)("EMPRESA") & "' AND IDBANCOPROPIO = '" & dtBPropiosTotal.Rows(shcont)("idBancoPropio").ToString & "'")
            '        '    If rowsAux.Length > 0 Then
            '        '        row("IMPVENCIMIENTO") = rowsAux(0)("impVencimiento")
            '        '    End If
            '        'Else
            '        '    row("IMPDISPUESTO") = 0
            '        'End If
            '        '' Añadir al repositorio final
            '        'CerosFila(row)
            '        'dtOriginal.Rows.Add(row)
            '    Else
            '        ' Actualizar fila
            '        rows(0)("fechaprev") = Now
            '        ' Columnas de importes
            '        rows(0)("impriesgo") = dtBPropiosTotal.Rows(shcont)("riesgoPagare")
            '        ' Coger los datos de cobros
            '        If Not IsNothing(dtCobros) Then
            '            rowsAux = dtCobros.Select("EMPRESA='" & dtBPropiosTotal.Rows(shcont)("EMPRESA") & "' AND IDBANCOPROPIO = '" & dtBPropiosTotal.Rows(shcont)("idBancoPropio") & "'")
            '            If rowsAux.Length > 0 Then
            '                rows(0)("IMPDISPUESTO") = rowsAux(0)("impVencimiento")
            '            End If
            '        Else
            '            rows(0)("IMPDISPUESTO") = 0
            '        End If
            '        ' Coger los datos de pagos
            '        If Not IsNothing(dtPagos) Then
            '            rowsAux = dtPagos.Select("EMPRESA='" & dtBPropiosTotal.Rows(shcont)("EMPRESA") & "' AND IDBANCOPROPIO = '" & dtBPropiosTotal.Rows(shcont)("idBancoPropio") & "'")
            '            If rowsAux.Length > 0 Then
            '                rows(0)("IMPVENCIMIENTO") = rowsAux(0)("impVencimiento")
            '            End If
            '        Else
            '            rows(0)("IMPVENCIMIENTO") = 0
            '        End If
            '        CerosFila(row)
            '    End If
            'Next

            Return 1
        Catch ex As Exception
            MsgBox("Error " & ex.Message, MsgBoxStyle.Exclamation, "Error al asignar detalles de previsión.")
            Return -1
        End Try

    End Function
    ' Calculo cóbros de todas las entidades
    Private Function obtenerCobros(ByRef dtCobros As DataTable, ByVal sbd As String, ByVal sempresa As String) As Short
        'Coger los cobros
        Try
            Dim dtCobrosEmpresa As New DataTable
            Dim drow As DataRow
            Dim sConsulta As String = "SELECT max(descestado) AS Empresa,SUM(ImpVencimiento) AS impVencimiento,idBancoPropio " & _
                "FROM " & sbd & " .dbo.tbCobro," & sbd & ".dbo.tbMaestroEstadoCobro " & _
                "WHERE (" & sbd & " .dbo.tbCobro.FechaVencimiento <= ' " & CDate(Now) & " ') " & _
                "AND " & sbd & " .dbo.tbCobro.Situacion = " & sbd & ".dbo.tbMaestroEstadoCobro.IDEstado " & _
                "AND (" & sbd & ".dbo.tbMaestroEstadoCobro.DescEstado LIKE 'DESCONTAD%') GROUP BY IDBancoPropio"
            dtCobrosEmpresa = AdminData.GetData(sConsulta)
            If dtCobrosEmpresa.Rows.Count > 0 Then
                ' Control de otros pagos
                If IsNothing(dtCobros) Then
                    dtCobros = dtCobrosEmpresa.Clone
                End If
                For shcont As Short = 0 To dtCobrosEmpresa.Rows.Count - 1
                    drow = dtCobros.NewRow
                    drow.ItemArray = dtCobrosEmpresa.Rows(shcont).ItemArray
                    drow("Empresa") = sempresa
                    dtCobros.Rows.Add(drow)
                Next
            End If
        Catch ex As Exception
            MsgBox("Se produjo un error al obtener cobros:" & ex.Message, MsgBoxStyle.Critical, "Error en la obtención de cobros " & sempresa)
        End Try

        ' Bien 
        Return 1
    End Function
    ' Calculo págos
    Private Function obtenerPagos(ByRef dtPagos As DataTable, ByVal sbd As String, ByVal sempresa As String) As Short
        'Coger los pagos hasta la fecha actual
        Try
            Dim dtPagosEmpresa As New DataTable
            Dim drow As DataRow
            ' LA CONSULTA CUMPLE CON LO Q ME HAN DICHO PERO NO HYA PAGOS CON EL ESTADO ENVIADO
            Dim sConsulta As String = "SELECT max(descestado) AS Empresa,SUM(ImpVencimiento) AS impVencimiento,idBancoPropio " & _
                "FROM " & sbd & " .dbo.tbPago," & sbd & ".dbo.tbMaestroEstadoPago " & _
                "WHERE (" & sbd & " .dbo.tbPago.FechaVencimiento <= ' " & CDate(Now) & " ') " & _
                "AND " & sbd & " .dbo.tbPago.situacion = " & sbd & ".dbo.tbMaestroEstadoPago.IDEstado " & _
                "AND (" & sbd & ".dbo.tbMaestroEstadoPago.DescEstado LIKE 'ENVIAD%') GROUP BY IDBancoPropio"
            'Pruebas con otro estado con datos
            'MsgBox("Datos de cobro obtenidos del estado:PAGARE")
            ' Obtener resultados
            dtPagosEmpresa = AdminData.GetData(sConsulta)
            If dtPagosEmpresa.Rows.Count > 0 Then
                ' Control de otros pagos
                If IsNothing(dtPagos) Then
                    dtPagos = dtPagosEmpresa.Clone
                End If
                For shcont As Short = 0 To dtPagosEmpresa.Rows.Count - 1
                    drow = dtPagos.NewRow
                    drow.ItemArray = dtPagosEmpresa.Rows(shcont).ItemArray
                    drow("Empresa") = sempresa
                    dtPagos.Rows.Add(drow)
                Next
            End If
        Catch ex As Exception
            MsgBox("Se produjo un error al obtener pagos:" & ex.Message, MsgBoxStyle.Critical, "Error en la obtención de pagos " & sempresa)
            Return -1
        End Try

        ' Bien 
        Return 1
    End Function
    ' Resetear valores
    Private Sub CerosFila(ByRef drow As DataRow)
        For shcont As Short = 0 To drow.ItemArray.Length - 1
            If Strings.Left(drow.Table.Columns(shcont).ColumnName, 3) = "IMP" Then
                If IsDBNull(drow(shcont)) Then drow(shcont) = 0
            End If
        Next

    End Sub
End Class
