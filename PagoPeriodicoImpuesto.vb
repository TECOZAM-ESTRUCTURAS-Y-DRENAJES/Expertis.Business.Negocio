Public Class PagoPeriodicoImpuesto
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

#Region "Constructor"

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbPagoPeriodicoImpuesto"

#End Region

#Region "Eventos PagoPeriodicoImpuesto"


    Private Function CalculoPorcentajesElementos(ByVal StrIDInmovilizado As String, _
                                                 ByVal StrIDElemento As String, _
                                                 ByVal IntAño As Integer, _
                                                 ByVal IntMes As Integer, _
                                                 ByVal DblFechaFin As Date) As DataTable
        '****************************************************************************************************
        '   - strIDInmovilizado -> Código del inmovilizado del Leasing
        '   - strIDElemento -> No se usa
        '   - intAño -> Año que estoy calculando
        '   - intMes -> Mes que estoy calculando
        '   - dblFechaFin -> Fecha Fin hasta donde debe llegar el cálculo.
        '   - Return datatable -> Devuelve un datatable con los porcentajes correspondientes a ese mes para todos los elementos
        '****************************************************************************************************
        Dim ClsElemAmort As New ElementoAmortizable
        Dim Dt As New DataTable
        Dim DtElemAmort As New DataTable
        Dim DblValorTotal As Double
        Dim DteFecha As Date
        Dim IntDia As Integer = Date.DaysInMonth(IntAño, IntMes)

        Dt.Columns.Add("IDElemento", GetType(String))
        Dt.Columns.Add("Porcentaje", GetType(String))
        Dt.Columns.Add("Mes", GetType(String))
        Dt.Columns.Add("Año", GetType(String))

        DteFecha = New Date(IntAño, IntMes, IntDia)
        'DteFecha = CDate(IntDia & "/" & IntMes & "/" & IntAño)

        '****************************************************************************************************
        'Selecciono todos los elementos amortizables del inmovilizado que están activos para ese mes y año.
        'Estos elementos serán aquellos que no están de baja o que en el caso de estarlo, su fecha de baja es
        'es posterior a la fecha que estamos analizando.
        '****************************************************************************************************
        Dim FilAmort As New Filter
        FilAmort.Add("IDInmovilizado", FilterOperator.Equal, StrIDInmovilizado, FilterType.String)
        FilAmort.Add(New BooleanFilterItem("Baja", True))
        FilAmort.Add("FechaBaja", FilterOperator.GreaterThan, DteFecha, FilterType.DateTime)
        Dim FilAmort2 As New Filter(FilterUnionOperator.Or)
        FilAmort2.Add(FilAmort)
        FilAmort2.Add(New BooleanFilterItem("Baja", False))
        DtElemAmort = ClsElemAmort.Filter(FilAmort2)
        'DtElemAmort = ClsElemAmort.Filter(, "IDInmovilizado = '" & StrIDInmovilizado & "' AND ((Baja = 1 and FechaBaja > '" & DteFecha & "') OR Baja = 0)")
        If Not DtElemAmort Is Nothing AndAlso DtElemAmort.Rows.Count > 0 Then
            For Each Dr As DataRow In DtElemAmort.Select
                DblValorTotal += Dr("ValorTotalElementoA")
            Next
            If Not DtElemAmort Is Nothing AndAlso DtElemAmort.Rows.Count > 0 Then
                For Each Dr As DataRow In DtElemAmort.Select
                    If DblValorTotal > 0 Then
                        Dim DrNew As DataRow = Dt.NewRow()
                        DrNew("IDElemento") = dr("IDElemento")
                        DrNew("Año") = IntAño
                        DrNew("Mes") = IntMes
                        DrNew("Porcentaje") = dr("ValorTotalElementoA") * 100 / DblValorTotal
                        Dt.Rows.Add(DrNew)
                    End If
                Next
            End If
        End If
        Return Dt
    End Function

  
  
    Private Function ObtenerAmortMensAnticipado(ByVal intAño As Integer, ByVal intMesCalculo As Integer, ByVal dtAmort As DataTable, _
                                            ByVal dteFechaUltimaContabilizacion As Date, Optional ByVal dtAmortHijo As DataTable = Nothing, _
                                             Optional ByVal dteFechaConversion As Date = cnMinDate, Optional ByVal dteFechaConversionActual As Date = cnMinDate) As DataTable
        'Obtiene de mrcsAmort la amortizacion por meses para un año determinado.

        Dim dtAmortAño As DataTable
        Dim dtMeses As DataTable
        Dim strMesesContable As String
        Dim strMesesRealizada As String
        Dim intMes As Integer
        Dim intMesIni As Integer
        Dim dteFechaAnalisis As Date
        Dim strMesesContableHijo As String
        Dim strMesesRealizadaHijo As String
        Dim blnTraspasado As Boolean


        'Creacion del dt que va a contener los datos
        dtMeses = New DataTable

        dtMeses.Columns.Add("Mes", GetType(Integer))
        dtMeses.Columns.Add("Amortizacion", GetType(Double))

        blnTraspasado = False
        If intAño <> 0 Then
            'Carga de los datos en el dt
            If Not dtAmort Is Nothing Then
                If dtAmort.Rows.Count <> 0 Then
                    Dim dvAmort As New DataView(dtAmort)
                    dvAmort.RowFilter = "Año=" & intAño
                    If dvAmort.Count > 0 Then
                        strMesesContable = dvAmort(0).Row("AmortContableMensual")
                        strMesesRealizada = dvAmort(0).Row("AmortRealizadaMensual")
                    End If
                End If
            End If

            If Nz(dteFechaConversion, Date.MinValue) <> Date.MinValue Then
                If intAño >= Year(dteFechaConversion) Then
                    blnTraspasado = True
                    If Not dtAmortHijo Is Nothing Then
                        If dtAmortHijo.Rows.Count <> 0 Then
                            Dim dvAmortHijo As New DataView(dtAmortHijo)
                            dvAmortHijo.RowFilter = "Año=" & intAño
                            If dvAmortHijo.Count > 0 Then
                                strMesesContableHijo = dvAmortHijo(0).Row("AmortContableMensual")
                                strMesesRealizadaHijo = dvAmortHijo(0).Row("AmortRealizadaMensual")
                            End If
                        End If
                    End If
                End If
            End If

            'Descomponemos la cadena y la metemos en el grid
            If strMesesContable <> String.Empty Or strMesesRealizada <> String.Empty Then
                For intMes = 1 To intMesCalculo
                    dteFechaAnalisis = DateSerial(intAño, intMes, 1)
                    Dim drMeses As DataRow = dtMeses.NewRow
                    'Esta última condición que se ha puesto es para que en aquellos elementos con padres,
                    'no se tenga en cuenta el primer mes ya que en ese mes iría todo lo amortizado
                    'por el elemento padre
                    If dteFechaConversionActual > dteFechaAnalisis Then
                        drMeses("Mes") = intMes
                        drMeses("Amortizacion") = 0
                    Else
                        If Nz(dteFechaConversion, Date.MinValue) > dteFechaAnalisis Or Nz(dteFechaConversion, Date.MinValue) = Date.MinValue Then
                            'TODO GetPropertyValue
                            'If GetPropertyValue(strMesesRealizada, "Mes" & intMes) <> String.Empty And GetPropertyValue(strMesesRealizada, "Amortizacion" & intMes) <> 0 Then
                            '    drMeses("Mes") = GetPropertyValue(strMesesRealizada, "Mes" & intMes)
                            '    drMeses("Amortizacion") = GetPropertyValue(strMesesRealizada, "Amortizacion" & intMes)
                            'ElseIf GetPropertyValue(strMesesContable, "Mes" & intMes) <> String.Empty Or GetPropertyValue(strMesesContable, "Amortizacion" & intMes) <> 0 Then
                            '    If (dteFechaAnalisis > dteFechaUltimaContabilizacion) Then
                            '        drMeses("Mes") = GetPropertyValue(strMesesContable, "Mes" & intMes)
                            '        drMeses("Amortizacion") = GetPropertyValue(strMesesContable, "Amortizacion" & intMes)
                            '    End If
                            'Else
                            '    drMeses("Mes") = intMes
                            '    drMeses("Amortizacion") = 0
                            'End If
                        Else
                            'TODO GetPropertyValue
                            'If GetPropertyValue(strMesesRealizadaHijo, "Mes" & intMes) <> String.Empty And GetPropertyValue(strMesesRealizadaHijo, "Amortizacion" & intMes) <> 0 Then
                            '    drMeses("Mes") = GetPropertyValue(strMesesRealizadaHijo, "Mes" & intMes)
                            '    drMeses("Amortizacion") = GetPropertyValue(strMesesRealizadaHijo, "Amortizacion" & intMes)
                            'ElseIf GetPropertyValue(strMesesContableHijo, "Mes" & intMes) <> String.Empty Or GetPropertyValue(strMesesContableHijo, "Amortizacion" & intMes) <> 0 Then
                            '    If (dteFechaAnalisis > dteFechaUltimaContabilizacion) Then
                            '        drMeses("Mes") = GetPropertyValue(strMesesContableHijo, "Mes" & intMes)
                            '        drMeses("Amortizacion") = GetPropertyValue(strMesesContableHijo, "Amortizacion" & intMes)
                            '    End If
                            'Else
                            '    drMeses("Mes") = intMes
                            '    drMeses("Amortizacion") = 0
                            'End If
                        End If
                    End If
                    dtMeses.Rows.Add(drMeses)
                Next
            End If
        End If

        ObtenerAmortMensAnticipado = dtMeses

    End Function
#End Region

End Class