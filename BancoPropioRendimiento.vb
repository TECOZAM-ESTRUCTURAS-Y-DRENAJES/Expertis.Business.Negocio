Public Class BancoPropioRendimiento

#Region "Constructor"

    Inherits BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBancoPropioRendimiento"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarIntervalo)
    End Sub

    <Task()> Public Shared Sub ValidarIntervalo(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim F As Filter
        If data.RowState = DataRowState.Modified AndAlso data("FechaDesde", DataRowVersion.Original) = data("FechaDesde") AndAlso _
           data("FechaHasta", DataRowVersion.Original) = data("FechaHasta") Then
            Exit Sub
        Else
            Dim F1 As New Filter
            F1.Add("FechaHasta", FilterOperator.GreaterThanOrEqual, data("FechaDesde"), FilterType.DateTime)
            F1.Add("FechaHasta", FilterOperator.LessThanOrEqual, data("FechaHasta"), FilterType.DateTime)
            F1.Add("IdBancoPropio", FilterOperator.Equal, data("IdBancoPropio"), FilterType.String)
            F1.Add("FechaDesde", FilterOperator.NotEqual, data("FechaDesde"), FilterType.DateTime)

            Dim F2 As New Filter
            F2.Add("FechaDesde", FilterOperator.GreaterThanOrEqual, data("FechaDesde"), FilterType.DateTime)
            F2.Add("FechaDesde", FilterOperator.LessThanOrEqual, data("FechaHasta"), FilterType.DateTime)
            F2.Add("IdBancoPropio", FilterOperator.Equal, data("IdBancoPropio"), FilterType.String)
            F2.Add("FechaDesde", FilterOperator.NotEqual, data("FechaDesde"), FilterType.DateTime)

            Dim F3 As New Filter
            F3.Add("FechaDesde", FilterOperator.LessThanOrEqual, data("FechaDesde"), FilterType.DateTime)
            F3.Add("FechaHasta", FilterOperator.GreaterThanOrEqual, data("FechaHasta"), FilterType.DateTime)
            F3.Add("IdBancoPropio", FilterOperator.Equal, data("IdBancoPropio"), FilterType.String)
            F3.Add("FechaDesde", FilterOperator.NotEqual, data("FechaDesde"), FilterType.DateTime)

            F = New Filter
            F.UnionOperator = FilterUnionOperator.Or
            F.Add(F1)
            F.Add(F2)
            F.Add(F3)
            Dim dtt As DataTable = New BancoPropioRendimiento().Filter(F)
            If dtt.Rows.Count > 0 Then ApplicationService.GenerateError("El intervalo de fecha especificado contiene fechas de otros intervalos ya especificados para el Banco.")
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Serializable()> _
    Public Class DatosCalcRend
        Public FechaDesde As Date
        Public FechaHasta As Date
        Public IDBancoPropio As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal FechaDesde As Date, ByVal FechaHasta As Date, ByVal IDBancoPropio As String)
            Me.FechaDesde = FechaDesde
            Me.FechaHasta = FechaHasta
            Me.IDBancoPropio = IDBancoPropio
        End Sub
    End Class

    <Serializable()> _
    Public Class DatosCalcIntereses
        Public DtNewRow As DataTable
        Public DtNewRowAnt As DataTable
        Public FechaHasta As Date
        Public Dias As Integer

        Public Sub New(ByVal DtNewRow As DataTable, ByVal DtNewRowAnt As DataTable, ByVal FechaHasta As Date, ByVal Dias As Integer)
            Me.DtNewRow = DtNewRow
            Me.DtNewRowAnt = DtNewRowAnt
            Me.FechaHasta = FechaHasta
            Me.Dias = Dias
        End Sub
    End Class

    <Task()> Public Shared Function CalcularRendimiento(ByVal data As DatosCalcRend, ByVal services As ServiceProvider) As DataTable
        Dim BancoPr As New BancoPropio
        Dim BancoInteres As New BancoPropioRendimiento
        Dim dttBancoInteres As DataTable
        Dim dttBancos As DataTable
        Dim dttApuntes As DataTable
        Dim dttRend As DataTable
        Dim F As New Filter
        Dim dblSaldoInicial As Double
        Dim dblSaldoAcum As Double
        Dim CampoFecha As String
        Dim colInteres() As DataRow
        Dim NewRow As DataRow
        Dim NewRowAnt As DataRow
        Dim FechaSaldoAnt As Date
        Dim SaldoAnt As Double
        Dim dias As Integer
        Dim Param As New Parametro
        Dim sql As String

        If Param.InteresesUsarFechaValor Then
            CampoFecha = "FechaValor"
        Else
            CampoFecha = "FechaApunte"
        End If

        dttRend = New DataTable
        dttRend.Columns.Add("Fecha", GetType(Date))
        dttRend.Columns.Add("FechaHasta", GetType(Date))
        dttRend.Columns.Add("IDBancoPropio", GetType(String))
        dttRend.Columns.Add("DescBancoPropio", GetType(String))
        dttRend.Columns.Add("Saldo", GetType(Double))
        dttRend.Columns.Add("Interes", GetType(Double))
        dttRend.Columns("Interes").DefaultValue = 0
        dttRend.Columns.Add("Rendimiento", GetType(Double))
        dttRend.Columns.Add("dias", GetType(Double))

        If Not data.IDBancoPropio Is Nothing Then
            F.Add("IDBancoPropio", FilterOperator.Equal, data.IDBancoPropio, FilterType.String)
        End If
        dttBancos = BancoPr.Filter(F)
        For Each row As DataRow In dttBancos.Rows
            If row("CContable") Is System.DBNull.Value Then Exit For
            dblSaldoAcum = 0
            dblSaldoInicial = 0
            'Calculo del saldo de partida para cada banco propio
            F = New Filter
            F.Add(CampoFecha, FilterOperator.GreaterThanOrEqual, New Date(data.FechaDesde.Year, 1, 1), FilterType.DateTime)
            F.Add(CampoFecha, FilterOperator.LessThan, data.FechaDesde, FilterType.DateTime)
            F.Add("IDCContable", FilterOperator.Equal, row("CContable"), FilterType.String)
            dttApuntes = New BE.DataEngine().Filter("vNegBancoPropioRendimiento", F, "Sum(ImpApunteA) as Saldo")
            If Not dttApuntes Is Nothing Then
                If dttApuntes.Rows.Count > 0 Then
                    If Not dttApuntes.Rows(0)("Saldo") Is System.DBNull.Value Then dblSaldoInicial = dttApuntes.Rows(0)("Saldo")
                    dblSaldoAcum = dblSaldoInicial
                End If
            End If
            'Fin de Calculo del saldo de partida para cada banco propio


            'Calculo de saldo y aplicación de intereses para cada día
            F = New Filter
            F.Add("IDBancoPropio", FilterOperator.Equal, row("IdBancoPropio"), FilterType.String)
            F.Add("FechaDesde", FilterOperator.LessThanOrEqual, data.FechaHasta, FilterType.DateTime)
            F.Add("FechaHasta", FilterOperator.GreaterThanOrEqual, data.FechaDesde, FilterType.DateTime)
            dttBancoInteres = BancoInteres.Filter(F, "FechaDesde ASC")

            F = New Filter
            F.Add(CampoFecha, FilterOperator.GreaterThanOrEqual, data.FechaDesde, FilterType.DateTime)
            F.Add(CampoFecha, FilterOperator.LessThanOrEqual, data.FechaHasta, FilterType.DateTime)
            F.Add("IDCContable", FilterOperator.Equal, row("CContable"), FilterType.String)
            sql = "SELECT " & CampoFecha & ", SUM(ImpApunteA) AS SaldoDia" & _
                  " FROM vNegBancoPropioRendimiento" & _
                  " WHERE " & AdminData.ComposeFilter(F) & _
                  " GROUP BY " & CampoFecha & _
                  " ORDER BY " & CampoFecha
            dttApuntes = AdminData.Execute(sql, ExecuteCommand.ExecuteReader)

            If dttApuntes.Rows.Count > 0 Then
                If dttApuntes.Rows(0)(CampoFecha) <> data.FechaDesde Then
                    F = New Filter
                    F.Add("FechaHasta", FilterOperator.GreaterThanOrEqual, data.FechaDesde, FilterType.DateTime)
                    F.Add("FechaDesde", FilterOperator.LessThan, dttApuntes.Rows(0)(CampoFecha), FilterType.DateTime)
                    Dim WhereFechas As String = F.Compose(New AdoFilterComposer)
                    colInteres = dttBancoInteres.Select(WhereFechas)

                    For i As Integer = 0 To colInteres.GetLength(0) - 1
                        NewRow = dttRend.NewRow
                        If colInteres(0)("FechaDesde") < data.FechaDesde Then
                            NewRow("Fecha") = data.FechaDesde
                        Else
                            NewRow("Fecha") = colInteres(0)("FechaDesde")
                        End If

                        NewRow("IDBancoPropio") = row("IdBancoPropio")
                        NewRow("DescBancoPropio") = row("DescBancoPropio")
                        NewRow("Saldo") = dblSaldoInicial
                        If NewRow("Saldo") > 0 Then
                            NewRow("Interes") = colInteres(0)("InteresAcreedor")
                        Else
                            NewRow("Interes") = colInteres(0)("InteresDeudor")
                        End If
                        If i < colInteres.GetLength(0) - 1 Then
                            dias = CType(CType(colInteres.GetValue(i + 1), DataRow)("FechaDesde"), Date).DayOfYear - CType(NewRow("Fecha"), Date).DayOfYear
                            Dim DtNew As DataTable = NewRow.Table.Clone
                            DtNew.ImportRow(NewRow)
                            Dim DtNewAnt As DataTable = NewRowAnt.Table.Clone
                            DtNewAnt.ImportRow(NewRowAnt)
                            Dim StDatos As New DatosCalcIntereses(DtNew, DtNew, CType(colInteres.GetValue(i + 1), DataRow)("FechaDesde"), dias)
                            ProcessServer.ExecuteTask(Of DatosCalcIntereses)(AddressOf CalcularInteres, StDatos, services)
                            NewRow.ItemArray = StDatos.DtNewRow.Rows(0).ItemArray
                            NewRowAnt.ItemArray = StDatos.DtNewRowAnt.Rows(0).ItemArray
                        Else
                            dias = CType(dttApuntes.Rows(0)(CampoFecha), Date).DayOfYear - CType(NewRow("Fecha"), Date).DayOfYear
                            Dim DtNew As DataTable = NewRow.Table.Clone
                            DtNew.ImportRow(NewRow)
                            Dim DtNewAnt As DataTable = NewRowAnt.Table.Clone
                            DtNewAnt.ImportRow(NewRowAnt)
                            Dim StDatos As New DatosCalcIntereses(DtNew, DtNew, dttApuntes.Rows(0)(CampoFecha), dias)
                            ProcessServer.ExecuteTask(Of DatosCalcIntereses)(AddressOf CalcularInteres, StDatos, services)
                            NewRow.ItemArray = StDatos.DtNewRow.Rows(0).ItemArray
                            NewRowAnt.ItemArray = StDatos.DtNewRowAnt.Rows(0).ItemArray
                        End If
                        dttRend.Rows.Add(NewRow)
                    Next
                End If
                FechaSaldoAnt = dttApuntes.Rows(0)(CampoFecha)
            End If
            For Each row1 As DataRow In dttApuntes.Rows
                F = New Filter
                If FechaSaldoAnt <> row1(CampoFecha) AndAlso row1(CampoFecha) > FechaSaldoAnt.AddDays(1) Then
                    F.Add("FechaDesde", FilterOperator.GreaterThan, FechaSaldoAnt, FilterType.DateTime)
                    F.Add("FechaDesde", FilterOperator.LessThan, row1(CampoFecha), FilterType.DateTime)
                    Dim WhereFechasPeriodo As String = F.Compose(New AdoFilterComposer)
                    colInteres = dttBancoInteres.Select(WhereFechasPeriodo)

                    If colInteres.GetLength(0) > 0 Then
                        dias = CType(colInteres(0)("FechaDesde"), Date).DayOfYear - CType(NewRowAnt("Fecha"), Date).DayOfYear
                        Dim DtNew As DataTable = NewRow.Table.Clone
                        DtNew.ImportRow(NewRow)
                        Dim DtNewAnt As DataTable = NewRowAnt.Table.Clone
                        DtNewAnt.ImportRow(NewRowAnt)
                        Dim StDatos As New DatosCalcIntereses(DtNew, DtNewAnt, colInteres(0)("FechaDesde"), dias)
                        ProcessServer.ExecuteTask(Of DatosCalcIntereses)(AddressOf CalcularInteres, StDatos, services)
                        NewRow.ItemArray = StDatos.DtNewRow.Rows(0).ItemArray
                        NewRowAnt.ItemArray = StDatos.DtNewRowAnt.Rows(0).ItemArray
                    Else
                        dias = CType(row1(CampoFecha), Date).DayOfYear - CType(NewRowAnt("Fecha"), Date).DayOfYear
                        Dim DtNew As DataTable = NewRow.Table.Clone
                        DtNew.ImportRow(NewRow)
                        Dim DtNewAnt As DataTable = NewRowAnt.Table.Clone
                        DtNewAnt.ImportRow(NewRowAnt)
                        Dim StDatos As New DatosCalcIntereses(DtNew, DtNewAnt, row1(CampoFecha), dias)
                        ProcessServer.ExecuteTask(Of DatosCalcIntereses)(AddressOf CalcularInteres, StDatos, services)
                        NewRow.ItemArray = StDatos.DtNewRow.Rows(0).ItemArray
                        NewRowAnt.ItemArray = StDatos.DtNewRowAnt.Rows(0).ItemArray
                    End If
                    For i As Integer = 0 To colInteres.GetLength(0) - 1
                        NewRow = dttRend.NewRow
                        NewRow("Fecha") = colInteres(0)("FechaDesde")
                        NewRow("IDBancoPropio") = row("IdBancoPropio")
                        NewRow("DescBancoPropio") = row("DescBancoPropio")
                        NewRow("Saldo") = SaldoAnt
                        If NewRow("Saldo") > 0 Then
                            NewRow("Interes") = colInteres(0)("InteresAcreedor")
                        Else
                            NewRow("Interes") = colInteres(0)("InteresDeudor")
                        End If
                        If i < colInteres.GetLength(0) - 1 Then
                            dias = CType(CType(colInteres.GetValue(i + 1), DataRow)("FechaDesde"), Date).DayOfYear - CType(NewRow("Fecha"), Date).DayOfYear
                            Dim DtNew As DataTable = NewRow.Table.Clone
                            DtNew.ImportRow(NewRow)
                            Dim DtNewAnt As DataTable = NewRowAnt.Table.Clone
                            DtNewAnt.ImportRow(NewRowAnt)
                            Dim StDatos As New DatosCalcIntereses(DtNew, DtNewAnt, CType(colInteres.GetValue(i + 1), DataRow)("FechaDesde"), dias)
                            ProcessServer.ExecuteTask(Of DatosCalcIntereses)(AddressOf CalcularInteres, StDatos, services)
                            NewRow.ItemArray = StDatos.DtNewRow.Rows(0).ItemArray
                            NewRowAnt.ItemArray = StDatos.DtNewRowAnt.Rows(0).ItemArray
                        Else
                            dias = CType(row1(CampoFecha), Date).DayOfYear - CType(NewRow("Fecha"), Date).DayOfYear
                            Dim DtNew As DataTable = NewRow.Table.Clone
                            DtNew.ImportRow(NewRow)
                            Dim DtNewAnt As DataTable = NewRowAnt.Table.Clone
                            DtNewAnt.ImportRow(NewRowAnt)
                            Dim StDatos As New DatosCalcIntereses(DtNew, DtNewAnt, row1(CampoFecha), dias)
                            ProcessServer.ExecuteTask(Of DatosCalcIntereses)(AddressOf CalcularInteres, StDatos, services)
                            NewRow.ItemArray = StDatos.DtNewRow.Rows(0).ItemArray
                            NewRowAnt.ItemArray = StDatos.DtNewRowAnt.Rows(0).ItemArray
                        End If
                        dttRend.Rows.Add(NewRow)
                        FechaSaldoAnt = CType(NewRow("Fecha"), Date)
                    Next

                ElseIf Not NewRowAnt Is Nothing Then
                    dias = CType(row1(CampoFecha), Date).DayOfYear - CType(NewRowAnt("Fecha"), Date).DayOfYear
                    Dim DtNew As DataTable = NewRow.Table.Clone
                    DtNew.ImportRow(NewRow)
                    Dim DtNewAnt As DataTable = NewRowAnt.Table.Clone
                    DtNewAnt.ImportRow(NewRowAnt)
                    Dim StDatos As New DatosCalcIntereses(DtNew, DtNewAnt, row1(CampoFecha), dias)
                    ProcessServer.ExecuteTask(Of DatosCalcIntereses)(AddressOf CalcularInteres, StDatos, services)
                    NewRow.ItemArray = StDatos.DtNewRow.Rows(0).ItemArray
                    NewRowAnt.ItemArray = StDatos.DtNewRowAnt.Rows(0).ItemArray
                End If

                F = New Filter
                F.Add("FechaDesde", FilterOperator.LessThanOrEqual, row1(CampoFecha), FilterType.DateTime)
                F.Add("FechaHasta", FilterOperator.GreaterThanOrEqual, row1(CampoFecha), FilterType.DateTime)
                Dim WhereFechas As String = F.Compose(New AdoFilterComposer)
                colInteres = dttBancoInteres.Select(WhereFechas)
                NewRow = dttRend.NewRow
                NewRow("Fecha") = row1(CampoFecha)
                NewRow("IDBancoPropio") = row("IdBancoPropio")
                NewRow("DescBancoPropio") = row("DescBancoPropio")
                dblSaldoAcum = dblSaldoAcum + row1("SaldoDia")
                NewRow("Saldo") = dblSaldoAcum
                If colInteres.GetLength(0) > 0 Then
                    If NewRow("Saldo") > 0 Then
                        NewRow("Interes") = colInteres(0)("InteresAcreedor")
                    Else
                        NewRow("Interes") = colInteres(0)("InteresDeudor")
                    End If
                    NewRow("Rendimiento") = xRound((NewRow("Saldo") * NewRow("Interes")) / 100, 2)
                Else
                    NewRow("Rendimiento") = 0
                End If
                dttRend.Rows.Add(NewRow)
                NewRowAnt = NewRow
                FechaSaldoAnt = CType(row1(CampoFecha), Date)
                SaldoAnt = dblSaldoAcum
            Next
            If Not NewRowAnt Is Nothing Then
                dias = data.FechaHasta.DayOfYear - CType(NewRowAnt("Fecha"), Date).DayOfYear + 1
                Dim DtNew As DataTable = NewRow.Table.Clone
                DtNew.ImportRow(NewRow)
                Dim DtNewAnt As DataTable = NewRowAnt.Table.Clone
                DtNewAnt.ImportRow(NewRowAnt)
                Dim StDatos As New DatosCalcIntereses(DtNew, DtNewAnt, data.FechaHasta, dias)
                ProcessServer.ExecuteTask(Of DatosCalcIntereses)(AddressOf CalcularInteres, StDatos, services)
                NewRow.ItemArray = StDatos.DtNewRow.Rows(0).ItemArray
                NewRowAnt.ItemArray = StDatos.DtNewRowAnt.Rows(0).ItemArray
            End If
        Next
        Return dttRend
    End Function

    <Task()> Public Shared Sub CalcularInteres(ByVal data As DatosCalcIntereses, ByVal service As ServiceProvider)
        data.DtNewRow.Rows(0)("FechaHasta") = data.FechaHasta
        data.DtNewRowAnt.Rows(0)("dias") = data.Dias
        data.DtNewRowAnt.Rows(0)("Rendimiento") = xRound((data.DtNewRowAnt.Rows(0)("Saldo") * data.DtNewRowAnt.Rows(0)("Interes") * data.Dias) / (100 * 365), 2)
    End Sub

#End Region

End Class