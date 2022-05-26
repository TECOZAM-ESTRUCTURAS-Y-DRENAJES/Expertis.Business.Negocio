Public Class CobroDevolucion
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper
    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbCobroDevolucion"

#Region " RegisterAddNewTasks "

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarIdentificador, data, services)
    End Sub

#End Region

#Region " Update "

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarIdentificador)
        updateProcess.AddTask(Of DataRow)(AddressOf ActualizarARepercutirEnCobro)
    End Sub

    <Task()> Public Shared Sub AsignarIdentificador(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IdDevolucion")) = 0 Then data("IdDevolucion") = AdminData.GetAutoNumeric
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarARepercutirEnCobro(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Modified Then
            Dim dblARepercutirOLD As Double = Nz(data("ARepercutir", DataRowVersion.Original), 0)
            Dim dblARepercutirAOLD As Double = Nz(data("ARepercutirA", DataRowVersion.Original), 0)
            Dim dblARepercutirBOLD As Double = Nz(data("ARepercutirB", DataRowVersion.Original), 0)

            Dim dtCobro As DataTable = New Cobro().SelOnPrimaryKey(data("IdCobro"))
            If Not IsNothing(dtCobro) AndAlso dtCobro.Rows.Count > 0 Then
                dtCobro.Rows(0)("ARepercutir") = dtCobro.Rows(0)("ARepercutir") + data("ARepercutir") - dblARepercutirOLD
                dtCobro.Rows(0)("ARepercutirA") = dtCobro.Rows(0)("ARepercutirA") + data("ARepercutirA") - dblARepercutirAOLD
                dtCobro.Rows(0)("ARepercutirB") = dtCobro.Rows(0)("ARepercutirB") + data("ARepercutirB") - dblARepercutirBOLD

                BusinessHelper.UpdateTable(dtCobro)
            End If
        End If
    End Sub

#End Region

#Region " BusinessRules "

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBRL As New BusinessRules
        oBRL.Add("GastoA", AddressOf CambioGastoA)
        oBRL.Add("ComisionA", AddressOf CambioComisionA)
        oBRL.Add("ARepercutirA", AddressOf CambioARepercutirA)
        Return oBRL
    End Function

    <Task()> Public Shared Sub CambioGastoA(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Not IsNumeric(data.Value) Then ApplicationService.GenerateError("El campo debe ser numérico.")
        data.Current(data.ColumnName) = data.Value
        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
        Dim MonInfoA As MonedaInfo = Monedas.MonedaA
        Dim MonInfoB As MonedaInfo = Monedas.MonedaB

        '//No utilizamos el MantenimientoValoresAyB, por que se recalcula desde GastoA
        data.Current("GastoA") = xRound(data.Current("GastoA"), MonInfoA.NDecimalesImporte)
        If Length(data.Current("IDMoneda")) > 0 AndAlso Length(data.Current("CambioA")) > 0 AndAlso data.Current("CambioA") <> 0 Then
            Dim MonInfo As MonedaInfo = Monedas.GetMoneda(data.Current("IDMoneda"))
            data.Current("Gasto") = xRound(CDbl(data.Current("GastoA")) / CDbl(data.Current("CambioA")), MonInfo.NDecimalesImporte)
            data.Current("GastoB") = xRound(CDbl(data.Current("Gasto")) * CDbl(data.Current("CambioB")), MonInfoB.NDecimalesImporte)
        Else
            data.Current("Gasto") = 0
            data.Current("GastoB") = 0
        End If

        Dim dataReper As New BusinessRuleData("ARepercutirA", Nz(data.Current("GastoA"), 0) + Nz(data.Current("ComisionA"), 0), data.Current, data.Context)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CambioARepercutirA, dataReper, services)
    End Sub

    <Task()> Public Shared Sub CambioComisionA(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Not IsNumeric(data.Value) Then ApplicationService.GenerateError("El campo debe ser numérico.")

        data.Current(data.ColumnName) = data.Value
        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
        Dim MonInfoA As MonedaInfo = Monedas.MonedaA
        Dim MonInfoB As MonedaInfo = Monedas.MonedaB

        '//No utilizamos el MantenimientoValoresAyB, por que se recalcula desde ComisionA
        data.Current("ComisionA") = xRound(data.Current("ComisionA"), MonInfoA.NDecimalesImporte)
        If Length(data.Current("IDMoneda")) > 0 AndAlso Length(data.Current("CambioA")) > 0 AndAlso data.Current("CambioA") <> 0 Then
            Dim MonInfo As MonedaInfo = Monedas.GetMoneda(data.Current("IDMoneda"))
            data.Current("Comision") = xRound(CDbl(data.Current("ComisionA")) / CDbl(data.Current("CambioA")), MonInfo.NDecimalesImporte)
            data.Current("ComisionB") = xRound(CDbl(data.Current("Comision")) * CDbl(data.Current("CambioB")), MonInfoB.NDecimalesImporte)
        Else
            data.Current("Comision") = 0
            data.Current("ComisionB") = 0
        End If

        Dim dataReper As New BusinessRuleData("ARepercutirA", Nz(data.Current("GastoA"), 0) + Nz(data.Current("ComisionA"), 0), data.Current, data.Context)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CambioARepercutirA, dataReper, services)
    End Sub

    <Task()> Public Shared Sub CambioARepercutirA(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Not IsNumeric(data.Value) Then ApplicationService.GenerateError("El campo debe ser numérico.")

        data.Current(data.ColumnName) = data.Value
        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
        Dim MonInfoA As MonedaInfo = Monedas.MonedaA
        Dim MonInfoB As MonedaInfo = Monedas.MonedaB
        data.Current("ARepercutirA") = xRound(data.Current("ARepercutirA"), MonInfoA.NDecimalesImporte)
        If Length(data.Current("IDMoneda")) > 0 AndAlso Length(data.Current("CambioA")) > 0 AndAlso data.Current("CambioA") <> 0 Then
            Dim MonInfo As MonedaInfo = Monedas.GetMoneda(data.Current("IDMoneda"))
            data.Current("ARepercutir") = xRound(CDbl(data.Current("ARepercutirA")) / CDbl(data.Current("CambioA")), MonInfo.NDecimalesImporte)
            data.Current("ARepercutirB") = xRound(CDbl(data.Current("ARepercutir")) * CDbl(data.Current("CambioB")), MonInfoB.NDecimalesImporte)
        Else
            data.Current("ARepercutir") = 0
            data.Current("ARepercutirB") = 0
        End If
    End Sub

#End Region

    <Serializable()> _
    Public Class DataBorrarDevolucion
        Public IDCobroDevolucion As Integer
        Public IDEstado As Integer
        Public IDRemesa As Integer
        Public IDFacturaCompra As Integer
    End Class

    <Task()> Public Shared Function ComprobarFacturaDevolucionABorrar(ByVal data As DataBorrarDevolucion, ByVal services As ServiceProvider) As Integer
        Dim CD As New CobroDevolucion
        Dim dtCobroDev As DataTable = CD.SelOnPrimaryKey(data.IDCobroDevolucion)
        If dtCobroDev.Rows.Count = 0 Then Return -1
        If Length(Nz(dtCobroDev.Rows(0)("IDFacturaCompra"))) > 1 Then
            Return dtCobroDev.Rows(0)("IDFacturaCompra") 'La factura asociada a la devolución agrupa más de un cobro.
        Else
            Return -1
        End If
    End Function

    <Task()> Public Shared Function BorrarDevolucion(ByVal data As DataBorrarDevolucion, ByVal services As ServiceProvider) As Boolean
        Dim AppParams As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()

        AdminData.BeginTx()
        Dim CD As New CobroDevolucion
        Dim dtCobroDev As DataTable


        Dim ClsFact As New FacturaCompraCabecera
        Dim DtFact As DataTable
        Dim blnEliminarFra As Boolean = False

        Dim IntIDCobro As Integer
        If data.IDFacturaCompra = -1 Then
            dtCobroDev = CD.SelOnPrimaryKey(data.IDCobroDevolucion)
        Else
            'Si hay factura generada buscamos todos los cobros devueltos agrupados en esa factura. Deberán ser borrados todos. 
            dtCobroDev = CD.Filter(New StringFilterItem("IDFacturaCompra", data.IDFacturaCompra))
        End If
        If dtCobroDev.Rows.Count = 0 Then Exit Function

        If dtCobroDev.Rows.Count > 0 Then
            For Each dr As DataRow In dtCobroDev.Rows
                IntIDCobro = dr("IDCobro")
                Dim IDEjercicio As String = dr("IDEjercicio") & String.Empty
                Dim NAsiento As Integer = Nz(dr("NAsiento"), -1)

                Dim FechaCobro As Date
                If Nz(dr("FechaCobro")) <> cnMinDate Then FechaCobro = dr("FechaCobro")

                Dim IDEjercicioTributario As String = dr("IDEjercicioTributario") & String.Empty
                Dim NAsientoTributario As Integer = Nz(dr("NAsientoTributario"), -1)

                If Length(IDEjercicio) = 0 AndAlso Length(IDEjercicioTributario) = 0 Then
                    '//Este if es para que se pueden hacer devoluciones sin tener información contable en la devolución como teníamos hasta ahora.
                    '//Tiene el inconveniente, q varias devoluciones en la misma fecha del mismo cobro, se harían incorrectamente las devoluciones.
                    Return ProcessServer.ExecuteTask(Of DataBorrarDevolucion, Boolean)(AddressOf BorrarDevolucionSinInfoContable, data, services)
                End If

                Dim IDFacturaCompra As Integer = Nz(dr("IDFacturaCompra"), -1)
                Dim IDRemesaAnterior As Integer = Nz(dr("IDRemesaAnterior"), -1)
                Dim DblARepercutir As Double = Nz(dr("ARepercutir"), 0)
                Dim DblARepercutirA As Double = Nz(dr("ARepercutirA"), 0)
                Dim DblARepercutirB As Double = Nz(dr("ARepercutirB"), 0)

                Dim FilCobro As New Filter
                ' FilCobro.Add("IDDocumento", FilterOperator.Equal, IntIDCobro, FilterType.Numeric)
                FilCobro.Add("IDTipoApunte", FilterOperator.Equal, enumDiarioTipoApunte.DevolucionRemesa, FilterType.Numeric)

                Dim fAsientos As New Filter(FilterUnionOperator.Or)
                If Length(IDEjercicio) > 0 Then
                    Dim fAsientoNIIF As New Filter
                    fAsientoNIIF.Add(New StringFilterItem("IDEjercicio", IDEjercicio))
                    fAsientoNIIF.Add(New NumberFilterItem("NAsiento", NAsiento))
                    fAsientos.Add(fAsientoNIIF)
                End If

                If Length(IDEjercicioTributario) > 0 Then
                    Dim fAsientoTributario As New Filter
                    fAsientoTributario.Add(New StringFilterItem("IDEjercicioTributario", IDEjercicioTributario))
                    fAsientoTributario.Add(New NumberFilterItem("NAsientoTributario", NAsientoTributario))
                    fAsientos.Add(fAsientoTributario)
                End If
                FilCobro.Add(fAsientos)

                Dim ClsDiario As Object = BusinessHelper.CreateBusinessObject("DiarioContable")
                Dim DtDiario As DataTable = ClsDiario.Filter(FilCobro)
                If Not DtDiario Is Nothing AndAlso DtDiario.Rows.Count > 0 Then

                    If IDFacturaCompra <> -1 Then
                        DtFact = ClsFact.SelOnPrimaryKey(IDFacturaCompra)
                        If Not DtFact Is Nothing AndAlso DtFact.Rows.Count > 0 Then
                            DtFact.Rows(0)("Estado") = enumfccEstado.fccNoContabilizado
                            blnEliminarFra = True
                        End If
                    End If

                    ClsDiario.Delete(DtDiario)
                End If

                CD.Delete(dr)
                Dim ClsCobro As New Cobro
                Dim DtCobro As DataTable = ClsCobro.SelOnPrimaryKey(dr("IDCobro"))
                If data.IDRemesa <> 0 Then
                    DtCobro.Rows(0)("IDRemesa") = data.IDRemesa
                Else
                    If IDRemesaAnterior <> -1 Then
                        DtCobro.Rows(0)("IDRemesa") = IDRemesaAnterior
                    End If
                End If

                If Length(DtCobro.Rows(0)("IDRemesa")) > 0 Then
                    '//Cambiar el tipo del asiento de remesa y ponerlo de tipo de liquidación (si es que tenia) (AL CREAR LA DEVOLUCION HACEMOS EL PROCESO INVERSO).
                    '// Sólo debemos modificar el apunte del cobro (podemos tener IDCobro en el IDDocumento del apunte del banco, 
                    '//en este caso el apunte del banco no debemos cambiarlo, para podernos apoyar en él por que en él tenemos en el NDocumento el NºRemesa)
                    Dim datLiquidacion As New Cobro.DataTipoLiquidacionATipoRemesa(IntIDCobro, DtCobro.Rows(0)("IDRemesa"))
                    datLiquidacion.Inversa = True  '// De Remesa a Liquidación
                    DtDiario = ProcessServer.ExecuteTask(Of Cobro.DataTipoLiquidacionATipoRemesa, DataTable)(AddressOf Cobro.TipoLiquidacionATipoRemesa, datLiquidacion, services)
                    BusinessHelper.UpdateTable(DtDiario)
                End If

                DtCobro.Rows(0)("ARepercutir") -= DblARepercutir
                DtCobro.Rows(0)("ARepercutirA") -= DblARepercutirA
                DtCobro.Rows(0)("ARepercutirB") -= DblARepercutirB
                DtCobro.Rows(0)("Situacion") = data.IDEstado
                If FechaCobro <> cnMinDate Then DtCobro.Rows(0)("FechaCobro") = FechaCobro

                If DtCobro.Rows(0)("Situacion") = enumCobroSituacion.Cobrado OrElse DtCobro.Rows(0)("Situacion") = enumCobroSituacion.Descontado Then
                    DtCobro.Rows(0)("Liquidado") = True
                End If
                DtCobro.Rows(0)("Contabilizado") = enumContabilizado.Contabilizado
                ClsCobro.Update(DtCobro)
            Next

            If blnEliminarFra AndAlso DtFact.Rows.Count > 0 Then
                DtFact = ClsFact.Update(DtFact)
                ClsFact.Delete(DtFact)
            End If
            Return True
        End If
        Return False
    End Function


    <Task()> Public Shared Function BorrarDevolucionSinInfoContable(ByVal data As DataBorrarDevolucion, ByVal services As ServiceProvider) As Boolean
        Dim AppParams As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
        AdminData.BeginTx()
        Dim CD As New CobroDevolucion
        Dim dtCobroDev As DataTable = CD.SelOnPrimaryKey(data.IDCobroDevolucion)
        If dtCobroDev.Rows.Count = 0 Then Exit Function

        Dim IntIDCobro As Integer = dtCobroDev.Rows(0)("IDCobro")
        Dim FechaCobro As Date
        If Nz(dtCobroDev.Rows(0)("FechaCobro")) <> cnMinDate Then FechaCobro = dtCobroDev.Rows(0)("FechaCobro")
        Dim DblARepercutir, DblARepercutirA, DblARepercutirB As Double
        DblARepercutir = dtCobroDev.Rows(0)("ARepercutir")
        DblARepercutirA = dtCobroDev.Rows(0)("ARepercutirA")
        DblARepercutirB = dtCobroDev.Rows(0)("ARepercutirB")
        Dim FilCobro As New Filter
        FilCobro.Add("IDDocumento", FilterOperator.Equal, IntIDCobro, FilterType.Numeric)
        FilCobro.Add("IDTipoApunte", FilterOperator.Equal, enumDiarioTipoApunte.DevolucionRemesa, FilterType.Numeric)
        Dim ClsDiario As Object = BusinessHelper.CreateBusinessObject("DiarioContable")
        Dim DtDiario As DataTable = ClsDiario.Filter(FilCobro)
        If Not DtDiario Is Nothing AndAlso DtDiario.Rows.Count > 0 Then
            Dim intNAsiento As Integer = DtDiario.Rows(0)("NAsiento")
            Dim StrIDEjercicio As String = DtDiario.Rows(0)("IDEjercicio")
            Dim IDEjercicioTributario As String
            If AppParams.ContabilidadMultiple Then
                Dim f As New Filter
                f.Add(New StringFilterItem("IDEjercicio", StrIDEjercicio))
                f.Add(New IsNullFilterItem("IDEjercicioTributario", False))
                Dim dtEjercicioTrib As DataTable = New BE.DataEngine().Filter("tbMaestroEjercicio", f)
                If dtEjercicioTrib.Rows.Count > 0 Then
                    IDEjercicioTributario = dtEjercicioTrib.Rows(0)("IDEjercicioTributario")
                End If
            End If
            Dim FilDef As New Filter
            FilDef.Add("NAsiento", FilterOperator.Equal, intNAsiento, FilterType.Numeric)
            FilDef.Add("IDEjercicio", FilterOperator.Equal, StrIDEjercicio, FilterType.String)
            FilDef.Add("IDTipoApunte", FilterOperator.Equal, enumDiarioTipoApunte.DevolucionRemesa, FilterType.Numeric)
            Dim DtDiarioDef As DataTable = ClsDiario.Filter(FilDef, "IDApunte")
            If Not DtDiarioDef Is Nothing AndAlso DtDiarioDef.Rows.Count > 0 Then
                If Length(DtDiarioDef.Rows(DtDiarioDef.Rows.Count - 1)("IDDocumento")) > 0 Then
                    Dim ClsFact As New FacturaCompraCabecera
                    Dim DtFact As DataTable = ClsFact.Filter(New FilterItem("IDFactura", FilterOperator.Equal, DtDiarioDef.Rows(DtDiarioDef.Rows.Count - 1)("IDDocumento"), FilterType.String))
                    If Not DtFact Is Nothing AndAlso DtFact.Rows.Count > 0 Then
                        DtFact.Rows(0)("Estado") = enumfccEstado.fccNoContabilizado
                        DtFact = ClsFact.Update(DtFact)
                        ClsFact.Delete(DtFact)
                    End If
                End If

                If AppParams.ContabilidadMultiple AndAlso Length(IDEjercicioTributario) > 0 Then
                    Dim fTributario As New Filter
                    fTributario.Add("IDDocumento", FilterOperator.Equal, IntIDCobro, FilterType.Numeric)
                    fTributario.Add("IDTipoApunte", FilterOperator.Equal, enumDiarioTipoApunte.DevolucionRemesa, FilterType.Numeric)
                    fTributario.Add("IDEjercicio", FilterOperator.Equal, IDEjercicioTributario, FilterType.String)
                    Dim dtAstoTributario As DataTable = ClsDiario.Filter(fTributario)
                    ClsDiario.Delete(dtAstoTributario)
                End If

                ClsDiario.Delete(DtDiarioDef)
                CD.Delete(dtCobroDev)
                Dim ClsCobro As New Cobro
                Dim DtCobro As DataTable = ClsCobro.SelOnPrimaryKey(IntIDCobro)
                DtCobro.Rows(0)("ARepercutir") -= DblARepercutir
                DtCobro.Rows(0)("ARepercutirA") -= DblARepercutirA
                DtCobro.Rows(0)("ARepercutirB") -= DblARepercutirB
                DtCobro.Rows(0)("Situacion") = data.IDEstado
                If FechaCobro <> cnMinDate Then DtCobro.Rows(0)("FechaCobro") = FechaCobro
                If data.IDRemesa <> 0 Then DtCobro.Rows(0)("IDRemesa") = data.IDRemesa
                DtCobro.Rows(0)("Contabilizado") = enumContabilizado.Contabilizado
                ClsCobro.Update(DtCobro)
                Return True
            End If

        End If
        Return False
    End Function

End Class