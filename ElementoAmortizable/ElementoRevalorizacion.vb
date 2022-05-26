Public Class ElementoRevalorizacion

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbElementoRevalorizacion"

#End Region

#Region "Eventos RegisterDeleteTasks"

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf Comunes.DeleteEntityRow)
        deleteProcess.AddTask(Of DataRow)(AddressOf Comunes.MarcarComoEliminado)
        deleteProcess.AddTask(Of DataRow)(AddressOf ActualizarElementoAmortizable)
    End Sub

    <Task()> Public Shared Sub ActualizarElementoAmortizable(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim DtReval As DataTable = New ElementoRevalorizacion().Filter(New FilterItem("IDElemento", FilterOperator.Equal, data("IDElemento"), FilterType.String), "FechaRevalorizacion DESC")
        If Not DtReval Is Nothing AndAlso DtReval.Rows.Count > 0 Then
            Dim DtElem As DataTable = New ElementoAmortizable().SelOnPrimaryKey(data("IDElemento"))
            If Not DtElem Is Nothing AndAlso DtElem.Rows.Count > 0 Then
                DtElem.Rows(0)("FechaUltimaRevalorizacion") = DtReval.Rows(0)("FechaRevalorizacion")
                DtElem.Rows(0)("ValorTotalRevalElementoA") = DtReval.Rows(0)("ValorCompraFechaA")
                Dim StValor As New ElementoAmortizable.DataCalcValorNetoContable(DtElem.Rows(0)("ValorTotalRevalElementoA"), DtElem.Rows(0)("ValorAmortizadoElementoA"), DtElem.Rows(0)("ValorResidualA"))
                DtElem.Rows(0)("ValorNetoContableElementoA") = ProcessServer.ExecuteTask(Of ElementoAmortizable.DataCalcValorNetoContable, Double)(AddressOf ElementoAmortizable.CalcularValorNetoContable, StValor, services)

                DtElem.Rows(0)("ValorTotalPlusvaliaA") = DtReval.Rows(0)("ValorPlusvaliaFechaA")
                DtElem.Rows(0)("ValorAmortizadoPlusvaliaA") = DtReval.Rows(0)("ValorAmortizadoPlusvaliaFechaA")
                DtElem.Rows(0)("ValorNetoContablePlusvaliaA") = DtElem.Rows(0)("ValorTotalPlusvaliaA") - DtElem.Rows(0)("ValorAmortizadoPlusvaliaA")

                DtElem.Rows(0)("IDCodigoAmortizacionContable") = DtReval.Rows(0)("IDTipoAmortizacionFecha")
                BusinessHelper.UpdateTable(DtElem)
            End If
        End If
    End Sub

#End Region

#Region "Eventos RegisterUpdateTask"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarIdentificador)
    End Sub

    <Task()> Public Shared Sub AsignarIdentificador(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added AndAlso Length(data("IDLineaRevalorizacion")) = 0 Then
            data("IDLineaRevalorizacion") = AdminData.GetAutoNumeric
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Function ActualizarLineasRevalorizacion(ByVal dr As DataRow, ByVal services As ServiceProvider) As DataTable
        If Not dr Is Nothing Then
            Dim strFiltro As String = String.Empty
            Dim StData As New DataCrearDtElemeReval(dr.Table, dr("FechaUltimaRevalorizacion"))
            Dim dtNuevoElemReval As DataTable = ProcessServer.ExecuteTask(Of DataCrearDtElemeReval, DataTable)(AddressOf CrearDtElemReval, StData, services)
            If Length(dr("AmortizacionAutomatica")) = 0 OrElse dr("AmortizacionAutomatica") = False Then 'No se ha iniciado la amortiz.auto.
                If dr("FechaUltimaRevalorizacion") < dr("FechaInicioContabilizacion") Then
                    If dr.RowState = DataRowState.Modified AndAlso dr("ValorAmortizadoElementoA") = 0 Then
                        dr("FechaUltimaRevalorizacion") = dr("FechaInicioContabilizacion")
                        dtNuevoElemReval.Rows(0)("FechaRevalorizacion") = dr("FechaInicioContabilizacion")
                    Else : ApplicationService.GenerateError("La fecha de revalorizacion no puede ser menor que la fecha de compra.")
                    End If
                Else : dtNuevoElemReval.Rows(0)("FechaRevalorizacion") = dr("FechaUltimaRevalorizacion")
                End If
                If (Not dr.IsNull("AñoUltimoContabilizado") AndAlso dr("AñoUltimoContabilizado") > 0) AndAlso dr("AñoUltimoContabilizado") < dtNuevoElemReval.Rows(0)("AñoRevalorizacion") Then
                    'La fecha de Compra o fecha de revalorizacion no puede ser mayor que la fecha de Ultima Contabilizacion
                    ApplicationService.GenerateError("La fecha de Compra o fecha de revalorizacion no puede ser mayor que la fecha de Ultima Contabilizacion")
                ElseIf Not AreEquals(dr("AñoUltimoContabilizado"), 0) And AreEquals(dr("AñoUltimoContabilizado"), dtNuevoElemReval.Rows(0)("AñoRevalorizacion")) Then
                    If Not AreEquals(dr("MesUltimoContabilizado"), 0) And dr("MesUltimoContabilizado") < dtNuevoElemReval.Rows(0)("MesRevalorizacion") Then
                        'La fecha de Compra o fecha de revalorizacion no puede ser mayor que la fecha de Ultima Contabilizacion
                        ApplicationService.GenerateError("La fecha de Compra o fecha de revalorizacion no puede ser mayor que la fecha de Ultima Contabilizacion")
                    Else : strFiltro = "IdElemento = '" & dr("IdElemento") & "'"
                    End If
                Else : strFiltro = "IdElemento = '" & dr("IdElemento") & "'"
                End If
            Else
                'EL PROGRAMA HA EMPEZADO A AMORTIZAR...
                If dr("FechaUltimaRevalorizacion") < dr("FechaUltimaContabilizacion") Then
                    ApplicationService.GenerateError("La fecha de revalorizacion no puede ser menor que la fecha de ultima contabilizacion")
                End If
                'Si se ha empezado a amortizar, se eliminan los que tengan FechaReval > FechaUltimaContab (solo habra 1)

                If dtNuevoElemReval.Rows(0)("AñoRevalorizacion") > dr("AñoUltimoContabilizado") Then
                    strFiltro = "IdElemento = '" & dr("IdElemento") & "' AND (AñoRevalorizacion >" & dr("AñoUltimoContabilizado") _
                            & " OR (AñoRevalorizacion=" & dr("AñoUltimoContabilizado") & " AND MesRevalorizacion > " & dr("MesUltimoContabilizado") & "))"
                ElseIf dtNuevoElemReval.Rows(0)("AñoRevalorizacion") = dr("AñoUltimoContabilizado") Then
                    If dtNuevoElemReval.Rows(0)("MesRevalorizacion") >= dr("MesUltimoContabilizado") Then
                        strFiltro = "IdElemento = '" & dr("IdElemento") & "' AND (AñoRevalorizacion >" & dr("AñoUltimoContabilizado") _
                            & " OR (AñoRevalorizacion=" & dr("AñoUltimoContabilizado") & " AND MesRevalorizacion > " & dr("MesUltimoContabilizado") _
                            & ") OR (AñoRevalorizacion=" & dr("AñoUltimoContabilizado") & " AND MesRevalorizacion = " & dr("MesUltimoContabilizado") _
                            & " AND FechaRevalorizacion >='" & dr("FechaUltimaContabilizacion") & "'))"
                    Else : ApplicationService.GenerateError("La fecha de Revalorizacion no puede ser menor que la fecha de Ultima Contabilizacion. Debe deshacer lo amortizado hasta la fecha de Revalorizacion")
                    End If
                Else : ApplicationService.GenerateError("La fecha de Revalorizacion no puede ser menor que la fecha de Ultima Contabilizacion. Debe deshacer lo amortizado hasta la fecha de Revalorizacion")
                End If
            End If

            If strFiltro.Length > 0 AndAlso dr.RowState = DataRowState.Modified Then
                Dim ClsReval As New ElementoRevalorizacion
                Dim dtResult As DataTable = ClsReval.Filter(, strFiltro, "IdLineaRevalorizacion DESC")
                If Not dtResult Is Nothing AndAlso dtResult.Rows.Count > 0 Then
                    For i As Integer = dtResult.Rows.Count - 1 To 0 Step -1
                        ClsReval.Delete(dtResult.Rows(i))
                    Next
                End If
            End If
            dtNuevoElemReval.Rows(0)("IDLineaRevalorizacion") = AdminData.GetAutoNumeric
            Return dtNuevoElemReval
        End If
    End Function

    <Serializable()> _
    Public Class DataCrearDtElemeReval
        Public Dt As DataTable
        Public FechaReval As Date

        Public Sub New()
        End Sub

        Public Sub New(ByVal Dt As DataTable, ByVal FechaReval As Date)
            Me.Dt = Dt
            Me.FechaReval = FechaReval
        End Sub
    End Class

    <Task()> Public Shared Function CrearDtElemReval(ByVal data As DataCrearDtElemeReval, ByVal services As ServiceProvider) As DataTable
        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
        Dim MonInfoA As MonedaInfo = Monedas.MonedaA
        Dim MonInfoB As MonedaInfo = Monedas.MonedaB

        'Creamos el rcs y cargamos los datos
        Dim dtCambio As DataTable = New ElementoRevalorizacion().AddNew
        Dim drNew As DataRow = dtCambio.NewRow

        drNew("IdElemento") = data.Dt.Rows(0)("IdElemento")
        drNew("ValorCompraFechaA") = xRound(data.Dt.Rows(0)("ValorTotalRevalElementoA"), MonInfoA.NDecimalesImporte)
        drNew("ValorCompraFechaB") = xRound(drNew("ValorCompraFechaA") * MonInfoA.CambioB, MonInfoB.NDecimalesImporte)
        drNew("NFactura") = data.Dt.Rows(0)("NFactura")

        If data.Dt.Rows(0).IsNull("ValorNetoContableElementoA") Then data.Dt.Rows(0)("ValorNetoContableElementoA") = 0
        drNew("ValorNetoFechaA") = xRound(data.Dt.Rows(0)("ValorNetoContableElementoA"), MonInfoA.NDecimalesImporte)
        drNew("ValorNetoFechaB") = xRound(drNew("ValorNetoFechaA") * MonInfoA.CambioB, MonInfoB.NDecimalesImporte)
        If data.Dt.Rows(0).IsNull("ValorAmortizadoElementoA") Then data.Dt.Rows(0)("ValorAmortizadoElementoA") = 0
        drNew("ValorAmortizadoFechaA") = xRound(data.Dt.Rows(0)("ValorAmortizadoElementoA"), MonInfoA.NDecimalesImporte)
        drNew("ValorAmortizadoFechaB") = xRound(drNew("ValorAmortizadoFechaA") * MonInfoA.CambioB, MonInfoB.NDecimalesImporte)
        drNew("ValorResidualFechaA") = xRound(data.Dt.Rows(0)("ValorResidualA"), MonInfoA.NDecimalesImporte)
        drNew("ValorResidualFechaB") = xRound(drNew("ValorResidualFechaA") * MonInfoA.CambioB, MonInfoB.NDecimalesImporte)
        If Length(data.Dt.Rows(0)("ValorTotalPlusvaliaA")) = 0 Then data.Dt.Rows(0)("ValorTotalPlusvaliaA") = 0
        drNew("ValorPlusvaliaFechaA") = xRound(data.Dt.Rows(0)("ValorTotalPlusvaliaA"), MonInfoA.NDecimalesImporte)
        drNew("ValorPlusvaliaFechaB") = xRound(drNew("ValorPlusvaliaFechaA") * MonInfoA.CambioB, MonInfoB.NDecimalesImporte)
        If Length(data.Dt.Rows(0)("ValorNetoContablePlusvaliaA")) = 0 Then data.Dt.Rows(0)("ValorNetoContablePlusvaliaA") = 0
        drNew("ValorNetoContablePlusvaliaFechaA") = xRound(data.Dt.Rows(0)("ValorNetoContablePlusvaliaA"), MonInfoA.NDecimalesImporte)
        drNew("ValorNetoContablePlusvaliaFechaB") = xRound(drNew("ValorNetoContablePlusvaliaFechaA") * MonInfoA.CambioB, MonInfoB.NDecimalesImporte)
        If Length(data.Dt.Rows(0)("ValorAmortizadoPlusvaliaA")) = 0 Then data.Dt.Rows(0)("ValorAmortizadoPlusvaliaA") = 0
        drNew("ValorAmortizadoPlusvaliaFechaA") = xRound(data.Dt.Rows(0)("ValorAmortizadoPlusvaliaA"), MonInfoA.NDecimalesImporte)
        drNew("ValorAmortizadoPlusvaliaFechaB") = xRound(drNew("ValorAmortizadoPlusvaliaFechaA") * MonInfoA.CambioB, MonInfoB.NDecimalesImporte)
        drNew("FechaRevalorizacion") = data.FechaReval
        drNew("MesRevalorizacion") = Month(data.FechaReval)
        drNew("AñoRevalorizacion") = Year(data.FechaReval)
        drNew("VidaUtilFecha") = data.Dt.Rows(0)("VidaContableElemento")
        drNew("IDTipoAmortizacionFecha") = data.Dt.Rows(0)("IDCodigoAmortizacionContable")
        drNew("NFactura") = data.Dt.Rows(0)("NFactura")
        dtCambio.Rows.Add(drNew)
        Return dtCambio
    End Function

    <Task()> Public Shared Function CrearRevalorizacionesElementosFactura(ByVal dtElementos As DataTable, ByVal services As ServiceProvider) As DataTable
        If Not dtElementos Is Nothing AndAlso dtElementos.Rows.Count > 0 Then
            Dim dtResult As DataTable = New ElementoRevalorizacion().AddNew
            For Each dr As DataRow In dtElementos.Select
                Dim drNew As DataRow = dtResult.NewRow
                drNew("IdLineaRevalorizacion") = AdminData.GetAutoNumeric
                drNew("IdElemento") = dr("IdElemento")
                drNew("ValorCompraFechaA") = dr("ValorTotalRevalElementoA")
                drNew("ValorNetoFechaA") = dr("ValorNetoContableElementoA")
                drNew("ValorAmortizadoFechaA") = dr("ValorAmortizadoElementoA")
                drNew("ValorResidualFechaA") = dr("ValorResidualA")
                drNew("ValorCompraFechaB") = dr("ValorTotalRevalElementoB")
                drNew("ValorNetoFechaB") = dr("ValorNetoContableElementoB")
                drNew("ValorAmortizadoFechaB") = dr("ValorAmortizadoElementoB")
                drNew("ValorResidualFechaB") = dr("ValorResidualB")
                drNew("ValorPlusvaliaFechaA") = dr("ValorTotalPlusvaliaA")
                drNew("ValorAmortizadoPlusvaliaFechaA") = dr("ValorAmortizadoPlusvaliaA")
                drNew("ValorNetoContablePlusvaliaFechaA") = dr("ValorNetoContablePlusvaliaA")
                drNew("ValorPlusvaliaFechaB") = dr("ValorTotalPlusvaliaB")
                drNew("ValorAmortizadoPlusvaliaFechaB") = dr("ValorAmortizadoPlusvaliaB")
                drNew("ValorNetoContablePlusvaliaFechaB") = dr("ValorNetoContablePlusvaliaB")
                drNew("FechaRevalorizacion") = dr("FechaInicioContabilizacion")
                drNew("MesRevalorizacion") = CDate(dr("FechaCompra")).Month
                drNew("AñoRevalorizacion") = CDate(dr("FechaCompra")).Year
                drNew("VidaUtilFecha") = dr("VidaContableElemento")
                drNew("DotacionFechaA") = dr("DotacionContableElementoA")
                drNew("DotacionFechaB") = dr("DotacionContableElementoB")
                drNew("IDTipoAmortizacionFecha") = dr("IDCodigoAmortizacionContable")
                drNew("PorcentajeFecha") = dr("PorcentajeAnualContable")
                drNew("NFactura") = dr("NFactura")
                dtResult.Rows.Add(drNew)
            Next
            Return dtResult
        End If
    End Function

    <Task()> Public Shared Sub InsertarElementosAutomatico(ByVal dtElementos As DataTable, ByVal services As ServiceProvider)
        If Not dtElementos Is Nothing AndAlso dtElementos.Rows.Count > 0 Then
            Dim ClsReval As New ElementoRevalorizacion
            Dim dtReval As DataTable = ClsReval.AddNew()
            For Each drElementos As DataRow In dtElementos.Rows
                Dim drReval As DataRow = dtReval.NewRow
                drReval("IdElemento") = drElementos("IdElemento")
                drReval("ValorCompraFechaA") = drElementos("ValorTotalRevalElementoA")
                drReval("ValorNetoFechaA") = drElementos("ValorNetoContableElementoA")
                drReval("ValorAmortizadoFechaA") = drElementos("ValorAmortizadoElementoA")
                drReval("ValorResidualFechaA") = drElementos("ValorResidualA")
                drReval("ValorCompraFechaB") = drElementos("ValorTotalRevalElementoB")
                drReval("ValorNetoFechaB") = drElementos("ValorNetoContableElementoB")
                drReval("ValorAmortizadoFechaB") = drElementos("ValorAmortizadoElementoB")
                drReval("ValorResidualFechaB") = drElementos("ValorResidualB")
                drReval("ValorPlusvaliaFechaA") = drElementos("ValorTotalPlusvaliaA")
                drReval("ValorAmortizadoPlusvaliaFechaA") = drElementos("ValorAmortizadoPlusvaliaA")
                drReval("ValorNetoContablePlusvaliaFechaA") = drElementos("ValorNetoContablePlusvaliaA")
                drReval("ValorPlusvaliaFechaB") = drElementos("ValorTotalPlusvaliaB")
                drReval("ValorAmortizadoPlusvaliaFechaB") = drElementos("ValorAmortizadoPlusvaliaB")
                drReval("ValorNetoContablePlusvaliaFechaB") = drElementos("ValorNetoContablePlusvaliaB")
                drReval("MesRevalorizacion") = Month(drElementos("FechaCompra"))
                drReval("AñoRevalorizacion") = Year(drElementos("FechaCompra"))
                drReval("FechaRevalorizacion") = drElementos("FechaCompra")
                drReval("VidaUtilFecha") = drElementos("VidaContableElemento")
                drReval("DotacionFechaA") = drElementos("DotacionContableElementoA")
                drReval("DotacionFechaB") = drElementos("DotacionContableElementoB")
                drReval("IDTipoAmortizacionFecha") = drElementos("IDCodigoAmortizacionContable")
                drReval("PorcentajeFecha") = drElementos("PorcentajeAnualContable")
                drReval("NFactura") = drElementos("NFactura")
                dtReval.Rows.Add(drReval)
            Next
            ClsReval.Update(dtReval)
        Else : ApplicationService.GenerateError("Error en el proceso.")
        End If
    End Sub

    <Task()> Public Shared Function RevalorizacionModificableBorrable(ByVal IntIDElemReval As Integer, ByVal services As ServiceProvider) As Boolean
        Dim DtReval As DataTable = New BE.DataEngine().Filter("vElementoRevalorizacionAmortizado", New FilterItem("IDLineaRevalorizacion", FilterOperator.Equal, IntIDElemReval, FilterType.Numeric), , "IDElemento, IDLineaRevalorizacion")
        If Not DtReval Is Nothing AndAlso DtReval.Rows.Count > 0 Then
            If Nz(DtReval.Rows(0)("TotalAmortizadoA"), 0) = 0 Then
                Return True
            Else : Return False
            End If
        End If
    End Function

    <Task()> Public Shared Function RevalorizacionNueva(ByVal StrIDElemento As String, ByVal services As ServiceProvider) As Boolean
        Dim DtReval As DataTable = New BE.DataEngine().Filter("vElementoRevalorizacionAmortizado", New FilterItem("IDElemento", FilterOperator.Equal, StrIDElemento, FilterType.String), , "IDElemento, IDLineaRevalorizacion")
        If Not DtReval Is Nothing AndAlso DtReval.Rows.Count > 0 Then
            For Each drReval As DataRow In DtReval.Select
                If Nz(drReval("TotalAmortizadoA"), 0) = 0 Then
                    Return False
                End If
            Next
            Return True
        Else : Return True
        End If
    End Function

    <Serializable()> _
    Public Class DataActualizarCondElem
        Public DtElem As DataTable
        Public CodContable As String
        Public CodTecnica As String
        Public CodFiscal As String
        Public PorMeses As Boolean

        Public Sub New()
        End Sub

        Public Sub New(ByVal DtElem As DataTable, ByVal CodContable As String, ByVal CodTecnica As String, ByVal CodFiscal As String, ByVal PorMeses As Boolean)
            Me.DtElem = DtElem
            Me.CodContable = CodContable
            Me.CodTecnica = CodTecnica
            Me.CodFiscal = CodFiscal
            Me.PorMeses = PorMeses
        End Sub
    End Class

    <Task()> Public Shared Sub ActualizarCondicionesElemento(ByVal data As DataActualizarCondElem, ByVal services As ServiceProvider)
        For Each Dr As DataRow In data.DtElem.Select
            If ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf RevalorizacionNueva, Dr("IDElemento"), services) Then
                Dr("IDCodigoAmortizacionContable") = data.CodContable
                If Length(data.CodTecnica) > 0 Then
                    Dr("IDCodigoAmortizacionTecnica") = data.CodTecnica
                End If
                If Length(data.CodFiscal) > 0 Then
                    Dr("IDCodigoAmortizacionFiscal") = data.CodFiscal
                End If
                Dr("FechaUltimaRevalorizacion") = Dr("FechaUltimaContabilizacion")
                Dr("PorMeses") = data.PorMeses
            End If
        Next
        Dim ClsElemAmort As New ElementoAmortizable
        ClsElemAmort.Update(data.DtElem)
    End Sub

    <Task()> Public Shared Sub BorrarRevalorizaciones(ByVal StrElem As String, ByVal services As ServiceProvider)
        Dim ClsReval As New ElementoRevalorizacion
        Dim DtReval As DataTable = ClsReval.Filter(New FilterItem("IDElemento", FilterOperator.Equal, StrElem, FilterType.String), "FechaRevalorizacion DESC")
        If Not DtReval Is Nothing AndAlso DtReval.Rows.Count > 1 Then
            ClsReval.Delete(DtReval.Rows(0))
        End If
    End Sub

    <Task()> Public Shared Sub BorrarRevalorizacionesArray(ByVal StrElem() As String, ByVal services As ServiceProvider)
        Dim ClsReval As New ElementoRevalorizacion
        For i As Integer = 0 To StrElem.Length - 1
            Dim DtReval As DataTable = ClsReval.Filter(New FilterItem("IDElemento", FilterOperator.Equal, StrElem(i), FilterType.String), "FechaRevalorizacion DESC")
            If Not DtReval Is Nothing AndAlso DtReval.Rows.Count > 1 Then
                ClsReval.Delete(DtReval.Rows(0))
            End If
        Next
    End Sub

#End Region

End Class