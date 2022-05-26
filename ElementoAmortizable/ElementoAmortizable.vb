Public Class ElementoAmortizable

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroElementoAmortizable"

#End Region

    <Serializable()> _
    Public Class StDatosAmort
        Public CodAmort As String
        Public Vida As Integer
        Public Porcentaje As Integer
    End Class

    <Serializable()> _
    Public Class StMesAño
        Public Mes As Integer
        Public Año As Integer
    End Class

#Region "Eventos RegisterDeleteTasks"

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf Comunes.BeginTransaction)
        deleteProcess.AddTask(Of DataRow)(AddressOf ComprobarAmortizaciones)
        deleteProcess.AddTask(Of DataRow)(AddressOf ComprobarFacturaCompra)
        deleteProcess.AddTask(Of DataRow)(AddressOf ComprobarRevalorizaciones)
        deleteProcess.AddTask(Of DataRow)(AddressOf ComprobarUbicaciones)
        deleteProcess.AddTask(Of DataRow)(AddressOf ComprobarAnalitica)
        deleteProcess.AddTask(Of DataRow)(AddressOf ComprobarElementoSubvencion)
        'deleteProcess.AddTask(Of DataRow)(AddressOf Comunes.DeleteEntityRow)
    End Sub

    <Task()> Public Shared Sub ComprobarAmortizaciones(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim DtAmortizacion As DataTable = New AmortizacionRegistro().Filter(New FilterItem("IDElemento", FilterOperator.Equal, data("IDElemento")))
        If Not DtAmortizacion Is Nothing AndAlso DtAmortizacion.Rows.Count > 0 Then
            ApplicationService.GenerateError("No se puede eliminar un elemento con amortizaciones.")
        End If
    End Sub

    <Task()> Public Shared Sub ComprobarFacturaCompra(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim DtLineasActualizar As DataTable
        Dim DtElementoFCL As DataTable = New ElementoAmortizableFCL().Filter(New FilterItem("IDElemento", FilterOperator.Equal, data("IDElemento")))
        For Each Dr As DataRow In DtElementoFCL.Select
            Dim DtFCLin As DataTable = New FacturaCompraLinea().SelOnPrimaryKey(Dr("IdLineaFactura"))
            If Not DtFCLin Is Nothing AndAlso DtFCLin.Rows.Count > 0 Then
                If DtLineasActualizar Is Nothing Then
                    DtLineasActualizar = DtFCLin.Clone()
                End If
                DtFCLin.Rows(0)("EstadoInmovilizado") = False
                DtLineasActualizar.ImportRow(DtFCLin.Rows(0))
            End If
        Next
        If Not DtLineasActualizar Is Nothing Then BusinessHelper.UpdateTable(DtLineasActualizar)
    End Sub

    <Task()> Public Shared Sub ComprobarRevalorizaciones(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim ClsReval As New ElementoRevalorizacion
        Dim DtReval As DataTable = ClsReval.Filter(New FilterItem("IDElemento", FilterOperator.Equal, data("IDElemento")))
        If Not DtReval Is Nothing AndAlso DtReval.Rows.Count > 0 Then
            ClsReval.Delete(DtReval)
        End If
    End Sub

    <Task()> Public Shared Sub ComprobarUbicaciones(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim ClsUbicacion As New ElementoAmortizUbicacion
        Dim DtUbicacion As DataTable = ClsUbicacion.Filter(New FilterItem("IDElemento", FilterOperator.Equal, data("IDElemento")))
        If Not DtUbicacion Is Nothing AndAlso DtUbicacion.Rows.Count > 0 Then
            ClsUbicacion.Delete(DtUbicacion)
        End If
    End Sub

    <Task()> Public Shared Sub ComprobarAnalitica(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim ClsElemAmortAnal As New ElementoAmortizAnalitica
        Dim DtElemAmortAnal As DataTable = ClsElemAmortAnal.Filter(New FilterItem("IDElemento", FilterOperator.Equal, data("IDElemento")))
        If Not DtElemAmortAnal Is Nothing AndAlso DtElemAmortAnal.Rows.Count > 0 Then
            ClsElemAmortAnal.Delete(DtElemAmortAnal)
        End If
    End Sub

    <Task()> Public Shared Sub ComprobarElementoSubvencion(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim DtElemento As DataTable = New BE.DataEngine().Filter("frmMntoElemAmortizables", New FilterItem("IDElemento", FilterOperator.Equal, data("IDElemento")))
        If Not DtElemento Is Nothing AndAlso DtElemento.Rows.Count > 0 Then
            Dim ClsElSubvencion As New ElementoSubvencion
            Dim DtSubvencion As New DataTable
            Dim EsSubvencion As Boolean
            If Not DtElemento.Rows(0).IsNull("Subvencion") AndAlso _
                DtElemento.Rows(0)("Subvencion") = True Then
                EsSubvencion = True
            End If
            If EsSubvencion Then    'IDElemento hace de enlace e IdSubvencion almacena datos
                DtSubvencion = ClsElSubvencion.Filter(New StringFilterItem("IDElemento", data("IDElemento")))
            Else                    'IdSubvencion hace de enlace e IDElemento almacena datos
                DtSubvencion = ClsElSubvencion.Filter(New StringFilterItem("IDSubvencion", data("IDElemento")))
            End If
            If Not DtSubvencion Is Nothing AndAlso DtSubvencion.Rows.Count > 0 Then
                For Each dr As DataRow In DtSubvencion.Select
                    dr.Delete()
                Next
                ClsElSubvencion.Delete(DtSubvencion)
            End If
        End If
    End Sub

#End Region

#Region "Eventos RegisterAddNewTasks"

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
        addnewProcess.AddTask(Of DataRow)(AddressOf AsignarCentroGestion)
        addnewProcess.AddTask(Of DataRow)(AddressOf AsignarContadorAdd)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        data("ValorTotalElementoA") = 0
        data("ValorTotalRevalElementoA") = 0
        data("ValorAmortizadoElementoA") = 0
        data("ValorResidualA") = 0
        data("ValorTotalPlusvaliaA") = 0
        data("ValorNetoContablePlusvaliaA") = 0
        data("ValorAmortizadoPlusvaliaA") = 0
        data("ValorUltimoContabilizadoA") = 0
        data("ValorReposicionA") = 0
        data("ValorReposicionActualizadoA") = 0
        data("Baja") = 0
        data("IDActivo") = New Parametro().ActivoPredeterminado()
    End Sub

    <Task()> Public Shared Sub AsignarCentroGestion(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim cgu As New UsuarioCentroGestion.UsuarioCentroGestionInfo
        cgu = ProcessServer.ExecuteTask(Of UsuarioCentroGestion.UsuarioCentroGestionInfo, UsuarioCentroGestion.UsuarioCentroGestionInfo)(AddressOf UsuarioCentroGestion.ObtenerUsuarioCentroGestion, cgu, services)
        data("IDCentroGestion") = cgu.IDCentroGestion
    End Sub

    <Task()> Public Shared Sub AsignarContadorAdd(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim DtContadores As DataTable = ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf Contador.CounterDt, "ElementoAmortizable", services)
        Dim DtContadorPred As DataTable = ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf Contador.CounterDefault, "ElementoAmortizable", services)
        Dim StrContador As String = String.Empty
        If Not DtContadorPred Is Nothing AndAlso DtContadorPred.Rows.Count > 0 Then
            StrContador = DtContadorPred.Rows(0)("IDContador")
        End If
        If StrContador.Length > 0 Then
            Dim Dr As DataRow() = DtContadores.Select("IDContador = '" & StrContador & "'")
            If Not Dr Is Nothing Then
                data("IDElemento") = Dr(0)("ValorProvisional")
                data("IDContador") = Dr(0)("IDContador")
            End If
        End If
    End Sub

#End Region

#Region "Eventos GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBrl As New BusinessRules
        oBrl.Add("IDGrupoAmortizacion", AddressOf CambioGrupoAmortizacion)
        oBrl.Add("ValorTotalRevalElementoA", AddressOf CambioTotalReval)
        oBrl.Add("ValorAmortizadoElementoA", AddressOf CambioAmortizado)
        oBrl.Add("ValorResidualA", AddressOf CambioValorResidualA)
        oBrl.Add("ValorResidualB", AddressOf CambioValorResidualB)
        oBrl.Add("IDElementoOrigen", AddressOf CambioElementoOrigen)
        Return oBrl
    End Function

    <Task()> Public Shared Sub CambioGrupoAmortizacion(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If data.Context("Nuevo") = True Then
            If Length(data.Current("IDGrupoAmortizacion")) > 0 Then
                Dim dtGrupoAmortiz As DataTable = New GrupoAmortizacion().SelOnPrimaryKey(data.Current("IDGrupoAmortizacion"))
                If dtGrupoAmortiz.Rows.Count > 0 Then
                    data.Current("IDCodigoAmortizacionContable") = dtGrupoAmortiz.Rows(0)("IDTipoAmortiz")
                    data.Current("IDCodigoAmortizacionTecnica") = dtGrupoAmortiz.Rows(0)("IDTipoAmortiz")
                    data.Current("IDCodigoAmortizacionFiscal") = dtGrupoAmortiz.Rows(0)("IDTipoAmortiz")
                End If
            Else
                data.Current("IDCodigoAmortizacionContable") = System.DBNull.Value
                data.Current("IDCodigoAmortizacionTecnica") = System.DBNull.Value
                data.Current("IDCodigoAmortizacionFiscal") = System.DBNull.Value
            End If
        End If
        If Length(data.Value) > 0 Then
            Dim drGrupoAmortiz As DataRow = New GrupoAmortizacion().GetItemRow(data.Value)
            If data.Current.ContainsKey("Subvencion") Then data.Current("Subvencion") = drGrupoAmortiz("Subvencion")
        Else
            If data.Current.ContainsKey("Subvencion") Then data.Current("Subvencion") = System.DBNull.Value
        End If
    End Sub

    <Task()> Public Shared Sub CambioTotalReval(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        Dim StValor As New DataCalcValorNetoContable(data.Value, data.Current("ValorAmortizadoElementoA"), data.Current("ValorResidualA"))
        data.Current("ValorNetoContableElementoA") = ProcessServer.ExecuteTask(Of DataCalcValorNetoContable, Double)(AddressOf CalcularValorNetoContable, StValor, services)
    End Sub

    <Task()> Public Shared Sub CambioAmortizado(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        Dim StValor As New DataCalcValorNetoContable(data.Current("ValorTotalRevalElementoA"), data.Value, data.Current("ValorResidualA"))
        data.Current("ValorNetoContableElementoA") = ProcessServer.ExecuteTask(Of DataCalcValorNetoContable, Double)(AddressOf CalcularValorNetoContable, StValor, services)
    End Sub

    <Task()> Public Shared Sub CambioValorResidualA(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        Dim StValor As New DataCalcValorNetoContable(data.Current("ValorTotalRevalElementoA"), data.Current("ValorAmortizadoElementoA"), data.Value)
        data.Current("ValorNetoContableElementoA") = ProcessServer.ExecuteTask(Of DataCalcValorNetoContable, Double)(AddressOf CalcularValorNetoContable, StValor, services)
        If data.Value Is System.DBNull.Value Then
            data.Current("ValorResidualB") = 0
        Else
            Dim StResidual As New DataCalcValorResidualMoneda(data.ColumnName, data.Value)
            data.Current("ValorResidualB") = ProcessServer.ExecuteTask(Of DataCalcValorResidualMoneda, Double)(AddressOf CalcularValorResidualMoneda, StResidual, services)
        End If
    End Sub

    <Task()> Public Shared Sub CambioValorResidualB(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If data.Value Is System.DBNull.Value Then
            data.Current("ValorResidualA") = 0
        Else
            Dim StResidual As New DataCalcValorResidualMoneda(data.ColumnName, data.Value)
            data.Current("ValorResidualA") = ProcessServer.ExecuteTask(Of DataCalcValorResidualMoneda, Double)(AddressOf CalcularValorResidualMoneda, StResidual, services)
        End If
    End Sub

    <Task()> Public Shared Sub CambioElementoOrigen(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If data.Context.ContainsKey("CheckCargarDesc") Then
            If data.Context("CheckCargarDesc") = True Then
                If data.Value Is System.DBNull.Value Then
                    data.Current("DescElemento") = String.Empty
                Else
                    Dim dr As DataRow = New ElementoAmortizable().GetItemRow(data.Value)
                    data.Current("DescElemento") = dr("DescElemento")
                End If
            End If
        End If
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf ValidarFechaInicioContabilizacion)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarContador)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarDatosAmortizaciones)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarDatosTotales)
        updateProcess.AddTask(Of DataRow)(AddressOf Comunes.UpdateEntityRow)
        updateProcess.AddTask(Of DataRow)(AddressOf Comunes.MarcarComoActualizado)
        updateProcess.AddTask(Of DataRow)(AddressOf InsertarEntidadesSecundarias)
        updateProcess.AddTask(Of DataRow)(AddressOf CambiosSubvenciones)
    End Sub

    <Task()> Public Shared Sub ValidarFechaInicioContabilizacion(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("FechaInicioContabilizacion")) = 0 AndAlso Length(data("FechaCompra")) > 0 Then
            data("FechaInicioContabilizacion") = data("FechaCompra")
        End If
    End Sub

    <Task()> Public Shared Sub AsignarContador(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If New Parametro().ObtenerInmovAuto() = 1 Then
                If Length(data("IdContador")) > 0 Then
                    data("IdElemento") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, data("IDContador"), services)
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarDatosAmortizaciones(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim InsertarRegistroRevalorizacion As Boolean
        If data.RowState = DataRowState.Added Then
            InsertarRegistroRevalorizacion = True
            data("FechaUltimaRevalorizacion") = data("FechaInicioContabilizacion")
            Dim StAmortCont As New DataGetVidaPorcenDot(data.Table, data("IDCodigoAmortizacionContable").ToString(), data("FechaInicioContabilizacion"))
            Dim DtAmortCont As DataTable = ProcessServer.ExecuteTask(Of DataGetVidaPorcenDot, DataTable)(AddressOf GetVidaPorcentajeDotacion, StAmortCont, services)
            If Not DtAmortCont Is Nothing AndAlso DtAmortCont.Rows.Count > 0 Then
                data("VidaContableElemento") = DtAmortCont.Rows(0)("Vida")
                data("PorcentajeAnualContable") = DtAmortCont.Rows(0)("Porcentaje")
                data("DotacionContableElementoA") = DtAmortCont.Rows(0)("Dotacion")
            End If
            Dim StAmortTec As New DataGetVidaPorcenDot(data.Table, data("IDCodigoAmortizacionTecnica").ToString())
            Dim DtAmortTec As DataTable = ProcessServer.ExecuteTask(Of DataGetVidaPorcenDot, DataTable)(AddressOf GetVidaPorcentajeDotacion, StAmortTec, services)
            If Not DtAmortTec Is Nothing AndAlso DtAmortTec.Rows.Count > 0 Then
                data("VidaTecnicaElemento") = DtAmortTec.Rows(0)("Vida")
                data("PorcentajeAnualTecnico") = DtAmortTec.Rows(0)("Porcentaje")
                data("DotacionTecnicaElementoA") = DtAmortTec.Rows(0)("Dotacion")
            End If
            Dim StAmortFis As New DataGetVidaPorcenDot(data.Table, data("IDCodigoAmortizacionFiscal").ToString())
            Dim DtAmortFis As DataTable = ProcessServer.ExecuteTask(Of DataGetVidaPorcenDot, DataTable)(AddressOf GetVidaPorcentajeDotacion, StAmortFis, services)
            If Not DtAmortFis Is Nothing AndAlso DtAmortFis.Rows.Count > 0 Then
                data("VidaFiscalElemento") = DtAmortFis.Rows(0)("Vida")
                data("PorcentajeAnualFiscal") = DtAmortFis.Rows(0)("Porcentaje")
                data("DotacionFiscalElementoA") = DtAmortFis.Rows(0)("Dotacion")
            End If
            Dim StData As New DataObtenerDotPrimerMes(data.Table, enTipoAmort.enContable)
            data("DotacionContableElementoA") = ProcessServer.ExecuteTask(Of DataObtenerDotPrimerMes, Double)(AddressOf ObtenerDotacionPrimerMes, StData, services)
            StData.Tipo = enTipoAmort.enTecnica
            data("DotacionTecnicaElementoA") = ProcessServer.ExecuteTask(Of DataObtenerDotPrimerMes, Double)(AddressOf ObtenerDotacionPrimerMes, StData, services)
            StData.Tipo = enTipoAmort.enFiscal
            data("DotacionFiscalElementoA") = ProcessServer.ExecuteTask(Of DataObtenerDotPrimerMes, Double)(AddressOf ObtenerDotacionPrimerMes, StData, services)
        ElseIf data.RowState = DataRowState.Modified Then
            If Not ComparaDr(data, "FechaInicioContabilizacion") Or _
                Not ComparaDr(data, "ValorTotalRevalElementoA") Or _
                Not ComparaDr(data, "ValorResidualA") Or _
                Not ComparaDr(data, "ValorTotalPlusvaliaA") Or _
                Not ComparaDr(data, "IDCodigoAmortizacionContable") Then
                InsertarRegistroRevalorizacion = True
                If Not ComparaDr(data, "FechaInicioContabilizacion") Then
                    data("FechaUltimaRevalorizacion") = data("FechaInicioContabilizacion")
                End If
            End If
            If Not ComparaDr(data, "ValorAmortizadoElementoA") Or _
                Not ComparaDr(data, "ValorAmortizadoPlusvaliaA") Then
                InsertarRegistroRevalorizacion = True
            End If
        End If
        If InsertarRegistroRevalorizacion Then
            Dim StAmortCont As New DataGetVidaPorcenDot(data.Table, data("IDCodigoAmortizacionContable"), data("FechaUltimaRevalorizacion"))
            Dim DtAmortCont As DataTable = ProcessServer.ExecuteTask(Of DataGetVidaPorcenDot, DataTable)(AddressOf GetVidaPorcentajeDotacion, StAmortCont, services)
            If Not DtAmortCont Is Nothing AndAlso DtAmortCont.Rows.Count > 0 Then
                data("VidaContableElemento") = DtAmortCont.Rows(0)("Vida")
                data("PorcentajeAnualContable") = DtAmortCont.Rows(0)("porcentaje")
                data("DotacionContableElementoA") = DtAmortCont.Rows(0)("Dotacion")
            End If
            Dim StAmortTec As New DataGetVidaPorcenDot(data.Table, data("IDCodigoAmortizacionTecnica"))
            Dim DtAmortTec As DataTable = ProcessServer.ExecuteTask(Of DataGetVidaPorcenDot, DataTable)(AddressOf GetVidaPorcentajeDotacion, StAmortTec, services)
            If Not DtAmortTec Is Nothing AndAlso DtAmortTec.Rows.Count > 0 Then
                data("VidaTecnicaElemento") = DtAmortTec.Rows(0)("Vida")
                data("PorcentajeAnualTecnico") = DtAmortTec.Rows(0)("Porcentaje")
                data("DotacionTecnicaElementoA") = DtAmortTec.Rows(0)("Dotacion")
            End If
            Dim StAmortFis As New DataGetVidaPorcenDot(data.Table, data("IDCodigoAmortizacionFiscal"))
            Dim DtAmortFis As DataTable = ProcessServer.ExecuteTask(Of DataGetVidaPorcenDot, DataTable)(AddressOf GetVidaPorcentajeDotacion, StAmortFis, services)
            If Not DtAmortFis Is Nothing AndAlso DtAmortFis.Rows.Count > 0 Then
                data("VidaFiscalElemento") = DtAmortFis.Rows(0)("Vida")
                data("PorcentajeAnualFiscal") = DtAmortFis.Rows(0)("Porcentaje")
                data("DotacionFiscalElementoA") = DtAmortFis.Rows(0)("Dotacion")
            End If
            data("IDEjercicio") = data("IDEjercicio")
            data("NAsiento") = data("NAsiento")
        End If
    End Sub

    <Task()> Public Shared Sub AsignarDatosTotales(ByVal Data As DataRow, ByVal services As ServiceProvider)
        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
        Dim MonInfoA As MonedaInfo = Monedas.MonedaA
        Dim MonInfoB As MonedaInfo = Monedas.MonedaB

        Data("ValorTotalElementoB") = xRound(Data("ValorTotalElementoA") * MonInfoA.CambioB, MonInfoB.NDecimalesImporte)
        Data("ValorTotalRevalElementoB") = xRound(Data("ValorTotalRevalElementoA") * MonInfoA.CambioB, MonInfoB.NDecimalesImporte)
        If Length(Data("ValorAmortizadoElementoA")) > 0 Then Data("ValorAmortizadoElementoB") = xRound(Data("ValorAmortizadoElementoA") * MonInfoA.CambioB, MonInfoB.NDecimalesImporte)
        If Length(Data("ValorNetoContableElementoA")) > 0 Then Data("ValorNetoContableElementoB") = xRound(Data("ValorNetoContableElementoA") * MonInfoA.CambioB, MonInfoB.NDecimalesImporte)
        If Length(Data("ValorTotalPlusvaliaA")) > 0 Then Data("ValorTotalPlusvaliaB") = xRound(Data("ValorTotalPlusvaliaA") * MonInfoA.CambioB, MonInfoB.NDecimalesImporte)
        If Length(Data("ValorAmortizadoPlusvaliaA")) > 0 Then Data("ValorAmortizadoPlusvaliaB") = xRound(Data("ValorAmortizadoPlusvaliaA") * MonInfoA.CambioB, MonInfoB.NDecimalesImporte)
        If Length(Data("ValorNetoContablePlusvaliaA")) > 0 Then Data("ValorNetoContablePlusvaliaB") = xRound(Data("ValorNetoContablePlusvaliaA") * MonInfoA.CambioB, MonInfoB.NDecimalesImporte)
        If Length(Data("ValorUltimoContabilizadoA")) > 0 Then
            Data("ValorUltimoContabilizadoB") = xRound(Data("ValorUltimoContabilizadoA") * MonInfoA.CambioB, MonInfoB.NDecimalesImporte)
        Else : Data("ValorUltimoContabilizadoB") = System.DBNull.Value
        End If
        Data("ValorResidualB") = xRound(Data("ValorResidualA") * MonInfoA.CambioB, MonInfoB.NDecimalesImporte)
        If Length(Data("ValorReposicionA")) > 0 Then
            Data("ValorReposicionB") = xRound(Data("ValorReposicionA") * MonInfoA.CambioB, MonInfoB.NDecimalesImporte)
        Else : Data("ValorReposicionB") = System.DBNull.Value
        End If
        If Length(Data("ValorReposicionActualizadoA")) > 0 Then
            Data("ValorReposicionActualizadoB") = xRound(Data("ValorReposicionActualizadoA") * MonInfoA.CambioB, MonInfoB.NDecimalesImporte)
        Else : Data("ValorReposicionActualizadoB") = System.DBNull.Value
        End If
        'Mes y Año Inicio se corresponden con el dia uno del primer mes completo desde
        'la fecha de compra.
        If Length(Data("FechaInicioContabilizacion")) > 0 Then
            Data("MesInicioContabilizado") = 0
            Data("AñoInicioContabilizado") = 0
            Dim dteFechaInicio As Date
            If CDate(Data("FechaInicioContabilizacion")).Day = 1 Then
                Data("MesInicioContabilizado") = CDate(Data("FechaInicioContabilizacion")).Month
                Data("AñoInicioContabilizado") = CDate(Data("FechaInicioContabilizacion")).Year
            Else
                dteFechaInicio = CDate(Data("FechaInicioContabilizacion")).AddMonths(1)
                Data("MesInicioContabilizado") = dteFechaInicio.Month
                Data("AñoInicioContabilizado") = dteFechaInicio.Year
            End If
        End If
        'Mes y Año que corresponden con la FechaUltimaContabilizacion
        If Not Data.IsNull("FechaUltimaContabilizacion") Then
            Data("MesUltimoContabilizado") = Month(Data("FechaUltimaContabilizacion"))
            Data("AñoUltimoContabilizado") = Year(Data("FechaUltimaContabilizacion"))
        Else
            Data("MesUltimoContabilizado") = 0
            Data("AñoUltimoContabilizado") = 0
        End If
    End Sub

    <Task()> Public Shared Sub InsertarEntidadesSecundarias(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim DblValorAmort As Double
        Dim DblValorAmortPlus As Double
        If Length(data("ValorAmortizadoElementoA")) = 0 Then
            DblValorAmort = 0
        Else : DblValorAmort = data("ValorAmortizadoElementoA")
        End If
        If Length(data("ValorAmortizadoPlusvaliaA")) = 0 Then
            DblValorAmortPlus = 0
        Else : DblValorAmortPlus = data("ValorAmortizadoPlusvaliaA")
        End If
        Dim InsertarRegistroRevalorizacion As Boolean
        Dim InsertarAmortizacion As Boolean
        If data.RowState = DataRowState.Added Then
            InsertarRegistroRevalorizacion = True
            If Length(data("FechaUltimaContabilizacion")) > 0 And (DblValorAmort > 0 Or DblValorAmortPlus > 0) Then InsertarAmortizacion = True
        Else
            If Not ComparaDr(data, "FechaInicioContabilizacion") Or _
                Not ComparaDr(data, "ValorTotalRevalElementoA") Or _
                Not ComparaDr(data, "ValorResidualA") Or _
                Not ComparaDr(data, "ValorTotalPlusvaliaA") Or _
                Not ComparaDr(data, "IDCodigoAmortizacionContable") Then
                InsertarRegistroRevalorizacion = True
            End If
            If Not ComparaDr(data, "ValorAmortizadoElementoA") Or _
                Not ComparaDr(data, "ValorAmortizadoPlusvaliaA") Then
                InsertarRegistroRevalorizacion = True
                InsertarAmortizacion = True
            End If
            If Not ComparaDr(data, "FechaUltimaContabilizacion") Then InsertarAmortizacion = True
        End If
        If data.RowState = DataRowState.Added Then
            Dim AppParams As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
            If AppParams.Analitica.AplicarAnalitica AndAlso AppParams.Analitica.AnaliticaCentroGestion Then
                'Analitica por centro de gestión
                Dim DtAnalitica As DataTable = ProcessServer.ExecuteTask(Of DataRow, DataTable)(AddressOf InsertarAnaliticaCentroGestion, data, services)
                If Not DtAnalitica Is Nothing AndAlso DtAnalitica.Rows.Count > 0 Then
                    Dim ClsElemAnEl As New ElementoAmortizAnalitica
                    Dim DtAnaliticaNew As DataTable = ClsElemAnEl.AddNew
                    For Each DrAn As DataRow In DtAnalitica.Select
                        DtAnaliticaNew.ImportRow(DrAn)
                    Next
                    If Not DtAnaliticaNew Is Nothing AndAlso DtAnaliticaNew.Rows.Count > 0 Then
                        ClsElemAnEl.Update(DtAnaliticaNew)
                    End If
                End If
            End If
        End If
        If InsertarRegistroRevalorizacion Then
            Dim DtReval As DataTable = ProcessServer.ExecuteTask(Of DataRow, DataTable)(AddressOf ElementoRevalorizacion.ActualizarLineasRevalorizacion, data, services)
            If Not DtReval Is Nothing AndAlso DtReval.Rows.Count > 0 Then
                Dim ClsElemReval As New ElementoRevalorizacion
                Dim DtRevalNew As DataTable = ClsElemReval.AddNew
                For Each drReval As DataRow In DtReval.Rows
                    DtRevalNew.ImportRow(drReval)
                Next
                If Not DtRevalNew Is Nothing AndAlso DtRevalNew.Rows.Count > 0 Then
                    ClsElemReval.Update(DtRevalNew)
                End If
            End If
        End If
        If InsertarAmortizacion Then
            If data.RowState = DataRowState.Modified And Length(data("FechaUltimaContabilizacion")) = 0 Then
                data("MesUltimoContabilizado") = 0
                data("AñoUltimoContabilizado") = 0
            End If
            Dim DtAmort As DataTable = ProcessServer.ExecuteTask(Of DataRow, DataTable)(AddressOf AmortizacionRegistro.InsertAmortizacionRegistro, data, services)
            If Not DtAmort Is Nothing AndAlso DtAmort.Rows.Count > 0 Then
                Dim ClsAmortizacion As New AmortizacionRegistro
                Dim DtAmortNew As DataTable = ClsAmortizacion.AddNew
                For Each drAm As DataRow In DtAmort.Select
                    DtAmortNew.ImportRow(drAm)
                Next
                If Not DtAmortNew Is Nothing AndAlso DtAmortNew.Rows.Count > 0 Then
                    ClsAmortizacion.Update(DtAmortNew)
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambiosSubvenciones(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Modified AndAlso data("IDGrupoAmortizacion", DataRowVersion.Original) <> data("IDGrupoAmortizacion") Then
            Dim StData As New DataCambioSubvencionGrupoAmort(data("IDElemento"), data("IDGrupoAmortizacion", DataRowVersion.Original), data("IDGrupoAmortizacion"))
            If ProcessServer.ExecuteTask(Of DataCambioSubvencionGrupoAmort, Boolean)(AddressOf EsCambioSubvencionGrupoAmortizacion, StData, services) Then
                Dim DtElemSubv As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf CrearDtCambiosSubvencionBienPpal, Nothing, services)
                Dim DrElemNew As DataRow = DtElemSubv.NewRow
                DrElemNew("IDElemento") = data("IDElemento")
                DrElemNew("IDGrupoAmortizacionOld") = data("IDGrupoAmortizacion", DataRowVersion.Original)
                DrElemNew("IDGrupoAmortizacionNew") = data("IDGrupoAmortizacion")
                DtElemSubv.Rows.Add(DrElemNew)
                For Each DrCambio As DataRow In DtElemSubv.Select
                    Dim StAmort As New DataCambioGrupoAmort(DrCambio("IDElemento"), DrCambio("IDGrupoAmortizacionOld"), DrCambio("IDGrupoAmortizacionNew"))
                    ProcessServer.ExecuteTask(Of DataCambioGrupoAmort)(AddressOf CambioSubvencionGrupoAmortizacion, StAmort, services)
                Next
            End If
        End If
    End Sub

#Region " Cambio Grupo Amortización "

    <Task()> Public Shared Function CrearDtCambiosSubvencionBienPpal(ByVal data As Object, ByVal services As ServiceProvider) As DataTable
        Dim DtNew As New DataTable
        DtNew.Columns.Add("IDElemento")
        DtNew.Columns.Add("IDGrupoAmortizacionOld")
        DtNew.Columns.Add("IDGrupoAmortizacionNew")
        Return DtNew
    End Function

    <Serializable()> _
    Public Class DataCambioSubvencionGrupoAmort
        Public IDElemento As String
        Public IDGrupoOld As String
        Public IDGrupoNew As String
        Public SubvencionGrupoOld As Boolean

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDElemento As String, ByVal IDGrupoOld As String, ByVal IDGrupoNew As String, Optional ByVal SubvencionGrupoOld As Boolean = False)
            Me.IDElemento = IDElemento
            Me.IDGrupoOld = IDGrupoOld
            Me.IDGrupoNew = IDGrupoNew
            Me.SubvencionGrupoOld = SubvencionGrupoOld
        End Sub
    End Class

    <Task()> Public Shared Function EsCambioSubvencionGrupoAmortizacion(ByVal data As DataCambioSubvencionGrupoAmort, ByVal services As ServiceProvider) As Boolean
        Dim blnSubvencionGrupoNew As Boolean
        Dim blnEsCambio As Boolean = False
        If Length(data.IDGrupoOld) > 0 Then
            Dim objNegGrupoAmortiz As New GrupoAmortizacion
            Dim dtGrupoOld As DataTable = objNegGrupoAmortiz.SelOnPrimaryKey(data.IDGrupoOld)
            Dim dtGrupoNew As DataTable = objNegGrupoAmortiz.SelOnPrimaryKey(data.IDGrupoNew)
            If Not IsNothing(dtGrupoOld) AndAlso dtGrupoOld.Rows.Count > 0 Then data.SubvencionGrupoOld = Nz(dtGrupoOld.Rows(0)("Subvencion"), False)
            If Not IsNothing(dtGrupoNew) AndAlso dtGrupoNew.Rows.Count > 0 Then blnSubvencionGrupoNew = Nz(dtGrupoNew.Rows(0)("Subvencion"), False)
            If data.SubvencionGrupoOld <> blnSubvencionGrupoNew Then blnEsCambio = True
        End If
        Return blnEsCambio
    End Function

    <Serializable()> _
    Public Class DataCambioGrupoAmort
        Public IDElemento As String
        Public IDGrupoOld As String
        Public IDGrupoNew As String

        Public Sub New(ByVal IDElemento As String, ByVal IDGrupoOld As String, ByVal IDGrupoNew As String)
            Me.IDElemento = IDElemento
            Me.IDGrupoOld = IDGrupoOld
            Me.IDGrupoNew = IDGrupoNew
        End Sub
    End Class

    <Task()> Public Shared Sub CambioSubvencionGrupoAmortizacion(ByVal data As DataCambioGrupoAmort, ByVal services As ServiceProvider)
        Dim StData As New DataCambioSubvencionGrupoAmort(data.IDElemento, data.IDGrupoOld, data.IDGrupoNew, False)
        If ProcessServer.ExecuteTask(Of DataCambioSubvencionGrupoAmort, Boolean)(AddressOf EsCambioSubvencionGrupoAmortizacion, StData, services) Then
            Dim objFilter As New Filter
            If StData.SubvencionGrupoOld Then
                objFilter.Add("IDSubvencion", FilterOperator.Equal, data.IDElemento)
            Else : objFilter.Add("IDElemento", FilterOperator.Equal, data.IDElemento)
            End If
            Dim objNegElementoSubvencion As New ElementoSubvencion
            Dim dtElementoSubvencion As DataTable = objNegElementoSubvencion.Filter(objFilter)
            objNegElementoSubvencion.Delete(dtElementoSubvencion)
        End If
    End Sub

#End Region

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatoriosCabecera)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarFamiliaElemento)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarValores)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarBajaElemento)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarRelacionesFechas)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarValorAmortizado)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatoriosCabecera(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim StrCampos As String = String.Empty
        If Length(data("IDCodigoAmortizacionContable")) = 0 Then StrCampos &= "IDCodigoAmortizacionContable, "
        If Length(data("ValorTotalElementoA")) = 0 OrElse data("ValorTotalElementoA") = 0 Then StrCampos &= "ValorTotalElementoA, "
        If Length(data("IDCodigoAmortizacionFiscal")) = 0 Then StrCampos &= "IDCodigoAmortizacionFiscal, "
        If Length(data("IDCodigoAmortizacionTecnica")) = 0 Then StrCampos &= "IDCodigoAmortizacionTecnica, "
        If Length(data("FechaInicioContabilizacion")) = 0 Then StrCampos &= "FechaInicioContabilizacion, "
        If Length(data("ValorTotalRevalElementoA")) = 0 Then StrCampos &= "ValorTotalRevalElementoA, "
        If Length(data("FechaCompra")) = 0 Then StrCampos &= "FechaCompra, "
        If Length(data("DescElemento")) = 0 Then StrCampos &= "DescElemento, "
        If Length(data("IDEstado")) = 0 Then StrCampos &= "IDEstado, "
        If Length(data("IDGrupoAmortizacion")) = 0 Then StrCampos &= "IDGrupoAmortizacion, "
        If Length(data("IDInmovilizado")) = 0 Then StrCampos &= "IDInmovilizado, "
        If Length(data("IDActivo")) = 0 Then StrCampos &= "IDActivo."
        If StrCampos.Length > 0 Then ApplicationService.GenerateError("Debe establecer valor a | ", StrCampos.Substring(0, StrCampos.Length - 2))
    End Sub

    <Task()> Public Shared Sub ValidarFamiliaElemento(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDFamiliaElemento")) > 0 AndAlso Length(data("IDTipoElemento")) > 0 Then
            Dim FilFam As New Filter
            FilFam.Add("IDTipoElemento", FilterOperator.Equal, data("IDTipoElemento"))
            FilFam.Add("IDFamiliaElemento", FilterOperator.Equal, data("IDFamiliaElemento"))
            Dim Dt As DataTable = New FamiliaElemento().Filter(FilFam)
            If Dt Is Nothing OrElse Dt.Rows.Count = 0 Then
                ApplicationService.GenerateError("No existe la familia del elemento para el tipo de elemento elegido")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarValores(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("ValorResidualA")) > 0 AndAlso data("ValorResidualA") < 0 Then ApplicationService.GenerateError("El valor residual A no es válido")
        If Length(data("ValorResidualB")) > 0 AndAlso data("ValorResidualB") < 0 Then ApplicationService.GenerateError("El valor residual B no es válido")
        If Length(data("ValorTotalElementoA")) > 0 AndAlso data("ValorTotalElementoA") < 0 Then ApplicationService.GenerateError("El valor total A no es válido")
        If Length(data("ValorTotalElementoB")) > 0 AndAlso data("ValorTotalElementoB") < 0 Then ApplicationService.GenerateError("El valor total B no es válido")
        If Length(data("ValorTotalRevalElementoA")) > 0 AndAlso data("ValorTotalRevalElementoA") < 0 Then ApplicationService.GenerateError("El valor total de revalorización A no es válido")
        If Length(data("ValorTotalRevalElementoB")) > 0 AndAlso data("ValorTotalRevalElementoB") < 0 Then ApplicationService.GenerateError("El valor total de revalorización B no es válido")
        If Length(data("ValorNetoContableElementoA")) > 0 AndAlso data("ValorNetoContableElementoA") < 0 Then ApplicationService.GenerateError("El valor neto contable A no es válido")
        If Length(data("ValorNetoContableElementoB")) > 0 AndAlso data("ValorNetoContableElementoB") < 0 Then ApplicationService.GenerateError("El valor neto contable B no es válido")
    End Sub

    <Task()> Public Shared Sub ValidarBajaElemento(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.HasVersion(DataRowVersion.Original) Then
            If data("Baja", DataRowVersion.Original) = True AndAlso (Not data.IsNull("Baja") AndAlso data("Baja") = True) Then
                ApplicationService.GenerateError("Un Elemento dado de Baja no puede ser modificado.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarRelacionesFechas(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("FechaInicioContabilizacion")) > 0 AndAlso Length(data("FechaCompra")) > 0 Then
            Dim FechaCompra As Date = CDate(data("FechaCompra"))
            Dim FechaInicio As Date = CDate(data("FechaInicioContabilizacion"))
            If FechaInicio < FechaCompra Then
                ApplicationService.GenerateError("La Fecha de Inicio de Contabilización no puede ser anterior a la Fecha de Compra del Elemento.")
            End If
        End If
        If Length(data("FechaInicioContabilizacion")) > 0 AndAlso Length(data("FechaUltimaContabilizacion")) > 0 Then
            Dim FechaInicio As Date = CDate(data("FechaInicioContabilizacion"))
            Dim FechaFin As Date = CDate(data("FechaUltimaContabilizacion"))
            If FechaInicio > FechaFin Then
                ApplicationService.GenerateError("La Fecha de Inicio de Contabilización no puede ser mayor que la Fecha de Ultima Contabilizacion.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarValorAmortizado(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim DblValorAmort As Double
        Dim DblValorAmortPlus As Double
        If Length(data("ValorAmortizadoElementoA")) = 0 Then
            DblValorAmort = 0
        Else : DblValorAmort = data("ValorAmortizadoElementoA")
        End If
        If Length(data("ValorAmortizadoPlusvaliaA")) = 0 Then
            DblValorAmortPlus = 0
        Else : DblValorAmortPlus = data("ValorAmortizadoPlusvaliaA")
        End If
        If DblValorAmort < 0 OrElse DblValorAmortPlus < 0 Then
            ApplicationService.GenerateError("El valor amortizado no puede ser inferior a 0")
        ElseIf DblValorAmort = 0 And DblValorAmortPlus = 0 Then
            If Length(data("FechaUltimaContabilizacion")) > 0 Then
                ApplicationService.GenerateError("Debe establecer el Valor Amortizado o borrar la Fecha de Ultima Amortizacion")
            End If
        ElseIf DblValorAmort > 0 OrElse DblValorAmortPlus > 0 Then
            If Length(data("FechaUltimaContabilizacion")) = 0 Then
                ApplicationService.GenerateError("Debe establecer valor en Fecha de Ultima Amortizacion o borrar el Valor Amortizado")
            End If
        End If
    End Sub

#End Region

#Region "Funciones ElementoAmortizable"

    <Serializable()> _
    Public Class udtElementoAmortizable
        Public dtElemento As DataTable
        Public dtAmortizacion As DataTable
        Public dtRevalorizacion As DataTable
        Public lngResultado As Boolean
    End Class

    <Serializable()> _
    Public Class DatosCuotaAmortizacion
        Public DotacionElementoA As Double
        Public DotacionElementoB As Double
        Public VidaUtil As Integer
        Public PorcentajeAnual As Double
    End Class

    <Task()> Public Shared Function CrearDtAmort(ByVal data As Object, ByVal services As ServiceProvider) As DataTable
        Dim DtAmort As New DataTable
        'Creamos el rs donde se guardaran los datos de la amortizacion
        DtAmort.Columns.Add("IDElemento", GetType(String))
        DtAmort.Columns.Add("Año", GetType(Integer))
        DtAmort.Columns.Add("AmortContable", GetType(Double))
        DtAmort.Columns.Add("AmortFiscal", GetType(Double))
        DtAmort.Columns.Add("AmortTecnica", GetType(Double))
        DtAmort.Columns.Add("AmortRealizada", GetType(Double))
        DtAmort.Columns.Add("AmortSimAño", GetType(Double))
        DtAmort.Columns.Add("ValorNeto", GetType(Double))
        DtAmort.Columns.Add("AmortContableMensual", GetType(String))
        DtAmort.Columns.Add("AmortTecnicaMensual", GetType(String))
        DtAmort.Columns.Add("AmortFiscalMensual", GetType(String))
        DtAmort.Columns.Add("AmortRealizadaMensual", GetType(String))
        DtAmort.Columns.Add("AmortAño", GetType(String))
        Return DtAmort
    End Function

    <Task()> Public Shared Function DeshacerBaja(ByVal DtElementos As DataTable, ByVal services As ServiceProvider) As Integer
        If Not DtElementos Is Nothing AndAlso DtElementos.Rows.Count > 0 Then
            DtElementos.Rows(0)("Baja") = False
            DtElementos.Rows(0)("FechaBaja") = System.DBNull.Value
            DtElementos.Rows(0)("IdEjercicioBaja") = System.DBNull.Value
            Dim ClsElem As New ElementoAmortizable
            ClsElem.Update(DtElementos)
        End If
    End Function

    Private Shared Function ComparaDr(ByVal Dr As DataRow, ByVal Campo As String) As Boolean
        'True si son iguales versión actual y orginal. False en otro caso.
        If Dr.HasVersion(DataRowVersion.Original) Then
            If Dr.IsNull(Dr.Table.Columns(Campo), DataRowVersion.Original) Then
                If Dr.IsNull(Dr.Table.Columns(Campo)) Then
                    Return True
                Else : Return False
                End If
            Else
                If Dr.IsNull(Dr.Table.Columns(Campo)) Then
                    Return False
                Else
                    If Dr(Campo, DataRowVersion.Original) = Dr(Campo) Then
                        Return True
                    Else : Return False
                    End If
                End If
            End If
        Else : Return False
        End If
    End Function


    <Serializable()> _
    Public Class DataValidarDeshacerAmortizacionGrupo
        Public Elementos() As String
        Public dtDesamortizar As DataTable

        Public Sub New(ByVal Elementos() As String, ByVal dtDesamortizar As DataTable)
            Me.Elementos = Elementos
            Me.dtDesamortizar = dtDesamortizar
        End Sub
    End Class

    <Task()> Public Shared Sub ValidarDeshacerAmortizacionGrupo(ByVal data As DataValidarDeshacerAmortizacionGrupo, ByVal services As ServiceProvider)
        If data Is Nothing Then Exit Sub
        Dim AppParams As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
        If AppParams.AmortizacionPorGrupo AndAlso Not data.Elementos Is Nothing AndAlso data.Elementos.Length > 0 Then



            Dim IDEjercicioAnt As String
            Dim NAsientoAnt As Integer

            For Each drAmortizacion As DataRow In data.dtDesamortizar.Select(Nothing, "IDEjercicio DESC, NAsiento DESC")

                If IDEjercicioAnt <> drAmortizacion("IDEjercicio") OrElse NAsientoAnt <> drAmortizacion("NAsiento") Then
                    Dim fDC As New Filter
                    fDC.Add(New StringFilterItem("IDEjercicio", drAmortizacion("IDEjercicio")))
                    fDC.Add(New NumberFilterItem("NAsiento", drAmortizacion("NAsiento")))


                    Dim dtAmortizacionReg As DataTable = New AmortizacionRegistro().Filter(fDC)
                    Dim ElementosAmorEnAsiento As List(Of Object) = (From c In dtAmortizacionReg _
                                                                     Where c("IDEjercicio") = drAmortizacion("IDEjercicio") _
                                                                       AndAlso c("NAsiento") = drAmortizacion("NAsiento") _
                                                                     Select c("IDElemento")).ToList()
                    If ElementosAmorEnAsiento.Count > 0 Then


                        For Each IDElemento As String In ElementosAmorEnAsiento

                            '//Validamos si no hemos marcado para desamortizar (debido a los filtros de la consulta), todos los elementos que fueron contabilizados en el mismo asiento
                            Dim NoExisteElemento As List(Of Object) = (From c In data.dtDesamortizar _
                                                                Where c("IDElemento") = IDElemento _
                                                                Select c("IDElemento")).ToList()
                            If Not NoExisteElemento Is Nothing AndAlso NoExisteElemento.Count = 0 Then
                                ApplicationService.GenerateError("No se han seleccionado todos los elementos para Desamortizar del Asiento {0} del Ejercicio {1}.{2}Deberían seleccionarse los Elementos {3}.", _
                                                                 Quoted(drAmortizacion("NAsiento")), Quoted(drAmortizacion("IDEjercicio")), vbNewLine, Strings.Join(ElementosAmorEnAsiento.ToArray, ","))
                            End If



                            '//Validamos que aun habiendo marcado los elementos para desamortizar, no todas sus amortizaciones se han agrupado con otro elemento, y no podemos desamortizarlo.
                            Dim NoOKElemento As List(Of DataRow) = (From c In data.dtDesamortizar _
                                                                        Where c("IDElemento") = IDElemento _
                                                                           AndAlso (c("IDEjercicio") <> drAmortizacion("IDEjercicio") _
                                                                           OrElse c("NAsiento") <> drAmortizacion("NAsiento")) _
                                                                        Select c).ToList()
                            If Not NoOKElemento Is Nothing AndAlso NoOKElemento.Count > 0 Then
                                ApplicationService.GenerateError("El Elemento {0} se ha amortizado en el Asiento {1} del Ejercicio {2}.", _
                                                                   Quoted(NoOKElemento(0)("IDElemento")), Quoted(NoOKElemento(0)("NAsiento")), Quoted(NoOKElemento(0)("IDEjercicio")))
                            Else
                                '//Si no viene entre los elementos a amortizar
                                Dim dtElemento As DataTable = AdminData.GetData("tbMaestroElementoAmortizable", New StringFilterItem("IDElemento", IDElemento))
                                NoOKElemento = (From c In dtElemento _
                                                                      Where c("IDElemento") = IDElemento _
                                                                         AndAlso (c("IDEjercicio") <> drAmortizacion("IDEjercicio") _
                                                                         OrElse c("NAsiento") <> drAmortizacion("NAsiento")) _
                                                                      Select c).ToList()
                                If Not NoOKElemento Is Nothing AndAlso NoOKElemento.Count > 0 Then
                                    ApplicationService.GenerateError("El Elemento {0} se ha amortizado en el Asiento {1} del Ejercicio {2}.", _
                                                                       Quoted(NoOKElemento(0)("IDElemento")), Quoted(NoOKElemento(0)("NAsiento")), Quoted(NoOKElemento(0)("IDEjercicio")))

                                End If
                            End If

                        Next


                    End If


                    IDEjercicioAnt = drAmortizacion("IDEjercicio")
                    NAsientoAnt = drAmortizacion("NAsiento")
                End If
            Next
        End If
    End Sub

    <Task()> Public Shared Sub DeshacerAmortizacion(ByVal Dt As DataTable, ByVal services As ServiceProvider)
        If Not Dt Is Nothing AndAlso Dt.Rows.Count > 0 Then
            Dim StrElems(Dt.Rows.Count - 1) As String
            Dim i As Integer = 0
            For Each Dr As DataRow In Dt.Select
                StrElems(i) = Dr("IdElemento")
                i += 1
            Next
            If StrElems.Length > 0 Then

                Dim datValidar As New DataValidarDeshacerAmortizacionGrupo(StrElems, Dt)
                ProcessServer.ExecuteTask(Of DataValidarDeshacerAmortizacionGrupo)(AddressOf ValidarDeshacerAmortizacionGrupo, datValidar, services)

                Dim DtElementos As DataTable = New ElementoAmortizable().Filter(New InListFilterItem("IDElemento", StrElems, FilterType.String))
                ProcessServer.ExecuteTask(Of Object)(AddressOf Comunes.BeginTransaction, Nothing, services)
                Dim DtAmortizacionReg As DataTable = New AmortizacionRegistro().Filter(New InListFilterItem("IDElemento", StrElems, FilterType.String), " AñoContabilizacion DESC, MesContabilizacion DESC, IDAmortizacionRegistro DESC")
                Dim DtRegistrosBorrar As New DataTable
                DtRegistrosBorrar.Columns.Add("IDAmortizacionRegistro", GetType(Integer))
                If Not DtElementos Is Nothing AndAlso DtElementos.Rows.Count > 0 Then
                    Dim DblImporteBajaA As Double = 0
                    Dim DblImporteBajaB As Double = 0
                    Dim DblPlusvaliaA As Double = 0
                    Dim DblPlusvaliaB As Double = 0
                    For Each Dr As DataRow In DtElementos.Select
                        Dim DrAmort() As DataRow = DtAmortizacionReg.Select("IDElemento='" & Dr("IdElemento") & "'")
                        If Not DrAmort Is Nothing AndAlso DrAmort.Length > 0 Then
                            Dim DrNew As DataRow = DtRegistrosBorrar.NewRow()
                            DrNew("IDAmortizacionRegistro") = DrAmort(0)("IDAmortizacionRegistro")
                            DtRegistrosBorrar.Rows.Add(DrNew)

                            Dim BlnBajaParcial As Boolean = False
                            Dim BlnPlusvalia As Boolean = False
                            If Length(DrAmort(0)("IDEjercicio")) > 0 AndAlso Length(DrAmort(0)("NAsiento")) > 0 Then
                                Dim ClsDiario As BusinessHelper = BusinessHelper.CreateBusinessObject("DiarioContable")
                                Dim f As New Filter
                                f.Add(New StringFilterItem("IDEjercicio", DrAmort(0)("IDEjercicio")))
                                f.Add(New NumberFilterItem("NAsiento", DrAmort(0)("NAsiento")))
                                '   f.Add(New NumberFilterItem("IDTipoApunte", enumDiarioTipoApunte.BajaParcialElemento))
                                Dim DtDiario As DataTable = ClsDiario.Filter(f)
                                Dim fTipoAsiento As New Filter

                                If Not DtDiario Is Nothing AndAlso DtDiario.Rows.Count > 0 Then
                                    'Dim StrDescApunte As String = "Inmovilizado inmaterial/material, Baja Parcial"
                                    If DtDiario.Rows(0)("IDTipoApunte") = enumDiarioTipoApunte.BajaParcialElemento Then
                                        fTipoAsiento.Add(New NumberFilterItem("IDTipoApunte", enumDiarioTipoApunte.BajaParcialElemento))
                                        Dim DrDiario() As DataRow = DtDiario.Select("DH ='H'")
                                        If Not DrDiario Is Nothing AndAlso DrDiario.Length Then
                                            BlnBajaParcial = True
                                            DblImporteBajaA = DrDiario(0)("ImpApunteA")
                                            DblImporteBajaB = DrDiario(0)("ImpApunteB")
                                        End If
                                    End If
                                    If DtDiario.Rows(0)("IDTipoApunte") = enumDiarioTipoApunte.Plusvalia Then
                                        fTipoAsiento.Add(New NumberFilterItem("IDTipoApunte", enumDiarioTipoApunte.Plusvalia))
                                        Dim DrDiario() As DataRow = DtDiario.Select("DH ='D'")
                                        If Not DrDiario Is Nothing AndAlso DrDiario.Length Then
                                            BlnPlusvalia = True
                                            DblPlusvaliaA = DrDiario(0)("ImpApunteA")
                                            DblPlusvaliaB = DrDiario(0)("ImpApunteB")
                                        End If
                                    End If
                                    If DtDiario.Rows(0)("IDTipoApunte") = enumDiarioTipoApunte.Amortiz Then
                                        fTipoAsiento.Add(New NumberFilterItem("IDTipoApunte", enumDiarioTipoApunte.Amortiz))
                                    End If

                                    '//CUIDADO. No quitar. Si al Renumerar asientos se cambian de número, puede que no esté actualizado el NAsiento.
                                    If fTipoAsiento.Count > 0 Then
                                        Dim fDelete As New Filter
                                        fDelete.Add(New StringFilterItem("IDEjercicio", DrAmort(0)("IDEjercicio")))
                                        fDelete.Add(New NumberFilterItem("NAsiento", DrAmort(0)("NAsiento")))
                                        fDelete.Add(fTipoAsiento)

                                        Dim AppParams As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
                                        If AppParams.AmortizacionPorGrupo Then
                                            fDelete.Add(New StringFilterItem("NDocumento", Dr("IDGrupoAmortizacion")))
                                        Else
                                            fDelete.Add(New StringFilterItem("NDocumento", Dr("IdElemento")))
                                        End If

                                        Dim dtDiarioComprueba As DataTable = ClsDiario.Filter(f)
                                        If dtDiarioComprueba.Rows.Count = 0 Then
                                            ApplicationService.GenerateError("Revise el Asiento {0} del Ejercicio {1}. No es posible eliminarlo.", DrAmort(0)("NAsiento"), DrAmort(0)("IDEjercicio"))
                                        Else
                                            NegocioGeneral.DeleteWhere(DrAmort(0)("IDEjercicio"), fDelete)
                                        End If
                                    Else
                                        ApplicationService.GenerateError("Revise el Asiento {0} del Ejercicio {1}. No es posible eliminarlo.", DrAmort(0)("NAsiento"), DrAmort(0)("IDEjercicio"))
                                    End If
                                End If
                            End If

                            If BlnBajaParcial Then
                                Dr("ValorTotalRevalElementoA") -= DblImporteBajaA
                                Dr("ValorTotalRevalElementoB") -= DblImporteBajaB
                                Dr("ValorAmortizadoElementoA") -= DrAmort(0)("ValorAmortizadoA")
                                Dr("ValorAmortizadoElementoB") -= DrAmort(0)("ValorAmortizadoB")
                                'Dr("ValorNetoContableElementoA") += Dr("ValorAmortizadoElementoA")
                                'Dr("ValorNetoContableElementoB") += Dr("ValorAmortizadoElementoB")
                                Dr("ValorNetoContableElementoA") = Dr("ValorTotalRevalElementoA") - Dr("ValorAmortizadoElementoA")
                                Dr("ValorNetoContableElementoB") = Dr("ValorTotalRevalElementoB") - Dr("ValorAmortizadoElementoB")

                                Dim f As New Filter
                                f.Add(New StringFilterItem("IDElemento", Dr("IdElemento")))
                                f.Add(New IsNullFilterItem("IDLineaRevalBaja", False))
                                Dim BEDataEngine As New BE.DataEngine
                                Dim dtElemReval As DataTable = BEDataEngine.Filter("tbElementoRevalorizacion", f, "TOP 1 IDLineaRevalorizacion, IDLineaRevalBaja, IDElemento", "IDLineaRevalBaja DESC")
                                If Not dtElemReval Is Nothing AndAlso dtElemReval.Rows.Count > 0 Then
                                    'Dim ER As New ElementoRevalorizacion
                                    dtElemReval.TableName = GetType(ElementoRevalorizacion).Name
                                    'ER.Delete(dtElemReval)
                                    dtElemReval.Rows(0).Delete()
                                    BusinessHelper.UpdateTable(dtElemReval)
                                End If
                            ElseIf BlnPlusvalia Then
                                Dr("ValorTotalPlusvaliaA") -= DblPlusvaliaA
                                Dr("ValorTotalPlusvaliaB") -= DblPlusvaliaB
                                Dr("ValorAmortizadoPlusvaliaA") -= DrAmort(0)("ValorAmortizadoPlusvaliaA")
                                Dr("ValorAmortizadoElementoB") -= DrAmort(0)("ValorAmortizadoPlusvaliaB")
                                'Dr("ValorNetoContableElementoA") += Dr("ValorAmortizadoElementoA")
                                'Dr("ValorNetoContableElementoB") += Dr("ValorAmortizadoElementoB")
                                Dr("ValorNetoContablePlusvaliaA") = Dr("ValorTotalPlusvaliaA") - Dr("ValorAmortizadoPlusvaliaA")
                                Dr("ValorNetoContablePlusvaliaB") = Dr("ValorTotalPlusvaliaB") - Dr("ValorAmortizadoPlusvaliaB")


                                Dim f As New Filter
                                f.Add(New StringFilterItem("IDElemento", Dr("IdElemento")))
                                f.Add(New IsNullFilterItem("IDLineaRevalBaja", False))
                                Dim BEDataEngine As New BE.DataEngine
                                Dim dtElemReval As DataTable = BEDataEngine.Filter("tbElementoRevalorizacion", f, "TOP 1 IDLineaRevalorizacion, IDLineaRevalBaja, IDElemento", "IDLineaRevalBaja DESC")
                                If Not dtElemReval Is Nothing AndAlso dtElemReval.Rows.Count > 0 Then
                                    'Dim ER As New ElementoRevalorizacion
                                    dtElemReval.TableName = GetType(ElementoRevalorizacion).Name
                                    'ER.Delete(dtElemReval)
                                    dtElemReval.Rows(0).Delete()
                                    BusinessHelper.UpdateTable(dtElemReval)
                                End If
                            Else
                                Dr("ValorAmortizadoElementoA") -= DrAmort(0)("ValorAmortizadoA")
                                Dr("ValorAmortizadoElementoB") -= DrAmort(0)("ValorAmortizadoB")
                                Dr("ValorNetoContableElementoA") = Dr("ValorTotalRevalElementoA") - Dr("ValorAmortizadoElementoA")
                                Dr("ValorNetoContableElementoB") = Dr("ValorTotalRevalElementoB") - Dr("ValorAmortizadoElementoB")

                                Dr("ValorAmortizadoPlusvaliaA") -= DrAmort(0)("ValorAmortizadoPlusvaliaA")
                                Dr("ValorAmortizadoPlusvaliaB") -= DrAmort(0)("ValorAmortizadoPlusvaliaB")
                                Dr("ValorNetoContablePlusvaliaA") = Dr("ValorTotalPlusvaliaA") - Dr("ValorAmortizadoPlusvaliaA")
                                Dr("ValorNetoContablePlusvaliaB") = Dr("ValorTotalPlusvaliaB") - Dr("ValorAmortizadoPlusvaliaB")
                            End If


                            If DrAmort.Length = 1 Then
                                Dr("MesUltimoContabilizado") = 0
                                Dr("AñoUltimoContabilizado") = 0
                                Dr("FechaUltimaContabilizacion") = System.DBNull.Value
                                Dr("ValorUltimoContabilizadoA") = 0
                                Dr("ValorUltimoContabilizadoB") = 0
                                Dr("AmortizacionAutomatica") = False
                                Dr("IDEjercicio") = DBNull.Value
                                Dr("NAsiento") = DBNull.Value
                            Else
                                Dr("MesUltimoContabilizado") = DrAmort(1)("MesContabilizacion")
                                Dr("AñoUltimoContabilizado") = DrAmort(1)("AñoContabilizacion")
                                Dr("FechaUltimaContabilizacion") = DrAmort(1)("FechaContabilizacion")
                                Dr("ValorUltimoContabilizadoA") = DrAmort(1)("ValorAmortizadoA")
                                Dr("ValorUltimoContabilizadoB") = DrAmort(1)("ValorAmortizadoB")
                                Dr("IDEjercicio") = DrAmort(1)("IDEjercicio")
                                Dr("NAsiento") = DrAmort(1)("NAsiento")
                                If Length(Dr("IDEjercicio")) = 0 AndAlso Length(Dr("NAsiento")) = 0 Then
                                    Dr("AmortizacionAutomatica") = False
                                End If
                            End If
                        End If

                    Next
                End If
                BusinessHelper.UpdateTable(DtElementos)
                Dim ClsAmortReg As New AmortizacionRegistro
                ClsAmortReg.Delete(DtRegistrosBorrar)
            End If
        End If
    End Sub

    <Task()> Public Shared Function InsertarAnaliticaCentroGestion(ByVal Dr As DataRow, ByVal services As ServiceProvider) As DataTable
        Dim DtResult As DataTable = New ElementoAmortizAnalitica().AddNew
        Dim drNew As DataRow = DtResult.NewRow
        drNew("IdElemento") = Dr("IdElemento")
        If Dr("IDCentroGestion").ToString.Trim.Length = 0 Then
            drNew("IDCentroGestion") = New Parametro().CGestionPredet
        Else : drNew("IDCentroGestion") = Dr("IDCentroGestion")
        End If
        drNew("IDCentroCoste") = New Parametro().ObtenerPredeterminado("CCOSTEPRED")
        drNew("porcentaje") = 100
        DtResult.Rows.Add(drNew)
        Return DtResult
    End Function

    <Task()> Public Shared Function InsertDtAutomatico(ByVal dt As DataTable, ByVal services As ServiceProvider) As DataTable
        ProcessServer.ExecuteTask(Of Object)(AddressOf Comunes.BeginTransaction, Nothing, services)
        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
        Dim MonInfoA As MonedaInfo = Monedas.MonedaA
        Dim dblCambioAB As Double = MonInfoA.CambioB
        Dim IntDecA As Integer = MonInfoA.NDecimalesImporte
        Dim MonInfoB As MonedaInfo = Monedas.MonedaB
        Dim IntDecB As Integer = MonInfoB.NDecimalesImporte

        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            Dim dtNew As DataTable = New ElementoAmortizable().AddNew()
            Dim TipoAmortizLinea As New TipoAmortizacionLinea

            For Each dr As DataRow In dt.Rows
                Dim drNew As DataRow = dtNew.NewRow
                For Each dc As DataColumn In dt.Columns
                    If dc.ColumnName <> "LineasFactura" Then
                        If dc.ColumnName = "FechaCompra" Then
                            drNew(dc.ColumnName) = dr(dc.ColumnName)
                            drNew("FechaInicioContabilizacion") = dr(dc.ColumnName)
                            dr("FechaInicioContabilizacion") = dr(dc.ColumnName)
                        ElseIf dc.ColumnName = "DescElemento" Then
                            drNew(dc.ColumnName) = Left(dr(dc.ColumnName), 300)
                        Else : drNew(dc.ColumnName) = dr(dc.ColumnName)
                        End If
                    End If
                Next
                drNew("ValorAmortizadoElementoA") = 0
                drNew("ValorNetoContableElementoA") = drNew("ValorTotalElementoA")
                drNew("ValorTotalPlusvaliaA") = 0
                drNew("ValorAmortizadoPlusvaliaA") = 0
                drNew("ValorNetoContablePlusvaliaA") = 0
                drNew("ValorAmortizadoElementoB") = 0
                drNew("ValorNetoContableElementoB") = drNew("ValorTotalElementoB")
                drNew("ValorTotalPlusvaliaB") = 0
                drNew("ValorAmortizadoPlusvaliaB") = 0
                drNew("ValorNetoContablePlusvaliaB") = 0

                Dim drGAmortiz As DataRow = New GrupoAmortizacion().GetItemRow(dr("IDGrupoAmortizacion"))
                If Not drGAmortiz Is Nothing Then
                    Dim drTipoAmortiz As DataRow
                    If Length(dr("IDCodigoAmortizacionContable")) > 0 Then
                        drTipoAmortiz = New TipoAmortizacionCabecera().GetItemRow(dr("IDCodigoAmortizacionContable"))
                    Else : drTipoAmortiz = New TipoAmortizacionCabecera().GetItemRow(drGAmortiz("IDTipoAmortiz"))
                    End If

                    drNew("IDCodigoAmortizacionContable") = drTipoAmortiz("IdTipoAmortizacion")
                    drNew("IDCodigoAmortizacionTecnica") = drTipoAmortiz("IdTipoAmortizacion")
                    drNew("IDCodigoAmortizacionFiscal") = drTipoAmortiz("IdTipoAmortizacion")
                    drNew("VidaContableElemento") = drTipoAmortiz("VidaUtil")
                    drNew("VidaTecnicaElemento") = drTipoAmortiz("VidaUtil")
                    drNew("VidaFiscalElemento") = drTipoAmortiz("VidaUtil")

                    Dim filAmortizLinea As New Filter
                    filAmortizLinea.Add("IDTipoAmortizacion", FilterOperator.Equal, drTipoAmortiz("IdTipoAmortizacion"), FilterType.String)
                    Dim dtTipoAmortizLinea As DataTable = New TipoAmortizacionLinea().Filter(filAmortizLinea, "NAño", "TOP 1 PorcentajeAmortizar")
                    Dim DblPorcentaje As Double
                    If Not IsNothing(dtTipoAmortizLinea) AndAlso dtTipoAmortizLinea.Rows.Count > 0 Then
                        DblPorcentaje = dtTipoAmortizLinea.Rows(0)("PorcentajeAmortizar")
                    Else : DblPorcentaje = 0
                    End If

                    drNew("PorcentajeAnualContable") = DblPorcentaje
                    drNew("PorcentajeAnualTecnico") = DblPorcentaje
                    drNew("PorcentajeAnualFiscal") = DblPorcentaje

                    Dim DblValorAmortizA, DblDotacionA, DblDotacionB As Double
                    DblValorAmortizA = drNew("ValorTotalRevalElementoA") - drNew("ValorResidualA")
                    DblDotacionA = (DblValorAmortizA * (DblPorcentaje / 100)) / 12
                    DblDotacionB = DblDotacionA * dblCambioAB

                    DblDotacionA = xRound(DblDotacionA, IntDecA)
                    DblDotacionB = xRound(DblDotacionB, IntDecB)

                    drNew("DotacionContableElementoA") = DblDotacionA
                    drNew("DotacionTecnicaElementoA") = DblDotacionA
                    drNew("DotacionFiscalElementoA") = DblDotacionA
                    drNew("DotacionContableElementoB") = DblDotacionB
                    drNew("DotacionTecnicaElementoB") = DblDotacionB
                    drNew("DotacionFiscalElementoB") = DblDotacionB
                End If
                drNew("MesUltimoContabilizado") = 0
                drNew("AñoUltimoContabilizado") = 0
                drNew("FechaUltimaContabilizacion") = System.DBNull.Value
                drNew("FechaUltimaRevalorizacion") = dr("FechaCompra")
                drNew("ValorUltimoContabilizadoA") = 0
                drNew("ValorUltimoContabilizadoB") = 0
                drNew("ValorReposicionA") = 0
                drNew("ValorReposicionActualizadoA") = 0
                drNew("ValorReposicionB") = 0
                drNew("ValorReposicionActualizadoB") = 0
                dtNew.Rows.Add(drNew)
            Next

            If Not IsNothing(dtNew) AndAlso dtNew.Rows.Count > 0 Then
                BusinessHelper.UpdateTable(dtNew)
                Dim dtElemFCL As DataTable = New ElementoAmortizableFCL().AddNew
                Dim ElemANA As New ElementoAmortizAnalitica
                For Each dr As DataRow In dt.Rows
                    If Length(dr("IDElementoOrigen") & String.Empty) = 0 Then
                        Dim dtElemAna As DataTable = ElemANA.AddNew()
                        Dim DblImpTotalANA As Double = 0
                        Dim strSQLinversion As String = "SELECT IDCentroCoste, IDCentroGestion, SUM(tbFacturaCompraAnalitica.ImporteA) AS ImporteA" & vbCrLf _
                                                           & "FROM tbFacturaCompraLinea INNER JOIN tbFacturaCompraAnalitica ON tbFacturaCompraLinea.IDLineaFactura = tbFacturaCompraAnalitica.IDLineaFactura" & vbCrLf _
                                                                               & "WHERE tbFacturaCompraLinea.IDLineaFactura IN (" & Replace(dr("LineasFactura"), ";", ",") & ")" & vbCrLf _
                                                                               & "GROUP BY IDCentroCoste, IDCentroGestion"
                        Dim dtInversionAna As DataTable = AdminData.Execute(strSQLinversion, ExecuteCommand.ExecuteReader)
                        If Not IsNothing(dtInversionAna) AndAlso dtInversionAna.Rows.Count > 0 Then
                            For Each drInversionAna As DataRow In dtInversionAna.Rows
                                DblImpTotalANA = DblImpTotalANA + drInversionAna("ImporteA")
                            Next
                        End If
                        Dim drElemAna As DataRow
                        For Each drInversionAna As DataRow In dtInversionAna.Rows
                            drElemAna = dtElemAna.NewRow
                            drElemAna("IdElemento") = dr("IdElemento")
                            drElemAna("IDCentroGestion") = drInversionAna("IDCentroGestion")
                            drElemAna("IDCentroCoste") = drInversionAna("IDCentroCoste")
                            drElemAna("porcentaje") = xRound(100 * drInversionAna("ImporteA") / DblImpTotalANA, 2)
                            dtElemAna.Rows.Add(drElemAna)
                        Next
                        ElemANA.Update(dtElemAna)

                        Dim arrIDs() As String = Split(dr("LineasFactura"), ";", , vbTextCompare)
                        Dim drElemFCL As DataRow
                        For i As Integer = 0 To UBound(arrIDs)
                            Dim drFCL As DataRow = New FacturaCompraLinea().GetItemRow(CLng(arrIDs(i)))
                            If Not IsNothing(drFCL) Then
                                drFCL("EstadoInmovilizado") = True
                                BusinessHelper.UpdateTable(drFCL.Table)
                                drElemFCL = dtElemFCL.NewRow
                                drElemFCL("IdLinea") = AdminData.GetAutoNumeric
                                drElemFCL("IdLineaFactura") = arrIDs(i)
                                drElemFCL("IdElemento") = dr("IDElemento")
                                drElemFCL("NFactura") = dr("NFactura")
                                drElemFCL("IDCContable") = drFCL("CContable")
                                drElemFCL("ImporteA") = drFCL("ImporteA")
                                drElemFCL("ImporteB") = drFCL("ImporteB")
                                dtElemFCL.Rows.Add(drElemFCL)
                            End If
                        Next
                    End If
                Next
                Dim ClsFCL As New ElementoAmortizableFCL
                ClsFCL.Update(dtElemFCL)
                ProcessServer.ExecuteTask(Of DataTable)(AddressOf ElementoRevalorizacion.InsertarElementosAutomatico, dtNew, services)
                Return dtNew
            End If
        Else : ApplicationService.GenerateError("Error en el proceso.")
        End If
        ProcessServer.ExecuteTask(Of Object)(AddressOf Comunes.CommitTransaction, Nothing, services)
    End Function

    <Task()> Public Shared Function InsertElementoMaterial(ByVal dt As DataTable, ByVal services As ServiceProvider) As DataTable
        Dim DblValorAmortizA, dblDotacionA, dblDotacionB, dblPorcentaje, DblCambioAB As Double
        ProcessServer.ExecuteTask(Of Object)(AddressOf Comunes.BeginTransaction, Nothing, services)

        Dim LngDecimalesA, LngDecimalesB As Integer
        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
        Dim MonInfoA As MonedaInfo = Monedas.MonedaA
        DblCambioAB = MonInfoA.CambioB
        LngDecimalesA = MonInfoA.NDecimalesImporte
        Dim MonInfoB As MonedaInfo = Monedas.MonedaB
        LngDecimalesB = MonInfoB.NDecimalesImporte

        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
            Dim dtNew As DataTable = New ElementoAmortizable().AddNew()
            For Each dr As DataRow In dt.Rows
                Dim drNew As DataRow = dtNew.NewRow
                For Each dc As DataColumn In dt.Columns
                    If dc.ColumnName <> "LineasFactura" Then
                        If dc.ColumnName = "IDElemento" Then
                            If dr(dc.ColumnName) = "" Then
                                drNew(dc.ColumnName) = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, dr("IDContador"), services)
                            Else : drNew(dc.ColumnName) = dr(dc.ColumnName)
                            End If
                        ElseIf dc.ColumnName = "FechaCompra" Then
                            drNew(dc.ColumnName) = dr(dc.ColumnName)
                            drNew("FechaInicioContabilizacion") = dr(dc.ColumnName)
                        ElseIf dc.ColumnName = "FechaBaja" Or dc.ColumnName = "IDEjercicioBaja" Then
                            drNew(dc.ColumnName) = System.DBNull.Value
                        ElseIf dc.ColumnName = "DescElemento" Then
                            drNew(dc.ColumnName) = Left(dr(dc.ColumnName), 300)
                        Else : drNew(dc.ColumnName) = dr(dc.ColumnName)
                        End If
                    End If
                Next
                drNew("ValorAmortizadoElementoA") = 0
                drNew("ValorTotalPlusvaliaA") = 0
                drNew("ValorAmortizadoPlusvaliaA") = 0
                drNew("ValorNetoContablePlusvaliaA") = 0
                drNew("ValorAmortizadoElementoB") = 0
                drNew("ValorNetoContableElementoB") = 0
                drNew("ValorTotalPlusvaliaB") = 0
                drNew("ValorAmortizadoPlusvaliaB") = 0
                drNew("ValorNetoContablePlusvaliaB") = 0

                Dim drGAmortiz As DataRow = New GrupoAmortizacion().GetItemRow(dr("IDGrupoAmortizacion"))
                If Not IsNothing(drGAmortiz) Then
                    Dim drTipoAmortiz As DataRow
                    If Length(dr("IDCodigoAmortizacionContable")) > 0 Then
                        drTipoAmortiz = New TipoAmortizacionCabecera().GetItemRow(dr("IDCodigoAmortizacionContable"))
                    Else : drTipoAmortiz = New TipoAmortizacionCabecera().GetItemRow(drGAmortiz("IDTipoAmortiz"))
                    End If
                    drNew("IDCodigoAmortizacionContable") = drTipoAmortiz("IdTipoAmortizacion")
                    drNew("IDCodigoAmortizacionTecnica") = drTipoAmortiz("IdTipoAmortizacion")
                    drNew("IDCodigoAmortizacionFiscal") = drTipoAmortiz("IdTipoAmortizacion")
                    drNew("VidaContableElemento") = drTipoAmortiz("VidaUtil")
                    drNew("VidaTecnicaElemento") = drTipoAmortiz("VidaUtil")
                    drNew("VidaFiscalElemento") = drTipoAmortiz("VidaUtil")

                    Dim filAmortizLinea As New Filter
                    filAmortizLinea.Add("IDTipoAmortizacion", FilterOperator.Equal, drTipoAmortiz("IdTipoAmortizacion"), FilterType.String)
                    Dim dtTipoAmortizLinea As DataTable = New TipoAmortizacionLinea().Filter(filAmortizLinea, "NAño", "TOP 1 PorcentajeAmortizar")
                    If Not IsNothing(dtTipoAmortizLinea) AndAlso dtTipoAmortizLinea.Rows.Count > 0 Then
                        dblPorcentaje = dtTipoAmortizLinea.Rows(0)("PorcentajeAmortizar")
                    Else : dblPorcentaje = 0
                    End If

                    drNew("PorcentajeAnualContable") = dblPorcentaje
                    drNew("PorcentajeAnualTecnico") = dblPorcentaje
                    drNew("PorcentajeAnualFiscal") = dblPorcentaje

                    DblValorAmortizA = drNew("ValorTotalRevalElementoA") - drNew("ValorResidualA")
                    dblDotacionA = (DblValorAmortizA * (dblPorcentaje / 100)) / 12
                    dblDotacionB = dblDotacionA * DblCambioAB

                    dblDotacionA = xRound(dblDotacionA, LngDecimalesA)
                    dblDotacionB = xRound(dblDotacionB, LngDecimalesB)

                    drNew("DotacionContableElementoA") = dblDotacionA
                    drNew("DotacionTecnicaElementoA") = dblDotacionA
                    drNew("DotacionFiscalElementoA") = dblDotacionA
                    drNew("DotacionContableElementoB") = dblDotacionB
                    drNew("DotacionTecnicaElementoB") = dblDotacionB
                    drNew("DotacionFiscalElementoB") = dblDotacionB
                End If
                drNew("MesUltimoContabilizado") = 0
                drNew("AñoUltimoContabilizado") = 0
                drNew("FechaUltimaContabilizacion") = System.DBNull.Value
                drNew("ValorUltimoContabilizadoA") = 0
                drNew("ValorUltimoContabilizadoB") = 0
                drNew("ValorReposicionA") = 0
                drNew("ValorReposicionActualizadoA") = 0
                drNew("ValorReposicionB") = 0
                drNew("ValorReposicionActualizadoB") = 0
                dtNew.Rows.Add(drNew)
            Next
            If Not dtNew Is Nothing AndAlso dtNew.Rows.Count > 0 Then BusinessHelper.UpdateTable(dtNew)
        End If
        ProcessServer.ExecuteTask(Of Object)(AddressOf Comunes.CommitTransaction, Nothing, services)
    End Function

    <Serializable()> _
    Public Class DataGetVidaPorcenDot
        Public Dt As DataTable
        Public NuevoTipoAmort As String
        Public FechaReval As Date

        Public Sub New(ByVal Dt As DataTable, ByVal NuevoTipoAmort As String, Optional ByVal FechaReval As Date = cnMinDate)
            Me.Dt = Dt
            Me.NuevoTipoAmort = NuevoTipoAmort
            Me.FechaReval = FechaReval
        End Sub
    End Class

    <Task()> Public Shared Function GetVidaPorcentajeDotacion(ByVal data As DataGetVidaPorcenDot, ByVal services As ServiceProvider) As DataTable
        Dim dt As New DataTable
        dt.Columns.Add("Vida", GetType(Integer))
        dt.Columns.Add("Porcentaje", GetType(Double))
        dt.Columns.Add("Dotacion", GetType(Double))

        If Not data.Dt.Rows(0) Is Nothing Then
            Dim StData As New DataObtenerDatosTipoAmort(String.Empty, data.NuevoTipoAmort)
            Dim DatosAmort As StDatosAmort = ProcessServer.ExecuteTask(Of DataObtenerDatosTipoAmort, StDatosAmort)(AddressOf ObtenerDatosTipoAmort, StData, services)
            Dim IntVida As Integer = DatosAmort.Vida
            Dim DblPorcen As Double = CDbl(DatosAmort.Porcentaje)
            Dim StDotValor As New DataDotacionPorValor(data.Dt.Rows(0)("FechaInicioContabilizacion"), IIf(data.FechaReval <> cnMinDate, data.FechaReval, Today), _
                                                       data.Dt.Rows(0)("ValorTotalRevalElementoA") - data.Dt.Rows(0)("ValorResidualA"), data.NuevoTipoAmort, IntVida, False)
            Dim DblDotacion As Double = ProcessServer.ExecuteTask(Of DataDotacionPorValor, Double)(AddressOf DotacionPorValor, StDotValor, services)
            Dim drNew As DataRow = dt.NewRow
            drNew("Vida") = IntVida
            drNew("Porcentaje") = DblPorcen
            drNew("Dotacion") = DblDotacion
            dt.Rows.Add(drNew)
        End If
        Return dt
    End Function

    <Serializable()> _
    Public Class DataSalvarAmortizPdte
        Public DtApdte As DataTable
        Public Maquina As String
        Public IDPrograma As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal DtApdte As DataTable, ByVal Maquina As String, ByVal IDPrograma As String)
            Me.DtApdte = DtApdte
            Me.Maquina = Maquina
            Me.IDPrograma = IDPrograma
        End Sub
    End Class

    <Task()> Public Shared Sub SalvarAmortizPendiente(ByVal data As DataSalvarAmortizPdte, ByVal services As ServiceProvider)
        Dim StrClave As String
        For Each Dr As DataRow In data.DtApdte.Select
            StrClave = StrClave & Dr("IdElemento") & "@" & Trim(Str(Dr("AmortizacionPdteA"))) _
                       & "@" & Trim(Str(Dr("AjusteRevalorizacionA"))) & "@" & _
                       Trim(Str(Dr("AmortizacionPlusPdteA"))) & "@" & Trim(Str(Dr("AmortizacionPdteB"))) _
                       & "@" & Trim(Str(Dr("AjusteRevalorizacionB"))) & "@" & _
                       Trim(Str(Dr("AmortizacionPlusPdteB"))) & "#"
        Next
        If Length(StrClave) Then StrClave = Left(StrClave, Length(StrClave) - 1)
    End Sub

    <Serializable()> _
    Public Class DataDotPorElem
        Public IDElemento As String
        Public FechaDotacion As Date
        Public PorDia As Boolean

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDElemento As String, ByVal FechaDotacion As Date, Optional ByVal PorDia As Boolean = False)
            Me.IDElemento = IDElemento
            Me.FechaDotacion = FechaDotacion
            Me.PorDia = PorDia
        End Sub
    End Class

    <Task()> Public Shared Function DotacionesPorElemento(ByVal data As DataDotPorElem, ByVal services As ServiceProvider) As String
        'Comienzo del Cuerpo de la Función
        DotacionesPorElemento = "Contable=0;Tecnica=0;Fiscal=0;"
        Dim DtElemento As DataTable = New ElementoAmortizable().SelOnPrimaryKey(data.IDElemento)
        If Not DtElemento Is Nothing Then
            Dim DtTipoAmort As DataTable = New TipoAmortizacionCabecera().SelOnPrimaryKey(DtElemento.Rows(0)("IDCodigoAmortizacionContable"))
            If Not DtTipoAmort.Rows.Count > 0 Then
                Dim StDot As New DataDotacionPorValor(DtElemento.Rows(0)("FechaInicioContabilizacion"), data.FechaDotacion, DtElemento.Rows(0)("ValorTotalRevalElementoA") - DtElemento.Rows(0)("ValorResidualA"), DtElemento.Rows(0)("IDCodigoAmortizacionContable"), DtTipoAmort.Rows(0)("VidaUtil"), data.PorDia)
                Dim DblDotacionContable As Double = ProcessServer.ExecuteTask(Of DataDotacionPorValor, Double)(AddressOf DotacionPorValor, StDot, services)
                StDot.TipoAmortizacion = DtElemento.Rows(0)("IDCodigoAmortizacionTecnica")
                Dim DblDotacionTecnica As Double = ProcessServer.ExecuteTask(Of DataDotacionPorValor, Double)(AddressOf DotacionPorValor, StDot, services)
                StDot.TipoAmortizacion = DtElemento.Rows(0)("IDCodigoAmortizacionFiscal")
                Dim DblDotacionFiscal As Double = ProcessServer.ExecuteTask(Of DataDotacionPorValor, Double)(AddressOf DotacionPorValor, StDot, services)
                DotacionesPorElemento = Replace(DotacionesPorElemento, "Contable=0", "Contable=" & DblDotacionContable)
                DotacionesPorElemento = Replace(DotacionesPorElemento, "Tecnica=0", "Tecnica=" & DblDotacionTecnica)
                DotacionesPorElemento = Replace(DotacionesPorElemento, "Fiscal=0", "Fiscal=" & DblDotacionFiscal)
            End If

        End If
    End Function

    <Serializable()> _
    Public Class DataDotacionPorValor
        Public FechaInicio As Date
        Public FechaDotacion As Date
        Public ValorTotal As Double
        Public TipoAmortizacion As String
        Public Vida As Integer
        Public PorDia As Boolean

        Public Sub New()
        End Sub

        Public Sub New(ByVal FechaInicio As Date, ByVal FechaDotacion As Date, ByVal ValorTotal As Double, ByVal TipoAmortizacion As String, ByVal Vida As Integer, Optional ByVal PorDia As Boolean = False)
            Me.FechaInicio = FechaInicio
            Me.FechaDotacion = FechaDotacion
            Me.ValorTotal = ValorTotal
            Me.TipoAmortizacion = TipoAmortizacion
            Me.Vida = Vida
            Me.PorDia = PorDia
        End Sub
    End Class

    <Task()> Public Shared Function DotacionPorValor(ByVal data As DataDotacionPorValor, ByVal services As ServiceProvider) As Double
        'Mes de cálculo de la dotación. En base al número de meses pasados desde el inicio de la dotación 
        'hasta la fecha de cálculo de la misma.
        'Si la amortización es por día, representa el número de días pasados desde la fecha de inicio.
        Dim IntMesDotacion As Integer = DateDiff(DateInterval.Month, data.FechaInicio, data.FechaDotacion) + 1
        If data.FechaInicio.Day <> 1 And Not data.PorDia Then
            'Si no es 1 la fecha de inicio, entonces la fecha de inicio es el 1 del mes siguiente.
            data.FechaInicio = New Date(data.FechaInicio.AddMonths(1).Year, data.FechaInicio.AddMonths(1).Month, 1)
            IntMesDotacion = DateDiff(DateInterval.Month, data.FechaInicio, data.FechaDotacion) + 1
        ElseIf data.PorDia Then
            IntMesDotacion = DateDiff(DateInterval.Day, data.FechaInicio, data.FechaDotacion)
        End If

        If IntMesDotacion > data.Vida Or IntMesDotacion < 0 Then
            Return 0
        Else
            Dim IntFactorDia As Integer = 0
            If data.PorDia Then
                IntFactorDia = DateTime.DaysInMonth(Year(data.FechaDotacion), Month(data.FechaDotacion))
            Else : IntFactorDia = 1
            End If
            'Obtenemos las monedas internas para poder aplicar los decimales correctamente.
            Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
            Dim MonInfoA As MonedaInfo = Monedas.MonedaA

            Dim DblDec As Double = MonInfoA.NDecimalesImporte
            Dim DtTipoAmortLinea As DataTable = New TipoAmortizacionLinea().Filter(New FilterItem("IdTipoAmortizacion", FilterOperator.Equal, data.TipoAmortizacion), "NAño")
            If Not DtTipoAmortLinea Is Nothing AndAlso DtTipoAmortLinea.Rows.Count > 0 Then
                Dim IntPorcenAux As Integer = 0
                Dim IntMesAux As Integer = 0
                For Each Dr As DataRow In DtTipoAmortLinea.Select
                    If (IntMesAux + 12) >= data.Vida Then             'Estamos en el último año. Amortización del resto
                        Return xRound(data.ValorTotal * (100 - IntPorcenAux) / ((data.Vida - IntMesAux) * 100 * IntFactorDia), DblDec)
                    ElseIf (IntMesAux + 12) >= IntMesDotacion Then  'Estamos en el año del mes de la dotación
                        Return xRound(data.ValorTotal * Dr("PorcentajeAmortizar") / (12 * 100 * IntFactorDia), DblDec)
                    Else
                        IntPorcenAux += Dr("PorcentajeAmortizar")
                        IntMesAux += 12
                    End If
                Next
            End If
        End If
    End Function

    <Serializable()> _
    Public Class DataDotacionPorElem
        Public IDElemento As String
        Public FechaDotacion As Date
        Public PorDia As Boolean

        Public Sub New(ByVal IDElemento As String, ByVal FechaDotacion As Date, Optional ByVal PorDia As Boolean = False)
            Me.IDElemento = IDElemento
            Me.FechaDotacion = FechaDotacion
            Me.PorDia = PorDia
        End Sub
    End Class

    <Task()> Public Shared Function DotacionPorElemento(ByVal data As DataDotacionPorElem, ByVal services As ServiceProvider) As Double
        DotacionPorElemento = 0
        Dim DtElemento As DataTable = New ElementoAmortizable().SelOnPrimaryKey(data.IDElemento)
        If Not DtElemento Is Nothing Then
            Dim DtTipoAmort As DataTable = New TipoAmortizacionCabecera().SelOnPrimaryKey(DtElemento.Rows(0)("IDCodigoAmortizacionContable"))
            If Not DtTipoAmort Is Nothing Then
                Dim StDotPorValor As New DataDotacionPorValor(DtElemento.Rows(0)("FechaInicioContabilizacion"), data.FechaDotacion, DtElemento.Rows(0)("ValorTotalRevalElementoA") - DtElemento.Rows(0)("ValorResidualA"), DtElemento.Rows(0)("IDCodigoAmortizacionContable"), DtTipoAmort.Rows(0)("VidaUtil"), data.PorDia)
                DotacionPorElemento = ProcessServer.ExecuteTask(Of DataDotacionPorValor, Double)(AddressOf DotacionPorValor, StDotPorValor, services)
            End If
        End If
    End Function

#End Region

#Region "Calculos Amortizacion"

    <Serializable()> _
    Public Class DataCalcAmort
        Public IDElem As String
        Public Año As Boolean

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDElem As String, Optional ByVal Año As Boolean = False)
            Me.IDElem = IDElem
            Me.Año = Año
        End Sub
    End Class

    <Task()> Public Shared Function CalcularAmortizacion(ByVal data As DataCalcAmort, ByVal services As ServiceProvider) As DataTable
        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
        Dim MonInfoA As MonedaInfo = Monedas.MonedaA
        Dim LngDecimalesA As Integer = MonInfoA.NDecimalesImporte
        Dim DtAmort As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf CrearDtAmort, Nothing, services)
        Dim StrTexto As String = "Mes1 =1; Amortizacion1= @Mes1@;Mes2 = 2; Amortizacion2= @Mes2@; " & _
        "Mes3 =3; Amortizacion3= @Mes3@;Mes4 = 4; Amortizacion4= @Mes4@; " & _
        "Mes5 =5; Amortizacion5= @Mes5@;Mes6 = 6; Amortizacion6= @Mes6@; " & _
        "Mes7 =7; Amortizacion7= @Mes7@;Mes8 = 8; Amortizacion8= @Mes8@; " & _
        "Mes9 =9; Amortizacion9= @Mes9@;Mes10 = 10; Amortizacion10= @Mes10@; " & _
        "Mes11 =11; Amortizacion11= @Mes11@;Mes12 = 12; Amortizacion12= @Mes12@"
        Dim DtmFecha, DtmFechaInicio, DtmFechaFin As DateTime
        Dim StrAmortContMensual, StrAmortContFiscalMensual, _
        StrAmortContTecnicaMensual, StrAmortContRealMensual As String
        Dim DblValor, DblAcumAmortMes As Double
        Dim IntVidaAmort As Integer
        Dim IntVidaUtil As Short
        Dim DblValorAmortizado, DblValorAmortizar, DblAcumValorAmortizar, DblAcumValorRealizada As Double
        Dim DblValorAmortizarFiscal, DblAcumValorAmortFiscal, DblValorAmortizadoFiscal As Double
        Dim DblValorAmortizarTecnica, DblAcumValorAmortTecnica, DblValorAmortizadoTecnica As Double
        Dim DblValorSimReal As Double
        Dim BlnFin As Boolean
        Dim DtElem As DataTable = New ElementoAmortizable().SelOnPrimaryKey(data.IDElem)
        Dim ClsAmortReg As New AmortizacionRegistro
        Dim DteFechaInicioCont, DteFechaMax, DtmFechaBaja As Date
        Dim DblAcumValorAmortizarFiscal, DblAcumValorAmortizarTecnica As Double
        Dim BlnBaja As Boolean

        If Not DtElem Is Nothing AndAlso DtElem.Rows.Count > 0 Then
            Dim DtTotalReal As DataTable = New BE.DataEngine().Filter("vCtlTotalAmortizadoPorAño", "*", "IDElemento = '" & data.IDElem & "'")
            Dim DtCambio As DataTable = New BE.DataEngine().Filter("tbElementoRevalorizacion", "*", "IDElemento = '" & data.IDElem & "'", "IDLineaRevalorizacion DESC")
            If Not DtCambio Is Nothing AndAlso DtCambio.Rows.Count > 0 Then
                DtmFecha = DtElem.Rows(0)("FechaInicioContabilizacion")

                Dim MesAño As New StMesAño
                Dim LngMesActual, LngAñoActual As Integer

                MesAño = ObtenerMesAño(DtmFecha, DtElem.Rows(0)("PorMeses"))
                LngMesActual = CInt(MesAño.Mes)
                LngAñoActual = CInt(MesAño.Año)
                If DtElem.Rows(0)("PorMeses") Then
                    DtmFechaInicio = New Date(LngAñoActual, LngMesActual, 1)
                Else
                    DtmFechaInicio = DtmFecha
                End If
                DteFechaInicioCont = DtmFechaInicio
                DtmFechaFin = New DateTime(DtmFechaInicio.Year, DtmFechaInicio.Month, DateTime.DaysInMonth(DtmFechaInicio.Year, DtmFechaInicio.Month))

                StrAmortContMensual = StrTexto
                StrAmortContFiscalMensual = StrTexto
                StrAmortContTecnicaMensual = StrTexto
                StrAmortContRealMensual = StrTexto

                IntVidaAmort = DtCambio.Rows(0)("VidaUtilFecha")
                If Not DtElem.Rows(0)("PorMeses") Then
                    If DtmFecha.Day > 1 Then IntVidaAmort += 1
                End If
                If data.Año Then DtmFechaFin = New DateTime(DtmFechaInicio.Year, 12, 31)

                DteFechaMax = DteFechaInicioCont.AddMonths(IntVidaAmort)

                If Nz(DtElem.Rows(0)("baja"), False) Then
                    BlnBaja = True
                    DtmFechaBaja = DtElem.Rows(0)("FechaBaja")
                End If

                While IntVidaAmort > 0
                    If BlnBaja AndAlso DtmFechaBaja < DtmFechaInicio Then
                        DblValorAmortizar = 0
                        DblValorAmortizarTecnica = 0
                        DblValorAmortizarFiscal = 0
                        IntVidaAmort = 0
                    Else
                        If DtElem.Rows(0)("PorMeses") Then
                            If Not BlnFin Then
                                Dim StAmort As New DataCalcPropuestas(DtmFecha, DtmFechaInicio, DtmFechaFin, DtElem.Rows(0)("ValorTotalRevalElementoA"), DblValorAmortizado, _
                                DtElem.Rows(0)("ValorResidualA"), DtCambio.Rows(0)("IDTipoAmortizacionFecha"), DtCambio.Rows(0)("VidaUtilFecha"))
                                DblValorAmortizar = ProcessServer.ExecuteTask(Of DataCalcPropuestas, Double)(AddressOf CalcularPropuestaPorMeses, StAmort, services)
                                Dim StAmortTec As New DataCalcPropuestas(DtmFecha, DtmFechaInicio, DtmFechaFin, DtElem.Rows(0)("ValorTotalRevalElementoA"), DblValorAmortizadoTecnica, _
                                DtElem.Rows(0)("ValorResidualA"), DtElem.Rows(0)("IDCodigoAmortizacionTecnica"), DtElem.Rows(0)("VidaTecnicaElemento"))
                                DblValorAmortizarTecnica = ProcessServer.ExecuteTask(Of DataCalcPropuestas, Double)(AddressOf CalcularPropuestaPorMeses, StAmortTec, services)
                                Dim StAmortFis As New DataCalcPropuestas(DtmFecha, DtmFechaInicio, DtmFechaFin, DtElem.Rows(0)("ValorTotalRevalElementoA"), DblValorAmortizadoFiscal, _
                                DtElem.Rows(0)("ValorResidualA"), DtElem.Rows(0)("IDCodigoAmortizacionFiscal"), DtElem.Rows(0)("VidaFiscalElemento"))
                                DblValorAmortizarFiscal = ProcessServer.ExecuteTask(Of DataCalcPropuestas, Double)(AddressOf CalcularPropuestaPorMeses, StAmortFis, services)
                            Else
                                DblValorAmortizar = DtElem.Rows(0)("ValorTotalRevalElementoA") - DtElem.Rows(0)("ValorResidualA") - DblValorAmortizado
                                DblValorAmortizarTecnica = DtElem.Rows(0)("ValorTotalRevalElementoA") - DtElem.Rows(0)("ValorResidualA") - DblValorAmortizadoTecnica
                                DblValorAmortizarFiscal = DtElem.Rows(0)("ValorTotalRevalElementoA") - DtElem.Rows(0)("ValorResidualA") - DblValorAmortizadoFiscal
                            End If
                        Else
                            If Not BlnFin Then
                                Dim StAmort As New DataCalcPropuestas(DtmFecha, DtmFechaInicio, DtmFechaFin, DtElem.Rows(0)("ValorTotalRevalElementoA"), DblValorAmortizado, _
                                DtElem.Rows(0)("ValorResidualA"), DtCambio.Rows(0)("IDTipoAmortizacionFecha"), DtCambio.Rows(0)("VidaUtilFecha"))
                                DblValorAmortizar = ProcessServer.ExecuteTask(Of DataCalcPropuestas, Double)(AddressOf CalcularPropuestaPorDias, StAmort, services)
                                Dim StAmortTec As New DataCalcPropuestas(DtmFecha, DtmFechaInicio, DtmFechaFin, DtElem.Rows(0)("ValorTotalRevalElementoA"), DblValorAmortizadoTecnica, _
                                DtElem.Rows(0)("ValorResidualA"), DtElem.Rows(0)("IDCodigoAmortizacionTecnica"), DtElem.Rows(0)("VidaTecnicaElemento"))
                                DblValorAmortizarTecnica = ProcessServer.ExecuteTask(Of DataCalcPropuestas, Double)(AddressOf CalcularPropuestaPorDias, StAmortTec, services)
                                Dim StAmortFis As New DataCalcPropuestas(DtmFecha, DtmFechaInicio, DtmFechaFin, DtElem.Rows(0)("ValorTotalRevalElementoA"), DblValorAmortizadoFiscal, _
                                DtElem.Rows(0)("ValorResidualA"), DtElem.Rows(0)("IDCodigoAmortizacionFiscal"), DtElem.Rows(0)("VidaFiscalElemento"))
                                DblValorAmortizarFiscal = ProcessServer.ExecuteTask(Of DataCalcPropuestas, Double)(AddressOf CalcularPropuestaPorDias, StAmortFis, services)
                            Else
                                DblValorAmortizar = DtElem.Rows(0)("ValorTotalRevalElementoA") - DtElem.Rows(0)("ValorResidualA") - DblValorAmortizado
                                DblValorAmortizarTecnica = DtElem.Rows(0)("ValorTotalRevalElementoA") - DtElem.Rows(0)("ValorResidualA") - DblValorAmortizadoTecnica
                                DblValorAmortizarFiscal = DtElem.Rows(0)("ValorTotalRevalElementoA") - DtElem.Rows(0)("ValorResidualA") - DblValorAmortizadoFiscal
                            End If
                        End If
                        DblValorAmortizar = xRound(DblValorAmortizar, LngDecimalesA)
                        DblValorAmortizarFiscal = xRound(DblValorAmortizarFiscal, LngDecimalesA)
                        DblValorAmortizarTecnica = xRound(DblValorAmortizarTecnica, LngDecimalesA)

                        DblAcumValorAmortizar = DblAcumValorAmortizar + DblValorAmortizar
                        DblAcumValorAmortizarFiscal = DblAcumValorAmortizarFiscal + DblValorAmortizarFiscal
                        DblAcumValorAmortizarTecnica = DblAcumValorAmortizarTecnica + DblValorAmortizarTecnica

                        DblValorAmortizado = DblValorAmortizado + DblValorAmortizar
                        DblValorAmortizadoTecnica = DblValorAmortizadoTecnica + DblValorAmortizarTecnica
                        DblValorAmortizadoFiscal = DblValorAmortizadoFiscal + DblValorAmortizarFiscal

                        StrAmortContMensual = Replace(StrAmortContMensual, "@Mes" & Month(DtmFechaInicio) & "@", DblValorAmortizar)
                        StrAmortContTecnicaMensual = Replace(StrAmortContTecnicaMensual, "@Mes" & Month(DtmFechaInicio) & "@", DblValorAmortizarTecnica)
                        StrAmortContFiscalMensual = Replace(StrAmortContFiscalMensual, "@Mes" & Month(DtmFechaInicio) & "@", DblValorAmortizarFiscal)

                        Dim DrTReal() As DataRow
                        If Not DtTotalReal Is Nothing AndAlso DtTotalReal.Rows.Count > 0 Then
                            DrTReal = DtTotalReal.Select("AñoContabilizacion = " & DtmFechaInicio.Year)
                            If DrTReal.Length > 0 Then
                                If DrTReal(0)("UltimoMesContabilizacion") < DtmFechaFin.Month Then
                                    DblValorSimReal += DblValorAmortizar
                                End If
                            Else
                                DblValorSimReal += DblValorAmortizar
                            End If
                        End If


                        Dim FilReg As New Filter
                        FilReg.Add("IDElemento", FilterOperator.Equal, data.IDElem, FilterType.String)
                        FilReg.Add("AñoContabilizacion", FilterOperator.Equal, DtmFechaInicio.Year, FilterType.Numeric)
                        FilReg.Add("MesContabilizacion", FilterOperator.Equal, DtmFechaInicio.Month, FilterType.Numeric)
                        Dim DtAmortReg As DataTable = ClsAmortReg.Filter(FilReg)
                        If Not DtAmortReg Is Nothing Then
                            If DtAmortReg.Rows.Count > 0 Then
                                DblAcumAmortMes = 0
                                For Each Dr As DataRow In DtAmortReg.Select
                                    DblAcumAmortMes += Dr("ValorAmortizadoA")
                                    DblAcumValorRealizada += Dr("ValorAmortizadoA")
                                Next
                                StrAmortContRealMensual = Replace(StrAmortContRealMensual, "@Mes" & DtmFechaInicio.Month & "@", CStr(DblAcumAmortMes))
                            Else
                                StrAmortContRealMensual = Replace(StrAmortContRealMensual, "@Mes" & DtmFechaInicio.Month & "@", CStr(0))
                            End If
                        End If
                        If DtmFechaFin.Month = 12 OrElse IntVidaAmort = 1 Then
                            If Not DrTReal Is Nothing AndAlso DrTReal.Length > 0 Then
                                DblValorSimReal += Nz(DrTReal(0)("TotalAmortizadoAñoA"), 0)
                            End If
                            Dim DrNew As DataRow = DtAmort.NewRow()
                            DrNew("IDElemento") = data.IDElem
                            DrNew("Año") = DtmFechaInicio.Year
                            DrNew("AmortContable") = DblAcumValorAmortizar
                            DrNew("AmortFiscal") = DblAcumValorAmortizarFiscal
                            DrNew("AmortTecnica") = DblAcumValorAmortizarTecnica
                            DrNew("AmortRealizada") = DblAcumValorRealizada
                            DrNew("AmortSimAño") = DblValorSimReal
                            DrNew("ValorNeto") = 0
                            DrNew("AmortAño") = DblAcumValorAmortizar
                            StrAmortContMensual = Replace(StrAmortContMensual, "@Mes1@", CStr(0))
                            StrAmortContMensual = Replace(StrAmortContMensual, "@Mes2@", CStr(0))
                            StrAmortContMensual = Replace(StrAmortContMensual, "@Mes3@", CStr(0))
                            StrAmortContMensual = Replace(StrAmortContMensual, "@Mes4@", CStr(0))
                            StrAmortContMensual = Replace(StrAmortContMensual, "@Mes5@", CStr(0))
                            StrAmortContMensual = Replace(StrAmortContMensual, "@Mes6@", CStr(0))
                            StrAmortContMensual = Replace(StrAmortContMensual, "@Mes7@", CStr(0))
                            StrAmortContMensual = Replace(StrAmortContMensual, "@Mes8@", CStr(0))
                            StrAmortContMensual = Replace(StrAmortContMensual, "@Mes9@", CStr(0))
                            StrAmortContMensual = Replace(StrAmortContMensual, "@Mes10@", CStr(0))
                            StrAmortContMensual = Replace(StrAmortContMensual, "@Mes11@", CStr(0))
                            StrAmortContMensual = Replace(StrAmortContMensual, "@Mes12@", CStr(0))
                            DrNew("AmortContableMensual") = StrAmortContMensual

                            StrAmortContTecnicaMensual = Replace(StrAmortContTecnicaMensual, "@Mes1@", 0)
                            StrAmortContTecnicaMensual = Replace(StrAmortContTecnicaMensual, "@Mes2@", 0)
                            StrAmortContTecnicaMensual = Replace(StrAmortContTecnicaMensual, "@Mes3@", 0)
                            StrAmortContTecnicaMensual = Replace(StrAmortContTecnicaMensual, "@Mes4@", 0)
                            StrAmortContTecnicaMensual = Replace(StrAmortContTecnicaMensual, "@Mes5@", 0)
                            StrAmortContTecnicaMensual = Replace(StrAmortContTecnicaMensual, "@Mes6@", 0)
                            StrAmortContTecnicaMensual = Replace(StrAmortContTecnicaMensual, "@Mes7@", 0)
                            StrAmortContTecnicaMensual = Replace(StrAmortContTecnicaMensual, "@Mes8@", 0)
                            StrAmortContTecnicaMensual = Replace(StrAmortContTecnicaMensual, "@Mes9@", 0)
                            StrAmortContTecnicaMensual = Replace(StrAmortContTecnicaMensual, "@Mes10@", 0)
                            StrAmortContTecnicaMensual = Replace(StrAmortContTecnicaMensual, "@Mes11@", 0)
                            StrAmortContTecnicaMensual = Replace(StrAmortContTecnicaMensual, "@Mes12@", 0)
                            DrNew("AmortTecnicaMensual") = StrAmortContTecnicaMensual

                            StrAmortContFiscalMensual = Replace(StrAmortContFiscalMensual, "@Mes1@", 0)
                            StrAmortContFiscalMensual = Replace(StrAmortContFiscalMensual, "@Mes2@", 0)
                            StrAmortContFiscalMensual = Replace(StrAmortContFiscalMensual, "@Mes3@", 0)
                            StrAmortContFiscalMensual = Replace(StrAmortContFiscalMensual, "@Mes4@", 0)
                            StrAmortContFiscalMensual = Replace(StrAmortContFiscalMensual, "@Mes5@", 0)
                            StrAmortContFiscalMensual = Replace(StrAmortContFiscalMensual, "@Mes6@", 0)
                            StrAmortContFiscalMensual = Replace(StrAmortContFiscalMensual, "@Mes7@", 0)
                            StrAmortContFiscalMensual = Replace(StrAmortContFiscalMensual, "@Mes8@", 0)
                            StrAmortContFiscalMensual = Replace(StrAmortContFiscalMensual, "@Mes9@", 0)
                            StrAmortContFiscalMensual = Replace(StrAmortContFiscalMensual, "@Mes10@", 0)
                            StrAmortContFiscalMensual = Replace(StrAmortContFiscalMensual, "@Mes11@", 0)
                            StrAmortContFiscalMensual = Replace(StrAmortContFiscalMensual, "@Mes12@", 0)
                            DrNew("AmortFiscalMensual") = StrAmortContFiscalMensual

                            StrAmortContRealMensual = Replace(StrAmortContRealMensual, "@Mes1@", CStr(0))
                            StrAmortContRealMensual = Replace(StrAmortContRealMensual, "@Mes2@", CStr(0))
                            StrAmortContRealMensual = Replace(StrAmortContRealMensual, "@Mes3@", CStr(0))
                            StrAmortContRealMensual = Replace(StrAmortContRealMensual, "@Mes4@", CStr(0))
                            StrAmortContRealMensual = Replace(StrAmortContRealMensual, "@Mes5@", CStr(0))
                            StrAmortContRealMensual = Replace(StrAmortContRealMensual, "@Mes6@", CStr(0))
                            StrAmortContRealMensual = Replace(StrAmortContRealMensual, "@Mes7@", CStr(0))
                            StrAmortContRealMensual = Replace(StrAmortContRealMensual, "@Mes8@", CStr(0))
                            StrAmortContRealMensual = Replace(StrAmortContRealMensual, "@Mes9@", CStr(0))
                            StrAmortContRealMensual = Replace(StrAmortContRealMensual, "@Mes10@", CStr(0))
                            StrAmortContRealMensual = Replace(StrAmortContRealMensual, "@Mes11@", CStr(0))
                            StrAmortContRealMensual = Replace(StrAmortContRealMensual, "@Mes12@", CStr(0))
                            DrNew("AmortRealizadaMensual") = StrAmortContRealMensual

                            DblAcumValorAmortizar = 0
                            DblAcumValorRealizada = 0
                            DblValorSimReal = 0
                            StrAmortContMensual = StrTexto
                            StrAmortContRealMensual = StrTexto
                            StrAmortContTecnicaMensual = StrTexto
                            StrAmortContFiscalMensual = StrTexto
                            DtAmort.Rows.Add(DrNew)
                        End If
                        If data.Año Then
                            If IntVidaAmort = 1 Then
                                IntVidaAmort = 0
                            Else
                                IntVidaAmort -= DateDiff(DateInterval.Month, DtmFechaInicio, DtmFechaFin) - 1
                            End If
                            DtmFechaInicio = DtmFechaFin.AddDays(1)
                            If IntVidaAmort > 0 AndAlso IntVidaAmort <= 12 Then
                                IntVidaAmort = 1
                                DtmFechaFin = CDate(DtElem.Rows(0)("FechaInicioContabilizacion")).AddMonths(DtCambio.Rows(0)("VidaUtilFecha"))
                                BlnFin = True
                            Else
                                Dim DtmFechaNueva As New Date(DtmFechaInicio.Year, 12, 31)
                                DtmFechaFin = DtmFechaNueva
                            End If
                        Else
                            IntVidaAmort -= 1
                            DtmFechaInicio = DtmFechaFin.AddDays(1)
                            If IntVidaAmort = 1 Then
                                DtmFechaFin = CDate(DtElem.Rows(0)("FechaInicioContabilizacion")).AddMonths(DtCambio.Rows(0)("VidaUtilFecha"))
                                BlnFin = True
                            Else
                                DtmFechaFin = DtmFechaFin.AddMonths(1)
                                Dim DtmFechaNueva As New Date(DtmFechaFin.Year, DtmFechaFin.Month, DtmFechaFin.DaysInMonth(DtmFechaFin.Year, DtmFechaFin.Month))
                                DtmFechaFin = DtmFechaNueva
                            End If
                        End If
                    End If
                End While
            End If
        End If
        Return DtAmort
    End Function

    <Serializable()> _
    Public Class DataCalcPropuestas
        Public FInicioCont As Date
        Public FechaInicial As Date
        Public FechaFinal As Date
        Public ValorRevalorizado As Double
        Public ValorAmortizado As Double
        Public ValorResidual As Double
        Public IDTipoAmortizacion As String
        Public VidaUtil As Integer

        Public Sub New()
        End Sub

        Public Sub New(ByVal FInicioCont As Date, ByVal FechaInicial As Date, ByVal FechaFinal As Date, _
                       ByVal ValorRevalorizado As Double, ByVal ValorAmortizado As Double, ByVal ValorResidual As Double, _
                       ByVal IDTipoAmortizacion As String, ByVal VidaUtil As Integer)
            Me.FInicioCont = FInicioCont
            Me.FechaInicial = FechaInicial
            Me.FechaFinal = FechaFinal
            Me.ValorAmortizado = ValorAmortizado
            Me.ValorRevalorizado = ValorRevalorizado
            Me.ValorResidual = ValorResidual
            Me.IDTipoAmortizacion = IDTipoAmortizacion
            Me.VidaUtil = VidaUtil
        End Sub
    End Class

    <Task()> Public Shared Function CalcularPropuestaPorMeses(ByVal data As DataCalcPropuestas, ByVal services As ServiceProvider) As Double
        Dim j As Integer 'Contador
        Dim i As Integer 'Contador de años=Posicion en el dt de TipoAmortizacionLinea
        Dim v As Integer 'vida util
        Dim Am As Double 'amortizacion correspondiente a un mes
        Dim AP As Double 'Amortizacion pendiente
        Dim VR As Double 'Valor Revalorizado(=ValorTotalrevalorizado - ValorResidual)
        Dim VA As Double 'Valor Amortizado
        Dim ma, aa As Integer 'Dia/mes/Año de alta del elemento amortizable
        Dim mi, ai As Integer 'Dia/mes/Año iniciales del calculo
        Dim mf, af As Integer 'Dia/mes/Año finales de contabilizacion
        Dim a, m As Integer 'mes/Año utilizados en el proceso como vbles auxiliares
        Dim LngPeriodoPropuesta As Integer
        Dim LngDiasMes As Integer
        Dim DblDif As Double 'Diferencia entre la fecha inicio y la de alta del elemento

        Dim DteFechaMax As Date
        Dim DtTALinea As New DataTable
        Dim ClsTALinea As New TipoAmortizacionLinea

        Am = 0 : AP = 0


        VR = data.ValorRevalorizado - data.ValorResidual
        VA = data.ValorAmortizado

        'Primero miramos si se trata o no de amortización lineal
        Dim dblPorcen As Double = ProcessServer.ExecuteTask(Of String, Double)(AddressOf TipoAmortizacionCabecera.ObtenerAmortizacionLineal, data.IDTipoAmortizacion, services)
        If dblPorcen <> 0 Then
            LngPeriodoPropuesta = DateDiff(DateInterval.Month, data.FechaInicial, data.FechaFinal, FirstDayOfWeek.System, FirstWeekOfYear.System)
            If data.FechaInicial.Day = 1 Then
                LngPeriodoPropuesta = LngPeriodoPropuesta + 1
            End If
            If LngPeriodoPropuesta < 0 Then
                LngPeriodoPropuesta = 0
            End If
         
            Am = (VR * (dblPorcen / 100)) / 12
            AP = Am * LngPeriodoPropuesta
            ' Para los elementos comprados en negativo
            If AP > VR - VA And VR > 0 Then
                AP = VR - VA
            End If
        Else
            'Este sería el caso en que la amortización no es lineal
            DtTALinea = ClsTALinea.Filter(New FilterItem("IDTipoAmortizacion", FilterOperator.Equal, data.IDTipoAmortizacion, FilterType.String), "NAño")
            If Not DtTALinea Is Nothing AndAlso DtTALinea.Rows.Count > 0 Then
                ma = data.FInicioCont.Month : aa = data.FInicioCont.Year
                mi = data.FechaInicial.Month : ai = data.FechaInicial.Year
                mf = data.FechaFinal.Month : af = data.FechaFinal.Year
                v = data.VidaUtil

                DteFechaMax = data.FInicioCont.AddMonths(v)
                LngPeriodoPropuesta = DateDiff(DateInterval.Month, data.FechaInicial, data.FechaFinal, FirstDayOfWeek.System, FirstWeekOfYear.System)
                If LngPeriodoPropuesta >= 0 Then
                    If data.FechaInicial.AddMonths(LngPeriodoPropuesta) <= DteFechaMax Then
                        'Localizar que porcentaje hay que utilizar para calcular la propuesta de amortizacion
                        DblDif = DateDiff(DateInterval.Month, data.FInicioCont, data.FechaInicial, FirstDayOfWeek.System, FirstWeekOfYear.System)
                        'La fecha inicial del calculo debe ser mayor que la fecha de alta del elemento
                        If DblDif >= 0 Then
                            i = Int(DblDif / 12)
                            If DtTALinea.Rows.Count - 1 >= i Then
                                m = mi
                                a = ai
                                'rcsTALinea.AbsolutePosition = i
                                Am = (VR * (DtTALinea.Rows(i)("PorcentajeAmortizar") / 100)) / 12
                                DblDif = DblDif + 1
                                For j = DblDif To v
                                    If j > 12 * i Then
                                        If i < DtTALinea.Rows.Count Then
                                            Am = (VR * (DtTALinea.Rows(i)("PorcentajeAmortizar") / 100)) / 12
                                        Else
                                            Exit For
                                        End If
                                        i += 1
                                    End If
                                    If i = v / 12 And j = v Then Am = VR - AP
                                    If System.Math.Abs(AP) + Math.Abs(Am) > Math.Abs(VR) Then Am = VR - AP
                                    AP = AP + Am
                                    If System.Math.Abs(AP) + Math.Abs(VA) > Math.Abs(VR) Then AP = VR - VA : Exit For
                                    m = m + 1
                                    If m > 12 Then
                                        m = 1
                                        a += 1
                                    End If
                                    If a * 12 + m > af * 12 + mf Then Exit For
                                Next j
                            End If
                        End If
                    Else
                        ' Para los elementos comprados en negativo


                        If VR > 0 Then
                            AP = VR - VA
                        End If
                    End If
                End If
            End If
        End If
        Return AP
    End Function

    <Task()> Public Shared Function CalcularPropuestaPorDias(ByVal data As DataCalcPropuestas, ByVal services As ServiceProvider) As Double
        Dim j As Integer 'contador
        Dim k As Short 'contador
        Dim i As Integer 'contador de periodos de 12 meses completos
        Dim v As Integer 'vida util
        Dim Ad As Double 'amortizacion correspondiente a un dia
        Dim AP As Double 'Amortizacion pendiente
        Dim VR As Double 'Valor Revalorizado(=ValorTotalrevalorizado - ValorResidual)
        Dim VA As Double 'Valor Amortizado
        Dim ma, da, aa As Integer 'Dia/mes/Año de alta del elemento mortizable
        Dim mi, di, ai As Integer 'Dia/mes/Año iniciales del calculo
        Dim mf, df, af As Integer 'Dia/mes/Año finales de contabilizacion
        Dim LngPeriodoPropuesta As Integer 'Numero total de dias de la propuesta
        Dim LngDiasRestantes As Integer '1 o 0 segun el año es bisiesto o no
        Dim DifMeses As Integer 'Diferencia entre la fecha inicio y la de alta del elemento
        Dim DiasPeriodo As Integer 'Numero de dias que se llevan contabilizados dentro del periodo de 12 meses
        Dim DteFechaMax As Date
        Dim DteFechaActual As Date
        Dim DtTALinea As New DataTable
        Dim ClsTALinea As New TipoAmortizacionLinea
        Dim LngDias As Integer

        Ad = 0 : AP = 0
        VR = data.ValorRevalorizado - data.ValorResidual
        VA = data.ValorAmortizado

        Dim DtAmortizacionLineal As DataTable = AdminData.Filter("vAmortizacionLineal", "*", "IDTipoAmortizacion='" & data.IDTipoAmortizacion & "'")
        Dim dblPorcen As Double = 0

        If Not DtAmortizacionLineal Is Nothing AndAlso DtAmortizacionLineal.Rows.Count = 1 Then
            dblPorcen = DtAmortizacionLineal.Rows(0)("PorcentajeAmortizar")
        End If
        If dblPorcen <> 0 Then
            LngPeriodoPropuesta = DateDiff(Microsoft.VisualBasic.DateInterval.Day, data.FechaInicial, data.FechaFinal, FirstDayOfWeek.System, FirstWeekOfYear.System) + 1
            If LngPeriodoPropuesta >= 0 Then
                Dim intAños As Integer = DateDiff(Microsoft.VisualBasic.DateInterval.Year, data.FechaInicial, data.FechaFinal)
                Dim FechaTratar As Date = data.FechaInicial
                FechaTratar = New Date(FechaTratar.Year, 12, 31)
                For i = 0 To intAños
                    If FechaTratar > data.FechaFinal Then
                        FechaTratar = data.FechaFinal
                    End If
                    LngDias = IIf(DateTime.IsLeapYear(FechaTratar.Year), 366, 365)

                    LngPeriodoPropuesta = DateDiff(Microsoft.VisualBasic.DateInterval.Day, data.FechaInicial, FechaTratar, FirstDayOfWeek.System, FirstWeekOfYear.System) + 1
                    Ad = (VR * (dblPorcen / 100)) / LngDias
                    AP = AP + Ad * LngPeriodoPropuesta

                    data.FechaInicial = DateAdd(DateInterval.Day, 1, FechaTratar)
                    FechaTratar = New Date(data.FechaInicial.Year, 12, 31)
                Next
                '  AP = Ad * LngPeriodoPropuesta
             
                ' Para los elementos comprados en negativo
                If AP > VR - VA And VR > 0 Then
                    AP = VR - VA
                End If
            Else : AP = 0
            End If
        Else
            DtTALinea = ClsTALinea.Filter(New FilterItem("IDTipoAmortizacion", FilterOperator.Equal, data.IDTipoAmortizacion, FilterType.String), "NAño")
            If Not DtTALinea Is Nothing Then
                If DtTALinea.Rows.Count > 0 Then
                    da = data.FInicioCont.Day : ma = data.FInicioCont.Month : aa = data.FInicioCont.Year
                    di = data.FechaInicial.Day : mi = data.FechaInicial.Month : ai = data.FechaInicial.Year
                    df = data.FechaFinal.Day : mf = data.FechaFinal.Month : af = data.FechaFinal.Year
                    v = data.VidaUtil
                    VR = data.ValorRevalorizado - data.ValorResidual
                    VA = data.ValorAmortizado
                    DteFechaMax = DateAdd(Microsoft.VisualBasic.DateInterval.Month, v, data.FInicioCont)
                    LngPeriodoPropuesta = DateDiff(Microsoft.VisualBasic.DateInterval.Day, data.FechaInicial, data.FechaFinal, FirstDayOfWeek.System, FirstWeekOfYear.System) + 1
                    If LngPeriodoPropuesta >= 1 Then
                        If data.FechaInicial.AddDays(LngPeriodoPropuesta) < DteFechaMax Then
                            'Tener en cuenta:
                            '1.Los años bisiestos se tienen en cuenta en el caso de que termine un periodo completo de 12 meses (o 365 dias)
                            '2.Si el elemento no ha sido contabilizado nunca, la contabilizacion comienza el mismo dia.
                            '3.Para elementos que ya han sido contabilizados, el dia de inicio es el dia siguiente al dia que establece la fecha de ultima contabilizacion
                            '4.Dentro de esta funcion se supone que la fecha de inicio de contabilizacion ya tiene en cuenta los puntos 2 y 3.
                            'Localizar que porcentaje hay que utilizar para calcular la propuesta de amortizacion
                            DifMeses = DateDiff(Microsoft.VisualBasic.DateInterval.Month, DateAdd(DateInterval.Day, 1, data.FInicioCont), data.FechaInicial, FirstDayOfWeek.System, FirstWeekOfYear.System)
                            If DifMeses >= 0 Then
                                i = Int(DifMeses / 12)
                                DiasPeriodo = DateDiff(Microsoft.VisualBasic.DateInterval.Day, DateAdd(DateInterval.Day, 1, DateAdd(Microsoft.VisualBasic.DateInterval.Month, i * 12, data.FInicioCont)), data.FechaInicial, FirstDayOfWeek.System, FirstWeekOfYear.System)

                                If i >= 0 Then
                                    If DateTime.IsLeapYear(data.FechaInicial.Year) Then
                                        LngDias = 366
                                    Else
                                        LngDias = 365
                                    End If
                                    Ad = (VR * (DtTALinea.Rows(i)("PorcentajeAmortizar") / 100)) / LngDias

                                    '//si I=0 (menos de un año) y se ha amortizado algo
                                    If i = 0 AndAlso data.FInicioCont <> data.FechaInicial AndAlso DateDiff(Microsoft.VisualBasic.DateInterval.Month, data.FInicioCont, data.FechaInicial, FirstDayOfWeek.System, FirstWeekOfYear.System) < LngDias Then
                                        ' LngDias = DateDiff(Microsoft.VisualBasic.DateInterval.Day, DateDiff(Microsoft.VisualBasic.DateInterval.Year, DteFechaInicial, DteFInicioCont), DteFechaInicial, FirstDayOfWeek.System, FirstWeekOfYear.System)
                                        DiasPeriodo = DateDiff(Microsoft.VisualBasic.DateInterval.Day, data.FInicioCont, data.FechaInicial, FirstDayOfWeek.System, FirstWeekOfYear.System)
                                    End If
                                Else
                                    Ad = 0
                                End If

                                DteFechaActual = data.FechaInicial
                                For j = 1 To LngPeriodoPropuesta
                                    AP += Ad
                                    DiasPeriodo += 1
                                    DteFechaActual = DateAdd(Microsoft.VisualBasic.DateInterval.Day, 1, DteFechaActual)
                                    If DiasPeriodo > 0 AndAlso (DiasPeriodo Mod LngDias) = 0 Then
                                        DiasPeriodo = 0
                                        DteFechaActual = DateAdd(Microsoft.VisualBasic.DateInterval.Day, LngDiasRestantes, DteFechaActual)
                                        'Cambio de periodo
                                        i += 1
                                        'rcsTALinea.AbsolutePosition = i
                                        'If rcsTALinea.AbsolutePosition > 0 Then
                                        If i < DtTALinea.Rows.Count Then
                                            If DateTime.IsLeapYear(Year(DteFechaActual)) Then
                                                LngDias = 366
                                            Else
                                                LngDias = 365
                                            End If
                                            Ad = (VR * (DtTALinea.Rows(i)("PorcentajeAmortizar") / 100)) / LngDias
                                        Else
                                            Ad = 0
                                        End If
                                    End If
                                    If Math.Abs(AP) + Math.Abs(VA) > Math.Abs(VR) Then AP = VR - VA : Exit For
                                Next j
                            End If
                        Else
                            AP = VR - VA
                        End If
                    End If
                End If
			End If
        End If

        Return AP
    End Function

    <Serializable()> _
    Public Class DataCalcAmortTipo
        Public IDElem As String
        Public Tipo As enTipoAmort
        Public Año As Boolean

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDElem As String, ByVal Tipo As enTipoAmort, Optional ByVal Año As Boolean = False)
            Me.IDElem = IDElem
            Me.Tipo = Tipo
            Me.Año = Año
        End Sub
    End Class

    <Task()> Public Shared Function GetAmortizacionElementoSimulacion(ByVal IDElemento As String, ByVal services As ServiceProvider) As DataTable
        Dim StAmort As New ElementoAmortizable.DataCalcAmortTipo(IDElemento, BusinessEnum.enTipoAmort.enContable)
        Dim DtCont As DataTable = ProcessServer.ExecuteTask(Of ElementoAmortizable.DataCalcAmortTipo, DataTable)(AddressOf CalcularAmortizacionTipo, StAmort, services)
        Dim StTec As New ElementoAmortizable.DataCalcAmortTipo(IDElemento, BusinessEnum.enTipoAmort.enTecnica)
        Dim DtTec As DataTable = ProcessServer.ExecuteTask(Of ElementoAmortizable.DataCalcAmortTipo, DataTable)(AddressOf CalcularAmortizacionTipo, StTec, services)
        Dim StFis As New ElementoAmortizable.DataCalcAmortTipo(IDElemento, BusinessEnum.enTipoAmort.enFiscal)
        Dim DtFisc As DataTable = ProcessServer.ExecuteTask(Of ElementoAmortizable.DataCalcAmortTipo, DataTable)(AddressOf CalcularAmortizacionTipo, StFis, services)

        Dim dtAmortizaciones As New DataTable
        dtAmortizaciones.Columns.Add("Año", GetType(Integer))
        dtAmortizaciones.Columns.Add("AmortContable", GetType(Double))
        dtAmortizaciones.Columns("AmortContable").DefaultValue = 0
        dtAmortizaciones.Columns.Add("AmortFiscal", GetType(Double))
        dtAmortizaciones.Columns("AmortFiscal").DefaultValue = 0
        dtAmortizaciones.Columns.Add("AmortTecnica", GetType(Double))
        dtAmortizaciones.Columns("AmortTecnica").DefaultValue = 0
        dtAmortizaciones.Columns.Add("AmortRealizada", GetType(Double))
        dtAmortizaciones.Columns("AmortRealizada").DefaultValue = 0
        dtAmortizaciones.Columns.Add("AmortContableMensual", GetType(String))
        dtAmortizaciones.Columns.Add("AmortTecnicaMensual", GetType(String))
        dtAmortizaciones.Columns.Add("AmortFiscalMensual", GetType(String))
        dtAmortizaciones.Columns.Add("AmortRealizadaMensual", GetType(String))


        If Not DtCont Is Nothing Then
            For Each Dr As DataRow In DtCont.Select
                Dim Drs() As DataRow = dtAmortizaciones.Select("Año=" & Dr("Año"))
                If Drs.GetLength(0) = 0 Then
                    Dim DrNew As DataRow = dtAmortizaciones.NewRow
                    DrNew("Año") = Dr("Año")
                    DrNew("AmortContable") = Dr("AmortContable")
                    DrNew("AmortRealizada") = Dr("AmortRealizada")
                    DrNew("AmortContableMensual") = Dr("AmortContableMensual")
                    DrNew("AmortRealizadaMensual") = Dr("AmortRealizadaMensual")
                    dtAmortizaciones.Rows.Add(DrNew)
                Else
                    Drs(0)("AmortContable") = Dr("AmortContable")
                    Drs(0)("AmortRealizada") = Dr("AmortRealizada")
                    Drs(0)("AmortContableMensual") = Dr("AmortContableMensual")
                    Drs(0)("AmortRealizadaMensual") = Dr("AmortRealizadaMensual")
                End If
            Next
        End If
        If Not DtTec Is Nothing Then
            For Each dr As DataRow In DtTec.Select
                Dim Drs() As DataRow = dtAmortizaciones.Select("Año=" & dr("Año"))
                If Drs.GetLength(0) = 0 Then
                    Dim DrNew As DataRow = dtAmortizaciones.NewRow
                    DrNew("Año") = dr("Año")
                    DrNew("AmortTecnica") = dr("AmortTecnica")
                    DrNew("AmortTecnicaMensual") = dr("AmortTecnicaMensual")
                    dtAmortizaciones.Rows.Add(DrNew)
                Else
                    Drs(0)("AmortTecnica") = dr("AmortTecnica")
                    Drs(0)("AmortTecnicaMensual") = dr("AmortTecnicaMensual")
                End If
            Next
        End If
        If Not DtFisc Is Nothing Then
            For Each dr As DataRow In DtFisc.Select
                Dim Drs() As DataRow = dtAmortizaciones.Select("Año=" & dr("Año"))
                If Drs.GetLength(0) = 0 Then
                    Dim DrNew As DataRow = dtAmortizaciones.NewRow
                    DrNew("Año") = dr("Año")
                    DrNew("AmortFiscal") = dr("AmortFiscal")
                    DrNew("AmortFiscalMensual") = dr("AmortFiscalMensual")
                    dtAmortizaciones.Rows.Add(DrNew)
                Else
                    Drs(0)("AmortFiscal") = dr("AmortFiscal")
                    Drs(0)("AmortFiscalMensual") = dr("AmortFiscalMensual")
                End If
            Next
        End If

        Return dtAmortizaciones
    End Function


    <Serializable()> _
    Public Class DataCalcularFechas
        Public dtElementoAmortizable As DataTable
        Public Tipo As BusinessEnum.enTipoAmort

        Public Sub New(ByVal dtElementoAmortizable As DataTable, ByVal Tipo As BusinessEnum.enTipoAmort)
            Me.dtElementoAmortizable = dtElementoAmortizable
            Me.Tipo = Tipo
        End Sub
    End Class

    <Serializable()> _
    Public Class DataFechas
        Public Fecha As Date        '// Fecha Inicio Contabilización (tbMaestroElementoAmortizable)
        Public FechaInicio As Date  '// Fecha Inicio en que se empezará la próxima amortización
        Public Fechafin As Date     '// Fecha Fin hasta donde se hará la próxima amortización
    End Class

    <Task()> Public Shared Function CalcularFechasContabilizacion(ByVal data As DataCalcularFechas, ByVal services As ServiceProvider) As DataFechas
        '// Calculamos fechas de inicio y de fin de contabilización...

        Dim datFechas As New DataFechas
        Dim blnPorMeses As Boolean
        If data.dtElementoAmortizable.Rows(0)("PorMeses") Then
            'datFechas.Fecha = data.dtElementoAmortizable.Rows(0)("FechaInicioContabilizacion")
            datFechas.Fecha = FechaPrimeroMesSig(data.dtElementoAmortizable.Rows(0)("FechaCompra"), data.dtElementoAmortizable.Rows(0)("FechaInicioContabilizacion"))
        Else
            datFechas.Fecha = data.dtElementoAmortizable.Rows(0)("FechaInicioContabilizacion")
        End If
        blnPorMeses = data.dtElementoAmortizable.Rows(0)("PorMeses")

        '//Devuelve el Mes y el año en el que se empezará la próxima amortización
        Dim MesAño As StMesAño = ObtenerMesAño(datFechas.Fecha, blnPorMeses)
        Dim intMesActual As Integer = CInt(MesAño.Mes)
        Dim intAñoActual As Integer = CInt(MesAño.Año)

        If blnPorMeses Then
            datFechas.FechaInicio = New DateTime(intAñoActual, intMesActual, 1)
        Else
            datFechas.FechaInicio = datFechas.Fecha
        End If
        datFechas.Fechafin = New DateTime(datFechas.FechaInicio.Year, datFechas.FechaInicio.Month, DateTime.DaysInMonth(datFechas.FechaInicio.Year, datFechas.FechaInicio.Month))
        Return datFechas
    End Function

    <Task()> Public Shared Function CalcularAmortizacionTipo(ByVal data As DataCalcAmortTipo, ByVal services As ServiceProvider) As DataTable
        'Funcion que devuelve un rcs con la Amortizacion Contable, Fiscal, Tecnica y Realizada
        'del ElementoAmortizable que recibe como parametro

        Dim strAmortMensual As String
        Dim strAmortRealMensual As String


        Dim strTexto As String = "Mes1 =1; Amortizacion1= @Mes1@;Mes2 = 2; Amortizacion2= @Mes2@; " & _
                                 "Mes3 =3; Amortizacion3= @Mes3@;Mes4 = 4; Amortizacion4= @Mes4@; " & _
                                 "Mes5 =5; Amortizacion5= @Mes5@;Mes6 = 6; Amortizacion6= @Mes6@; " & _
                                 "Mes7 =7; Amortizacion7= @Mes7@;Mes8 = 8; Amortizacion8= @Mes8@; " & _
                                 "Mes9 =9; Amortizacion9= @Mes9@;Mes10 = 10; Amortizacion10= @Mes10@; " & _
                                 "Mes11 =11; Amortizacion11= @Mes11@;Mes12 = 12; Amortizacion12= @Mes12@"

        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
        Dim MonInfoA As MonedaInfo = Monedas.MonedaA
        Dim dtAmort As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf CrearDtAmort, Nothing, services)
        Dim DblValorAmortizado As Double
        Dim DblValorAmortizar, DblAcumValorAmortizar, DblAcumValorRealizada As Double
        Dim DblAcumAmortMes As Double
        Dim BlnFin As Boolean

        'Cargamos un Dt con el Elemento que se va a amortizar
        Dim dtElem As DataTable = New ElementoAmortizable().SelOnPrimaryKey(data.IDElem)
        If Not dtElem Is Nothing AndAlso dtElem.Rows.Count > 0 Then
            Dim dblValorTotalRevalElemento As Double = Nz(dtElem.Rows(0)("ValorTotalRevalElementoA"), 0)
            Dim dblValorResidualElemento As Double = Nz(dtElem.Rows(0)("ValorResidualA"), 0)

            Dim datCalFechas As New DataCalcularFechas(dtElem, data.Tipo)
            Dim datFechas As DataFechas = ProcessServer.ExecuteTask(Of DataCalcularFechas, DataFechas)(AddressOf CalcularFechasContabilizacion, datCalFechas, services)

            Dim blnAmortizacionPorMeses As Boolean = Nz(dtElem.Rows(0)("PorMeses"), True)
            Dim ViewRegistroAmortizacion As String = "tbAmortizacionRegistro"
            Dim IDTipoAmortizacion As String
            Select Case data.Tipo
                Case BusinessEnum.enTipoAmort.enFiscal
                    strAmortMensual = strTexto
                    strAmortRealMensual = strTexto
                    IDTipoAmortizacion = dtElem.Rows(0)("IDCodigoAmortizacionFiscal")
                Case BusinessEnum.enTipoAmort.enTecnica
                    strAmortMensual = strTexto
                    IDTipoAmortizacion = dtElem.Rows(0)("IDCodigoAmortizacionTecnica")
                Case Else
                    strAmortMensual = strTexto
                    strAmortRealMensual = strTexto
                    IDTipoAmortizacion = dtElem.Rows(0)("IDCodigoAmortizacionContable")
            End Select


            Dim intVidaAmort As Integer
            Dim intVidaUtil As Integer

            Dim TAmortiz As New TipoAmortizacionCabecera
            Dim dtTipo As DataTable = TAmortiz.SelOnPrimaryKey(IDTipoAmortizacion)
            If dtTipo.Rows.Count > 0 Then
                intVidaAmort = Nz(dtTipo.Rows(0)("VidaUtil"), 0)
                intVidaUtil = intVidaAmort
            End If

            Dim dtUltimaAmortizacion As DataTable = AdminData.GetData(ViewRegistroAmortizacion, New StringFilterItem("IDElemento", data.IDElem), "Top 1 FechaContabilizacion", "FechaContabilizacion desc")
            If Not dtUltimaAmortizacion Is Nothing AndAlso dtUltimaAmortizacion.Rows.Count > 0 Then
                If intVidaAmort < DateDiff(DateInterval.Month, datFechas.Fecha, dtUltimaAmortizacion.Rows(0)("FechaContabilizacion")) Then
                    intVidaAmort = DateDiff(DateInterval.Month, datFechas.Fecha, dtUltimaAmortizacion.Rows(0)("FechaContabilizacion"))
                    If DateAndTime.Day(datFechas.Fecha) = 1 Then intVidaAmort += 1
                    intVidaUtil = intVidaAmort
                End If
            End If

            If Not blnAmortizacionPorMeses Then
                If datFechas.Fecha.Day > 1 Then intVidaAmort += 1
            End If
            If data.Año Then
                datFechas.Fechafin = New DateTime(datFechas.FechaInicio.Year, 12, 31)
            End If
            While intVidaAmort > 0
                If blnAmortizacionPorMeses Then
                    If Not BlnFin Then
                        Dim StAmort As New Negocio.ElementoAmortizable.DataCalcPropuestas(datFechas.Fecha, datFechas.FechaInicio, datFechas.Fechafin, dblValorTotalRevalElemento, DblValorAmortizado, _
                                           dblValorResidualElemento, IDTipoAmortizacion, intVidaUtil)
                        DblValorAmortizar = ProcessServer.ExecuteTask(Of Negocio.ElementoAmortizable.DataCalcPropuestas, Double)(AddressOf ElementoAmortizable.CalcularPropuestaPorMeses, StAmort, services)
                    Else
                        DblValorAmortizar = dblValorTotalRevalElemento - dblValorResidualElemento - DblValorAmortizado
                    End If
                Else
                    If Not BlnFin Then
                        Dim StAmort As New Negocio.ElementoAmortizable.DataCalcPropuestas(datFechas.Fecha, datFechas.FechaInicio, datFechas.Fechafin, _
                                            dblValorTotalRevalElemento, DblValorAmortizado, _
                                            dblValorResidualElemento, IDTipoAmortizacion, intVidaUtil)
                        DblValorAmortizar = ProcessServer.ExecuteTask(Of Negocio.ElementoAmortizable.DataCalcPropuestas, Double)(AddressOf ElementoAmortizable.CalcularPropuestaPorDias, StAmort, services)
                    Else
                        DblValorAmortizar = dblValorTotalRevalElemento - dblValorResidualElemento - DblValorAmortizado
                    End If
                End If

                DblValorAmortizar = xRound(DblValorAmortizar, MonInfoA.NDecimalesImporte)
                DblAcumValorAmortizar += DblValorAmortizar
                DblValorAmortizado += DblValorAmortizar

                strAmortMensual = Replace(strAmortMensual, "@Mes" & Month(datFechas.FechaInicio) & "@", DblValorAmortizar)
                Select Case data.Tipo
                    Case Else
                        'Registro de amortizaciones...
                        Dim FilReg As New Filter
                        FilReg.Add("IDElemento", FilterOperator.Equal, data.IDElem, FilterType.String)
                        FilReg.Add("AñoContabilizacion", FilterOperator.Equal, datFechas.FechaInicio.Year, FilterType.Numeric)
                        FilReg.Add("MesContabilizacion", FilterOperator.Equal, datFechas.FechaInicio.Month, FilterType.Numeric)
                        Dim dtAmortRegistro As DataTable = New AmortizacionRegistro().Filter(FilReg)
                        If dtAmortRegistro.Rows.Count > 0 Then
                            DblAcumAmortMes = 0
                            For Each Dr As DataRow In dtAmortRegistro.Select
                                DblAcumAmortMes += Dr("ValorAmortizadoA")
                                DblAcumValorRealizada += Dr("ValorAmortizadoA")
                            Next
                            strAmortRealMensual = Replace(strAmortRealMensual, "@Mes" & datFechas.FechaInicio.Month & "@", CStr(DblAcumAmortMes))
                        Else
                            strAmortRealMensual = Replace(strAmortRealMensual, "@Mes" & datFechas.FechaInicio.Month & "@", CStr(0))
                        End If

                End Select
                If datFechas.Fechafin.Month = 12 Or intVidaAmort = 1 Then
                    Dim DrNew As DataRow = dtAmort.NewRow()
                    DrNew("Año") = datFechas.FechaInicio.Year
                    DrNew("AmortContable") = 0 ' DblAcumValorAmortizar
                    DrNew("AmortFiscal") = 0 'DblAcumValorAmortizar 'DblAcumValorAmortizarFiscal
                    DrNew("AmortTecnica") = 0 ' DblAcumValorAmortizar 'DblAcumValorAmortizarTecnica
                    DrNew("AmortRealizada") = 0 ' DblAcumValorRealizada
                    DrNew("ValorNeto") = 0
                    Select Case data.Tipo
                        Case BusinessEnum.enTipoAmort.enFiscal
                            DrNew("AmortFiscal") = DblAcumValorAmortizar
                            DrNew("AmortAño") = 0
                        Case BusinessEnum.enTipoAmort.enTecnica
                            DrNew("AmortTecnica") = DblAcumValorAmortizar
                            DrNew("AmortAño") = DblAcumValorAmortizar
                        Case Else
                            DrNew("AmortContable") = DblAcumValorAmortizar
                            DrNew("AmortRealizada") = DblAcumValorRealizada

                            DrNew("AmortAño") = DblAcumValorAmortizar
                    End Select



                    strAmortMensual = Replace(strAmortMensual, "@Mes1@", CStr(0))
                    strAmortMensual = Replace(strAmortMensual, "@Mes2@", CStr(0))
                    strAmortMensual = Replace(strAmortMensual, "@Mes3@", CStr(0))
                    strAmortMensual = Replace(strAmortMensual, "@Mes4@", CStr(0))
                    strAmortMensual = Replace(strAmortMensual, "@Mes5@", CStr(0))
                    strAmortMensual = Replace(strAmortMensual, "@Mes6@", CStr(0))
                    strAmortMensual = Replace(strAmortMensual, "@Mes7@", CStr(0))
                    strAmortMensual = Replace(strAmortMensual, "@Mes8@", CStr(0))
                    strAmortMensual = Replace(strAmortMensual, "@Mes9@", CStr(0))
                    strAmortMensual = Replace(strAmortMensual, "@Mes10@", CStr(0))
                    strAmortMensual = Replace(strAmortMensual, "@Mes11@", CStr(0))
                    strAmortMensual = Replace(strAmortMensual, "@Mes12@", CStr(0))

                    Select Case data.Tipo
                        Case BusinessEnum.enTipoAmort.enFiscal
                            DrNew("AmortFiscalMensual") = strAmortMensual
                        Case BusinessEnum.enTipoAmort.enTecnica
                            DrNew("AmortTecnicaMensual") = strAmortMensual
                        Case Else
                            DrNew("AmortContableMensual") = strAmortMensual

                            strAmortRealMensual = Replace(strAmortRealMensual, "@Mes1@", CStr(0))
                            strAmortRealMensual = Replace(strAmortRealMensual, "@Mes2@", CStr(0))
                            strAmortRealMensual = Replace(strAmortRealMensual, "@Mes3@", CStr(0))
                            strAmortRealMensual = Replace(strAmortRealMensual, "@Mes4@", CStr(0))
                            strAmortRealMensual = Replace(strAmortRealMensual, "@Mes5@", CStr(0))
                            strAmortRealMensual = Replace(strAmortRealMensual, "@Mes6@", CStr(0))
                            strAmortRealMensual = Replace(strAmortRealMensual, "@Mes7@", CStr(0))
                            strAmortRealMensual = Replace(strAmortRealMensual, "@Mes8@", CStr(0))
                            strAmortRealMensual = Replace(strAmortRealMensual, "@Mes9@", CStr(0))
                            strAmortRealMensual = Replace(strAmortRealMensual, "@Mes10@", CStr(0))
                            strAmortRealMensual = Replace(strAmortRealMensual, "@Mes11@", CStr(0))
                            strAmortRealMensual = Replace(strAmortRealMensual, "@Mes12@", CStr(0))
                            DrNew("AmortRealizadaMensual") = strAmortRealMensual
                    End Select
                    dtAmort.Rows.Add(DrNew)
                    strAmortMensual = strTexto
                    DblAcumValorAmortizar = 0
                    DblAcumValorRealizada = 0

                    Select Case data.Tipo
                        Case BusinessEnum.enTipoAmort.enFiscal, BusinessEnum.enTipoAmort.enContable
                            strAmortRealMensual = strTexto
                    End Select
                End If
                'REALIZADA
                If data.Año Then
                    If intVidaAmort = 1 Then
                        intVidaAmort = 0
                    Else
                        intVidaAmort = intVidaAmort - DateDiff(Microsoft.VisualBasic.DateInterval.Month, datFechas.FechaInicio, datFechas.Fechafin) - 1
                    End If
                    datFechas.FechaInicio = datFechas.Fechafin.AddDays(1)
                    If intVidaAmort > 0 And intVidaAmort <= 12 Then
                        intVidaAmort = 1
                        datFechas.Fechafin = datFechas.Fecha.AddMonths(intVidaAmort)
                        BlnFin = True
                    Else
                        datFechas.Fechafin = New Date(datFechas.FechaInicio.Year, 12, 31)
                    End If
                Else
                    intVidaAmort -= 1
                    datFechas.FechaInicio = datFechas.Fechafin.AddDays(1)
                    If intVidaAmort = 1 Then
                        datFechas.Fechafin = datFechas.Fecha.AddMonths(intVidaAmort)
                        BlnFin = True
                    Else
                        datFechas.Fechafin = datFechas.Fechafin.AddMonths(1)
                        datFechas.Fechafin = New Date(datFechas.Fechafin.Year, datFechas.Fechafin.Month, datFechas.Fechafin.DaysInMonth(datFechas.Fechafin.Year, datFechas.Fechafin.Month))
                    End If
                End If
            End While
        End If

        Return dtAmort
    End Function


#End Region

#Region " Propuestas Amortización "

    <Serializable()> _
    Public Class DataPropuestasAmort
        Public DtDatos As DataTable
        Public FechaFinal As Date
        Public PorMeses As Boolean
        Public CalcularAjustes As Boolean

        Public Sub New()
        End Sub

        Public Sub New(ByVal DtDatos As DataTable, ByVal FechaFinal As Date, Optional ByVal PorMeses As Boolean = True, Optional ByVal CalcularAjustes As Boolean = False)
            Me.DtDatos = DtDatos
            Me.FechaFinal = FechaFinal
            Me.PorMeses = PorMeses
            Me.CalcularAjustes = CalcularAjustes
        End Sub
    End Class

    <Task()> Public Shared Function PropuestaAmortizacionSegunElemento(ByVal data As DataPropuestasAmort, ByVal services As ServiceProvider) As DataTable
        If Not data.DtDatos Is Nothing AndAlso data.DtDatos.Rows.Count > 0 Then
            Dim dtDatosMeses, dtDatosDias As DataTable
            dtDatosMeses = data.DtDatos.Clone
            dtDatosDias = data.DtDatos.Clone
            For Each drDatos As DataRow In data.DtDatos.Rows
                If drDatos("PorMeses") Then
                    dtDatosMeses.Rows.Add(drDatos.ItemArray)
                Else
                    dtDatosDias.Rows.Add(drDatos.ItemArray)
                End If
            Next

            Dim dtMeses As DataTable
            If Not dtDatosMeses Is Nothing AndAlso dtDatosMeses.Rows.Count > 0 Then
                Dim StPropMes As New DataPropuestasAmort(dtDatosMeses, data.FechaFinal, , data.CalcularAjustes)
                dtMeses = ProcessServer.ExecuteTask(Of DataPropuestasAmort, DataTable)(AddressOf PropuestaAmortizacionPorMeses, StPropMes, services)
            End If

            Dim dtDias As DataTable
            If Not dtDatosDias Is Nothing AndAlso dtDatosDias.Rows.Count > 0 Then
                Dim StPropDias As New DataPropuestasAmort(dtDatosDias, data.FechaFinal, , data.CalcularAjustes)
                dtDias = ProcessServer.ExecuteTask(Of DataPropuestasAmort, DataTable)(AddressOf PropuestaAmortizacionPorDias, StPropDias, services)
            End If
            If Not dtMeses Is Nothing AndAlso dtMeses.Rows.Count > 0 Then
                PropuestaAmortizacionSegunElemento = dtMeses
                If Not dtDias Is Nothing AndAlso dtDias.Rows.Count > 0 Then
                    For Each dr As DataRow In dtDias.Rows
                        PropuestaAmortizacionSegunElemento.Rows.Add(dr.ItemArray)
                    Next
                End If
            Else
                If Not dtDias Is Nothing AndAlso dtDias.Rows.Count > 0 Then
                    PropuestaAmortizacionSegunElemento = dtDias
                End If
            End If
        End If
    End Function

    <Task()> Public Shared Function PropuestaAmortizacion(ByVal data As DataPropuestasAmort, ByVal services As ServiceProvider) As DataTable
        If data.PorMeses Then
            Dim StData As New DataPropuestasAmort(data.DtDatos, data.FechaFinal, , data.CalcularAjustes)
            Return ProcessServer.ExecuteTask(Of DataPropuestasAmort, DataTable)(AddressOf PropuestaAmortizacionPorMeses, StData, services)
        Else
            Dim StData As New DataPropuestasAmort(data.DtDatos, data.FechaFinal, , data.CalcularAjustes)
            Return ProcessServer.ExecuteTask(Of DataPropuestasAmort, DataTable)(AddressOf PropuestaAmortizacionPorDias, StData, services)
        End If
    End Function

    <Task()> Public Shared Function PropuestaAmortizacionPorMeses(ByVal data As DataPropuestasAmort, ByVal services As ServiceProvider) As DataTable
        Dim DblAmortizAjusteA, DblAmortizPdteA As Double
        Dim DblAmortizAjusteA3, DblAmortizPdteA3 As Double
        'Fecha inicio contabilizacion fijada en el elemento
        Dim DteInicioContabilizacion As Date
        'Fecha de inicio de contabilizacion de este proceso.
        'Si no hay contabilizaciones previas coincide con dteInicioContabilizacion.
        'Si las hay se calcula a partir de la Fecha de ultima contabilizacion
        Dim DtAPdteA As New DataTable
        Dim DtAPdteAB As New DataTable
        Dim DtEReval As New DataTable
        Dim ClsEReval As New ElementoRevalorizacion
        Dim DteFechaUltCambio As Date

        If Not data.DtDatos Is Nothing AndAlso data.DtDatos.Rows.Count > 0 Then
            DtAPdteA.Columns.Add("IDElemento", GetType(String))
            DtAPdteA.Columns.Add("AmortizacionPdteA", GetType(Double))
            DtAPdteA.Columns.Add("AjusteRevalorizacionA", GetType(Double))
            DtAPdteA.Columns.Add("AmortizacionPlusPdteA", GetType(Double))
            DtAPdteAB.Columns.Add("IDElemento", GetType(String))
            DtAPdteAB.Columns.Add("AmortizacionPdteA", GetType(Double))
            DtAPdteAB.Columns.Add("AjusteRevalorizacionA", GetType(Double))
            DtAPdteAB.Columns.Add("AmortizacionPlusPdteA", GetType(Double))
            DtAPdteAB.Columns.Add("AmortizacionPdteB", GetType(Double))
            DtAPdteAB.Columns.Add("AjusteRevalorizacionB", GetType(Double))
            DtAPdteAB.Columns.Add("AmortizacionPlusPdteB", GetType(Double))
            'End With

            Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
            Dim MonInfoA As MonedaInfo = Monedas.MonedaA
            Dim MonInfoB As MonedaInfo = Monedas.MonedaB

            For Each Dr As DataRow In data.DtDatos.Select
                DblAmortizPdteA = 0
                DblAmortizAjusteA = 0
                DblAmortizPdteA3 = 0
                DblAmortizAjusteA3 = 0

                DtEReval = ClsEReval.Filter(New FilterItem("IDElemento", FilterOperator.Equal, Dr("IdElemento")), "FechaRevalorizacion DESC,IdLineaRevalorizacion DESC", "TOP 2 ValorCompraFechaA, ValorNetoFechaA, ValorAmortizadoFechaA, ValorResidualFechaA,FechaRevalorizacion ,VidaUtilFecha, IDTipoAmortizacionFecha")
                If Not DtEReval Is Nothing AndAlso DtEReval.Rows.Count > 0 Then
                    DteFechaUltCambio = DtEReval.Rows(0)("FechaRevalorizacion")
                    If DtEReval.Rows.Count = 1 Then
                        '**(I1)
                        '** No ha tenido cambio de condiciones, sus condiciones son las iniciales
                        '**(I1)
                        DteInicioContabilizacion = FechaPrimeroMesSig(Dr("FechaAltaElemento"), Dr("FechaInicioContabilizacion"))
                        If AreEquals(Dr("ValorRealAmortizadoA"), 0) Then
                            ''Caso A0
                            Dim StMesAmort As New DataCalcPropuestas(DteInicioContabilizacion, DteInicioContabilizacion, FechaFinMes(data.FechaFinal), DtEReval.Rows(0)("ValorCompraFechaA"), 0, _
                                                                     DtEReval.Rows(0)("ValorResidualFechaA"), DtEReval.Rows(0)("IDTipoAmortizacionFecha"), DtEReval.Rows(0)("VidaUtilFecha"))
                            DblAmortizPdteA = ProcessServer.ExecuteTask(Of DataCalcPropuestas, Double)(AddressOf CalcularPropuestaPorMeses, StMesAmort, services)
                        Else
                            ''Caso A1
                            If IsDBNull(Dr("FechaUltimaContabilizacion")) Then
                                ApplicationService.GenerateError("La Fecha de Última Amortización del Elemento Amortizable | no puede ser Nula.  ", Dr("IDElemento"))
                            End If
                            If data.CalcularAjustes Then
                                ''Aunque no haya tenido cambio de condiciones, puede que se haya elegido manualmente
                                ''amortizar menos de lo que debia, entonces calculo ese ajuste
                                Dim StCalcAj As New DataCalcPropuestas(DteInicioContabilizacion, DteInicioContabilizacion, FechaFinMes(Dr("FechaUltimaContabilizacion")), DtEReval.Rows(0)("ValorCompraFechaA"), 0, _
                                                                       DtEReval.Rows(0)("ValorResidualFechaA"), DtEReval.Rows(0)("IDTipoAmortizacionFecha"), DtEReval.Rows(0)("VidaUtilFecha"))
                                DblAmortizAjusteA = -Dr("ValorRealAmortizadoA") + ProcessServer.ExecuteTask(Of DataCalcPropuestas, Double)(AddressOf CalcularPropuestaPorMeses, StCalcAj, services)
                            End If
                            Dim StCalcPdte As New DataCalcPropuestas(DteInicioContabilizacion, FechaPrimeroMesSig(Dr("FechaUltimaContabilizacion")), FechaFinMes(data.FechaFinal), DtEReval.Rows(0)("ValorCompraFechaA"), Dr("ValorRealAmortizadoA") + DblAmortizAjusteA, _
                                                                     DtEReval.Rows(0)("ValorResidualFechaA"), DtEReval.Rows(0)("IDTipoAmortizacionFecha"), DtEReval.Rows(0)("VidaUtilFecha"))
                            DblAmortizPdteA = ProcessServer.ExecuteTask(Of DataCalcPropuestas, Double)(AddressOf CalcularPropuestaPorMeses, StCalcPdte, services)
                        End If
                        '**(F1)
                        '** No ha tenido cambio de condiciones, sus condiciones son las iniciales
                        '**(F1)
                    ElseIf DtEReval.Rows.Count > 1 Then
                        '**(I2)
                        '** Si ha tenido cambio de condiciones, sus condiciones han cambiado en algun parametro de los iniciales
                        '**(I2)
                        DteInicioContabilizacion = FechaPrimeroMesSig(Dr("FechaAltaElemento"), Dr("FechaInicioContabilizacion"))
                        ''Caso A2
                        ''Es el mismo caso que A1
                        Dim DrReval() As DataRow = DtEReval.Select("", "FechaRevalorizacion DESC")
                        If data.CalcularAjustes Then
                            'Si quiere calcular el ajuste entonces calculo lo que tendria que haber amortizado desde la fecha de alta
                            'del elemento hasta la ultima fecha de amortizacion, esto lo hago bajo las nuevas condiciones, y a esto
                            'le resto lo que ya tengo amortizado. Entonces el resultado será el ajuste de lo que ya he amortizado
                            Dim dteFechaFin As Date
                            If FechaFinMes(Dr("FechaUltimaContabilizacion")) <> Dr("FechaUltimaContabilizacion") Then
                                'Es que se han cambiado las consiciones de días a meses
                                dteFechaFin = FechaFinMes(DateAdd(DateInterval.Month, -1, Dr("FechaUltimaContabilizacion")))
                            Else
                                dteFechaFin = Dr("FechaUltimaContabilizacion")
                            End If
                            Dim StCalc As New DataCalcPropuestas(DteInicioContabilizacion, DteInicioContabilizacion, dteFechaFin, DrReval(0)("ValorCompraFechaA"), 0, _
                                                                 DrReval(0)("ValorResidualFechaA"), DrReval(0)("IDTipoAmortizacionFecha"), DrReval(0)("VidaUtilFecha"))
                            DblAmortizAjusteA = -Dr("ValorRealAmortizadoA") + ProcessServer.ExecuteTask(Of DataCalcPropuestas, Double)(AddressOf CalcularPropuestaPorMeses, StCalc, services)
                        End If
                        'Puede ser que todavia me quede por amortizar bajo las condiciones iniciales algunos meses. Esto ocurre
                        'cuando entre la fecha de ultima amortizacion y la fecha de cambio de condiciones hay algunos meses enteros

                        '*+*
                        'If FechaFinMes(Dr("FechaUltimaContabilizacion")) <> FechaFinMes(DteFechaUltCambio) Then
                        '    DblAmortizPdteA = CalcularPropuestaPorMeses(DteInicioContabilizacion, _
                        '                           FechaPrimeroMesSig(Dr("FechaUltimaContabilizacion")), _
                        '                           IIf(FechaFinMes(DteFechaFinal) < FechaPrimeroMesSig(DteFechaUltCambio), _
                        '                           FechaFinMes(DteFechaFinal), FechaFinMes(DteFechaUltCambio)), _
                        '                           DtEReval.Rows(0)("ValorCompraFechaA"), _
                        '                           Dr("ValorRealAmortizadoA"), _
                        '                           DtEReval.Rows(0)("ValorResidualFechaA"), _
                        '                           DtEReval.Rows(0)("IDTipoAmortizacionFecha"), _
                        '                           DtEReval.Rows(0)("VidaUtilFecha"))

                        'End If
                        ''Caso A3
                        ''Se calcula el ajuste sobre dblAmortizAjusteA + dblAmortizPdteA
                        If data.CalcularAjustes Then
                            If Length(data.DtDatos.Rows(0)("FechaUltimaContabilizacion")) > 0 Then
                                If FechaPrimeroMesSig(DteFechaUltCambio) > FechaFinMes(data.DtDatos.Rows(0)("FechaUltimaContabilizacion")) Then
                                    Dim StCalcA3 As New DataCalcPropuestas(FechaPrimeroMesSig(DteFechaUltCambio), _
                                    IIf(FechaPrimeroMesSig(DteFechaUltCambio) > FechaPrimeroMesSig(Dr("FechaUltimaContabilizacion")), _
                                    FechaPrimeroMesSig(DteFechaUltCambio), FechaPrimeroMesSig(Dr("FechaUltimaContabilizacion"))), _
                                    FechaFinMes(data.FechaFinal), DtEReval.Rows(0)("ValorCompraFechaA"), _
                                    Dr("ValorRealAmortizadoA") + DblAmortizAjusteA + DblAmortizPdteA, _
                                    DtEReval.Rows(0)("ValorResidualFechaA"), DtEReval.Rows(0)("IDTipoAmortizacionFecha"), _
                                    DtEReval.Rows(0)("VidaUtilFecha"))
                                    DblAmortizPdteA3 = ProcessServer.ExecuteTask(Of DataCalcPropuestas, Double)(AddressOf CalcularPropuestaPorMeses, StCalcA3, services)
                                    DblAmortizAjusteA += DblAmortizAjusteA3
                                End If
                            End If
                        End If

                        If FechaFinMes(data.FechaFinal) > FechaPrimeroMesSig(DteFechaUltCambio) Then
                            Dim StCalcPdteA3 As New DataCalcPropuestas(DteInicioContabilizacion, _
                            IIf(FechaPrimeroMesSig(DteFechaUltCambio) > FechaPrimeroMesSig(Dr("FechaUltimaContabilizacion")), _
                            FechaPrimeroMesSig(DteFechaUltCambio), FechaPrimeroMesSig(Dr("FechaUltimaContabilizacion"))), _
                            FechaFinMes(data.FechaFinal), DtEReval.Rows(0)("ValorCompraFechaA"), _
                            Dr("ValorRealAmortizadoA") + DblAmortizAjusteA + DblAmortizPdteA, DtEReval.Rows(0)("ValorResidualFechaA"), _
                            DtEReval.Rows(0)("IDTipoAmortizacionFecha"), DtEReval.Rows(0)("VidaUtilFecha"))
                            DblAmortizPdteA3 = ProcessServer.ExecuteTask(Of DataCalcPropuestas, Double)(AddressOf CalcularPropuestaPorMeses, StCalcPdteA3, services)
                            DblAmortizPdteA += DblAmortizPdteA3
                        End If
                        '**(F2)
                        '** Si ha tenido cambio de condiciones, sus condiciones han cambiado en algun parametro de las iniciales
                        '**(F2)
                        End If
                End If
                'Valores en la Moneda A (a Presentacion)
                Dim DrNewA As DataRow = DtAPdteA.NewRow()
                DrNewA("IDElemento") = Dr("IDElemento")

                ' Para los elementos comprados en negativo
                If Dr("ValorTotalRevalElementoA") > 0 Then
                    If DblAmortizPdteA > Dr("ValorNetoContableElementoA") And Not data.CalcularAjustes Then
                        DrNewA("AmortizacionPdteA") = Dr("ValorNetoContableElementoA")
                    Else
                        If Dr("ValorNetoContableElementoA") > 0 Then
                            DrNewA("AmortizacionPdteA") = xRound(DblAmortizPdteA, MonInfoA.NDecimalesImporte)
                        End If
                        DrNewA("AjusteRevalorizacionA") = xRound(DblAmortizAjusteA, MonInfoA.NDecimalesImporte)
                    End If


                    If Dr("ValorNetoContableElementoA") > 0 Then
                        '   DrNewA("AmortizacionPlusPdteA") = xRound((DblAmortizPdteA * Dr("ValorTotalPlusvaliaA") / Dr("ValorTotalRevalElementoA")), MonInfoA.NDecimalesImporte)
                        DrNewA("AmortizacionPlusPdteA") = xRound((DblAmortizPdteA * Dr("ValorNetoContablePlusvaliaA") / Dr("ValorNetoContableElementoA")), MonInfoA.NDecimalesImporte)
                    Else
                        DrNewA("AmortizacionPlusPdteA") = 0
                    End If
                Else
                    If DblAmortizPdteA < Dr("ValorNetoContableElementoA") And Not data.CalcularAjustes Then
                        DrNewA("AmortizacionPdteA") = Dr("ValorNetoContableElementoA")
                    Else
                        If Dr("ValorNetoContableElementoA") < 0 Then
                            DrNewA("AmortizacionPdteA") = xRound(DblAmortizPdteA, MonInfoA.NDecimalesImporte)
                        End If
                        DrNewA("AjusteRevalorizacionA") = xRound(DblAmortizAjusteA, MonInfoA.NDecimalesImporte)
                    End If


                    If Dr("ValorNetoContableElementoA") < 0 Then
                        '   DrNewA("AmortizacionPlusPdteA") = xRound((DblAmortizPdteA * Dr("ValorTotalPlusvaliaA") / Dr("ValorTotalRevalElementoA")), MonInfoA.NDecimalesImporte)
                        DrNewA("AmortizacionPlusPdteA") = xRound((DblAmortizPdteA * Dr("ValorNetoContablePlusvaliaA") / Dr("ValorNetoContableElementoA")), MonInfoA.NDecimalesImporte)
                    Else
                        DrNewA("AmortizacionPlusPdteA") = 0
                    End If
                End If
                DtAPdteA.Rows.Add(DrNewA)
                Dim DrNewAB As DataRow = DtAPdteAB.NewRow()
                DrNewAB("IDElemento") = Dr("IDElemento")
                DrNewAB("AmortizacionPdteA") = DrNewA("AmortizacionPdteA")
                DrNewAB("AjusteRevalorizacionA") = DrNewA("AjusteRevalorizacionA")
                DrNewAB("AmortizacionPlusPdteA") = DrNewA("AmortizacionPlusPdteA")
                DrNewAB("AmortizacionPdteB") = xRound(DblAmortizPdteA * MonInfoA.CambioB, MonInfoB.NDecimalesImporte)
                DrNewAB("AjusteRevalorizacionB") = xRound(DblAmortizAjusteA * MonInfoA.CambioB, MonInfoB.NDecimalesImporte)
                If Dr("ValorTotalRevalElementoB") > 0 Then DrNewAB("AmortizacionPlusPdteB") = xRound(DtAPdteA.Rows(0)("AmortizacionPlusPdteA") * MonInfoA.CambioB, MonInfoB.NDecimalesImporte)
                DtAPdteAB.Rows.Add(DrNewAB)
            Next
            If Not DtAPdteA Is Nothing Then
                Return DtAPdteA
            End If
        End If
    End Function

    <Task()> Public Shared Function PropuestaAmortizacionPorDias(ByVal data As DataPropuestasAmort, ByVal services As ServiceProvider) As DataTable
        Dim DblAmortizAjusteA, DblAmortizPdteA, DblAmortizAjusteA3, DblAmortizPdteA3 As Double
        'Fecha inicio contabilizacion fijada en el elemento
        Dim DteInicioContabilizacion As Date
        'Fecha de inicio de contabilizacion de este proceso.
        'Si no hay contabilizaciones previas coincide con dteInicioContabilizacion.
        'Si las hay se calcula a partir de la Fecha de ultima contabilizacion
        Dim DtAPdteA As New DataTable
        Dim DtAPdteAB As New DataTable
        Dim DtEReval As New DataTable
        Dim ClsEReval As New ElementoRevalorizacion
        Dim DteFechaUltCambio As Date

        If Not data.DtDatos Is Nothing Then
            If data.DtDatos.Rows.Count > 0 Then
                With DtAPdteA.Columns
                    .Add("IDElemento", GetType(String))
                    .Add("AmortizacionPdteA", GetType(Double))
                    .Add("AjusteRevalorizacionA", GetType(Double))
                    .Add("AmortizacionPlusPdteA", GetType(Double))
                End With
                With DtAPdteAB.Columns
                    .Add("IDElemento", GetType(String))
                    .Add("AmortizacionPdteA", GetType(Double))
                    .Add("AjusteRevalorizacionA", GetType(Double))
                    .Add("AmortizacionPlusPdteA", GetType(Double))
                    .Add("AmortizacionPdteB", GetType(Double))
                    .Add("AjusteRevalorizacionB", GetType(Double))
                    .Add("AmortizacionPlusPdteB", GetType(Double))
                End With

                Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
                Dim MonInfoA As MonedaInfo = Monedas.MonedaA
                Dim MonInfoB As MonedaInfo = Monedas.MonedaB

                For Each Dr As DataRow In data.DtDatos.Select
                    DblAmortizPdteA = 0
                    DblAmortizAjusteA = 0
                    DblAmortizPdteA3 = 0
                    DblAmortizAjusteA3 = 0

                    DtEReval = ClsEReval.Filter(New FilterItem("IDElemento", FilterOperator.Equal, Dr("IDElemento")), "FechaRevalorizacion DESC, IDLineaRevalorizacion DESC", "TOP 2 ValorCompraFechaA," & "ValorNetoFechaA, ValorAmortizadoFechaA, ValorResidualFechaA, FechaRevalorizacion ,VidaUtilFecha, IDTipoAmortizacionFecha")
                    If DtEReval.Rows.Count > 0 Then
                        DteFechaUltCambio = DtEReval.Rows(0)("FechaRevalorizacion")
                        If DtEReval.Rows.Count = 1 Then
                            '**(I1)
                            '** No ha tenido cambio de condiciones, sus condiciones son las iniciales
                            '**(I1)
                            DteInicioContabilizacion = Dr("FechaInicioContabilizacion")
                            If Dr("ValorRealAmortizadoA") = 0 Then
                                ''Caso A0
                                Dim StDias As New DataCalcPropuestas(DteInicioContabilizacion, DteInicioContabilizacion, _
                                                data.FechaFinal, DtEReval.Rows(0)("ValorCompraFechaA"), 0, DtEReval.Rows(0)("ValorResidualFechaA"), _
                                                DtEReval.Rows(0)("IDTipoAmortizacionFecha"), DtEReval.Rows(0)("VidaUtilFecha"))
                                DblAmortizPdteA = ProcessServer.ExecuteTask(Of DataCalcPropuestas, Double)(AddressOf CalcularPropuestaPorDias, StDias, services)
                            Else
                                ''Caso A1
                                If data.CalcularAjustes Then
                                    ''Aunque no haya tenido cambio de condiciones, puede que se haya elegido manualmente
                                    ''amortizar menos de lo que debia, entonces calculo ese ajuste
                                    Dim StAjDias As New DataCalcPropuestas(DteInicioContabilizacion, DteInicioContabilizacion, Dr("FechaUltimaContabilizacion"), DtEReval.Rows(0)("ValorCompraFechaA"), 0, _
                                                                           DtEReval.Rows(0)("ValorResidualFechaA"), DtEReval.Rows(0)("IDTipoAmortizacionFecha"), DtEReval.Rows(0)("VidaUtilFecha"))
                                    DblAmortizAjusteA = -Dr("ValorRealAmortizadoA") + ProcessServer.ExecuteTask(Of DataCalcPropuestas, Double)(AddressOf CalcularPropuestaPorDias, StAjDias, services)
                                End If
                                Dim StPdteDias As New DataCalcPropuestas(DteInicioContabilizacion, CDate(Dr("FechaUltimaContabilizacion")).AddDays(1), data.FechaFinal, _
                                                                         DtEReval.Rows(0)("ValorCompraFechaA"), Dr("ValorRealAmortizadoA") + DblAmortizAjusteA, _
                                                                         DtEReval.Rows(0)("ValorResidualFechaA"), DtEReval.Rows(0)("IDTipoAmortizacionFecha"), _
                                                                         DtEReval.Rows(0)("VidaUtilFecha"))
                                DblAmortizPdteA = ProcessServer.ExecuteTask(Of DataCalcPropuestas, Double)(AddressOf CalcularPropuestaPorDias, StPdteDias, services)
                            End If
                            '**(F1)
                            '** No ha tenido cambio de condiciones, sus condiciones son las iniciales
                            '**(F1)
                        ElseIf DtEReval.Rows.Count > 1 Then
                            '**(I2)
                            '** Si ha tenido cambio de condiciones, sus condiciones han cambiado en algun parametro de las iniciales
                            '**(I2)
                            DteInicioContabilizacion = Dr("FechaInicioContabilizacion")
                            ''Caso A2
                            ''Es el mismo caso que A1
                            Dim DrEval() As DataRow = DtEReval.Select("", "FechaRevalorizacion DESC")
                            If data.CalcularAjustes Then
                                Dim StRevalDias As New DataCalcPropuestas(DteInicioContabilizacion, DteInicioContabilizacion, Dr("FechaUltimaContabilizacion"), _
                                                                          DrEval(0)("ValorCompraFechaA"), 0, DrEval(0)("ValorResidualFechaA"), DrEval(0)("IDTipoAmortizacionFecha"), _
                                                                          DrEval(0)("VidaUtilFecha"))
                                DblAmortizAjusteA = -Dr("ValorRealAmortizadoA") + ProcessServer.ExecuteTask(Of DataCalcPropuestas, Double)(AddressOf CalcularPropuestaPorDias, StRevalDias, services)
                            End If
                            Dim StPdteDias As New DataCalcPropuestas(DteInicioContabilizacion, CDate(Dr("FechaUltimaContabilizacion")).AddDays(1), IIf(data.FechaFinal < DteFechaUltCambio.AddDays(-1), data.FechaFinal, _
                                                                     DteFechaUltCambio.AddDays(-1)), DtEReval.Rows(0)("ValorCompraFechaA"), Dr("ValorRealAmortizadoA"), DtEReval.Rows(0)("ValorResidualFechaA"), _
                                                                     DtEReval.Rows(0)("IDTipoAmortizacionFecha"), DtEReval.Rows(0)("VidaUtilFecha"))
                            DblAmortizPdteA = ProcessServer.ExecuteTask(Of DataCalcPropuestas, Double)(AddressOf CalcularPropuestaPorDias, StPdteDias, services)
                            If data.FechaFinal >= DteFechaUltCambio Then
                                ''Caso A3
                                ''Se calcula el ajuste sobre dblAmortizAjusteA + dblAmortizPdteA
                                If data.CalcularAjustes Then
                                    If DteFechaUltCambio > Dr("FechaUltimaContabilizacion") Then
                                        Dim StDiasUlt As New DataCalcPropuestas(DteInicioContabilizacion, DteInicioContabilizacion, DteFechaUltCambio.AddDays(-1), DrEval(0)("ValorCompraFechaA"), _
                                                                                Dr("ValorRealAmortizadoA") + DblAmortizAjusteA + DblAmortizPdteA, DrEval(0)("ValorResidualFechaA"), DrEval(0)("IDTipoAmortizacionFecha"), DrEval(0)("VidaUtilFecha"))
                                        DblAmortizAjusteA3 = ProcessServer.ExecuteTask(Of DataCalcPropuestas, Double)(AddressOf CalcularPropuestaPorDias, StDiasUlt, services)
                                        DblAmortizAjusteA += DblAmortizAjusteA3
                                    End If
                                End If
                                Dim StPdteUlt As New DataCalcPropuestas(DteInicioContabilizacion, CDate(Dr("FechaUltimaContabilizacion")).AddDays(1), data.FechaFinal, DrEval(0)("ValorCOmpraFechaA"), _
                                                                        Dr("ValorRealAmortizadoA") + DblAmortizAjusteA + DblAmortizPdteA, DrEval(0)("ValorResidualFechaA"), DrEval(0)("IDTipoAmortizacionFecha"), DrEval(0)("VidaUtilFecha"))
                                DblAmortizPdteA3 = ProcessServer.ExecuteTask(Of DataCalcPropuestas, Double)(AddressOf CalcularPropuestaPorDias, StPdteUlt, services)
                                DblAmortizPdteA += DblAmortizPdteA3
                            End If
                            '**(F2)
                            '** Si ha tenido cambio de condiciones, sus condiciones han cambiado en algun parametro de las iniciales
                            '**(F2)
                        End If
                    End If
                    'Valores en la moneda A (a Presentacion)
                    Dim DrNewA As DataRow = DtAPdteA.NewRow
                    DrNewA("IDElemento") = Dr("IDElemento")

                    ' Para los elementos comprados en negativo
                    If Dr("ValorTotalRevalElementoA") > 0 Then
                        If DblAmortizPdteA > Dr("ValorNetoContableElementoA") And Not data.CalcularAjustes Then
                            DrNewA("AmortizacionPdteA") = Dr("ValorNetoContableElementoA")
                        Else

                            If Dr("ValorNetoContableElementoA") > 0 Then
                                DrNewA("AmortizacionPdteA") = xRound(DblAmortizPdteA, MonInfoA.NDecimalesImporte)
                            End If
                            DrNewA("AjusteRevalorizacionA") = xRound(DblAmortizAjusteA, MonInfoA.NDecimalesImporte)
                        End If
                    Else
                        If DblAmortizPdteA < Dr("ValorNetoContableElementoA") And Not data.CalcularAjustes Then
                            DrNewA("AmortizacionPdteA") = Dr("ValorNetoContableElementoA")
                        Else

                            If Dr("ValorNetoContableElementoA") < 0 Then
                                DrNewA("AmortizacionPdteA") = xRound(DblAmortizPdteA, MonInfoA.NDecimalesImporte)
                            End If
                            DrNewA("AjusteRevalorizacionA") = xRound(DblAmortizAjusteA, MonInfoA.NDecimalesImporte)
                        End If
                    End If
                    ' DrNewA("AmortizacionPdteA") = xRound(DblAmortizPdteA, DtMonedaA.Rows(0)("NDecimalesImp"))
                    'DrNewA("AjusteRevalorizacionA") = xRound(DblAmortizAjusteA, DtMonedaA.Rows(0)("NDecimalesImp"))



                    If Dr("ValorNetoContableElementoA") > 0 Then
                        ' DrNewA("AmortizacionPlusPdteA") = xRound((DblAmortizPdteA * Dr("ValorTotalPlusvaliaA")) / Dr("ValorTotalRevalElementoA"), MonInfoA.NDecimalesImporte)
                        DrNewA("AmortizacionPlusPdteA") = xRound((DblAmortizPdteA * Dr("ValorNetoContablePlusvaliaA") / Dr("ValorNetoContableElementoA")), MonInfoA.NDecimalesImporte)
                    Else
                        DrNewA("AmortizacionPlusPdteA") = 0
                    End If
                    DtAPdteA.Rows.Add(DrNewA)
                    Dim DrNewAB As DataRow = DtAPdteAB.NewRow()
                    DrNewAB("IDElemento") = Dr("IDElemento")
                    DrNewAB("AmortizacionPdteA") = DrNewA("AmortizacionPdteA")
                    DrNewAB("AjusteRevalorizacionA") = DrNewA("AjusteRevalorizacionA")
                    DrNewAB("AmortizacionPlusPdteA") = DrNewA("AmortizacionPlusPdteA")
                    DrNewAB("AmortizacionPdteB") = xRound(DblAmortizPdteA * MonInfoA.CambioB, MonInfoB.NDecimalesImporte)
                    DrNewAB("AjusteRevalorizacionB") = xRound(DblAmortizAjusteA * MonInfoA.CambioB, MonInfoB.NDecimalesImporte)
                    If Dr("ValorTotalRevalElementoB") > 0 Then DrNewAB("AmortizacionPlusPdteB") = xRound(DtAPdteA.Rows(0)("AmortizacionPlusPdteA") * MonInfoA.CambioB, MonInfoB.NDecimalesImporte)
                    DtAPdteAB.Rows.Add(DrNewAB)
                Next
                If Not DtAPdteA Is Nothing Then
                    Return DtAPdteA
                End If
            End If
        End If
    End Function

#End Region

#Region "Generacion de elementos amortizables a partir de facturas de compra"

    <Serializable()> _
    Public Class DataMostrarElemFact
        Public DtContador As DataTable
        Public G As Guid
        Public Agrupar As Boolean

        Public Sub New()
        End Sub

        Public Sub New(ByVal DtContador As DataTable, ByVal G As Guid, ByVal Agrupar As Boolean)
            Me.DtContador = DtContador
            Me.G = G
            Me.Agrupar = Agrupar
        End Sub
    End Class

    <Task()> Public Shared Function MostrarElementosFacturas(ByVal data As DataMostrarElemFact, ByVal services As ServiceProvider) As DataTable
        Dim DtAux As DataTable = New BE.DataEngine().Filter("vFrmInversionFacturaCompra", New GuidFilterItem("IDProcess", data.G))
        Dim StData As New DataGenElem(DtAux, data.DtContador, data.Agrupar)
        Return ProcessServer.ExecuteTask(Of DataGenElem, DataTable)(AddressOf GenerarElementos, StData, services)
    End Function

    <Serializable()> _
    Public Class DataGenElem
        Public DtMarcados As DataTable
        Public DtContador As DataTable
        Public Agrupar As Boolean

        Public Sub New()
        End Sub

        Public Sub New(ByVal DtMarcados As DataTable, ByVal DtContador As DataTable, Optional ByVal Agrupar As Boolean = False)
            Me.DtMarcados = DtMarcados
            Me.DtContador = DtContador
            Me.Agrupar = Agrupar
        End Sub
    End Class

    <Task()> Public Shared Function GenerarElementos(ByVal data As DataGenElem, ByVal services As ServiceProvider) As DataTable
        Dim IntSize As Integer
        Dim ArrContadores() As String
        Dim DteFechaCompra As Date
        Dim Año, Mes As Integer
        Dim DtLineas As New DataTable
        If Not data.DtMarcados Is Nothing AndAlso data.DtMarcados.Rows.Count > 0 Then
            If data.Agrupar Then
                For Each Dr As DataRow In data.DtMarcados.Select
                    If Dr("IdLineaFactura").ToString.Length > IntSize Then
                        IntSize = Dr("IdLineaFactura").ToString.Length
                    End If
                Next
                IntSize = data.DtMarcados.Rows.Count * (IntSize + 1)
            Else : IntSize = 10
            End If
            DtLineas = data.DtMarcados.Copy
            DtLineas.Columns.Remove("IDProcess")
            DtLineas.Columns.Add("LineasFactura")
            DtLineas.Columns("LineasFactura").DataType = GetType(String)
            DtLineas.Columns("LineasFactura").MaxLength = IntSize

            'Si hay que agrupar por proveedor previamente hay que manipular el dt
            If data.Agrupar Then DtLineas = ProcessServer.ExecuteTask(Of DataTable, DataTable)(AddressOf AgruparPorProveedor, DtLineas, services)
            Dim DtElementos As DataTable = New ElementoAmortizable().AddNew 'New BE.DataEngine().Filter("vfrmCreacionElementosAmortiz", "*", "", , True)
            DtElementos.Columns.Add("LineasFactura")
            DtElementos.Columns("LineasFactura").DataType = GetType(String)
            DtElementos.Columns("LineasFactura").MaxLength = IntSize
            'DtElementos.Columns.Add("IDCodigoAmortizacionContable")
            'DtElementos.Columns.Add("IDCodigoAmortizacionTecnica")
            'DtElementos.Columns.Add("IDCodigoAmortizacionFiscal")
            'DtElementos.Columns.Add("ValorNetoContableElementoA", GetType(Double))

            If Not data.DtContador Is Nothing AndAlso data.DtContador.Rows.Count > 0 Then
                Dim StSim As New DataSimularContador(DtLineas.Rows.Count, data.DtContador.Rows(0)("IdContador"))
                ArrContadores = ProcessServer.ExecuteTask(Of DataSimularContador, String())(AddressOf SimularContador, StSim, services)
            Else : ApplicationService.GenerateError("El Contador no es válido.")
            End If

            Dim DrNew As DataRow
            Dim i As Integer
            For Each dr As DataRow In DtLineas.Select
                DrNew = DtElementos.NewRow
                DrNew("IdElemento") = ArrContadores(i)
                If Not data.Agrupar Then
                    DrNew("DescElemento") = "(" & xRound(CDbl(dr("Cantidad")), 2) & ") " & dr("DescArticulo")
                    DrNew("LineasFactura") = CStr(dr("IdLineaFactura"))
                Else
                    DrNew("DescElemento") = dr("DescArticulo")
                    DrNew("LineasFactura") = dr("LineasFactura")
                End If
                DrNew("FechaCompra") = dr("FechaFactura")
                DrNew("ValorTotalElementoA") = dr("ImporteA")
                DrNew("ValorTotalRevalElementoA") = dr("ImporteA")
                DrNew("ValorNetoContableElementoA") = dr("ImporteA")
                DrNew("ValorTotalElementoB") = dr("ImporteB")
                DrNew("ValorTotalRevalElementoB") = dr("ImporteB")
                DrNew("ValorResidualA") = 0
                DrNew("ValorResidualB") = 0
                DrNew("IDFactura") = dr("IDFactura")
                DrNew("NFactura") = dr("NFactura")
                DrNew("IDCentroGestion") = dr("IDCentroGestion")
                DrNew("IdContador") = data.DtContador.Rows(0)("IdContador")

                If Not dr.IsNull("IDGrupoAmortiz") Then
                    DrNew("IDGrupoAmortizacion") = dr("IDGrupoAmortiz")
                    Dim DtGrupoAmort As DataTable = New GrupoAmortizacion().SelOnPrimaryKey(DrNew("IDGrupoAmortizacion"))
                    If Not DtGrupoAmort Is Nothing AndAlso DtGrupoAmort.Rows.Count > 0 Then
                        If Length(DtGrupoAmort.Rows(0)("IDTipoAmortiz")) > 0 Then
                            DrNew("IDCodigoAmortizacionContable") = DtGrupoAmort.Rows(0)("IDTipoAmortiz")
                        End If
                    End If
                End If
                DteFechaCompra = CDate(DrNew("FechaCompra"))
                If DteFechaCompra.Day = 1 Then
                    Mes = DteFechaCompra.Month
                    Año = DteFechaCompra.Year
                Else
                    If DteFechaCompra.Month <> 12 Then
                        Mes = DteFechaCompra.Month + 1
                        Año = DteFechaCompra.Year
                    Else
                        Mes = 1
                        Año = DteFechaCompra.Year + 1
                    End If
                End If
                DrNew("MesInicioContabilizado") = Mes
                DrNew("AñoInicioContabilizado") = Año
                DrNew("IdCContableFactura") = Nz(dr("CContable"))
                DrNew("CantidadArticulo") = Nz(dr("Cantidad"))

                DtElementos.Rows.Add(DrNew)
                i += 1
            Next
            Return DtElementos
        End If

    End Function

    <Task()> Public Shared Function AgruparPorProveedor(ByVal DtLineas As DataTable, ByVal services As ServiceProvider) As DataTable
        'Acumulamos ImporteA e Importe B para cada par Proveedor-Centro de Gestión
        Const StrSeparador As String = ";"
        Dim DblSumaA, DblSumaB As Double
        Dim StrDescElemento As String = String.Empty
        Dim StrLineasFactura As String = String.Empty
        Dim DtAgrupados As New DataTable
        DtAgrupados = DtLineas.Clone
        Dim Dv As DataView = DtLineas.DefaultView
        Dv.Sort = "IdProveedor, IdCentroGestion"
        Dim StrPActual As String = String.Empty      'Proveedor actual
        Dim StrCGActual As String = String.Empty     'Centro Gestión actual
        Dim DrNew As DataRow
        Dim Cantidad As Double = 0
        For i As Integer = 0 To DtLineas.DefaultView.Count - 1
            If StrPActual <> Dv(i)("IDProveedor").ToString() OrElse _
                StrCGActual <> Dv(i)("IDCentroGestion").ToString() Then
                If Not DrNew Is Nothing AndAlso DrNew.RowState = DataRowState.Detached Then
                    DtAgrupados.Rows.Add(DrNew)
                End If
                'Cambio de proveedor o de centro de gestión.
                DrNew = DtAgrupados.NewRow()
                StrPActual = Dv(i)("IDProveedor").ToString()
                StrCGActual = Dv(i)("IDCentroGestion").ToString()
                DblSumaA = Dv(i)("ImporteA")
                DblSumaB = Dv(i)("ImporteB")
                Cantidad = Dv(i)("Cantidad")

                StrLineasFactura = String.Empty
                StrDescElemento = String.Empty
            Else
                DblSumaA += Dv(i)("ImporteA")
                DblSumaB += Dv(i)("ImporteB")
                Cantidad += Dv(i)("Cantidad")
            End If
            StrDescElemento = StrDescElemento & "(" & xRound(CDbl(Dv(i)("Cantidad")), 2) & ") " & Dv(i)("DescArticulo") & " , "
            StrLineasFactura = StrLineasFactura & Dv(i)("IdLineaFactura") & StrSeparador
            DrNew("IDFactura") = Dv(i)("IDFactura")
            DrNew("Cantidad") = Cantidad
            DrNew("DescArticulo") = Strings.Left((Left(StrDescElemento, Length(StrDescElemento) - 2)), 299)
            DrNew("FechaFactura") = Dv(i)("FechaFactura")
            DrNew("CContable") = Dv(i)("CContable")
            DrNew("ImporteA") = DblSumaA
            DrNew("ImporteB") = DblSumaB
            DrNew("EstadoInmovilizado") = False
            DrNew("IDProveedor") = StrPActual
            DrNew("IDGrupoAmortiz") = Dv(i)("IDGrupoAmortiz")
            DrNew("IDCentroGestion") = StrCGActual
            DrNew("LineasFactura") = Left(StrLineasFactura, Length(StrLineasFactura) - 1)
        Next
        If Not DrNew Is Nothing AndAlso DrNew.RowState = DataRowState.Detached Then
            DtAgrupados.Rows.Add(DrNew)
        End If
        Return DtAgrupados


    End Function

    <Serializable()> _
    Public Class DataSimularContador
        Public CountReg As Integer
        Public IDContador As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal CountReg As Integer, ByVal IDContador As String)
            Me.CountReg = CountReg
            Me.IDContador = IDContador
        End Sub
    End Class

    <Task()> Public Shared Function SimularContador(ByVal data As DataSimularContador, ByVal services As ServiceProvider) As String()
        Dim DtContador As DataTable = New Contador().SelOnPrimaryKey(data.IDContador)
        If Not DtContador Is Nothing AndAlso DtContador.Rows.Count > 0 Then
            Dim i As Integer
            Dim arrContadores(data.CountReg - 1) As String
            Dim lngCounter As Integer = DtContador.Rows(0)("Contador")
            Dim lngCounterLen As Integer = DtContador.Rows(0)("Longitud")
            Dim blnNumeric As Boolean = DtContador.Rows(0)("Numerico")
            Dim strCounterText As String = Nz(DtContador.Rows(0)("Texto"), String.Empty)
            Dim strCounter As String = CStr(lngCounter)
            Dim intPad As Integer
            If Not blnNumeric Then
                intPad = lngCounterLen - Len(strCounter) - Len(strCounterText)
            End If
            For i = 0 To data.CountReg - 1
                strCounter = CStr(lngCounter + i)
                If intPad > 0 Then
                    strCounter = New String("0", intPad) & strCounter
                End If
                arrContadores(i) = strCounterText & strCounter
            Next i
            Return arrContadores
        Else : ApplicationService.GenerateError("El contador no está configurado.")
        End If
    End Function

    <Task()> Public Shared Sub CrearElementosFacturas(ByVal Dt As DataTable, ByVal services As ServiceProvider)
        'Variables generales
        Const StrDelimeter As String = ";"
        Dim i As Integer
        'Para los nuevos elementos amortizables
        Dim DtNew As New DataTable
        'Para las líneas de factuas de compra modificadas
        Dim DtNewFCL As DataTable
        Dim ClsFCL As New FacturaCompraLinea
        Dim DtFCL As New DataTable
        'Para los nuevos registros en ElementoAmortizableFCL
        Dim ClsElemFCL As New ElementoAmortizableFCL
        Dim DtNewElementoAmortizableFCL As DataTable = ClsElemFCL.AddNew
        'Para los nuevos registros de analítica
        Dim ClsElemANA As New ElementoAmortizAnalitica
        Dim DtNewAnalitica As DataTable = ClsElemANA.AddNew
        Dim DtInversionANA As New DataTable
        Dim DblImpTotalANA As Double
        Dim StrSQLinversion As String
        'Para los nuevos registros en ElementoRevalorizacion
        Dim DtNewElementoRevalorizacion As New DataTable
        Dim ClsER As New ElementoRevalorizacion

        'Para obtener datos de las amortizaciones
        Dim DtGAmortiz As New DataTable
        Dim DtTipoAmortiz As New DataTable
        Dim DtTipoAmortizLinea As New DataTable
        Dim DblValorAmortizA As Double
        Dim DblDotacionA As Double
        Dim DblDotacionB As Double
        Dim DblPorcentaje As Double
        Dim ArrIDs() As String
        Dim ArrContadores() As Object

        '** DATOS DE LA MONEDA
        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
        Dim MonInfoA As MonedaInfo = Monedas.MonedaA
        Dim LngDecimalesA As Integer = MonInfoA.NDecimalesImporte
        Dim DblCambioAB As Double = MonInfoA.CambioB
        Dim MonInfoB As MonedaInfo = Monedas.MonedaB
        Dim LngDecimalesB As Integer = MonInfoB.NDecimalesImporte

        '***************************************
        '////Completar e insertar los elementos
        If Not Dt Is Nothing AndAlso Dt.Rows.Count > 0 Then
            DtNew = New ElementoAmortizable().AddNew()
            'Guardar contadores validos en un array
            ReDim ArrContadores(Dt.Rows.Count - 1)
            For j As Integer = 0 To ArrContadores.GetUpperBound(0)
                ArrContadores(j) = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, Dt.Rows(j)("IDContador"), services)
            Next
            Dim FwnGAmortiz As New GrupoAmortizacion
            Dim FwnTipoAmortiz As New TipoAmortizacionCabecera
            Dim FwnTipoAmortizLinea As New TipoAmortizacionLinea
            Dim DrNew As DataRow
            For Each dr As DataRow In Dt.Rows
                DrNew = DtNew.NewRow
                For Each dc As DataColumn In Dt.Columns
                    If dc.ColumnName <> "LineasFactura" Then
                        If dc.ColumnName = "IDElemento" Then
                            DrNew(dc.ColumnName) = ArrContadores(i)
                        ElseIf dc.ColumnName = "IdCentroGestion" Then
                            DrNew(dc.ColumnName) = dr(dc.ColumnName)
                        ElseIf dc.ColumnName = "FechaCompra" Then
                            DrNew(dc.ColumnName) = dr(dc.ColumnName)
                            DrNew("FechaInicioContabilizacion") = dr(dc.ColumnName)
                            DrNew("FechaUltimaRevalorizacion") = dr(dc.ColumnName)
                        Else
                            If dc.ColumnName = "DescElemento" Then
                                If dr(dc.ColumnName).ToString.Length > dc.MaxLength And dc.MaxLength <> 0 Then
                                    'Trunca la longitud
                                    dr(dc.ColumnName) = dr(dc.ColumnName).ToString.Substring(0, dc.MaxLength)
                                End If
                            End If
                            DrNew(dc.ColumnName) = dr(dc.ColumnName)
                        End If
                    End If
                Next

                DrNew("ValorAmortizadoElementoA") = 0
                If Not DrNew.IsNull("ValorResidualA") Then
                    DrNew("ValorNetoContableElementoA") = DrNew("ValorTotalElementoA") - DrNew("ValorResidualA")
                Else
                    DrNew("ValorNetoContableElementoA") = DrNew("ValorTotalElementoA")
                End If
                DrNew("ValorTotalPlusvaliaA") = 0
                DrNew("ValorAmortizadoPlusvaliaA") = 0
                DrNew("ValorNetoContablePlusvaliaA") = 0
                DrNew("ValorAmortizadoElementoB") = 0
                If Not dr.IsNull("ValorResidualB") Then
                    DrNew("ValorNetoContableElementoB") = DrNew("ValorTotalElementoB") - DrNew("ValorResidualB")
                Else
                    DrNew("ValorNetoContableElementoB") = DrNew("ValorTotalElementoB")
                End If
                DrNew("ValorTotalPlusvaliaB") = 0
                DrNew("ValorAmortizadoPlusvaliaB") = 0
                DrNew("ValorNetoContablePlusvaliaB") = 0

                DtGAmortiz = FwnGAmortiz.SelOnPrimaryKey(dr("IDGrupoAmortizacion").ToString())

                If Not DtGAmortiz Is Nothing AndAlso DtGAmortiz.Rows.Count > 0 Then
                    DtTipoAmortiz = FwnTipoAmortiz.SelOnPrimaryKey(Nz(dr("IDCodigoAmortizacionContable"), DtGAmortiz.Rows(0)("IdTipoAmortiz")))
                    If Not IsNothing(DtTipoAmortiz) AndAlso DtTipoAmortiz.Rows.Count > 0 Then
                        DrNew("IDCodigoAmortizacionContable") = DtTipoAmortiz.Rows(0)("IDTipoAmortizacion")
                        DrNew("IDCodigoAmortizacionTecnica") = DtTipoAmortiz.Rows(0)("IDTipoAmortizacion")
                        DrNew("IDCodigoAmortizacionFiscal") = DtTipoAmortiz.Rows(0)("IDTipoAmortizacion")
                        DrNew("VidaContableElemento") = DtTipoAmortiz.Rows(0)("VidaUtil")
                        DrNew("VidaTecnicaElemento") = DtTipoAmortiz.Rows(0)("VidaUtil")
                        DrNew("VidaFiscalElemento") = DtTipoAmortiz.Rows(0)("VidaUtil")
                    End If

                    DtTipoAmortizLinea = FwnTipoAmortizLinea.Filter(New FilterItem("IDTipoAmortizacion", FilterOperator.Equal, DtGAmortiz.Rows(0)("IdTipoAmortiz"), FilterType.String), "NAño")
                    If Not DtTipoAmortizLinea Is Nothing AndAlso DtTipoAmortizLinea.Rows.Count > 0 Then
                        DblPorcentaje = DtTipoAmortizLinea.Rows(0)("PorcentajeAmortizar")
                    Else
                        DblPorcentaje = 0
                    End If

                    DrNew("PorcentajeAnualContable") = DblPorcentaje
                    DrNew("PorcentajeAnualTecnico") = DblPorcentaje
                    DrNew("PorcentajeAnualFiscal") = DblPorcentaje

                    DblValorAmortizA = DrNew("ValorTotalRevalElementoA") - DrNew("ValorResidualA")
                    DblDotacionA = (DblValorAmortizA * (DblPorcentaje / 100)) / 12
                    DblDotacionB = DblDotacionA * DblCambioAB

                    'Decimales
                    DblDotacionA = xRound(DblDotacionA, LngDecimalesA)
                    DblDotacionB = xRound(DblDotacionB, LngDecimalesB)

                    DrNew("DotacionContableElementoA") = DblDotacionA
                    DrNew("DotacionTecnicaElementoA") = DblDotacionA
                    DrNew("DotacionFiscalElementoA") = DblDotacionA
                    DrNew("DotacionContableElementoB") = DblDotacionB
                    DrNew("DotacionTecnicaElementoB") = DblDotacionB
                    DrNew("DotacionFiscalElementoB") = DblDotacionB
                End If
                DrNew("MesUltimoContabilizado") = 0
                DrNew("AñoUltimoContabilizado") = 0
                DrNew("FechaUltimaContabilizacion") = System.DBNull.Value
                DrNew("ValorUltimoContabilizadoA") = 0
                DrNew("ValorUltimoContabilizadoB") = 0
                DrNew("ValorReposicionA") = 0
                DrNew("ValorReposicionActualizadoA") = 0
                DrNew("ValorReposicionB") = 0
                DrNew("ValorReposicionActualizadoB") = 0
                DtNew.Rows.Add(DrNew)

                '** Tablas relacionadas con la de ElementoAmortizable
                '///Actualizar el campo EstadoInmovilizado en las lineas de facturas compra
                '///Agregar línea a ElementoFCL
                ArrIDs = Split(dr("LineasFactura"), StrDelimeter, , CompareMethod.Text)
                For j As Integer = 0 To ArrIDs.GetUpperBound(0)
                    DtFCL = ClsFCL.SelOnPrimaryKey(CInt(ArrIDs(j)))
                    If Not DtFCL Is Nothing AndAlso DtFCL.Rows.Count > 0 Then
                        If DtNewFCL Is Nothing Then
                            DtNewFCL = DtFCL.Clone
                        End If
                        DtFCL.Rows(0)("EstadoInmovilizado") = True
                        DtNewFCL.ImportRow(DtFCL.Rows(0))
                        Dim DrNewElementoFCL As DataRow = DtNewElementoAmortizableFCL.NewRow
                        DrNewElementoFCL("IdLinea") = AdminData.GetAutoNumeric
                        DrNewElementoFCL("IdLineaFactura") = ArrIDs(j)
                        DrNewElementoFCL("IdElemento") = DrNew("IdElemento")
                        DrNewElementoFCL("NFactura") = DtFCL.Rows(0)("NFactura")
                        DrNewElementoFCL("IDCContable") = DtFCL.Rows(0)("CContable")
                        DrNewElementoFCL("ImporteA") = DtFCL.Rows(0)("ImporteA")
                        DrNewElementoFCL("ImporteB") = DtFCL.Rows(0)("ImporteB")
                        DtNewElementoAmortizableFCL.Rows.Add(DrNewElementoFCL)
                    End If
                Next

                '//// Analítica
                Dim ad As New AdminData
                StrSQLinversion = "SELECT IDCentroCoste, IDCentroGestion, SUM(tbFacturaCompraAnalitica.ImporteA) AS ImporteA" & vbCrLf & _
                "FROM tbFacturaCompraLinea INNER JOIN tbFacturaCompraAnalitica ON tbFacturaCompraLinea.IDLineaFactura = tbFacturaCompraAnalitica.IDLineaFactura" & vbCrLf & _
                "WHERE tbFacturaCompraLinea.IDLineaFactura IN (" & Replace(dr("LineasFactura").ToString(), StrDelimeter, ",") & ")" & vbCrLf & _
                "GROUP BY IDCentroCoste, IDCentroGestion"
                Dim CmdInversion As Common.DbCommand = AdminData.GetCommand
                CmdInversion.CommandType = CommandType.Text
                CmdInversion.CommandText = StrSQLinversion
                DtInversionANA = AdminData.Execute(CmdInversion, ExecuteCommand.ExecuteReader)
                DblImpTotalANA = 0
                For Each drAna As DataRow In DtInversionANA.Rows
                    DblImpTotalANA += drAna("ImporteA")
                Next
                For Each drAna As DataRow In DtInversionANA.Rows
                    Dim drElemAnaNew As DataRow = DtNewAnalitica.NewRow
                    drElemAnaNew("IdElemento") = DrNew("IdElemento")
                    drElemAnaNew("IDCentroGestion") = drAna("IDCentroGestion")
                    drElemAnaNew("IDCentroCoste") = drAna("IDCentroCoste")
                    drElemAnaNew("Porcentaje") = xRound(100 * (drAna("ImporteA") / DblImpTotalANA), 2)
                    DtNewAnalitica.Rows.Add(drElemAnaNew)
                Next
                i += 1
            Next
            'Insertar los elementos
            If Not DtNew Is Nothing Then
                If DtNew.Rows.Count > 0 Then
                    ProcessServer.ExecuteTask(Of Object)(AddressOf Comunes.BeginTransaction, Nothing, services)
                    BusinessHelper.UpdateTable(DtNew)
                    BusinessHelper.UpdateTable(DtNewFCL)
                    BusinessHelper.UpdateTable(DtNewElementoAmortizableFCL)
                    BusinessHelper.UpdateTable(DtNewAnalitica)
                    DtNewElementoRevalorizacion = ProcessServer.ExecuteTask(Of DataTable, DataTable)(AddressOf ElementoRevalorizacion.CrearRevalorizacionesElementosFactura, DtNew, services)
                    If Not DtNewElementoRevalorizacion Is Nothing AndAlso DtNewElementoRevalorizacion.Rows.Count > 0 Then
                        BusinessHelper.UpdateTable(DtNewElementoRevalorizacion)
                    End If
                End If
            End If
        End If
    End Sub

#End Region

#Region "Funciones Obtener"

    <Serializable()> _
    Public Class DataObtenerAmort
        Public ValorTotal As Double
        Public VidaAmort As Integer
        Public IDTipo As Integer
        Public DtAmort As DataTable
        Public FechaInicio As Date
        Public EnTipo As enTipoAmort

        Public Sub New()
        End Sub

        Public Sub New(ByVal ValorTotal As Double, ByVal VidaAmort As Integer, ByVal IDTipo As Integer, ByVal DtAmort As DataTable, ByVal FechaInicio As Date, ByVal EnTipo As enTipoAmort)
            Me.ValorTotal = ValorTotal
            Me.VidaAmort = VidaAmort
            Me.IDTipo = IDTipo
            Me.DtAmort = DtAmort
            Me.FechaInicio = FechaInicio
            Me.EnTipo = EnTipo
        End Sub
    End Class

    <Task()> Public Shared Function ObtenerAmort(ByVal data As DataObtenerAmort, ByVal services As ServiceProvider) As DataTable
        'Funcion que RECIBE: los datos del elemento, los meses que se han amortizado y el tipo de amortizacion
        'y DEVUELVE:un rcs con lo amortizado por años y meses
        Dim DblAmortTeorico As Double 'Cantidad Amortizada hasta el Mes Actual con el nuevo cambio
        Dim DblValorMes As Double 'Cantidad a amortizar al mes en el Año actual
        Dim LngMesActual As Integer 'Mes actual
        Dim LngAñoActual As Integer 'AñoActual
        Dim DblDiferencia As Double 'Diferencia entre el valor de añocon redondeo y sin el
        Dim DblValorAñoSin As Double 'Valor del año SIN redondeo
        Dim DblValorAñoCon As Double 'Valor del año CON redondeo
        Dim DblValorAmortAcum As Double 'Valor que se ha amortizado hasta el momento
        Dim ClsTipo As New TipoAmortizacionLinea
        Dim DtTipo As New DataTable
        Dim MesAño As StMesAño
        Dim intContadorMes As Short
        Dim dblValorTemp As Double 'Valor que guarda el ValorMes para que al cambio de año no se
        'arrastre el ValorMes del ultimo mes con el ajuste de los decimales

        'La amortizacion se muestra en la moneda A, obtenemos un rcs de la moneda para hacer
        'los redondeos
        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
        Dim MonInfoA As MonedaInfo = Monedas.MonedaA

        'Obtenemos el rcs con el TipoAmortizacion del Elemento
        DtTipo = ClsTipo.Filter(New FilterItem("IdTipoAmortizacion", FilterOperator.Equal, data.IDTipo, FilterType.String), "NAño ASC")
        If DtTipo Is Nothing Then Exit Function
        intContadorMes = 1

        'Inicializamos las variables que nos van a servir para saber el año y mes en que nos
        'encontramos en cada momento
        MesAño = ObtenerMesAño(data.FechaInicio)
        LngMesActual = CShort(MesAño.Mes)
        LngAñoActual = CShort(MesAño.Año)

        'Calculamos cual va a ser el valor que se va a amortizar al mes al comienzo
        DblValorMes = xRound((data.ValorTotal * (DtTipo.Rows(0)("PorcentajeAmortizar") / 100)) / 12, MonInfoA.NDecimalesImporte)
        dblValorTemp = DblValorMes

        'Actualizamos los datos para cada mes de amortizacion

        Do While data.VidaAmort > 0
            DblValorMes = dblValorTemp
            If LngMesActual = 12 Or data.VidaAmort = 1 Then
                'Si es el ultimo mes del año o si es el ultimo mes de amortizacion, sumamos
                'los decimales que se han ido arrastrando
                If data.EnTipo <> enTipoAmort.enTeorica Then
                    If data.FechaInicio.Year = LngAñoActual Then
                        'Si es el primer año de amortizacion, el numero de meses seran: 12-MesInicio+1
                        DblValorAñoSin = ((data.ValorTotal * (DtTipo.Rows(0)("PorcentajeAmortizar") / 100)) / 12) * (12 - data.FechaInicio.Month + 1)
                        DblValorAñoCon = DblValorMes * (12 - data.FechaInicio.Month + 1)
                    Else
                        'Si no, el numero de meses coincide con el MesActual, ya que si estamos
                        'en un año intermedio MesActual sera 12, y si es el ultimo año, MesActual
                        'sera el num de meses transcurridos en ese ultimo año
                        DblValorAñoSin = ((data.ValorTotal * (DtTipo.Rows(0)("PorcentajeAmortizar") / 100)) / 12) * LngMesActual
                        DblValorAñoCon = DblValorMes * LngMesActual 'Valor real amortizado ese año
                    End If

                    'Calculamos el ValorMes actual, sumandole los decimales que se han ido perdiendo
                    DblDiferencia = DblValorAñoSin - DblValorAñoCon
                    dblValorTemp = DblValorMes
                    DblValorMes = xRound(DblValorMes + DblDiferencia, MonInfoA.NDecimalesImporte)
                End If
            End If

            If data.EnTipo <> enTipoAmort.enTeorica Then
                'Si la amortizacion no es teorica, se cargan los datos en el rcs
                DblValorAmortAcum = DblValorAmortAcum + DblValorMes
                If data.VidaAmort = 1 Then
                    DblValorMes = (data.ValorTotal - DblValorAmortAcum) + DblValorMes
                End If
                Dim StCarga As New DataCargarDtAmort(DblValorMes, data.DtAmort, LngMesActual, LngAñoActual, data.FechaInicio, data.EnTipo)
                data.DtAmort = ProcessServer.ExecuteTask(Of DataCargarDtAmort, DataTable)(AddressOf CargarDtAmort, StCarga, services)
            Else
                'Si la amortizacion es Teorica, va acumulando la Amortizacion teorica
                DblAmortTeorico = DblValorMes + DblAmortTeorico
            End If

            'Si el mes ha sido diciembre actualizamos el mes y el año
            If LngMesActual = 12 Then
                LngMesActual = 1
                LngAñoActual += 1
            Else : LngMesActual += 1
            End If

            'Actualizamos el contador de los meses y el TipoAmortizacion si es necesario
            If intContadorMes = 12 Then
                'Si ya han pasado 12 meses, se actualiza el Valor de mes con el nuevo porcentaje
                'rcsTipo.MoveNext()
                'If rcsTipo.EOF Then rcsTipo.MoveFirst()
                DblValorMes = xRound((data.ValorTotal * (DtTipo.Rows(0)("PorcentajeAmortizar") / 100)) / 12, MonInfoA.NDecimalesImporte)
                dblValorTemp = DblValorMes
                intContadorMes = 1
            Else
                intContadorMes += 1
            End If

            'Actualizamos los meses que quedan por amortizar
            data.VidaAmort -= 1
        Loop

        'Cargamos el rcs que se va a devolver, con los datos obtetenidos
        If data.EnTipo = enTipoAmort.enTeorica Then
            Dim DrNew As DataRow = data.DtAmort.NewRow()
            DrNew("AmortTeorico") = DblAmortTeorico
            DrNew("ValorMes") = DblValorMes
            data.DtAmort.Rows.Add(DrNew)
        Else
            ''''        Select Case enTipo
            ''''            Case enTecnica
            ''''                strCampo = "AmortTecnica"
            ''''            Case enFISCAL
            ''''                strCampo = "AmortFiscal"
            ''''        End Select
            ''''
            ''''        'Calculamos la los decimales que se han ido arrastrando
            ''''        rcsAmort.MoveFirst
            ''''        Do While Not rcsAmort.EOF
            ''''            dblAcum = dblAcum + rcsAmort.Fields(strCampo)
            ''''            rcsAmort.MoveNext
            ''''        Loop
            ''''
            ''''        'Sumamos la diferencia al valor del ultimo año
            ''''        dblDiferencia = dblValorTotal - dblAcum
            ''''        rcsAmort.MoveLast
            ''''        rcsAmort.Fields(strCampo).Value = xRound(rcsAmort.Fields(strCampo).Value + dblDiferencia, lngNDecimalesA)
            '''''        strMeses=left(rcsAmort.Fields(strCampomes).Value
            '''''        intPos = InStr(rcsAmort.Fields(strCampoMes).Value, ";")
            '''''        strUltimoMes = Mid(rcsAmort.Fields(strCampo).Value, intPos)
            '''''        dblUltimoMes=getproperyvalue(rcsAmort.Fields(strCampo).Value,)
            '''''        dblValorMes dblDiferencia
            ''''
        End If
        Return data.DtAmort
    End Function

    <Serializable()> _
    Public Class DataObtenerAmortCont
        Public IDElem As String
        Public Dt As DataTable

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDElem As String, ByVal Dt As DataTable)
            Me.IDElem = IDElem
            Me.Dt = Dt
        End Sub
    End Class

    <Task()> Public Shared Function ObtenerAmortCont(ByVal data As DataObtenerAmortCont, ByVal services As ServiceProvider) As DataTable
        'Funcion que devuelve un rcs con la Amortizacion Contable por meses y años, de un Elemento
        Dim IntMesActual As Short 'Mes en curso de la amortizacion
        Dim IntAñoActual As Short 'Año en curso de la amortizacion
        Dim IntMesCambio As Short 'Mes del ultimo cambio
        Dim IntAñoCambio As Short 'Año del ultimo cambio
        Dim ClsCambio As New ElementoRevalorizacion
        Dim ClsTipo As New TipoAmortizacionLinea
        Dim DtCambio As New DataTable
        Dim DtElem As New DataTable
        Dim DtTipo As New DataTable

        Dim DblValorTotal As Double 'Total a amortizar
        Dim DblValorNeto As Double 'Cantidad que queda por amortizar
        Dim DblValorAmort As Double 'Cantidad Amortizada
        Dim DblValorMes As Double 'Cantidad a amortizar en el mes en curso
        Dim DblValorMesCambio As Double
        Dim DblDiferencia As Double 'Diferencia entre el valor de añocon redondeo y sin el
        Dim DblValorAñoSin As Double 'Valor del año SIN redondeo
        Dim DblValorAñoCon As Double 'Valor del año CON redondeo
        Dim DblAcum As Double

        Dim IntVidaUtil As Short 'Vida util total del elemento
        Dim IntVidaUtilRes As Short 'Vida util restante por amortizar

        Dim BlnCambio As Boolean 'Inidica si estamos en un mes de cambio o no
        Dim Result As New StDatosCambio
        Dim StrIDTipo As String
        Dim IntAñoTipo As Short
        Dim LngCadena As Integer
        Dim IntContadorMes As Short 'Contabiliza el numero de meses amortizados hasta el momento
        Dim IntContadorAño As Short 'Idem para el numero de años
        Dim DtmFecha As Date
        Dim MesAño As New StMesAño
        Dim DblValorTemp As Double 'Valor que guarda el ValorMes para que al cambio de año no se
        'arrastre el ValorMes del ultimo mes con el ajuste de los decimales
        Dim IntMesCambioSig As Short
        Dim IntAñoCambioSig As Short
        Dim blnUltimoCambio As Boolean

        'Obtenemos el Mes y el Año (de Fecha de Compra) desde el que se va a empezar la amortizacion
        DtElem = New ElementoAmortizable().SelOnPrimaryKey(data.IDElem)

        DtmFecha = DtElem.Rows(0)("FechaInicioContabilizacion")
        '    intMesActual = rcsElem.Fields("[MesInicioContabilizado").Value]
        '    intAñoActual = rcsElem.Fields("[AñoInicioContabilizado").Value]
        MesAño = ObtenerMesAño(DtmFecha)
        IntMesActual = CShort(MesAño.Mes)
        IntAñoActual = CShort(MesAño.Año)

        IntContadorMes = 1
        IntContadorAño = 1

        'Obtenemos la monedaA para coger los decimales con los que se va a redondear
        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
        Dim MonInfoA As MonedaInfo = Monedas.MonedaA

        'Se cogen los dos ultimos cambios del Elemento, que son el ultimo y el anterior a la
        'FechaUltimaContabilizacion del elemento
        DtCambio = ClsCambio.Filter(New FilterItem("IdElemento", FilterOperator.Equal, data.IDElem, FilterType.String), "FechaRevalorizacion ASC")

        'Inicializamos las variables con los datos del primer cambio
        If Not DtCambio Is Nothing Then
            If DtCambio.Rows.Count <> 0 Then
                'Se coge la fecha del siguiente cambio al actual, para saber en que fecha tendremos
                'que coger el siguiente cambio
                MesAño = ObtenerMesAño(DtCambio.Rows(0)("FechaRevalorizacion"))
                IntMesCambio = CShort(MesAño.Mes)
                IntAñoCambio = CShort(MesAño.Año)

                'Inicializamos los valores con los datos del primer cambio
                DblValorTotal = DtCambio.Rows(0)("ValorCompraFechaA") - DtCambio.Rows(0)("ValorResidualFechaA")
                DblValorNeto = DblValorTotal
                DblValorAmort = 0

                IntVidaUtil = DtCambio.Rows(0)("VidaUtilFecha")
                IntVidaUtilRes = IntVidaUtil

                DtTipo = ClsTipo.Filter(New FilterItem("IdTipoAmortizacion", FilterOperator.Equal, DtCambio.Rows(0)("IDTipoAmortizacionFecha"), FilterType.String), "NAño ASC")

                If Not DtTipo Is Nothing Then
                    If DtTipo.Rows.Count <> 0 Then
                        StrIDTipo = DtTipo.Rows(0)("idTipoAmortizacion")
                        DblValorTemp = xRound((DblValorTotal * (DtTipo.Rows(0)("PorcentajeAmortizar") / 100)) / 12, MonInfoA.NDecimalesImporte)
                    End If
                End If
            Else
                'Si el elem no tiene ningun cambio guardado, error
                Return data.Dt
                Exit Function
            End If
        End If

        BlnCambio = False

        Do While IntVidaUtilRes > 0
            If DtCambio.Rows.Count > 0 Then 'If rcsCambio.RecordCount > 1 Then

                DblValorMes = DblValorTemp

                'Si la fecha Actual de la amortizacion es igual a la del cambio, se coge el ste cambio
                If IntMesCambio = IntMesActual And IntAñoCambio = IntAñoActual Then
                    Dim StCambio As New DataObtenerDatosCambio(DtCambio, IntVidaUtil, DblValorAmort, IntVidaUtilRes, DtmFecha)
                    Result = ProcessServer.ExecuteTask(Of DataObtenerDatosCambio, StDatosCambio)(AddressOf ObtenerDatosCambio, StCambio, services)
                    If Not IsDBNull(Result) Then
                        'Cogemos los valores que nos ha devuelto la funcion ObtenerDatosCambio
                        DblValorTotal = CDbl(Result.ValorTotal)
                        DblValorMes = CDbl(Result.ValorMesCambio)
                        DblValorTemp = DblValorMes
                        IntVidaUtil = CShort(Result.VidaUtil)
                        IntVidaUtilRes = CShort(Result.VidaUtilRes)
                        If Result.IDTipo <> StrIDTipo Then
                            'Cargamos en un rcs el nuevo TipoAmortizacion
                            StrIDTipo = Result.IDTipo
                            Dim FilTipo As New Filter
                            FilTipo.Add("IdTipoAmortizacion", FilterOperator.Equal, StrIDTipo, FilterType.String)
                            FilTipo.Add("NAño", FilterOperator.Equal, IntContadorAño, FilterType.Numeric)
                            DtTipo = ClsTipo.Filter(FilTipo, "Naño ASC")
                        End If
                        BlnCambio = True
                    Else
                        BlnCambio = False
                    End If

                    'Cogemos la fecha del ste cambio
                    If Not DtCambio Is Nothing Then
                        MesAño = ObtenerMesAño(DtCambio.Rows(0)("FechaRevalorizacion"))
                        IntMesCambio = CShort(MesAño.Mes)
                        IntAñoCambio = CShort(MesAño.Año)

                        'Comprobamos que si ste cambio es en el mismo mes
                        For i As Integer = 0 To DtCambio.Rows.Count - 1
                            MesAño = ObtenerMesAño(DtCambio.Rows(i)("FechaRevalorizacion"))
                            IntMesCambioSig = CShort(MesAño.Mes)
                            IntAñoCambioSig = CShort(MesAño.Año)
                            If IntMesCambioSig = IntMesCambio AndAlso IntMesCambioSig = IntMesCambio Then
                                i += 1
                            Else
                                i -= 1
                                Exit For
                            End If
                        Next
                    End If
                End If
            End If

            If IntMesActual = 12 Or IntVidaUtilRes = 1 Then
                'Si es el ultimo mes del año o si es el ultimo mes de amortizacion, sumamos
                'los decimales que se han ido arrastrando
                If DtmFecha.Year = IntAñoActual Then
                    DblValorAñoSin = ((DblValorTotal * (DtTipo.Rows(0)("PorcentajeAmortizar") / 100)) / 12) * (12 - DtmFecha.Month)
                    If Not IsDBNull(Result) Then
                        DblValorAñoCon = CDbl(Result.ValorMes) * (12 - DtmFecha.Month)
                    Else
                        DblValorAñoCon = DblValorMes * (12 - DtmFecha.Month)
                    End If
                Else
                    DblValorAñoSin = ((DblValorTotal * (DtTipo.Rows(0)("PorcentajeAmortizar") / 100)) / 12) * IntMesActual
                    DblValorAñoCon = DblValorMes * IntMesActual 'Valor real amortizado ese año
                End If
                DblDiferencia = DblValorAñoSin - DblValorAñoCon
                DblValorTemp = DblValorMes 'Guardamos el ValorMes sin ajuste de decimales
                DblValorMes = xRound(DblValorMes + DblDiferencia, MonInfoA.NDecimalesImporte)
            End If

            'Actualizamos el valorNeto y el ValorAmort
            If Not BlnCambio Then
                DblValorAmort += DblValorMes
                DblValorNeto = DblValorTotal - DblValorAmort
            Else
                'Si ha habido un cambio, cogemos los valores del string que nos ha devuelto la funcion ObtenerAmortTeorica
                DblValorNeto = CDbl(Result.ValorNeto)
                DblValorAmort = CDbl(Result.ValorAmort)
                IntContadorMes = 1
            End If

            'Cargamos el Dt
            Dim StCarga As New DataCargarDtAmort(DblValorMes, data.Dt, IntMesActual, IntAñoActual, DtmFecha, enTipoAmort.enContable)
            data.Dt = ProcessServer.ExecuteTask(Of DataCargarDtAmort, DataTable)(AddressOf CargarDtAmort, StCarga, services)
            If IntContadorMes = 12 Then
                'Si cambia el año, cambia el porcentaje amortizar, por lo que hay que calcular de nuevo
                'el ValorMes
                DblValorMes = xRound(((DblValorTotal * DtTipo.Rows(0)("PorcentajeAmortizar")) / 100) / 12, MonInfoA.NDecimalesImporte)
                DblValorTemp = DblValorMes
                IntContadorMes = 1
                IntContadorAño += 1
            Else
                IntContadorMes += 1
            End If

            'Actualizamos las variables (Año, Mes, VidaUtilRes y ValorMes si ha habido un cambio)
            'para el ste mes
            If IntMesActual = 12 Then
                IntAñoActual += 1
                IntMesActual = 1
            Else
                IntMesActual += 1
            End If

            If BlnCambio Then
                DblValorMes = CDbl(Result.ValorMes)
                DblValorTemp = DblValorMes
                BlnCambio = False
            End If
            IntVidaUtilRes -= 1
        Loop

        'Calculamos la los decimales que se han ido arrastrando
        If Not data.Dt Is Nothing Then
            If Not data.Dt.Rows.Count > 0 Then
                For Each Dr As DataRow In data.Dt.Select
                    'Se va almacenanado en dblAcum el total amortizado
                    If Not AreEquals(Dr("AmortContable"), 0) Then
                        DblAcum += Dr("AmortContable")
                    Else
                        'Si es 0, es porque se ha llegado al final de la amortizacion y no tenemos que
                        'seguir avanzado en el rcs para mantener la posicion donde tenemos que añadir
                        'los decimales que se han perdido
                        Exit For
                    End If
                Next

                'Nos movemos al año de la ultima amortizacion, que es donde se hara el ajuste de los
                'decimales
                'If CDbl(rcs.AbsolutePosition) <> 1 Then
                'Si el rcs esta en la primera amortizacion, es porque todos los meses se amortizan 0 pts,
                'asi que no hay que realizar ajuste. Esto pasa cuando el Valor del elemento es 0.
                'If rcs.EOF Or CDbl(rcs.AbsolutePosition) = 1 Then
                '    rcs.MoveLast()
                'Else
                '    rcs.MovePrevious()
                'End If

                'Sumamos la diferencia al valor del ultimo año y al del mes del ultimo año
                DblDiferencia = DblValorTotal - DblAcum
                data.Dt.Rows(0)("AmortContable") = xRound(data.Dt.Rows(0)("AmortContable") + DblDiferencia, MonInfoA.NDecimalesImporte)

                'Tambien sumamos esta diferencia  al ultimo mes del ultimo año
                If IntMesActual = 1 Then
                    IntMesActual = 12
                Else
                    IntMesActual -= 1
                End If
                'TODO GetPropertyValue
                'LngCadena = Length(Dt.Rows(0)("AmortContableMensual")) - (Length(GetPropertyValue(Dt.Rows(0)("AmortContableMensual"), "Amortizacion" & IntMesActual)) + Length(";"))
                'DblValorMes = Nz(GetPropertyValue(Dt.Rows(0)("AmortContableMensual"), "Amortizacion" & IntMesActual), 0)
                data.Dt.Rows(0)("AmortContableMensual") = Left(data.Dt.Rows(0)("AmortContableMensual"), LngCadena)
                data.Dt.Rows(0)("AmortContableMensual") = data.Dt.Rows(0)("AmortContableMensual") & xRound(DblValorMes + DblDiferencia, MonInfoA.NDecimalesImporte) & ";"
            End If
        End If
        Return data.Dt
    End Function

    <Serializable()> _
    Public Class DataObtenerAmortReal
        Public IDElem As String
        Public ValorElem As Double
        Public DtAmort As DataTable

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDElem As String, ByVal ValorElem As Double, ByVal DtAmort As DataTable)
            Me.IDElem = IDElem
            Me.ValorElem = ValorElem
            Me.DtAmort = DtAmort
        End Sub
    End Class

    <Task()> Public Shared Function ObtenerAmortReal(ByVal data As DataObtenerAmortReal, ByVal services As ServiceProvider) As DataTable
        'Declaración de constantes, variables u objetos locales
        Dim DtAmortReg As New DataTable
        Dim StrSelect As String
        Dim DblAcum As Double
        Dim IntMes As Short
        Dim IntAño As Short
        Dim IntPos As Short
        Dim BlnFin As Boolean
        Dim DblValorNeto As Double
        Dim DblValorNetoAcum As Double

        'Comienzo del Cuerpo de la Función
        StrSelect = "MesContabilizacion, AñoContabilizacion, ValorAmortizadoA"
        DtAmortReg = AdminData.Filter("tbAmortizacionRegistro", StrSelect, "IdElemento='" & data.IDElem & "'", "MesContabilizacion, AñoContabilizacion")

        If Not data.DtAmort Is Nothing Then
            DblValorNeto = data.ValorElem
            DblValorNetoAcum = DblValorNeto

            'Recorremos cada año del  dt que contiene las amortizaciones
            For Each Dr As DataRow In data.DtAmort.Select
                IntAño = Dr("año")
                Dim DrRegA() As DataRow = DtAmortReg.Select("AñoContabilizacion=" & IntAño)
                DblAcum = 0

                For IntMes = 1 To 12
                    Dim DrRegAM() As DataRow = DtAmortReg.Select("AñoContabilizacion=" & IntAño & " AND MesContabilizacion=" & IntMes)
                    If DrRegAM.Length > 0 Then
                        For Each DrReg As DataRow In DrRegAM
                            DblAcum += DrReg("ValorAmortizadoA")
                        Next
                    End If
                    'IntPos = CShort(rcsAmort.AbsolutePosition)
                    Dim StCarga As New DataCargarDtAmort(DblAcum, data.DtAmort, IntMes, IntAño, Today, enTipoAmort.enRealizada, DblValorNeto)
                    data.DtAmort = ProcessServer.ExecuteTask(Of DataCargarDtAmort, DataTable)(AddressOf CargarDtAmort, StCarga, services)
                    DblValorNetoAcum -= DblAcum
                    DblAcum = 0
                Next
            Next
        End If
        Return data.DtAmort
    End Function

    <Serializable()> _
    Public Class DataObtenerDatosCambio
        Public DtCambio As DataTable
        Public VidaUtil As Integer
        Public ValorAmort As Double
        Public VidaRes As Integer
        Public FechaCompra As Date

        Public Sub New()
        End Sub

        Public Sub New(ByVal DtCambio As DataTable, ByVal VidaUtil As Integer, ByVal ValorAmort As Double, ByVal VidaRes As Integer, ByVal FechaCompra As Date)
            Me.DtCambio = DtCambio
            Me.VidaUtil = VidaUtil
            Me.ValorAmort = ValorAmort
            Me.VidaRes = VidaRes
            Me.FechaCompra = FechaCompra
        End Sub
    End Class

    <Task()> Public Shared Function ObtenerDatosCambio(ByVal data As DataObtenerDatosCambio, ByVal services As ServiceProvider) As StDatosCambio
        'Funcion que devuelve un string con los datos del cambio (ElementoRevalorizacion)
        Dim DblValorMesCambio As Double 'Valor del Mes del Cambio con la correccion
        Dim DblDifMes As Double 'Valor de la diferencia para hacer la correccion del Mes
        Dim DblValorMes As Double 'Valor de los Meses a partir del cambio
        Dim IntMesesAmort As Short 'Meses amortizados hasta el momento
        Dim DblValorTotal, DblValorNeto, StrIDTipo As String
        Dim DtResult As New DataTable

        'Creamos un rcs donde ObtenerAmort nos devolvera el ValorAmortTeorico y el nuevo ValorMes
        DtResult.Columns.Add("AmortTeorico", GetType(Double))
        DtResult.Columns.Add("ValorMes", GetType(Double))

        If Not data.DtCambio Is Nothing AndAlso data.DtCambio.Rows.Count > 0 Then
            Dim DatosCambio As New StDatosCambio
            DatosCambio.ValorTotal = data.DtCambio.Rows(0)("ValorCompraFechaA") - data.DtCambio.Rows(0)("ValorResidualFechaA")
            IntMesesAmort = data.VidaUtil - data.VidaRes 'Meses que se han amortizado hasta el momento
            DatosCambio.VidaUtil = data.DtCambio.Rows(0)("VidaUtilFecha") 'Nueva vida util
            DatosCambio.VidaUtilRes = DatosCambio.VidaUtil - IntMesesAmort 'Nueva vida restante

            'Obtenemos cuanto se habria amortizado hasta el momento con las nuevas condiciones AmortizacionTeorica
            Dim StAmort As New DataObtenerAmort(DatosCambio.ValorTotal, IntMesesAmort, data.DtCambio.Rows(0)("IDTipoAmortizacionFecha"), DtResult, data.FechaCompra, enTipoAmort.enTeorica)
            DtResult = ProcessServer.ExecuteTask(Of DataObtenerAmort, DataTable)(AddressOf ObtenerAmort, StAmort, services)
            DblDifMes = DtResult.Rows(0)("AmortTeorico") - data.ValorAmort
            DatosCambio.ValorMes = DtResult.Rows(0)("ValorMes")
            DatosCambio.ValorMesCambio = DblDifMes + DatosCambio.ValorMes
            DatosCambio.ValorAmort = data.ValorAmort + DatosCambio.ValorMesCambio
            DatosCambio.ValorNeto = DatosCambio.ValorTotal - DatosCambio.ValorAmort
            DatosCambio.IDTipo = data.DtCambio.Rows(0)("IDTipoAmortizacionFecha")
            Return DatosCambio
        Else : Return Nothing
        End If
    End Function

    <Serializable()> _
    Public Class StDatosCambio
        Public ValorTotal As Double
        Public ValorNeto As Double
        Public ValorAmort As Double
        Public ValorMes As Double
        Public ValorMesCambio As Double
        Public IDTipo As String
        Public VidaUtil As Integer
        Public VidaUtilRes As Integer
    End Class
    <Serializable()> _
    Public Class DataCargarDtAmort
        Public ValorMes As Double
        Public Dt As DataTable
        Public MesActual As Integer
        Public AñoActual As Integer
        Public FechaIni As Date
        Public EnTipo As enTipoAmort
        Public ValorNeto As Double

        Public Sub New()
        End Sub

        Public Sub New(ByVal ValorMes As Double, ByVal Dt As DataTable, ByVal MesActual As Integer, ByVal AñoActual As Integer, ByVal FechaIni As Date, ByVal EnTipo As enTipoAmort, Optional ByVal ValorNeto As Double = 0)
            Me.ValorMes = ValorMes
            Me.Dt = Dt
            Me.MesActual = MesActual
            Me.AñoActual = AñoActual
            Me.FechaIni = FechaIni
            Me.EnTipo = EnTipo
            Me.ValorNeto = ValorNeto
        End Sub
    End Class

    <Task()> Public Shared Function CargarDtAmort(ByVal data As DataCargarDtAmort, ByVal services As ServiceProvider) As DataTable
        'Funcion que actualiza el Dt de las amortizaciones. Se le llama una vez por cada mes amortizado
        Dim DblAmortAño As Double
        Dim StrMeses, StrAnual, StrMensual As String
        Dim DblDiferencia, DblValorAcum As Double
        'Guardamos el nombre de los campos del rcs en funcion del tipo
        Select Case data.EnTipo
            Case enTipoAmort.enContable
                StrAnual = "AmortContable"
                StrMensual = "AmortContableMensual"
            Case enTipoAmort.enTecnica
                StrAnual = "AmortTecnica"
                StrMensual = "AmortTecnicaMensual"
            Case enTipoAmort.enFiscal
                StrAnual = "AmortFiscal"
                StrMensual = "AmortFiscalMensual"
            Case enTipoAmort.enRealizada
                StrAnual = "AmortRealizada"
                StrMensual = "AmortRealizadaMensual"
        End Select
        'Inicializamos la cadena con la amortizacion por meses
        StrMeses = "Mes" & data.MesActual & "=" & data.MesActual & ";Amortizacion" & data.MesActual & "=" & data.ValorMes & ";"
        If Not data.Dt Is Nothing Then
            If data.Dt.Rows.Count = 0 Then
                'Si todavia no se ha creado el rcs, se inserta la primera linea
                Dim DrNew As DataRow = data.Dt.NewRow()
                DrNew("año") = data.AñoActual
                DrNew(StrAnual) = data.ValorMes
                DrNew(StrMensual) = StrMeses
                If data.EnTipo = enTipoAmort.enRealizada Then DrNew("ValorNeto") = data.ValorNeto
                data.Dt.Rows.Add(DrNew)
            Else
                If data.MesActual = 1 Then
                    'Si el mes es Enero, se añade una nueva linea
                    'If rcs.EOF Then
                    '    'Si el  registro actual no esta creado
                    '    rcs.AddNew()
                    '    rcs.Fields("año").Value = IntAñoActual
                    'End If
                    data.Dt.Rows(0)(StrAnual) = data.ValorMes
                    data.Dt.Rows(0)(StrMensual) = StrMeses
                    If data.EnTipo = enTipoAmort.enRealizada Then data.Dt.Rows(0)("ValorNeto") = data.ValorNeto
                Else
                    'Si el mes no es enero, se incrementa el valor del año con el actual ValorMes
                    'y se concatena en el string de los meses
                    data.Dt.Rows(0)(StrAnual) = data.ValorMes + data.Dt.Rows(0)(StrAnual)
                    data.Dt.Rows(0)(StrMensual) = data.Dt.Rows(0)(StrMensual) & StrMeses
                    If data.EnTipo = enTipoAmort.enRealizada Then data.Dt.Rows(0)("ValorNeto") = data.ValorNeto
                End If
            End If
            'Si estamos en Diciembre nos movemos al ste registro
            'If IntMesActual = 12 Then rcs.MoveNext()
        End If
        Return data.Dt
    End Function

    <Task()> Public Shared Function ObtenerDotacionActual(ByVal StrIdElem As String, ByVal services As ServiceProvider) As String
        'Obtiene la dotacion de un elemento amortizable en funcion de la fecha actual
        Dim ClsElemRev As New ElementoRevalorizacion
        Dim ClsTALinea As New TipoAmortizacionLinea
        Dim DtAmort As New DataTable
        Dim DtElemRev As New DataTable
        Dim DtTALinea As New DataTable
        Dim DtElem As New DataTable
        Dim IntMesActual, IntAñoActual, IntNAño As Short
        Dim IntAño, IntMes As Short
        Dim StrMeses, StrIdTipoContable As String
        Dim StrWhere As String
        Dim DblDotContable, DblDotTecnica, DblDotFiscal As Double
        Dim IntVidaContable, IntPorcContable As Double
        Dim IntPorcFiscal, IntPorcTecnico As Double
        Dim DblValor As Double
        Dim DtmFecha As Date
        Dim MesAño As New StMesAño

        'Comienzo del Cuerpo de la Función
        DtElem = New ElementoAmortizable().SelOnPrimaryKey(StrIdElem)

        If Not DtElem Is Nothing Then
            If DtElem.Rows.Count > 0 Then
                DtAmort = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf CrearDtAmort, Nothing, services)
                If (DtElem.Rows(0)("MesInicioContabilizado") <= Today.Month AndAlso AreEquals(DtElem.Rows(0)("AñoInicioContabilizado"), Today.Year)) OrElse DtElem.Rows(0)("AñoInicioContabilizado") <= Today.Year Then
                    'Calculamos la Dotacion si la fechaInicio no es mayor que la fecha actual
                    DtmFecha = DtElem.Rows(0)("FechaInicioContabilizacion")
                    DblValor = DtElem.Rows(0)("ValorTotalRevalElementoA") - DtElem.Rows(0)("ValorResidualA")
                    'TECNICA
                    Dim StAmort As New DataObtenerAmort(DblValor, DtElem.Rows(0)("VidaTecnicaElemento"), DtElem.Rows(0)("IDCodigoAmortizacionTecnica"), DtAmort, DtmFecha, enTipoAmort.enTecnica)
                    DtAmort = ProcessServer.ExecuteTask(Of DataObtenerAmort, DataTable)(AddressOf ObtenerAmort, StAmort, services)
                    'FISCAL
                    If DtAmort Is Nothing Then Exit Function
                    Dim StAmortFis As New DataObtenerAmort(DblValor, DtElem.Rows(0)("VidaFiscalElemento"), DtElem.Rows(0)("IDCodigoAmortizacionFiscal"), DtAmort, DtmFecha, enTipoAmort.enFiscal)
                    DtAmort = ProcessServer.ExecuteTask(Of DataObtenerAmort, DataTable)(AddressOf ObtenerAmort, StAmortFis, services)
                    'CONTABLE
                    Dim StAmortCont As New DataObtenerAmortCont(DtElem.Rows(0)("IdElemento"), DtAmort)
                    DtAmort = ProcessServer.ExecuteTask(Of DataObtenerAmortCont, DataTable)(AddressOf ObtenerAmortCont, StAmortCont, services)
                    IntMesActual = Today.Month
                    IntAñoActual = Today.Year

                    'Cogemos el valor de las dotaciones correspondiente al mes actual
                    If Not DtAmort Is Nothing Then
                        If DtAmort.Rows.Count > 0 Then
                            Dim DrAmort() As DataRow = DtAmort.Select("Año=" & IntAñoActual)
                            If DrAmort.Length = 0 Then
                                DblDotContable = 0
                            Else
                                StrMeses = DrAmort(0)("AmortContableMensual")
                                'TODO
                                'If GetPropertyValue(StrMeses, "Amortizacion" & IntMesActual) = String.Empty Then
                                '    DblDotContable = 0
                                'Else
                                '    DblDotContable = CDbl(GetPropertyValue(StrMeses, "Amortizacion" & IntMesActual))
                                'End If
                            End If
                            If DrAmort.Length = 0 Then
                                DblDotTecnica = 0
                            Else
                                StrMeses = DrAmort(0)("AmorttecnicaMensual")
                                'TODO GetPropertyValue
                                'If GetPropertyValue(StrMeses, "Amortizacion" & IntMesActual) = String.Empty Then
                                '    DblDotTecnica = 0
                                'Else
                                '    DblDotTecnica = CDbl(GetPropertyValue(StrMeses, "Amortizacion" & IntMesActual))
                                'End If
                            End If

                            If DrAmort.Length = 0 Then
                                DblDotFiscal = 0
                            Else
                                StrMeses = DrAmort(0)("AmortFiscalMensual")
                                'TODO GetPropertyValue
                                'If GetPropertyValue(StrMeses, "Amortizacion" & IntMesActual) = String.Empty Then
                                '    DblDotFiscal = 0
                                'Else
                                '    DblDotFiscal = CDbl(GetPropertyValue(StrMeses, "Amortizacion" & IntMesActual))
                                'End If
                            End If
                        Else
                            ApplicationService.GenerateError("Falta el Elemento Revalorización")
                            Exit Function
                        End If
                    Else
                        ApplicationService.GenerateError("Falta el Elemento Revalorización")
                        Exit Function
                    End If

                    'Cogemos el TipoAmortizacion y la vida Contable
                    MesAño = ObtenerMesAño(Today)
                    IntMes = CShort(MesAño.Mes)
                    IntAño = CShort(MesAño.Año)

                    'En este rcs tenemos el cambio que se aplica en la fecha actual
                    Dim oF1 As New Filter
                    oF1.Add("IdElemento", FilterOperator.Equal, StrIdElem)
                    oF1.Add("FechaRevalorizacion", FilterOperator.LessThanOrEqual, Today)

                    DtElemRev = ClsElemRev.Filter(oF1, "FechaRevalorizacion DESC, IdLineaRevalorizacion DESC", "TOP 1 *")
                    If Not DtElemRev Is Nothing Then
                        If DtElemRev.Rows.Count > 0 Then
                            StrIdTipoContable = DtElemRev.Rows(0)("IDTipoAmortizacionFecha")
                            IntVidaContable = DtElemRev.Rows(0)("VidaUtilFecha")
                        Else
                            'ERROR: Hay un error en los datos porque todos los Elementos deben tener al menos un
                            'Elemento Revalorizacion
                            ApplicationService.GenerateError("Falta el Elemento Revalorización")
                            Exit Function
                        End If
                    Else
                        'ERROR: Hay un error en los datos porque todos los Elementos deben tener al menos un
                        'Elemento Revalorizacion
                        ApplicationService.GenerateError("Falta el Elemento Revalorización")
                        Exit Function
                    End If

                    'Cogemos los porcentajes  correspondientes al año actual
                    '(El Porcentaje tecnico y fiscal sera el mismo)
                    DtTALinea = ClsTALinea.Filter(, "IdTipoAmortizacion='" & DtElem.Rows(0)("IDCodigoAmortizacionFiscal") & "'")
                    If Not DtTALinea Is Nothing Then
                        If DtTALinea.Rows.Count > 0 Then
                            IntNAño = (IntAñoActual - DtmFecha.Year) + 1
                            Dim DrLinea() As DataRow = DtTALinea.Select("NAño=" & IntNAño)
                            If DrLinea.Length > 0 Then
                                IntPorcTecnico = DrLinea(0)("PorcentajeAmortizar")
                                IntPorcFiscal = IntPorcTecnico
                            End If
                        End If
                    End If

                    If AreEquals(DtElem.Rows(0)("IDCodigoAmortizacionContable"), DtElem.Rows(0)("IDCodigoAmortizacionFiscal")) AndAlso DtElemRev.Rows.Count = 1 Then
                        'Si los tipos son iguales y el Elemento no tiene ningun cambio, los porcentajes seran los
                        'mismos en los 3 casos
                        IntPorcContable = IntPorcTecnico
                    Else
                        If Not AreEquals(DtElem.Rows(0)("IDCodigoAmortizacionContable"), DtElem.Rows(0)("IDCodigoAmortizacionFiscal")) Then
                            'Si los tipo son distintos, cogemos los datos del TipoContable
                            DtTALinea = ClsTALinea.Filter(New FilterItem("IdTipoAmortizacion", FilterOperator.Equal, StrIdTipoContable, FilterType.String))
                        End If
                        If Not DtTALinea Is Nothing Then
                            If DtTALinea.Rows.Count > 0 Then
                                If DtElemRev.Rows.Count > 0 Then
                                    IntNAño = (IntAñoActual - CDate(DtElemRev.Rows(0)("FechaRevalorizacion")).Year) + 1
                                    Dim DrLinea() As DataRow = DtTALinea.Select("NAño=" & IntNAño)
                                    If DrLinea.Length > 0 Then IntPorcContable = DrLinea(0)("PorcentajeAmortizar")
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Return "DotacionContable=" & DblDotContable & ";DotacionTecnica=" & DblDotTecnica & ";DotacionFiscal=" & DblDotFiscal & ";TipoAmortContable=" & StrIdTipoContable & ";VidaContable=" & IntVidaContable & ";PorcentajeContable=" & IntPorcContable & ";PorcentajeTecnico=" & IntPorcTecnico & ";PorcentajeFiscal=" & IntPorcFiscal
    End Function

    <Serializable()> _
    Public Class DataObtenerDotPrimerMes
        Public Dt As DataTable
        Public Tipo As enTipoAmort

        Public Sub New(ByVal Dt As DataTable, ByVal Tipo As enTipoAmort)
            Me.Dt = Dt
            Me.Tipo = Tipo
        End Sub
    End Class

    <Task()> Public Shared Function ObtenerDotacionPrimerMes(ByVal data As DataObtenerDotPrimerMes, ByVal services As ServiceProvider) As Double
        'Busca  el valor del primer mes distinto de 0
        Dim DblPorcentaje As Double = 0
        Select Case data.Tipo
            Case enTipoAmort.enContable
                DblPorcentaje = data.Dt.Rows(0)("PorcentajeAnualContable") / 100
            Case enTipoAmort.enFiscal
                DblPorcentaje = data.Dt.Rows(0)("PorcentajeAnualFiscal") / 100
            Case enTipoAmort.enTecnica
                DblPorcentaje = data.Dt.Rows(0)("PorcentajeAnualTecnico") / 100
        End Select

        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
        Dim MonInfoA As MonedaInfo = Monedas.MonedaA
        Dim DblValorAmort As Double = data.Dt.Rows(0)("ValorTotalRevalElementoA") - data.Dt.Rows(0)("ValorResidualA")
        Return xRound((DblValorAmort * DblPorcentaje) / 12, MonInfoA.NDecimalesImporte)
    End Function

    <Serializable()> _
    Public Class DataObtenerMesAño
        Public Fecha As Date
        Public PorMeses As Boolean

        Public Sub New()
        End Sub

        Public Sub New(ByVal Fecha As Date, ByVal PorMeses As Boolean)
            Me.PorMeses = PorMeses
        End Sub
    End Class

    <Task()> Public Shared Function ObtenerMesAño(ByVal DtmFecha As Date, Optional ByVal BlnPorMeses As Boolean = 1) As StMesAño
        'Devuelve el Mes y el año en los que se empezará la proxima amortizacion
        Dim IntMes, IntAño As Short
        If BlnPorMeses Then
            If DtmFecha.Day > 1 Then
                IntMes = DtmFecha.Month + 1
            Else
                IntMes = DtmFecha.Month
            End If
        Else
            IntMes = DtmFecha.Month
        End If
        If IntMes = 13 Then
            IntAño = DtmFecha.Year + 1
            IntMes = 1
        Else
            IntAño = DtmFecha.Year
        End If
        Dim DatosMes As New StMesAño
        DatosMes.Mes = IntMes
        DatosMes.Año = IntAño
        Return DatosMes
    End Function

    <Serializable()> _
    Public Class DataObtenerDatosTipoAmort
        Public IDGrupo As String
        Public IDTipo As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDGrupo As String, Optional ByVal IDTipo As String = "")
            Me.IDGrupo = IDGrupo
            Me.IDTipo = IDTipo
        End Sub
    End Class

    <Task()> Public Shared Function ObtenerDatosTipoAmort(ByVal data As DataObtenerDatosTipoAmort, ByVal services As ServiceProvider) As StDatosAmort
        'Funcion que devuelve en un string los datos (Tipo, Vida y Porcentaje) del TipoAmortizacion
        'del Elemento .Se le puede pasar el Grupo o el Tipo
        If data.IDGrupo & String.Empty <> String.Empty Then
            Dim DtGAmortiz As DataTable = New GrupoAmortizacion().SelOnPrimaryKey(data.IDGrupo)
            If Not DtGAmortiz Is Nothing AndAlso DtGAmortiz.Rows.Count Then
                data.IDTipo = DtGAmortiz.Rows(0)("IdTipoAmortiz")
            End If
        End If
        If data.IDTipo & String.Empty <> String.Empty Then
            Dim DtTipoAmortiz As DataTable = New TipoAmortizacionCabecera().SelOnPrimaryKey(data.IDTipo)
            Dim IntVida As Integer = DtTipoAmortiz.Rows(0)("VidaUtil")
            Dim DtTipoAmortizLinea As DataTable = New TipoAmortizacionLinea().Filter(New FilterItem("IDTipoAmortizacion", FilterOperator.Equal, data.IDTipo), "NAño", "TOP 1 PorcentajeAmortizar")
            Dim IntPorc As Integer = 0
            If DtTipoAmortizLinea.Rows.Count > 0 Then IntPorc = DtTipoAmortizLinea.Rows(0)("PorcentajeAmortizar")
            Dim DatosAmort As New StDatosAmort
            DatosAmort.CodAmort = data.IDTipo
            DatosAmort.Vida = IntVida
            DatosAmort.Porcentaje = IntPorc
            Return DatosAmort
        Else : Return Nothing
        End If
    End Function

#End Region

#Region "Funciones Calculos"

    <Serializable()> _
    Public Class DataCalcPrimeraCuota
        Public CodigoAmort As String
        Public ValorTotal As Double
        Public ValorResidual As Double

        Public Sub New()
        End Sub

        Public Sub New(ByVal CodigoAmort As String, ByVal ValorTotal As Double, ByVal ValorResidual As Double)
            Me.CodigoAmort = CodigoAmort
            Me.ValorTotal = ValorTotal
            Me.ValorResidual = ValorResidual
        End Sub
    End Class

    <Task()> Public Shared Function CalcularPrimeraCuota(ByVal data As DataCalcPrimeraCuota, ByVal services As ServiceProvider) As DatosCuotaAmortizacion
        Dim D As New DatosCuotaAmortizacion
        If data.CodigoAmort = String.Empty Then
            D.DotacionElementoA = 0
            D.DotacionElementoB = 0
            D.VidaUtil = 0
            D.PorcentajeAnual = 0
        Else
            Dim Dt As DataTable = New TipoAmortizacionCabecera().SelOnPrimaryKey(data.CodigoAmort)
            If Dt Is Nothing OrElse Dt.Rows.Count = 0 Then
                D.DotacionElementoA = 0
                D.DotacionElementoB = 0
                D.VidaUtil = 0
                D.PorcentajeAnual = 0
            Else
                D.VidaUtil = Dt.Rows(0)("VidaUtil")
                Dim ClsTipoLinea As New TipoAmortizacionLinea
                Dim DtLineas As DataTable = ClsTipoLinea.Filter(New StringFilterItem("IDTipoAmortizacion", _
                    Dt.Rows(0)("IDTipoAmortizacion")), "NAño")
                If Not DtLineas Is Nothing AndAlso DtLineas.Rows.Count > 0 Then
                    D.PorcentajeAnual = DtLineas.Rows(0)("PorcentajeAmortizar")
                    'MIRAR POR DÍAS...
                    If data.ValorTotal <> 0 Then
                        Dim totalAmortizar As Double
                        totalAmortizar = data.ValorTotal
                        totalAmortizar -= data.ValorResidual
                        D.DotacionElementoA = (totalAmortizar * (D.PorcentajeAnual / 100)) / 12
                    Else : D.DotacionElementoA = 0
                    End If
                Else
                    D.PorcentajeAnual = 0
                    D.DotacionElementoA = 0
                End If
            End If
        End If
        Return D
    End Function

    <Serializable()> _
    Public Class DataCalcValorNetoContable
        Public ValorTotal As Object
        Public ValorAmortizado As Object
        Public ValorResidual As Object

        Public Sub New()
        End Sub

        Public Sub New(ByVal ValorTotal As Object, ByVal ValorAmortizado As Object, ByVal ValorResidual As Object)
            Me.ValorTotal = ValorTotal
            Me.ValorAmortizado = ValorAmortizado
            Me.ValorResidual = ValorResidual
        End Sub
    End Class

    <Task()> Public Shared Function CalcularValorNetoContable(ByVal data As DataCalcValorNetoContable, ByVal services As ServiceProvider) As Double
        Dim Resultado As Double
        If data.ValorTotal Is System.DBNull.Value Then
            Resultado = 0
        Else : Resultado = CDbl(data.ValorTotal)
        End If
        If Not data.ValorAmortizado Is System.DBNull.Value Then
            Resultado -= CDbl(data.ValorAmortizado)
        End If
        If Not data.ValorResidual Is System.DBNull.Value Then
            Resultado -= CDbl(data.ValorResidual)
        End If
        Return Resultado
    End Function

    <Serializable()> _
    Public Class DataCalcValorResidualMoneda
        Public Columna As String
        Public Valor As Double

        Public Sub New()
        End Sub

        Public Sub New(ByVal Columna As String, ByVal Valor As Double)
            Me.Columna = Columna
            Me.Valor = Valor
        End Sub
    End Class

    <Task()> Public Shared Function CalcularValorResidualMoneda(ByVal data As DataCalcValorResidualMoneda, ByVal services As ServiceProvider) As Double
        Dim MDblCambio As Double = 0
        If data.Columna = "ValorResidualA" Then
            Dim DtMoneda As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf Business.General.Moneda.ObtenerMonedaA, Nothing, services)
            Dim StDatos As New Moneda.DatosObtenerMoneda
            StDatos.IDMoneda = DtMoneda.Rows(0)("IDMoneda")
            Dim Moneda As MonedaInfo = ProcessServer.ExecuteTask(Of Moneda.DatosObtenerMoneda, MonedaInfo)(AddressOf Business.General.Moneda.ObtenerMoneda, StDatos, services)
            MDblCambio = Moneda.CambioB
        ElseIf data.Columna = "ValorResidualB" Then
            Dim DtMoneda As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf Business.General.Moneda.ObtenerMonedaB, Nothing, services)
            Dim StDatos As New Moneda.DatosObtenerMoneda
            StDatos.IDMoneda = DtMoneda.Rows(0)("IDMoneda")
            Dim Moneda As MonedaInfo = ProcessServer.ExecuteTask(Of Moneda.DatosObtenerMoneda, MonedaInfo)(AddressOf Business.General.Moneda.ObtenerMoneda, StDatos, services)
            MDblCambio = Moneda.CambioA
        End If
        Return data.Valor * MDblCambio
    End Function

    <Serializable()> _
    Public Class DataCalcFechaInicioContab
        Public FaltaElemento As Date
        Public FechaInicioCont As Date
        Public CalculoPorDias As Boolean

        Public Sub New()
        End Sub

        Public Sub New(ByVal FaltaElemento As Date, ByVal FechaInicioCont As Date, ByVal CalculoPorDias As Boolean)
            Me.FaltaElemento = FaltaElemento
            Me.FechaInicioCont = FechaInicioCont
            Me.CalculoPorDias = CalculoPorDias
        End Sub
    End Class

    <Task()> Public Shared Function CalcularFechaIniContab(ByVal data As DataCalcFechaInicioContab, ByVal services As ServiceProvider) As Date
        Dim DteAux As Date
        If data.CalculoPorDias Then
            If DateDiff(DateInterval.Day, data.FaltaElemento, data.FechaInicioCont) >= 0 Then
                DteAux = data.FechaInicioCont
            Else
                'Este caso no deberia darse nunca, pero para no parar el proceso
                'se tomara por defecto la fecha de alta del elemento
                DteAux = data.FaltaElemento
            End If
        Else
            If DateDiff(DateInterval.Month, data.FaltaElemento, data.FechaInicioCont) = 0 Then
                If DateDiff(DateInterval.Day, data.FaltaElemento, data.FechaInicioCont) < 0 Then
                    DteAux = data.FaltaElemento
                ElseIf DateDiff(DateInterval.Day, data.FaltaElemento, data.FechaInicioCont) >= 0 Then
                    DteAux = data.FechaInicioCont
                End If
                If DteAux.Day > 1 Then
                    DteAux = DteAux.AddMonths(1)
                    DteAux = New Date(DteAux.Year, DteAux.Month, 1)
                End If
            ElseIf DateDiff(DateInterval.Month, data.FaltaElemento, data.FechaInicioCont) > 0 Then
                If data.FechaInicioCont.Day = 1 Then
                    DteAux = data.FechaInicioCont
                Else
                    DteAux = data.FechaInicioCont.AddMonths(1)
                    DteAux = New Date(DteAux.Year, DteAux.Month, 1)
                End If
            Else
                'Este caso no deberia darse nunca, pero para no parar el proceso
                'se tomara en cuenta la fecha de alta del elemento
                If data.FaltaElemento.Day = 1 Then
                    DteAux = data.FaltaElemento
                Else
                    DteAux = data.FaltaElemento.AddMonths(1)
                    DteAux = New Date(DteAux.Year, DteAux.Month, 1)
                End If
            End If
        End If
        Return DteAux
    End Function

#End Region

#Region "Funciones de Fechas"

    Public Shared Function EsFechaMayor(ByVal PMes1 As Integer, _
                                     ByVal PAño1 As Integer, _
                                     ByVal PMes2 As Integer, _
                                     ByVal PAño2 As Integer) As Boolean
        EsFechaMayor = False
        If PAño1 > PAño2 Then
            Return True
        ElseIf PAño1 = PAño2 Then
            If PMes1 > PMes2 Then
                Return True
            End If
        End If
    End Function

    Public Shared Function EsFechaMayorIgual(ByVal PMes1 As Integer, _
                                       ByVal PAño1 As Integer, _
                                       ByVal PMes2 As Integer, _
                                       ByVal PAño2 As Integer) As Boolean
        EsFechaMayorIgual = False
        If PAño1 > PAño2 Then
            Return True
        ElseIf PAño1 = PAño2 Then
            If PMes1 >= PMes2 Then
                Return True
            End If
        End If
    End Function

    Public Shared Function FechaFinMes(ByVal DteFecha As Date, _
                                    Optional ByVal DteFecha2 As Date = cnMinDate) As Date
        'El ultimo dia del mes de dteFecha
        If CLng(DteFecha2.ToOADate) <> 0 Then
            If DteFecha2 > DteFecha Then
                DteFecha = DteFecha2
            End If
        End If
        If DteFecha.AddDays(1).Day = 1 Then
            FechaFinMes = DteFecha
        Else
            DteFecha = DteFecha.AddMonths(1)
            Return New Date(DteFecha.Year, DteFecha.Month, 1).AddDays(-1)
        End If
    End Function

    Public Shared Function FechaPrimeroMesSig(ByVal DteFecha As Date, _
                                        Optional ByVal DteFecha2 As Date = cnMinDate, _
                                        Optional ByVal BlnPorDias As Boolean = False) As Date
        'El primer dia del mes siguiente de dteFecha
        If Not BlnPorDias Then
            If CLng(DteFecha2.ToOADate) <> 0 Then
                If DteFecha2 > DteFecha Then
                    DteFecha = DteFecha2
                End If
            End If
            If DteFecha.Day = 1 Then
                Return DteFecha
            Else
                DteFecha = DteFecha.AddMonths(1)
                Return New Date(DteFecha.Year, DteFecha.Month, 1)
            End If
        End If
    End Function

#End Region

End Class