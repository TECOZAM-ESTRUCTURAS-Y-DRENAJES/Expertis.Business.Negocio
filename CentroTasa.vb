Public Class CentroTasa

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroCentroTasa"

#End Region

#Region " Validate "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarCentroObligatorio)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarTasaObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarFechaDesdeObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarFechaHastaObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarFechaDesdeHasta)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarCentroTasaExistente)
    End Sub

    <Task()> Public Shared Sub ValidarCentroTasaExistente(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dtCentroTasa As DataTable = New CentroTasa().SelOnPrimaryKey(data("IDCentro"), data("IDTasa"), data("FechaDesde"))
            If Not dtCentroTasa Is Nothing AndAlso dtCentroTasa.Rows.Count > 0 Then
                ApplicationService.GenerateError("Ya existe la Tasa {0} para la Fecha Desde {1} en el Centro {2}.", Quoted(data("IDTasa")), Quoted(Format(data("FechaDesde"), "dd/MM/yyyy")), Quoted(data("IDCentro")))
            End If
        End If
    End Sub

#End Region

#Region " Update "

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AplicarMonedaTasa)
        updateProcess.AddTask(Of DataRow)(AddressOf Comunes.UpdateEntityRow)
        updateProcess.AddTask(Of DataRow)(AddressOf Comunes.MarcarComoActualizado)
        updateProcess.AddTask(Of DataRow)(AddressOf ActualizarCabeceraCentro)
    End Sub

    <Task()> Public Shared Sub AplicarMonedaTasa(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim strMonPred As String = New Parametro().MonedaPred
        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
        Dim MonPred As MonedaInfo = Monedas.GetMoneda(strMonPred)

        Dim ValAyB As New ValoresAyB(New DataRowPropertyAccessor(data), MonPred.ID, MonPred.CambioA, MonPred.CambioB)
        ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)
    End Sub

    <Task()> Public Shared Sub ActualizarCabeceraCentro(ByVal data As datarow, ByVal services As ServiceProvider)
        Dim DtCentros As DataTable = New Centro().SelOnPrimaryKey(data("IDCentro"))
        ProcessServer.ExecuteTask(Of DataTable)(AddressOf Centro.CentroActualizarTasa, DtCentros, services)
    End Sub


#End Region

#Region "Eventos GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBrl As New BusinessRules
        oBrl.Add("IDTasa", AddressOf CambioTasa)
        Return oBrl
    End Function

    <Task()> Public Shared Sub CambioTasa(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) = 0 Then
            Dim ff As New Filter
            Dim clsTasa As New Tasa
            ff.Add("IDTasa", FilterOperator.Equal, data.Value)
            Dim dtTasa As DataTable = clsTasa.Filter(ff)
            If Not IsNothing(dtTasa) AndAlso dtTasa.Rows.Count > 0 Then
                data.Current("Fiscal") = dtTasa.Rows(0)("Fiscal")
            Else
                data.Current("Fiscal") = False
            End If
        End If
    End Sub

#End Region

End Class