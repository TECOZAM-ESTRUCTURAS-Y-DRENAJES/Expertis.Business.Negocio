Public Class PresupuestoCosteOperacion
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbPresupuestoCosteOperacion"

#Region " RegisterDeleteTasks "

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf Business.General.Comunes.DeleteEntityRow)
        deleteProcess.AddTask(Of DataRow)(AddressOf Business.General.Comunes.MarcarComoEliminado)
        deleteProcess.AddTask(Of DataRow)(AddressOf ActualizarPresupuestoCosteVarios)
    End Sub

#End Region

#Region " RegisterValidateTasks "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidaDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidaDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDArticulo")) = 0 Then ApplicationService.GenerateError("El Artículo es un dato obligatorio.")
        If Length(data("IDOperacion")) = 0 Then ApplicationService.GenerateError("La Operación es un dato obligatorio.")
        If data("TipoOperacion") = enumtrTipoOperacion.trInterna Then
            If Length(data("IDCentro")) = 0 Then ApplicationService.GenerateError("El Centro es un dato obligatorio.")
            If data("FactorProduccion") <= 0 Then ApplicationService.GenerateError("El Factor de Producción ha de ser un valor superior a 0.")
            If data("LoteMinimo") <= 0 Then ApplicationService.GenerateError("El Lote Mínimo ha de ser un valor superior a 0.")
        End If
        If data("Secuencia") <= 0 Then ApplicationService.GenerateError("La Secuencia ha de ser un valor superior a 0.")
        If Length(data("Orden")) = 0 Then data("Orden") = 0
        If Length(data("Nivel")) = 0 Then data("Nivel") = 0
        If Length(data("UDTiempoPrep")) = 0 Then data("UDTiempoPrep") = 0
        If Length(data("TiempoPrep")) = 0 Then data("TiempoPrep") = 0
    End Sub

#End Region

#Region " RegisterUpdateTasks "

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarIdentificador)
        updateProcess.AddTask(Of DataRow)(AddressOf ActualizarImportesAyB)
        updateProcess.AddTask(Of DataRow)(AddressOf Business.General.Comunes.UpdateEntityRow)
        updateProcess.AddTask(Of DataRow)(AddressOf Business.General.Comunes.MarcarComoActualizado)
        updateProcess.AddTask(Of DataRow)(AddressOf ActualizarPresupuestoCosteVarios)
    End Sub

    <Task()> Public Shared Sub AsignarIdentificador(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDPresupOperacion")) = 0 Then data("IDPresupOperacion") = AdminData.GetAutoNumeric
    End Sub

    <Task()> Public Shared Sub ActualizarImportesAyB(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim dataImportesAB As IPropertyAccessor = New DataRowPropertyAccessor(data)
        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf PresupuestoCosteEstandar.CalcularImportesAyB, dataImportesAB, services)
    End Sub

    <Task()> Public Shared Sub ActualizarPresupuestoCosteVarios(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim dataActualizarCostesVarios As New PresupuestoCosteVarios.dataActualizarCostesVarios(data("IDPresupuesto"), PresupuestoCosteVarios.dataActualizarCostesVarios.enumOrigenActualizacionCostesVarios.Operaciones)
        ProcessServer.ExecuteTask(Of PresupuestoCosteVarios.dataActualizarCostesVarios)(AddressOf PresupuestoCosteVarios.ActualizarCostesVarios, dataActualizarCostesVarios, services)
    End Sub

#End Region

#Region " GetBusinessRules "

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBrl As New BusinessRules
        oBrl.Add("IDCentro", AddressOf CambioCentro)
        oBrl.Add("FactorHombre", AddressOf CalcularCosteOperacion)
        oBrl.Add("TiempoPrep", AddressOf CalcularCosteOperacion)
        oBrl.Add("UdTiempoPrep", AddressOf CalcularCosteOperacion)
        oBrl.Add("TiempoEjecUnit", AddressOf CalcularCosteOperacion)
        oBrl.Add("UdTiempoEjec", AddressOf CalcularCosteOperacion)
        oBrl.Add("FactorProduccion", AddressOf CalcularCosteOperacion)
        oBrl.Add("TasaEjecucionA", AddressOf CalcularCosteOperacion)
        oBrl.Add("TasaPreparacionA", AddressOf CalcularCosteOperacion)
        oBrl.Add("TasaMODA", AddressOf CalcularCosteOperacion)
        oBrl.Add("LoteMinimo", AddressOf CalcularCosteOperacion)
        Return oBrl
    End Function

    <Task()> Public Shared Sub CambioCentro(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            data.Current(data.ColumnName) = data.Value
            Dim dtCentro As DataTable = New Centro().SelOnPrimaryKey(data.Value)
            If Not dtCentro Is Nothing AndAlso dtCentro.Rows.Count > 0 Then
                data.Current("TasaPreparacionA") = dtCentro.Rows(0)("TasaPreparacionA")
                data.Current("TasaEjecucionA") = dtCentro.Rows(0)("TasaEjecucionA")
                data.Current("TasaMODA") = dtCentro.Rows(0)("TasaManoObraA")
                data.Current("FactorHombre") = dtCentro.Rows(0)("FactorHombre")

                data.Current = New PresupuestoCosteOperacion().ApplyBusinessRule("TasaPreparacionA", data.Current("TasaPreparacionA"), data.Current, data.Context)
            End If
        Else
            data.Current("TasaPreparacionA") = 0
            data.Current("TasaEjecucionA") = 0
            data.Current("TasaMODA") = 0
            data.Current("FactorHombre") = 0
            data.Current("CosteOperacionA") = 0
        End If
    End Sub

    <Task()> Public Shared Sub CalcularCosteOperacion(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            data.Current(data.ColumnName) = data.Value

            Dim ImportePrep As Double = 0
            Dim ImporteEjec As Double = 0
            Dim ImporteMOD As Double = 0
            If data.Current("FactorProduccion") > 0 Then
                '///Preparacion
                Dim Tiempo As Double
                If data.Current("LoteMinimo") > 0 Then
                    Dim infoTiempoOperacion As New dataTiempoOperacion(Nz(data.Current("TiempoPrep"), 0), Nz(data.Current("UDTiempoPrep"), 0), enumstdUdTiempo.Horas)
                    Tiempo = ProcessServer.ExecuteTask(Of dataTiempoOperacion, Double)(AddressOf TiempoOperacion, infoTiempoOperacion, services)
                    ImportePrep = ((Tiempo * Nz(data.Current("TasaPreparacionA"), 0)) / data.Current("LoteMinimo")) / data.Current("FactorProduccion")
                End If

                '///Ejecucion
                Dim infoTiempoEjecucion As New dataTiempoOperacion(Nz(data.Current("TiempoEjecUnit"), 0), Nz(data.Current("UDTiempoEjec"), 0), enumstdUdTiempo.Horas)
                Tiempo = ProcessServer.ExecuteTask(Of dataTiempoOperacion, Double)(AddressOf TiempoOperacion, infoTiempoEjecucion, services)
                ImporteEjec = (Tiempo * data.Current("TasaEjecucionA")) / data.Current("FactorProduccion")

                '///MOD
                ImporteMOD = (data.Current("FactorHombre") * data.Current("TasaMODA") * Tiempo) / data.Current("FactorProduccion")
            End If

            If Len(data.Current("CantidadAcumulada")) = 0 OrElse data.Current("CantidadAcumulada") = 0 Then
                data.Current("CantidadAcumulada") = 1
            End If
            data.Current("CosteOperacionA") = (ImportePrep + ImporteEjec + ImporteMOD) * Nz(data.Current("CantidadAcumulada"), 1)
            ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf PresupuestoCosteEstandar.CalcularImportesAyB, data.Current, services)
        Else
            Select Case data.ColumnName
                Case "FactorProduccion"
                    If data.Value <= 0 Then ApplicationService.GenerateError("El Factor de Producción ha de ser un valor superior a 0.")
                Case "LoteMinimo"
                    If data.Value <= 0 Then ApplicationService.GenerateError("El Lote Mínimo ha de ser un valor superior a 0.")
                Case "UDTiempoPrep"
                    If data.Value <= 0 Then ApplicationService.GenerateError("La Unidad del Tiempo Preparación es un dato necesario para el cálculo del Importe.")
                Case "UDTiempoEjec"
                    If data.Value <= 0 Then ApplicationService.GenerateError("La Unidad del Tiempo Ejecución es un dato necesario para el cálculo del Importe.")
            End Select
        End If
    End Sub

#Region " TiempoOperacion "

    <Serializable()> _
    Public Class dataTiempoOperacion
        Public Tiempo As Double
        Public UDTiempoOld As enumstdUdTiempo
        Public UDTiempoNew As enumstdUdTiempo
        Public Sub New(ByVal Tiempo As Double, ByVal UDTiempoOld As enumstdUdTiempo, ByVal UDTiempoNew As enumstdUdTiempo)
            Me.Tiempo = Tiempo
            Me.UDTiempoOld = UDTiempoOld
            Me.UDTiempoNew = UDTiempoNew
        End Sub
    End Class
    <Task()> Public Shared Function TiempoOperacion(ByVal data As dataTiempoOperacion, ByVal services As ServiceProvider) As Double
        Return (xRound(data.Tiempo * ProcessServer.ExecuteTask(Of dataTiempoOperacion, Double)(AddressOf FactorTiempo, data, services), 2))
    End Function

    <Task()> Public Shared Function FactorTiempo(ByVal data As dataTiempoOperacion, ByVal services As ServiceProvider) As Double
        Dim Factor As Double
        'Pasar de la unidad de tiempo 1 a la unidad de tiempo 2

        Select Case data.UDTiempoOld
            Case enumstdUdTiempo.Dias
                Factor = 24
            Case enumstdUdTiempo.Horas
                Factor = 1
            Case enumstdUdTiempo.Minutos
                Factor = 1 / 60
            Case enumstdUdTiempo.Segundos
                Factor = 1 / 3600
        End Select

        Select Case data.UDTiempoNew
            Case enumstdUdTiempo.Dias
                Factor = Factor / 24
            Case enumstdUdTiempo.Horas
                Factor = Factor / 1
            Case enumstdUdTiempo.Minutos
                Factor = Factor / (1 / 60)
            Case enumstdUdTiempo.Segundos
                Factor = Factor / (1 / 3600)
        End Select
        Return Factor
    End Function

#End Region

#End Region

End Class