Public Class PresupuestoCosteMaterial
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbPresupuestoCosteMaterial"

#Region " RegisterDeleteTask "

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
        If Length(data("IDPresupuesto")) = 0 Then ApplicationService.GenerateError("El Presupuesto es un dato obligatorio.")
        If Length(data("IDArticulo")) = 0 Then ApplicationService.GenerateError("El Artículo es un dato obligatorio.")
        If Length(data("IDComponente")) = 0 Then ApplicationService.GenerateError("El Componente es obligatorio")
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
        If data.RowState = DataRowState.Added OrElse data.RowState = DataRowState.Modified Then
            If Length(data("IDPresupMaterial")) = 0 Then data("IDPresupMaterial") = AdminData.GetAutoNumeric
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarImportesAyB(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim dataImportesAB As IPropertyAccessor = New DataRowPropertyAccessor(data)
        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf PresupuestoCosteEstandar.CalcularImportesAyB, dataImportesAB, services)
    End Sub

    <Task()> Public Shared Sub ActualizarPresupuestoCosteVarios(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim dataActualizarCostesVarios As New PresupuestoCosteVarios.dataActualizarCostesVarios(data("IDPresupuesto"), PresupuestoCosteVarios.dataActualizarCostesVarios.enumOrigenActualizacionCostesVarios.Materiales)
        ProcessServer.ExecuteTask(Of PresupuestoCosteVarios.dataActualizarCostesVarios)(AddressOf PresupuestoCosteVarios.ActualizarCostesVarios, dataActualizarCostesVarios, services)
    End Sub

#End Region

#Region " Business Rules "

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBRL As New BusinessRules
        oBRL.Add("IDComponente", AddressOf CalcularPrecio)
        oBRL.Add("Cantidad", AddressOf CalcularPrecio)
        oBRL.Add("Merma", AddressOf RecalcularCostes)
        oBRL.Add("PrecioStdA", AddressOf RecalcularCostes)
        oBRL.Add("PorcentajeMat", AddressOf RecalcularCostes)
        Return oBRL
    End Function

    <Task()> Public Shared Sub CalcularPrecio(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            data.Current(data.ColumnName) = data.Value
            If Length(data.Current("IDComponente")) > 0 AndAlso Length(data.Current("Cantidad")) > 0 Then
                Dim dataTarifa As New DataCalculoTarifaComercial(data.Current("IDComponente"), data.Current("Cantidad"), Date.Today)
                ProcessServer.ExecuteTask(Of DataCalculoTarifaComercial)(AddressOf ProcesoComercial.TarifaCosteArticulo, dataTarifa, services)
                If Not dataTarifa.DatosTarifa Is Nothing AndAlso dataTarifa.DatosTarifa.PrecioCosteA > 0 Then
                    data.Current("PrecioStdA") = dataTarifa.DatosTarifa.PrecioCosteA
                    data.Current = New PresupuestoCosteMaterial().ApplyBusinessRule("PrecioStdA", data.Current("PrecioStdA"), data.Current, data.Context)
                End If
                If Length(data.Current("PrecioStdA") > 0) Then
                    Dim CosteStdA As Double = data.Current("Cantidad") * (1 + (data.Current("Merma") / 100)) * data.Current("PrecioStdA")
                    CosteStdA = CosteStdA * (1 + (data.Current("PorcentajeMat") / 100))
                    data.Current("CosteStdA") = CosteStdA
                    ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf PresupuestoCosteEstandar.CalcularImportesAyB, data.Current, services)
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub RecalcularCostes(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If data.ColumnName = "Merma" Then
            If data.Value < 0 Or data.Value > 100 Then
                ApplicationService.GenerateError("La Merma ha de ser un valor comprendido entre 0 y 100.")
            End If
        ElseIf data.ColumnName = "PorcentajeMat" Then
            If data.Value < 0 Then ApplicationService.GenerateError("El Porcentaje no puede ser negativo.")
        End If
        data.Current(data.ColumnName) = data.Value

        Dim CosteStdA As Double = data.Current("Cantidad") * (1 + (data.Current("Merma") / 100)) * data.Current("PrecioStdA")
        CosteStdA = CosteStdA * (1 + (data.Current("PorcentajeMat") / 100))
        data.Current("CosteStdA") = CosteStdA
        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf PresupuestoCosteEstandar.CalcularImportesAyB, data.Current, services)
    End Sub

#End Region

End Class