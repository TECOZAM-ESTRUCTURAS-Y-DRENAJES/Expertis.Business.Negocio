Public Class PresupuestoCosteVarios
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbPresupuestoCosteVarios"

#Region " RegisterValidateTasks "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidaDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidaDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDPresupuesto")) = 0 Then ApplicationService.GenerateError("El Presupuesto es un dato obligatorio.")
        If Length(data("IDArticulo")) = 0 Then ApplicationService.GenerateError("El Artículo es un dato obligatorio.")
        If Length(data("IDVarios")) = 0 Then ApplicationService.GenerateError("Varios es un dato obligatorio.")
        If Length(data("Orden")) = 0 Then data("Orden") = 0
        If Length(data("Nivel")) = 0 Then data("Nivel") = 0
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarIdentificador)
        updateProcess.AddTask(Of DataRow)(AddressOf ActualizarImportesAyB)
    End Sub

    <Task()> Public Shared Sub AsignarIdentificador(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDPresupVarios")) = 0 Then data("IDPresupVarios") = AdminData.GetAutoNumeric
    End Sub

    <Task()> Public Shared Sub ActualizarImportesAyB(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim dataImportesAB As IPropertyAccessor = New DataRowPropertyAccessor(data)
        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf PresupuestoCosteEstandar.CalcularImportesAyB, dataImportesAB, services)
    End Sub

#End Region

#Region " GetBusinessRules "

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBrl As New BusinessRules
        oBrl.Add("Valor", AddressOf CambioTipoValor)
        oBrl.Add("Tipo", AddressOf CambioTipoValor)
        Return oBrl
    End Function

    <Task()> Public Shared Sub CambioTipoValor(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) Then
            If data.ColumnName = "Valor" AndAlso data.Current("Valor") < 0 Then
                ApplicationService.GenerateError("El Valor ha de ser un valor positivo.")
            End If
            data.Current(data.ColumnName) = data.Value
            Dim dblCosteMat As Double = 0
            Dim dblCosteOpe As Double = 0
            Dim dblCosteExt As Double = 0
            Dim dblPVP As Double = 0

            If data.Context.ContainsKey("CosteMatStdA") Then dblCosteMat = data.Context("CosteMatStdA")
            If data.Context.ContainsKey("CosteOpeStdA") Then dblCosteOpe = data.Context("CosteOpeStdA")
            If data.Context.ContainsKey("CosteExtStdA") Then dblCosteExt = data.Context("CosteExtStdA")
            If data.Context.ContainsKey("PVPA") Then dblPVP = data.Context("PVPA")

            Select Case data.Current("Tipo")
                Case enumCosteVarios.cvValor
                    data.Current("CosteVariosA") = data.Current("Valor")
                Case enumCosteVarios.cvPorMaterial
                    data.Current("CosteVariosA") = dblCosteMat * (data.Current("Valor") / 100)
                Case enumCosteVarios.cvPorInterno
                    data.Current("CosteVariosA") = dblCosteOpe * (data.Current("Valor") / 100)
                Case enumCosteVarios.cvPorExterno
                    data.Current("CosteVariosA") = dblCosteExt * (data.Current("Valor") / 100)
                Case enumCosteVarios.cvPorTotal
                    data.Current("CosteVariosA") = (dblCosteMat + dblCosteOpe + dblCosteExt) * (data.Current("Valor") / 100)
                Case enumCosteVarios.cvPorPVP
                    data.Current("CosteVariosA") = dblPVP * (data.Current("Valor") / 100)
            End Select
        End If
    End Sub

#End Region

#Region " ActualizarCostesVarios "

    <Serializable()> _
    Public Class dataActualizarCostesVarios
        Public IDPresupuesto As Integer
        Public Origen As enumOrigenActualizacionCostesVarios
        Public Enum enumOrigenActualizacionCostesVarios
            Operaciones
            Materiales
        End Enum

        Public Sub New(ByVal IDPresupuesto As Integer, ByVal Origen As enumOrigenActualizacionCostesVarios)
            Me.IDPresupuesto = IDPresupuesto
            Me.Origen = Origen
        End Sub
    End Class
    <Task()> Public Shared Sub ActualizarCostesVarios(ByVal data As dataActualizarCostesVarios, ByVal services As ServiceProvider)
        Select Case data.Origen
            Case dataActualizarCostesVarios.enumOrigenActualizacionCostesVarios.Operaciones
                Dim dataOPInterna As New dataTratarOperaciones(data.IDPresupuesto, enumtrTipoOperacion.trInterna)
                ProcessServer.ExecuteTask(Of dataTratarOperaciones)(AddressOf TratarOperaciones, dataOPInterna, services)

                Dim dataOPExterna As New dataTratarOperaciones(data.IDPresupuesto, enumtrTipoOperacion.trExterna)
                ProcessServer.ExecuteTask(Of dataTratarOperaciones)(AddressOf TratarOperaciones, dataOPExterna, services)
            Case dataActualizarCostesVarios.enumOrigenActualizacionCostesVarios.Materiales
                Dim dataOPMaterial As New dataTratarMateriales(data.IDPresupuesto)
                ProcessServer.ExecuteTask(Of dataTratarMateriales)(AddressOf TratarMateriales, dataOPMaterial, services)
        End Select
    End Sub

    <Serializable()> _
    Public Class dataTratarOperaciones
        Public IDPresupuesto As Integer
        Public TipoOperacion As enumtrTipoOperacion
        Public Tipo As enumCosteVarios
        Public Sub New(ByVal IDPresupuesto As Integer, ByVal TipoOperacion As enumtrTipoOperacion)
            Me.IDPresupuesto = IDPresupuesto
            Me.TipoOperacion = TipoOperacion
            If Me.TipoOperacion = enumtrTipoOperacion.trInterna Then
                Tipo = enumCosteVarios.cvPorInterno
            Else
                Tipo = enumCosteVarios.cvPorExterno
            End If
        End Sub
    End Class
    <Task()> Public Shared Sub TratarOperaciones(ByVal data As dataTratarOperaciones, ByVal services As ServiceProvider)
        Dim f As New Filter
        f.Add(New NumberFilterItem("IDPresupuesto", data.IDPresupuesto))
        f.Add(New NumberFilterItem("Tipo", data.Tipo))
        Dim CV As New PresupuestoCosteVarios
        Dim dtVarios As DataTable = CV.Filter(f)
        If Not dtVarios Is Nothing AndAlso dtVarios.Rows.Count > 0 Then
            f.Clear()
            f.Add(New NumberFilterItem("IDPresupuesto", data.IDPresupuesto))
            f.Add(New NumberFilterItem("TipoOperacion", data.TipoOperacion))
            Dim dtCosteOpe As DataTable = New PresupuestoCosteOperacion().Filter(f)
            If Not dtCosteOpe Is Nothing AndAlso dtCosteOpe.Rows.Count > 0 Then
                Dim TotalCosteOperacionA As Double = dtCosteOpe.Compute("SUM(CosteOperacionA)", Nothing)
                For Each drVarios As DataRow In dtVarios.Select
                    drVarios("CosteVariosA") = drVarios("Valor") * TotalCosteOperacionA / 100
                Next
            Else
                For Each drVarios As DataRow In dtVarios.Select
                    drVarios("CosteVariosA") = 0
                Next
            End If
            CV.Update(dtVarios)
        End If
    End Sub

    <Serializable()> _
    Public Class dataTratarMateriales
        Public IDPresupuesto As Integer
        Public Tipo As enumCosteVarios
        Friend Sub New(ByVal IDPresupuesto As Integer)
            Me.IDPresupuesto = IDPresupuesto
            Me.Tipo = enumCosteVarios.cvPorMaterial
        End Sub
    End Class
    <Task()> Public Shared Sub TratarMateriales(ByVal data As dataTratarMateriales, ByVal services As ServiceProvider)
        Dim f As New Filter
        f.Add(New NumberFilterItem("IDPresupuesto", data.IDPresupuesto))
        f.Add(New NumberFilterItem("Tipo", data.Tipo))
        Dim CV As New PresupuestoCosteVarios
        Dim dtVarios As DataTable = CV.Filter(f)
        If Not dtVarios Is Nothing AndAlso dtVarios.Rows.Count > 0 Then
            Dim dtCosteMat As DataTable = New PresupuestoCosteMaterial().Filter(New NumberFilterItem("IDPresupuesto", data.IDPresupuesto))
            If Not dtCosteMat Is Nothing AndAlso dtCosteMat.Rows.Count > 0 Then
                Dim TotalCosteStdA As Double = dtCosteMat.Compute("SUM(CosteStdA)", Nothing)
                For Each drVarios As DataRow In dtVarios.Select
                    drVarios("CosteVariosA") = drVarios("Valor") * TotalCosteStdA / 100
                Next
            Else
                For Each drVarios As DataRow In dtVarios.Select
                    drVarios("CosteVariosA") = 0
                Next
            End If
            CV.Update(dtVarios)
        End If
    End Sub

#End Region

End Class