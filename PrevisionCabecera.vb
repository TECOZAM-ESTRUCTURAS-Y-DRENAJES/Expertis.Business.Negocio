Public Class PrevisionCabecera

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbPrevisionCabecera"

#End Region

#Region "Eventos RegisterAddNewTasks"

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim StDatos As New Contador.DatosDefaultCounterValue(data, GetType(PrevisionCabecera).Name, "IDPrevision")
        ProcessServer.ExecuteTask(Of Contador.DatosDefaultCounterValue)(AddressOf Contador.LoadDefaultCounterValue, StDatos, services)
    End Sub

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("DescPrevision")) = 0 Then ApplicationService.GenerateError("Introduzca la Descripción de la previsión.")
        If Length(data("TipoPrevision")) = 0 Then ApplicationService.GenerateError("Introduzca el Tipo de Previsión.")
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClave)
    End Sub

    <Task()> Public Shared Sub AsignarClave(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDPrevision")) > 0 Then
                If Length(data("IdContador")) > 0 Then
                    data("IDPrevision") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, data("IDContador"), services)
                End If
                Dim dtTemp As DataTable = New PrevisionCabecera().SelOnPrimaryKey(data("IDPrevision"))
                If Not dtTemp Is Nothing AndAlso dtTemp.Rows.Count > 0 Then
                    ApplicationService.GenerateError("Ya existe una previsión con esa clave.")
                End If
            Else : ApplicationService.GenerateError("Introduzca el código de la previsión.")
            End If
        End If
    End Sub

#End Region

End Class