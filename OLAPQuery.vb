Public Class OLAPQuery

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Private Const cnEntidad As String = "tbOLAPQuery"

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarIdentificador)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarOrden)
    End Sub

    <Task()> Public Shared Sub AsignarIdentificador(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDQuery")) > 0 Then data("IDQuery") = Guid.NewGuid
        End If
    End Sub

    <Task()> Public Shared Sub AsignarOrden(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then data("Orden") = ProcessServer.ExecuteTask(Of Object, Integer)(AddressOf OLAPQuery.GetMaxOrden, Nothing, services)
    End Sub

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        data("IdQuery") = Guid.NewGuid
    End Sub

    <Task()> Public Shared Function GetMaxOrden(ByVal data As Object, ByVal services As ServiceProvider) As Integer
        Dim Dt As DataTable = New OLAPQuery().Filter("MAX(Orden) as OrdenMaximo", , )
        If Dt.Rows.Count > 0 Then
            Return Nz(Dt.Rows(0)("OrdenMaximo"), 0)
        End If
    End Function

#End Region

End Class