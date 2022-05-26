Public Class FormaEnvioDetalle

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    
    Private Const cnEntidad As String = "tbMaestroFormaEnvioDetalle"

#End Region

#Region "Eventos Entidad"

    
    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
    End Sub

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClavePrimaria)
        updateProcess.AddTask(Of DataRow)(AddressOf TratarFormaEnvioDetallePredeterminado)
    End Sub

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Sub AsignarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then data("IDFormaEnvioDetalle") = Guid.NewGuid
    End Sub

    <Task()> Public Shared Sub TratarFormaEnvioDetallePredeterminado(ByVal dr As DataRow, ByVal services As ServiceProvider)
        Dim f As New Filter
        f.Add("IDFormaEnvio", FilterOperator.Equal, dr("IDFormaEnvio"))
        f.Add("Predeterminado", FilterOperator.Equal, True)
        Dim dtFED As DataTable = New FormaEnvioDetalle().Filter(f)
        If IsNothing(dtFED) OrElse dtFED.Rows.Count = 0 Then
            ' No hay más almacenes para el articulo actual con lo cual será el predeterminado.
            dr("Predeterminado") = True
        Else
            ' Si el almacen ha sido marcado como predeterminado
            If dr("Predeterminado") Then
                If dr("IDFormaEnvioDetalle") <> dtFED.Rows(0)("IDFormaEnvioDetalle") Then
                    dtFED.Rows(0)("Predeterminado") = False
                    BusinessHelper.UpdateTable(dtFED)
                End If
            ElseIf dr.RowState = DataRowState.Modified AndAlso dr("Predeterminado") <> dr("Predeterminado", DataRowVersion.Original) Then
                dr("Predeterminado") = True
            End If
        End If
    End Sub

#End Region

End Class