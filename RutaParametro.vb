Public Class RutaParametro

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbRutaParametro"

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarPrimaryKey)
    End Sub

    <Task()> Public Shared Sub AsignarPrimaryKey(ByVal data As DataRow, ByVal services As ServiceProvider)

        If data.RowState = DataRowState.Added Then
            data("ID") = AdminData.GetAutoNumeric
        End If
    End Sub


#End Region

#Region "Procesos Públicos"

    <Task()> Public Shared Sub EliminarRutaParametros(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim ClsDel As New RutaParametro
        Dim DtDel As DataTable = ClsDel.Filter(New FilterItem("IDRutaOp", FilterOperator.Equal, data("IDRutaOp")))
        ClsDel.Delete(DtDel)
    End Sub

    <Task()> Public Shared Sub ActualizarRutaParametros(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim f As New Filter
        f.Add(New StringFilterItem("IDOperacion", data("IDOperacion")))
        f.Add("CopiarParametros", FilterOperator.Equal, True)

        Dim dtO As DataTable = New Operacion().Filter(f)
        If (Not (dtO Is Nothing)) AndAlso dtO.Rows.Count > 0 Then
            f.Clear()
            f.Add(New StringFilterItem("IDOperacion", data("IDOperacion")))
            Dim dtOD As DataTable = New OperacionDetalle().Filter(f)
            If Not dtOD Is Nothing AndAlso dtOD.Rows.Count > 0 Then
                Dim dtRP As DataTable = New RutaParametro().AddNew()
                For Each drOD As DataRow In dtOD.Rows
                    Dim drRP As DataRow = dtRP.NewRow
                    drRP("ID") = AdminData.GetAutoNumeric
                    drRP("IDRutaOp") = data("IDRutaOp")
                    drRP("IDParametro") = drOD("IDParametro")
                    drRP("DescParametro") = drOD("DescParametro")
                    drRP("Valor") = drOD("Valor")
                    drRP("Secuencia") = drOD("Secuencia")
                    drRP("GrupoParametro") = drOD("GrupoParametro")
                    dtRP.Rows.Add(drRP)
                Next
                BusinessHelper.UpdateTable(dtRP)
            End If
        End If
    End Sub

#End Region

End Class