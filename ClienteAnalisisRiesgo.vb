Public Class ClienteAnalisisRiesgo

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper
    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub
    Private Const cnEntidad As String = "tbClienteAnalisisRiesgo"

#End Region

#Region "Clases"

    <Serializable()> _
    Public Class DataCambioRegistro
        Public DtRegistro As DataTable
        Public DtRiesgo As DataTable
    End Class

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarPrimaryKey)
    End Sub

    <Task()> Public Shared Sub AsignarPrimaryKey(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If IsDBNull(data("IDClienteAnalisisRiesgo")) Then
                data("IDClienteAnalisisRiesgo") = AdminData.GetAutoNumeric
            End If
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Function ObtenerFechaUltimaActualizacionProceso(ByVal data As Object, ByVal services As ServiceProvider) As DateTime
        Dim dt As DataTable = AdminData.Filter("vCIClientesRiesgoAnalisisFechaUltimoLanzamiento")
        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            If Length(dt.Rows(0)("FechaUltimaActualizacionProceso")) > 0 Then
                Return DirectCast(dt.Rows(0)("FechaUltimaActualizacionProceso"), System.DateTime)
            Else
                Return New DateTime(0)
            End If
        Else
            Return New DateTime(0)
        End If
    End Function

    <Task()> Public Shared Sub CambioRegistros(ByVal data As DataCambioRegistro, ByVal services As ServiceProvider)
        If Not data.DtRiesgo Is Nothing AndAlso data.DtRiesgo.Rows.Count > 0 Then
            ' Borrar todos los registros que se van a autocalcular
            Dim ClsClie As New ClienteAnalisisRiesgo
            If Not data.DtRegistro Is Nothing AndAlso data.DtRegistro.Rows.Count > 0 Then
                Dim Filtro As New Filter
                Filtro.UnionOperator = FilterUnionOperator.Or
                For Each dr As DataRow In data.DtRegistro.Rows
                    Filtro.Add("IDCliente", FilterOperator.Equal, dr("IDCliente"))
                Next
                Dim dtCLAnRiesgo As DataTable = ClsClie.Filter(Filtro)
                If Not dtCLAnRiesgo Is Nothing AndAlso dtCLAnRiesgo.Rows.Count > 0 Then ClsClie.Delete(dtCLAnRiesgo)
            End If
            ' Actualizar la tabla con los nuevos registros.
            ClsClie.Update(data.DtRiesgo)
        End If
    End Sub

#End Region

End Class