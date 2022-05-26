Public Class EntidadContador

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbEntidadContador"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ComprobarPredeterminado)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data("Entidad").ToString.Trim.Length = 0 Then
            ApplicationService.GenerateError("La entidad es un campo obligatorio")
        Else
            Dim dt As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf ObtenerEntidades, Nothing, services)
            Dim drEnt() As DataRow
            drEnt = dt.Select("Entidad='" & data("Entidad").ToString & "'")
            If drEnt.GetLength(0) = 0 Then
                ApplicationService.GenerateError("La entidad introducida no existe")
            End If
        End If

        Dim dtTemp As DataTable
        If data.RowState = DataRowState.Modified Then
            dtTemp = New EntidadContador().SelOnPrimaryKey(data("Entidad"), data("IDContador"))
            If Not dtTemp Is Nothing AndAlso dtTemp.Rows.Count > 0 Then
                If data.IsNull("Predeterminado") OrElse data("Predeterminado") = False Then
                    If dtTemp.Rows(0)("Predeterminado") = False Then data.RejectChanges()
                Else
                    If dtTemp.Rows(0)("Predeterminado") = True Then data.RejectChanges()
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ComprobarPredeterminado(ByVal data As DataRow, ByVal services As ServiceProvider)
        If (Not data.IsNull("Predeterminado") AndAlso data("Predeterminado") = True) And _
                   (data.RowState = DataRowState.Added OrElse data("Predeterminado", DataRowVersion.Original) = False) Then
            Dim ClsEnt As New EntidadContador
            Dim f As New Filter(FilterUnionOperator.And)
            f.Add(New StringFilterItem("Entidad", data("Entidad")))
            f.Add(New BooleanFilterItem("Predeterminado", True))
            Dim dt As DataTable = ClsEnt.Filter(f)
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                dt.Rows(0)("Predeterminado") = False
                ClsEnt.Update(dt)
            End If
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Function ObtenerEntidades(ByVal data As Object, ByVal services As ServiceProvider) As DataTable
        Return New BE.DataEngine().Filter("xEntity", "*", "", "Entidad", , True)
    End Function

#End Region

End Class