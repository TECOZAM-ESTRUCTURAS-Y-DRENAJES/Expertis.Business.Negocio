Public Class ComisionCantidad

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbComisionCantidad"

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        updateProcess.AddTask(Of DataRow)(AddressOf ValidarCantidadDuplicada)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IdComisionRepresentante")) = 0 Then
            ApplicationService.GenerateError("La comisión de representante para la que se establece este tramo es obligatoria.")
        Else
            Dim DtCom As DataTable = New ComisionRepresentante().SelOnPrimaryKey(data("IdComisionRepresentante"))
            If DtCom Is Nothing OrElse DtCom.Rows.Count = 0 Then ApplicationService.GenerateError("La comisión de representante para la que se establece este tramo no existe")
        End If
        If Length(data("QDesde")) = 0 Then ApplicationService.GenerateError("El campo Cantidad ha de ser numérico.")
        If Length(data("Comision")) = 0 Then
            ApplicationService.GenerateError("El campo Comisión ha de ser numérico.")
        Else
            If Not data("Porcentaje") Is System.DBNull.Value AndAlso data("Porcentaje") = True Then
                If CDbl(data("Comision")) < 0 Or CDbl(data("Comision")) > 100 Then
                    ApplicationService.GenerateError("El campo Comisión es un porcentaje, introduzca valores entre 0 y 100.")
                End If
            End If
        End If

        ' Comprobamos que no esté establecida ya esta cantidad.
        Dim f As New Filter(FilterUnionOperator.And)
        f.Add(New NumberFilterItem("QDesde", FilterOperator.Equal, data("QDesde")))
        f.Add(New StringFilterItem("IDComisionRepresentante", FilterOperator.Equal, data("IDComisionRepresentante")))
        If Not data.RowState = DataRowState.Added Then f.Add(New NumberFilterItem("IDComisionCantidad", FilterOperator.NotEqual, data("IDComisionCantidad")))
        Dim DtDatos As DataTable = New ComisionCantidad().Filter(f)
        If Not DtDatos Is Nothing AndAlso DtDatos.Rows.Count > 0 Then ApplicationService.GenerateError("Ya existe esta cantidad para la comisión actual")
    End Sub

    <Task()> Public Shared Sub ValidarCantidadDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then data("IDComisionCantidad") = AdminData.GetAutoNumeric
    End Sub

#End Region

End Class