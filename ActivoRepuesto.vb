Public Class ActivoRepuesto

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbActivoRepuesto"

#End Region

#Region "Eventos GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBrl As New BusinessRules
        oBrl.Add("IDComponente", AddressOf ValidarRepuesto)
        oBrl.Add("Cantidad", AddressOf ValidarCantidad)
        Return oBrl
    End Function

    <Task()> Public Shared Sub ValidarRepuesto(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDComponente")) > 0 Then
            Dim objNegArticulo As New Articulo
            Dim dr As DataRow = objNegArticulo.GetItemRow(data.Current("IDComponente"))
            If Not IsNothing(dr) Then
                If data.Current.ContainsKey("DescComponente") Then data.Current("DescComponente") = dr("DescArticulo")
            Else
                If data.Current.ContainsKey("DescComponente") Then data.Current("DescComponente") = System.DBNull.Value
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarCantidad(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 AndAlso Not IsNumeric(data.Value) Then
            ApplicationService.GenerateError("Campo no numérico.")
        End If
    End Sub

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarRepuestoObligatorio)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClavePrimaria)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
    End Sub

    <Task()> Public Shared Sub ValidarRepuestoObligatorio(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDComponente")) = 0 Then ApplicationService.GenerateError("El Repuesto es obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim f As New Filter
        f.Add(New StringFilterItem("IDArticulo", data("IDComponente")))
        f.Add(New BooleanFilterItem("Compra", True))
        Dim dt As DataTable = New BE.DataEngine().Filter("advArticulo", f)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then ApplicationService.GenerateError("El Repuesto no existe.")
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim f As New Filter
            f.Add(New StringFilterItem("IDActivo", data("IDActivo")))
            f.Add(New StringFilterItem("IDComponente", data("IDComponente")))
            Dim dt As DataTable = New ActivoRepuesto().Filter(f)
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then ApplicationService.GenerateError("Repuesto duplicado para el activo actual")
        End If
    End Sub

#End Region

End Class