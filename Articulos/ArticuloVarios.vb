Public Class ArticuloVarios

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbArticuloVarios"

#End Region

#Region "Eventos GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBrl As New BusinessRules
        oBrl.Add("Valor", AddressOf CambioValor)
        Return oBrl
    End Function

    <Task()> Public Shared Sub CambioValor(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Not IsNumeric(data.Value) Then ApplicationService.GenerateError("Campo no numérico.")
    End Sub

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDArticulo")) = 0 Then ApplicationService.GenerateError("El Artículo es obligatorio.")
        If Length(data("IDVarios")) = 0 Then ApplicationService.GenerateError("Debe introducir un código de Varios.")
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = New ArticuloVarios().SelOnPrimaryKey(data("IDArticulo"), data("IDVarios"))
            If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("El código de Varios '|' ya está asignado en este artículo.", data("IDVarios"))
            End If
        End If
    End Sub

#End Region

End Class