Public Class CaracteristicaValor

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbCaracteristicaValor"

#End Region

#Region "Eventos GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBrl As New BusinessRules
        oBrl.Add("IDValor", AddressOf CambioValor)
        Return oBrl
    End Function

    <Task()> Public Shared Sub CambioValor(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            If data.Context("TipoDato") = enumTipoDato.Numerico AndAlso Not IsNumeric(data.Value) Then
                ApplicationService.GenerateError("Tipo de Valor incorrecto. Debe ser numérico")
            End If
        End If
    End Sub

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim dtCaracteristica As DataTable = New Caracteristica().SelOnPrimaryKey(data("IDCaracteristica"))
        If Not IsNothing(dtCaracteristica) AndAlso dtCaracteristica.Rows.Count > 0 Then
            If dtCaracteristica.Rows(0)("TipoDato") = enumTipoDato.Numerico AndAlso Not IsNumeric(data("IDValor")) Then
                ApplicationService.GenerateError("El valor debe ser numérico.")
            End If
        End If
        If Length(data("IDCaracteristica")) = 0 Then ApplicationService.GenerateError("La Característica es un dato obligatorio.")
        If Length(data("IDValor")) = 0 Then ApplicationService.GenerateError("El Valor es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dtA As DataTable = New CaracteristicaValor().SelOnPrimaryKey(data("IDCaracteristica"), data("IDValor"))
            If Not IsNothing(dtA) AndAlso dtA.Rows.Count > 0 Then
                ApplicationService.GenerateError("El Valor ya existe para esta Característica.")
            End If
        End If
    End Sub

#End Region

End Class