Public Class FamiliaCaracteristica

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroFamiliaCaracteristica"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDTipo")) = 0 Then ApplicationService.GenerateError("El código de Tipo es obligatorio")
        If Length(data("IDFamilia")) = 0 Then ApplicationService.GenerateError("La Familia es un dato obligatorio.")
        If Length(data("IDCaracteristica")) = 0 Then ApplicationService.GenerateError("El código de la Característica es obligatorio")
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = New FamiliaCaracteristica().SelOnPrimaryKey(data("IDTipo"), data("IDFamilia"), data("IDCaracteristica"))
            If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("Esta Característica ya está creada para el Tipo '|' y Familia '|'.", data("IDTipo"), data("IDFamilia"))
            End If
        End If
    End Sub

#End Region

End Class