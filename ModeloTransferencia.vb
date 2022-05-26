Public Class ModeloTransferencia

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbModeloTransferencia"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("Descripcion")) = 0 Then ApplicationService.GenerateError("Introduzca la descripción del modelo de transferencia.")
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("Tipo")) > 0 Then
                Dim dtTemp As DataTable = New ModeloTransferencia().SelOnPrimaryKey(data("Tipo"))
                If Not dtTemp Is Nothing AndAlso dtTemp.Rows.Count > 0 Then ApplicationService.GenerateError("Ya existe un modelo de transferencia con esa clave.")
            Else : ApplicationService.GenerateError("Introduzca el código del modelo de transferencia.")
            End If
        End If
    End Sub

#End Region

End Class