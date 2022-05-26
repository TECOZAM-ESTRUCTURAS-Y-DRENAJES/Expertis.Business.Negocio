Public Class CaracteristicaAgrupacion

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbCaracteristicaAgrupacion"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDAgrupacion")) = 0 Then ApplicationService.GenerateError("La Agrupaci�n es un dato obligatorio.")
        If Length(data("DescAgrupacion")) = 0 Then ApplicationService.GenerateError("La Descripci�n es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = New CaracteristicaAgrupacion().SelOnPrimaryKey(data("IDAgrupacion"))
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("Ese registro ya existe en la Base de Datos.")
            End If
        End If
    End Sub

#End Region

End Class