Public Class FormaEnvio

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroFormaEnvio"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("DescFormaEnvio")) = 0 Then ApplicationService.GenerateError("Introduzca la descripci�n de la forma de env�o")
        If Length(data("IDProveedor")) > 0 Then
            Dim dtProveedor As DataTable = New Proveedor().SelOnPrimaryKey(data("IDProveedor"))
            If dtProveedor Is Nothing Or dtProveedor.Rows.Count = 0 Then
                ApplicationService.GenerateError("El Proveedor no existe en la Base de Datos")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDFormaEnvio")) > 0 Then
                Dim dtTemp As DataTable = New FormaEnvio().SelOnPrimaryKey(data("IDFormaEnvio"))
                If Not dtTemp Is Nothing AndAlso dtTemp.Rows.Count > 0 Then
                    ApplicationService.GenerateError("Ya existe una forma de env�o con la misma clave.")
                End If
            Else : ApplicationService.GenerateError("Introduzca el c�digo de la forma de env�o.")
            End If
        End If
    End Sub

#End Region

End Class