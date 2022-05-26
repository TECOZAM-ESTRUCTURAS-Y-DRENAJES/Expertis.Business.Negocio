Public Class Idioma

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroIdioma"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("DescIdioma")) = 0 Then ApplicationService.GenerateError("Introduzca la descripci�n del idioma")
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDIdioma")) > 0 Then
                Dim dtTemp As DataTable = New Idioma().SelOnPrimaryKey(data("IDIdioma"))
                If Not dtTemp Is Nothing AndAlso dtTemp.Rows.Count > 0 Then
                    ApplicationService.GenerateError("Idioma duplicado.")
                End If
            Else : ApplicationService.GenerateError("Introduzca el c�digo del idioma.")
            End If
        End If
    End Sub

#End Region

End Class