Public Class Banco

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroBanco"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("DescBanco")) = 0 Then ApplicationService.GenerateError("Introduzca la descripción del banco")
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As datarow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDBanco")) > 0 Then
                Dim DtTemp As DataTable = New Banco().SelOnPrimaryKey(data("IDBanco"))
                If Not DtTemp Is Nothing AndAlso DtTemp.Rows.Count > 0 Then ApplicationService.GenerateError("Ya existe un banco con esa clave.")
            Else : ApplicationService.GenerateError("Introduzca el código del banco.")
            End If
        End If
    End Sub

#End Region

End Class