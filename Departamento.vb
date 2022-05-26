Public Class Departamento

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroDepartamento"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("DescDepartamento")) = 0 Then ApplicationService.GenerateError("Introduzca la descripci�n del departamento")
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDDepartamento")) > 0 Then
                Dim DtTemp As DataTable = New Departamento().SelOnPrimaryKey(data("IDDepartamento"))
                If Not DtTemp Is Nothing AndAlso DtTemp.Rows.Count > 0 Then
                    ApplicationService.GenerateError("Departamento duplicado.")
                End If
            Else : ApplicationService.GenerateError("Introduzca el c�digo del departament")
            End If
        End If
    End Sub

#End Region

End Class