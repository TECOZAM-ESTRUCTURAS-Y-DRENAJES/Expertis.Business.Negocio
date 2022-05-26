Public Class CaracteristicaArticulo2

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbCaracteristicaArticulo2"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveCaracteristica)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDCaracteristica2")) = 0 Then ApplicationService.GenerateError("La Clave de la Característica es un Dato Obligatorio.")
        If Length(data("DescCaracteristica2")) = 0 Then ApplicationService.GenerateError("La Descripción de la Característica es un Dato Obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarClaveCaracteristica(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim DtClave As DataTable = New CaracteristicaArticulo2().SelOnPrimaryKey(data("IDCaracteristica2"))
        If Not DtClave Is Nothing AndAlso DtClave.Rows.Count > 0 Then
            ApplicationService.GenerateError("La Clave de Característica introducida ya existe en la Base de Datos.")
        End If
    End Sub

#End Region

End Class