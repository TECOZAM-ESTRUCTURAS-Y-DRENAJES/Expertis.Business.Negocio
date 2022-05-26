Public Class Plantilla

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroPlantilla"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("Descripcion")) = 0 Then ApplicationService.GenerateError("La descripción es un dato obligatorio.")
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = New Plantilla().SelOnPrimaryKey(data("IDArticulo"), data("Descripcion"))
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("Ya existe una especificación con esta descripción para este artículo.")
            End If
        End If
    End Sub

#End Region

End Class