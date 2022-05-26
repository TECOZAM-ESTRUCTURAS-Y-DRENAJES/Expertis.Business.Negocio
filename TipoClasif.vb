Public Class TipoClasif
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroTipoClasif"

#Region " Validate "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarIdentificadorObligatorio)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDescripcionObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarTipoFacturaObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarTipoFacturaNoExistente)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarTipoClasificacionExistente)
    End Sub

    <Task()> Public Shared Sub ValidarIdentificadorObligatorio(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added AndAlso Length(data("IDTipoClasif")) = 0 Then
            ApplicationService.GenerateError("El Tipo de Clasificación es un dato obligatorio.")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarDescripcionObligatoria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("DescTipoClasif")) = 0 Then
            ApplicationService.GenerateError("La descripción del tipo de clasificación es un dato obligatorio.")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarTipoFacturaObligatoria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDTipoFactura")) = 0 Then
            ApplicationService.GenerateError("El Tipo de Factura es un dato obligatorio.")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarTipoFacturaNoExistente(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim dt As DataTable = New TipoFactura().SelOnPrimaryKey(data("IDTipoFactura"))
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then
            ApplicationService.GenerateError("El Tipo de Factura introducido no existe.")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarTipoClasificacionExistente(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = New TipoClasif().SelOnPrimaryKey(data("IDTipoClasif"))
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("El Tipo Clasificación introducido ya existe.")
            End If
        End If
    End Sub

#End Region

End Class