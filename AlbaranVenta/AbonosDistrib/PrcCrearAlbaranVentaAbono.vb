Public Class PrcCrearAlbaranVentaAbono
    Inherits Process(Of AlbCabVentaAlbaran, DocumentoAlbaranVenta)

    '//Crea la secuencia de Tareas a realizar
    Public Overrides Sub RegisterTasks()
        'Me.AddTask(Of DataPrcCrearAlbaranVentaAbono, AlbCabVenta)(AddressOf DatosIniciales)

        Me.AddTask(Of AlbCabVenta, DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.CrearDocumentoAlbaranVenta)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.AsignarDatosAlbaranOrigen)

        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoComunes.AsignarNumeroAlbaran)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.CalcularBasesImponibles)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoComunes.TotalDocumento)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.TotalPesos)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.ActualizarEstadoAlbaran)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.GrabarDocumento)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.AñadirAResultado)

        ' Me.AddTask(Of Object, AlbaranLogProcess)(AddressOf ProcesoComunes.ResultadoAlbaran)
    End Sub

    Protected Overrides Function OnException(ByVal exceptionArgs As Engine.BE.BusinessProcesses.ProcessExceptionArgs) As Engine.BE.BusinessProcesses.OnExceptionBehaviour
        AdminData.RollBackTx()

        Dim alog As AlbaranLogProcess = exceptionArgs.Services.GetService(Of AlbaranLogProcess)()
        If alog.CreateData Is Nothing Then alog.CreateData = New LogProcess
        ReDim Preserve alog.CreateData.Errors(alog.CreateData.Errors.Length)
        If TypeOf exceptionArgs.TaskData Is DocumentoAlbaranVenta Then
            Dim alb As AlbCabVenta = CType(exceptionArgs.TaskData, DocumentoAlbaranVenta).Cabecera
            alog.CreateData.Errors(alog.CreateData.Errors.Length - 1) = New ClassErrors(" Albarán: " & alb.NOrigen, exceptionArgs.Exception.Message)

        Else
            alog.CreateData.Errors(alog.CreateData.Errors.Length - 1) = New ClassErrors(String.Empty, exceptionArgs.Exception.Message)
        End If

        Return MyBase.OnException(exceptionArgs)
    End Function

End Class
