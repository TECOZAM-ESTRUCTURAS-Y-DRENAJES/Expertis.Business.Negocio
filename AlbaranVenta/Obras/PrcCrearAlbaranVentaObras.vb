Public Class PrcCrearAlbaranVentaObras
    Inherits Process

    '//Crea la secuencia de Tareas a realizar
    Public Overrides Sub RegisterTasks()
        Me.AddTask(Of AlbCabVentaObras, DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.CrearDocumentoAlbaranVenta)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.AsignarValoresPredeterminadosGenerales)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVentaObras.AsignarDatosCliente)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVentaObras.AsignarCentroGestion)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVentaObras.AsignarPedidoCliente)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.AsignarDireccion)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVentaObras.AsignarTexto)

        'Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.AsignarBanco)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoComercial.AsignarObservacionesComercial)
        'Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.AsignarAlmacen)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.AsignarContador)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoComunes.AsignarNumeroAlbaran)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoComercial.AsignarEjercicio)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoComunes.AsignarCambiosMoneda)

        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVentaObras.CrearLineasDesdeObras)  'Sin Kits
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.GestionArticulosKit)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.GestionArticulosFantasma)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoComunes.CalcularImporteLineasAlbaran)
        'Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.CopiarAnalitica)
        'Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.CopiarRepresentantes)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.AsignarAlbaranVentaLotes)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.CalcularBasesImponibles)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoComunes.TotalDocumento)
        ''  Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.CalcularImportesAlbaran)
        'Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.TotalPesos)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.ActualizarEstadoAlbaran)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.GrabarDocumento)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.AñadirAResultado)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.ActualizacionAutomaticaStock)
    End Sub

    Protected Overrides Function OnException(ByVal exceptionArgs As Engine.BE.BusinessProcesses.ProcessExceptionArgs) As Engine.BE.BusinessProcesses.OnExceptionBehaviour
        ProcessServer.ExecuteTask(Of Boolean)(AddressOf Comunes.RollbackTransaction, True, New ServiceProvider())

        Dim alog As AlbaranLogProcess = exceptionArgs.Services.GetService(Of AlbaranLogProcess)()
        If alog.CreateData Is Nothing Then alog.CreateData = New LogProcess
        ReDim Preserve alog.CreateData.Errors(alog.CreateData.Errors.Length)
        If TypeOf exceptionArgs.TaskData Is DocumentoAlbaranVenta Then

            Dim alb As AlbCabVenta = CType(exceptionArgs.TaskData, DocumentoAlbaranVenta).Cabecera
            If TypeOf alb Is AlbCabVentaPedido Then
                Select Case CType(alb, AlbCabVentaPedido).Agrupacion
                    Case enummcAgrupAlbaran.mcPedido
                        alog.CreateData.Errors(alog.CreateData.Errors.Length - 1) = New ClassErrors(" Pedido: " & alb.NOrigen, exceptionArgs.Exception.Message)
                    Case enummcAgrupAlbaran.mcCliente
                        alog.CreateData.Errors(alog.CreateData.Errors.Length - 1) = New ClassErrors(" Cliente: " & alb.IDCliente, exceptionArgs.Exception.Message)
                End Select
            ElseIf TypeOf alb Is AlbCabVentaObras Then
                alog.CreateData.Errors(alog.CreateData.Errors.Length - 1) = New ClassErrors(" Obra: " & alb.NOrigen, exceptionArgs.Exception.Message)
            End If
        Else
            alog.CreateData.Errors(alog.CreateData.Errors.Length - 1) = New ClassErrors(String.Empty, exceptionArgs.Exception.Message)
        End If
        Return MyBase.OnException(exceptionArgs)
    End Function

End Class
