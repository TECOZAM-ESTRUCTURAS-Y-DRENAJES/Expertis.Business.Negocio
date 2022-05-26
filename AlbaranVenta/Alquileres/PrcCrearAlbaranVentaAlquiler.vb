Public Class PrcCrearAlbaranVentaAlquiler
    Inherits Process

    '//Crea la secuencia de Tareas a realizar
    Public Overrides Sub RegisterTasks()
        Me.AddTask(Of AlbCabVentaAlquiler, DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.CrearDocumentoAlbaranVenta)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.AsignarValoresPredeterminadosGenerales)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVentaObras.AsignarDatosCliente)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVentaObras.AsignarCentroGestion)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVentaObras.AsignarPedidoCliente)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.AsignarDireccion)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVentaObras.AsignarTexto)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVentaAlquiler.AsignarAlmacenCabecera)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.AsignarContador)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVentaAlquiler.AsignarNumeroAlbaranAlquilerPropuesta)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoComercial.AsignarEjercicio)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoComunes.AsignarCambiosMoneda)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVentaAlquiler.AsignarEstado)

        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVentaObras.CrearLineasDesdeObras)  'Sin Kits
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.GestionArticulosKit)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.GestionArticulosFantasma)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoComunes.CalcularImporteLineasAlbaran)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.CalcularBasesImponibles)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoComunes.TotalDocumento)
        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.TotalPesos)

        Me.AddTask(Of DocumentoAlbaranVenta)(AddressOf AñadirAResultado)
    End Sub

    <Task()> Public Shared Sub AñadirAResultado(ByVal doc As DocumentoAlbaranVenta, ByVal services As ServiceProvider)
        Dim Albaranes As DataTable = services.GetService(Of ResultAlbaranAlquiler)().PropuestaAlbaranes
        Albaranes.Rows.Add(doc.HeaderRow.ItemArray)

        Dim arDocAlbaranes As ArrayList = services.GetService(Of ArrayList)()

        arDocAlbaranes.Add(doc)
    End Sub

    Protected Overrides Function OnException(ByVal exceptionArgs As Engine.BE.BusinessProcesses.ProcessExceptionArgs) As Engine.BE.BusinessProcesses.OnExceptionBehaviour
        ProcessServer.ExecuteTask(Of Boolean)(AddressOf Comunes.RollbackTransaction, True, New ServiceProvider())

        Dim alog As AlbaranLogProcess = exceptionArgs.Services.GetService(Of AlbaranLogProcess)()
        If alog.CreateData Is Nothing Then alog.CreateData = New LogProcess
        ReDim Preserve alog.CreateData.Errors(alog.CreateData.Errors.Length)
        If TypeOf exceptionArgs.TaskData Is DocumentoAlbaranVenta Then
            Dim alb As AlbCabVenta = CType(exceptionArgs.TaskData, DocumentoAlbaranVenta).Cabecera
            If TypeOf alb Is AlbCabVentaAlquiler Then
                alog.CreateData.Errors(alog.CreateData.Errors.Length - 1) = New ClassErrors(" Alquiler: " & alb.NOrigen, exceptionArgs.Exception.Message)
            End If
        Else
            alog.CreateData.Errors(alog.CreateData.Errors.Length - 1) = New ClassErrors(String.Empty, exceptionArgs.Exception.Message)
        End If

        Return MyBase.OnException(exceptionArgs)
    End Function

End Class
