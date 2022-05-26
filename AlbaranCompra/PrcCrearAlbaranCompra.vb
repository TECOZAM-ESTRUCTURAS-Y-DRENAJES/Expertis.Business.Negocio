Imports Solmicro.Expertis.Engine.BE.BusinessProcesses

Public Class PrcCrearAlbaranCompra
    Inherits Process

    '//Crea la secuencia de Tareas a realizar
    Public Overrides Sub RegisterTasks()
        Me.AddTask(Of AlbCabCompra, DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.CrearDocumentoAlbaranCompra)
        Me.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.AsignarValoresPredeterminadosGenerales)
        Me.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.AsignarDatosProveedor)
        Me.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.AsignarDireccion)
        Me.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoCompra.AsignarObservacionesCompra)
        Me.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoComunes.AsignarCentroGestion)
        Me.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.AsignarCentroCoste)
        Me.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoComunes.AsignarAlmacen)
        Me.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.AsignarContador)
        Me.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoComunes.AsignarNumeroAlbaran)
        Me.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoCompra.AsignarEjercicio)
        Me.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoComunes.AsignarCambiosMoneda)
        Me.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.CrearLineasDesdePedido)
        Me.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.AsignarDatosTransporte)
        '  Me.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.GestionArticulosKit)
        Me.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoComunes.CalcularImporteLineasAlbaran)
        Me.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.ValoracionSuministro)
        Me.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.CopiarAnalitica)
        Me.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.CopiarGastos)
        Me.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.AsignarAlbaranCompraLotes)
        Me.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.CalcularBasesImponibles)
        Me.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.TotalPesos)
        Me.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoComunes.TotalDocumento)
        Me.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.ActualizarEstadoAlbaran)

        Me.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.GrabarDocumento)
        Me.AddTask(Of DocumentoAlbaranCompra)(AddressOf AñadirAResultado)
        Me.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.ActualizacionAutomaticaStock)
    End Sub

    <Task()> Public Shared Function EsAlbaranSubcontratacion(ByVal Doc As DocumentoAlbaranCompra, ByVal services As ServiceProvider) As Boolean
        Dim AppParamsCompra As ParametroCompra = services.GetService(Of ParametroCompra)()
        Return (Doc.HeaderRow("IDTipoCompra") = AppParamsCompra.TipoCompraSubcontratacion)
    End Function

    <Task()> Public Shared Sub AñadirAResultado(ByVal oDocAlb As DocumentoAlbaranCompra, ByVal services As ServiceProvider)
        Dim alog As AlbaranLogProcess = services.GetService(Of AlbaranLogProcess)()
        If alog.CreateData Is Nothing Then alog.CreateData = New LogProcess
        ReDim Preserve alog.CreateData.CreatedElements(UBound(alog.CreateData.CreatedElements) + 1)
        alog.CreateData.CreatedElements(UBound(alog.CreateData.CreatedElements)) = New CreateElement
        alog.CreateData.CreatedElements(UBound(alog.CreateData.CreatedElements)).IDElement = oDocAlb.HeaderRow("IDAlbaran")
        alog.CreateData.CreatedElements(UBound(alog.CreateData.CreatedElements)).NElement = oDocAlb.HeaderRow("NAlbaran")
    End Sub

    Protected Overrides Function OnException(ByVal exceptionArgs As Engine.BE.BusinessProcesses.ProcessExceptionArgs) As Engine.BE.BusinessProcesses.OnExceptionBehaviour
        ProcessServer.ExecuteTask(Of Boolean)(AddressOf Comunes.RollbackTransaction, True, New ServiceProvider())

        Dim alog As AlbaranLogProcess = exceptionArgs.Services.GetService(Of AlbaranLogProcess)()
        If alog.CreateData Is Nothing Then alog.CreateData = New LogProcess
        ReDim Preserve alog.CreateData.Errors(alog.CreateData.Errors.Length)
        If TypeOf exceptionArgs.TaskData Is DocumentoAlbaranCompra Then
            Dim alb As AlbCabPedidoCompra = CType(exceptionArgs.TaskData, DocumentoAlbaranCompra).Cabecera
            Select Case alb.Agrupacion
                Case enummpAgrupAlbaran.mpPedido
                    alog.CreateData.Errors(alog.CreateData.Errors.Length - 1) = New ClassErrors(" Pedido: " & alb.NPedido, exceptionArgs.Exception.Message)
                Case enummpAgrupAlbaran.mpProveedor
                    alog.CreateData.Errors(alog.CreateData.Errors.Length - 1) = New ClassErrors(" Proveedor: " & alb.IDProveedor, exceptionArgs.Exception.Message)
            End Select
        Else
            alog.CreateData.Errors(alog.CreateData.Errors.Length - 1) = New ClassErrors(String.Empty, exceptionArgs.Exception.Message)
        End If
        Return MyBase.OnException(exceptionArgs)
    End Function

End Class
