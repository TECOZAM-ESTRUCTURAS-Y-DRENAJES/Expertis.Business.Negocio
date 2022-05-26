
Imports Solmicro.Expertis.Engine.BE.BusinessProcesses

Public Class PrcCrearPedidoVentaPrograma
    Inherits Process(Of DataPrcCrearPedidoVentaPrograma, LogProcess)

    'Lista de tareas a ejecutar
    Public Overrides Sub RegisterTasks()
        Me.AddTask(Of DataPrcCrearPedidoVentaPrograma)(AddressOf ValidacionesPrevias)
        Me.AddTask(Of DataPrcCrearPedidoVentaPrograma)(AddressOf DatosIniciales)
        Me.AddTask(Of DataPrcCrearPedidoVentaPrograma, PedCabPrograma())(AddressOf ProcesoPedidoVenta.AgruparProgramas)
        Me.AddTask(Of PedCabPrograma())(AddressOf ProcesoPedidoVenta.Ordenar)
        'Bucle para recorrer todos los documentos a factura a generar
        Me.AddForEachTask(Of PrcCrearPedidoVenta)(OnExceptionBehaviour.NextLoop)
        Me.AddTask(Of Object, LogProcess)(AddressOf ProcesoComunes.ResultadoLogProcess)
    End Sub

    'Registar en el services la información que vamos a compartir durante el proceso, para evitar accesos innecesarios
    <Task()> Public Shared Sub DatosIniciales(ByVal data As DataPrcCrearPedidoVentaPrograma, ByVal services As ServiceProvider)
        '//Prepara en el service información del proceso.
        services.RegisterService(New ProcessInfoPV(data.IDContador))

        '//Prepara en el service lo que va a retornar el proceso, para tenerlo disponible en todo el proceso.
        services.RegisterService(New LogProcess)
    End Sub

    <Task()> Public Shared Sub ValidacionesPrevias(ByVal data As DataPrcCrearPedidoVentaPrograma, ByVal services As ServiceProvider)
        'ProcessServer.ExecuteTask(Of String)(AddressOf ProcesoComunes.ValidarContadorObligatorio, data.IDContador & String.Empty, services)

        Dim datosValCont As New ProcesoComunes.DataValidarContadorEntidad(data.IDContador, GetType(PedidoVentaCabecera).Name)
        ProcessServer.ExecuteTask(Of ProcesoComunes.DataValidarContadorEntidad)(AddressOf ProcesoComunes.ValidarContadorEntidad, datosValCont, services)
    End Sub

  

End Class
