Public Class PrcCrearFacturaVentaMnto
    Inherits Process(Of DataPrcCrearFacturaVentaMnto, LogProcess)

    'Lista de tareas a ejecutar
    Public Overrides Sub RegisterTasks()
        Me.AddTask(Of DataPrcCrearFacturaVentaMnto)(AddressOf ValidacionesPrevias)
        Me.AddTask(Of DataPrcCrearFacturaVentaMnto)(AddressOf DatosIniciales)
        ' Me.AddTask(Of DataPrcCrearFacturaVentaMnto, FraCab())(AddressOf ProcesoFacturacionVenta.AgruparAlbaranes)
        Me.AddTask(Of FraCab())(AddressOf ProcesoFacturacionVenta.Ordenar)
        'Bucle para recorrer todos los documentos a factura a generar
        Me.AddForEachTask(Of PrcCrearFacturaVenta)(OnExceptionBehaviour.NextLoop)
        Me.AddTask(Of Object, LogProcess)(AddressOf ProcesoComunes.ResultadoLogProcess)
    End Sub

    <Task()> Public Shared Sub ValidacionesPrevias(ByVal data As DataPrcCrearFacturaVentaMnto, ByVal services As ServiceProvider)
        Dim datosValCont As New ProcesoComunes.DataValidarContadorEntidad(data.IDContador, GetType(FacturaVentaCabecera).Name)
        ProcessServer.ExecuteTask(Of ProcesoComunes.DataValidarContadorEntidad)(AddressOf ProcesoComunes.ValidarContadorEntidad, datosValCont, services)
    End Sub

    <Task()> Public Shared Sub DatosIniciales(ByVal data As DataPrcCrearFacturaVentaMnto, ByVal services As ServiceProvider)
        '//Prepara en el service información del proceso.
        ' services.RegisterService(New ProcessInfofv(data.IDContador, data.IDOperario))

        '//Prepara en el service lo que va a retornar el proceso, para tenerlo disponible en todo el proceso.
        services.RegisterService(New LogProcess)
    End Sub
End Class
