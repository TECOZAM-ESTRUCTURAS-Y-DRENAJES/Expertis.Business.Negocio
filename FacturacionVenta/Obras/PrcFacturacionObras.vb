Imports Solmicro.Expertis.Engine.BE.BusinessProcesses

Public Class PrcFacturacionObras
    Inherits Process(Of DataPrcFacturacionObras, ResultFacturacion)

    'Lista de tareas a ejecutar
    Public Overrides Sub RegisterTasks()
        Me.AddTask(Of DataPrcFacturacionObras)(AddressOf ValidacionesPrevias)
        Me.AddTask(Of DataPrcFacturacionObras)(AddressOf DatosIniciales)
        Me.AddTask(Of DataPrcFacturacionObras, FraCabObra())(AddressOf ProcesoFacturacionObras.AgruparObras)
        Me.AddTask(Of FraCabObra())(AddressOf ProcesoFacturacionObras.Ordenar)
        'Bucle para recorrer todos los documentos a factura a generar
        Me.AddForEachTask(Of PrcCrearFacturaVentaObra)(OnExceptionBehaviour.NextLoop)
        Me.AddTask(Of Object, ResultFacturacion)(AddressOf ProcesoFacturacionObras.Resultado)
    End Sub

    'Registar en el services la información que vamos a compartir durante el proceso, para evitar accesos innecesarios
    <Task()> Public Shared Sub DatosIniciales(ByVal data As DataPrcFacturacionObras, ByVal services As ServiceProvider)
        services.RegisterService(New ProcessInfoFraObras(data.IDContador, data.TipoFacturacion, data.TipoAgrupacion, data.CalculoSeguros))
        Dim Facturas As DataTable = New FacturaVentaCabecera().AddNew
        services.RegisterService(New ResultFacturacion(Facturas))
    End Sub

    <Task()> Public Shared Sub ValidacionesPrevias(ByVal data As DataPrcFacturacionObras, ByVal services As ServiceProvider)
        'ProcessServer.ExecuteTask(Of String)(AddressOf ProcesoComunes.ValidarContadorObligatorio, data.IDContador & String.Empty, services)

        Dim datosValCont As New ProcesoComunes.DataValidarContadorEntidad(data.IDContador, GetType(FacturaVentaCabecera).Name)
        ProcessServer.ExecuteTask(Of ProcesoComunes.DataValidarContadorEntidad)(AddressOf ProcesoComunes.ValidarContadorEntidad, datosValCont, services)
    End Sub

End Class