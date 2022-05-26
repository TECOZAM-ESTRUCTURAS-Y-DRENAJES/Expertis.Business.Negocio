'Proceso de AUTOFACTURACION establecido de forma estándar, relación de tareas
Imports Solmicro.Expertis.Engine.BE.BusinessProcesses

Public Class PrcAutofacturacionCompra
    Inherits Process(Of DataPrcAutofacturacionCompra, ResultFacturacion)

    'Lista de tareas a ejecutar
    Public Overrides Sub RegisterTasks()
        Me.AddTask(Of DataPrcAutofacturacionCompra)(AddressOf ValidacionesPrevias)
        Me.AddTask(Of DataPrcAutofacturacionCompra)(AddressOf DatosIniciales)
        Me.AddTask(Of DataPrcAutofacturacionCompra, FraCabCompra())(AddressOf ProcesoFacturacionCompra.AgruparAlbaranesAutoFraCompra)
        Me.AddTask(Of FraCabCompra())(AddressOf ProcesoFacturacionCompra.Ordenar)
        'Bucle para recorrer todos los documentos a factura a generar
        Me.AddForEachTask(Of PrcCrearFraCompra)(OnExceptionBehaviour.NextLoop)
        Me.AddTask(Of Object, ResultFacturacion)(AddressOf ProcesoFacturacionCompra.ResultadoPropuesta)
    End Sub

    'Registar en el services la información que vamos a compartir durante el proceso, para evitar accesos innecesarios
    <Task()> Public Shared Sub DatosIniciales(ByVal data As DataPrcAutofacturacionCompra, ByVal services As ServiceProvider)
        Dim TipoLineaDef As String = ProcessServer.ExecuteTask(Of Object, String)(AddressOf TipoLinea.TipoLineaPorDefecto, Nothing, services)
        services.RegisterService(New ProcessInfoFra(data.IDContador, TipoLineaDef, data.SuFactura))
        Dim Facturas As DataTable = New FacturaCompraCabecera().AddNew
        services.RegisterService(New ResultFacturacion(Facturas))
    End Sub

    <Task()> Public Shared Sub ValidacionesPrevias(ByVal data As DataPrcAutofacturacionCompra, ByVal services As ServiceProvider)
        'ProcessServer.ExecuteTask(Of String)(AddressOf ProcesoComunes.ValidarContadorObligatorio, data.IDContador & String.Empty, services)

        Dim datosValCont As New ProcesoComunes.DataValidarContadorEntidad(data.IDContador, GetType(FacturaCompraCabecera).Name)
        ProcessServer.ExecuteTask(Of ProcesoComunes.DataValidarContadorEntidad)(AddressOf ProcesoComunes.ValidarContadorEntidad, datosValCont, services)
    End Sub

End Class


