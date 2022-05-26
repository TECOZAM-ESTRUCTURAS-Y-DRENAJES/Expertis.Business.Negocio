'Proceso de AUTOFACTURACION establecido de forma estándar, relación de tareas
Imports Solmicro.Expertis.Engine.BE.BusinessProcesses

Public Class PrcAutofacturacion
    Inherits Process(Of DataPrcAutofacturacion, ResultFacturacion)
    'Lista de tareas a ejecutar
    Public Overrides Sub RegisterTasks()
        Me.AddTask(Of DataPrcAutofacturacion)(AddressOf ValidacionesPrevias)
        Me.AddTask(Of DataPrcAutofacturacion)(AddressOf DatosIniciales)
        Me.AddTask(Of DataPrcAutofacturacion, FraCabAlbaran())(AddressOf ProcesoFacturacionVenta.AgruparAlbaranesAutoFra)
        Me.AddTask(Of FraCabAlbaran())(AddressOf ProcesoFacturacionVenta.Ordenar)
        'Bucle para recorrer todos los documentos a factura a generar
        Me.AddForEachTask(Of PrcCrearFacturaVenta)(OnExceptionBehaviour.NextLoop)
        Me.AddTask(Of Object, ResultFacturacion)(AddressOf ProcesoFacturacionVenta.Resultado)
    End Sub

    'Registar en el services la información que vamos a compartir durante el proceso, para evitar accesos innecesarios
    <Task()> Public Shared Sub DatosIniciales(ByVal data As DataPrcAutofacturacion, ByVal services As ServiceProvider)
        Dim TipoLineaDef As String = ProcessServer.ExecuteTask(Of Object, String)(AddressOf TipoLinea.TipoLineaPorDefecto, Nothing, services)
        services.RegisterService(New ProcessInfoFra(data.IDContador, TipoLineaDef, , , , data.ConPropuesta))
        If data.ConPropuesta Then
            Dim Facturas As DataTable = New FacturaVentaCabecera().AddNew
            services.RegisterService(New ResultFacturacion(Facturas))
        End If
    End Sub

    <Task()> Public Shared Sub ValidacionesPrevias(ByVal data As DataPrcAutofacturacion, ByVal services As ServiceProvider)
        ' ProcessServer.ExecuteTask(Of String)(AddressOf ProcesoComunes.ValidarContadorObligatorio, data.IDContador & String.Empty, services)

        Dim datosValCont As New ProcesoComunes.DataValidarContadorEntidad(data.IDContador, GetType(FacturaVentaCabecera).Name)
        ProcessServer.ExecuteTask(Of ProcesoComunes.DataValidarContadorEntidad)(AddressOf ProcesoComunes.ValidarContadorEntidad, datosValCont, services)
    End Sub


End Class




