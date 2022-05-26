Imports Solmicro.Expertis.Engine.BE.BusinessProcesses

Public Class PrcAlbaranarObras
    Inherits Process

    '//Crea la secuencia de Tareas a realizar en el proceso de Albaranar los Pedidos
    Public Overrides Sub RegisterTasks()
        Me.AddTask(Of DataPrcAlbaranar)(AddressOf ProcesoAlbaranVenta.ValidacionesContador)
        Me.AddTask(Of DataPrcAlbaranar)(AddressOf ProcesoAlbaranVenta.DatosIniciales)
        Me.AddTask(Of DataPrcAlbaranar, AlbCabVentaObras())(AddressOf ProcesoAlbaranVentaObras.AgruparObras)
        Me.AddTask(Of AlbCabVentaObras(), AlbCabVentaObras())(AddressOf ProcesoAlbaranVenta.ValidacionesPreviasDesdeObras)
        '//Bucle para recorrer todos los documentos de Albarán de Venta a generar
        Me.AddForEachTask(Of PrcCrearAlbaranVentaObras)(OnExceptionBehaviour.NextLoop)
        Me.AddTask(Of Object, AlbaranLogProcess)(AddressOf ProcesoComunes.ResultadoAlbaran)
    End Sub

End Class
