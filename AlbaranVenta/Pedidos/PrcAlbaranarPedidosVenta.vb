Imports Solmicro.Expertis.Engine.BE.BusinessProcesses

Public Class PrcAlbaranarPedidosVenta
    Inherits Process(Of DataPrcAlbaranar, AlbaranLogProcess)

    '//Crea la secuencia de Tareas a realizar en el proceso de Albaranar los Pedidos
    Public Overrides Sub RegisterTasks()
        Me.AddTask(Of DataPrcAlbaranar)(AddressOf ProcesoAlbaranVenta.ValidacionesContador)
        Me.AddTask(Of DataPrcAlbaranar)(AddressOf ProcesoAlbaranVenta.DatosIniciales)
        Me.AddTask(Of DataPrcAlbaranar, AlbCabVentaPedido())(AddressOf ProcesoAlbaranVentaPedidos.AgruparPedidos)

        Me.AddTask(Of AlbCabVentaPedido(), AlbCabVentaPedido())(AddressOf ProcesoAlbaranVenta.ValidacionesPreviasDesdePedido)

        '//Bucle para recorrer todos los documentos de Albarán de Venta a generar
        Me.AddForEachTask(Of PrcCrearAlbaranVentaPedidos)(OnExceptionBehaviour.NextLoop)
        Me.AddTask(Of Object, AlbaranLogProcess)(AddressOf ProcesoComunes.ResultadoAlbaran)
    End Sub

End Class

