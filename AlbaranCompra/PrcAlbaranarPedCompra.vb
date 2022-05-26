Imports Solmicro.Expertis.Engine.BE.BusinessProcesses

Public Class PrcAlbaranarPedCompra
    Inherits Process

    '//Crea la secuencia de Tareas a realizar en el proceso de Albaranar los Pedidos
    Public Overrides Sub RegisterTasks()
        Me.AddTask(Of DataPrcAlbaranarPedCompra)(AddressOf ProcesoAlbaranCompra.ValidacionesContador)
        Me.AddTask(Of DataPrcAlbaranarPedCompra)(AddressOf ProcesoAlbaranCompra.DatosIniciales)
        Me.AddTask(Of DataPrcAlbaranarPedCompra, AlbCabPedidoCompra())(AddressOf ProcesoAlbaranCompra.AgruparPedidos)
        Me.AddTask(Of AlbCabPedidoCompra())(AddressOf ProcesoAlbaranCompra.ValidacionesPrevias)
        '//Bucle para recorrer todos los documentos de Albarán de Compra a generar
        Me.AddForEachTask(Of PrcCrearAlbaranCompra)(OnExceptionBehaviour.NextLoop)
        Me.AddTask(Of Object, AlbaranLogProcess)(AddressOf ProcesoComunes.ResultadoAlbaran)
    End Sub

End Class

