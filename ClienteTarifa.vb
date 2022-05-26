Public Class ClienteTarifa

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper
    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub
    Private Const cnEntidad As String = "tbClienteTarifa"

#End Region

#Region "Eventos RegisterDeleteTasks"

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf ActualizarOfertaComercialDetalle)
        'deleteProcess.AddTask(Of DataRow)(AddressOf Reordenar)
    End Sub

    <Task()> Public Shared Sub ActualizarOfertaComercialDetalle(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDLineaOfertaDetalle")) > 0 Then
            Dim ClsObj As BusinessHelper = CreateBusinessObject("OfertaComercialDetalle")
            Dim dtO As DataTable = ClsObj.Filter(New FilterItem("IDLineaOfertaDetalle", FilterOperator.Equal, data("IDLineaOfertaDetalle"), FilterType.Numeric), "IDLineaOfertaDetalle,EstadoCliente")
            If Not dtO Is Nothing AndAlso dtO.Rows.Count > 0 Then
                dtO.Rows(0)("EstadoCliente") = False
                BusinessHelper.UpdateTable(dtO)
            End If
        End If
    End Sub

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarCliente)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarIDTarifa)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarPrimaryKey)
    End Sub

    <Task()> Public Shared Sub ValidarCliente(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDCliente")) = 0 Then ApplicationService.GenerateError("El Cliente es un dato obligatorio.")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarIDTarifa(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IdTarifa")) = 0 Then ApplicationService.GenerateError("La Tarifa es un dato obligatorio.")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarPrimaryKey(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dtTemp As DataTable = New ClienteTarifa().SelOnPrimaryKey(data("IdCliente"), data("IdTarifa"))
            If Not dtTemp Is Nothing AndAlso dtTemp.Rows.Count > 0 Then
                ApplicationService.GenerateError("Este Cliente ya tiene asociada esa Tarifa")
            End If
        End If
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarOrden)
        ''updateProcess.AddTask(Of DataRow)(AddressOf ReOrdenar)
    End Sub

    <Task()> Public Shared Sub AsignarOrden(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Nz(data("Orden"), 0) = 0 Then data("Orden") = 1
    End Sub

#End Region

#Region "Funciones Privadas"

    <Task()> Public Shared Sub ReOrdenar(ByVal data As DataRow, ByVal services As ServiceProvider)
        'Funcion que establece el orden de una determinada columna, haciendo que sus valores sean correlativos y que esten ordenados.
        'Cuando se le llama desde el DELETE: Se le pasa el rs con un 0 en la columna que hay que reordenar.
        'La llamada a esta funcion se hace despues de hacer la insercion, acrualizacion o borrado.
        If Not data Is Nothing Then
            Dim f As New Filter
            f.Add("IDCliente", FilterOperator.Equal, data("IdCliente"))
            f.Add("IdTarifa", FilterOperator.NotEqual, data("IdTarifa"))
            Dim dtOrden As DataTable = New ClienteTarifa().Filter(f, "Orden")

            If Not dtOrden Is Nothing AndAlso dtOrden.Rows.Count > 0 Then
                Dim dtActual As DataTable
                If data("Orden") > dtOrden.Rows.Count + 1 Then
                    dtActual = New ClienteTarifa().SelOnPrimaryKey(data("IdCliente"), data("IdTarifa"))
                    dtActual.Rows(0)("Orden") = dtOrden.Rows.Count + 1

                    BusinessHelper.UpdateTable(dtActual)
                End If

                Dim intOrden As Integer = 1
                For Each drOrden As DataRow In dtOrden.Rows
                    dtActual = New ClienteTarifa().SelOnPrimaryKey(drOrden("IdCliente"), drOrden("IdTarifa"))
                    If AreEquals(intOrden, data("Orden")) Then intOrden = intOrden + 1

                    dtActual.Rows(0)("Orden") = intOrden
                    BusinessHelper.UpdateTable(dtActual)

                    intOrden = intOrden + 1
                Next
            End If
        End If
    End Sub

#End Region

End Class