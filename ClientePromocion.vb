Public Class ClientePromocion

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper
    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub
    Private Const cnEntidad As String = "tbClientePromocion"

#End Region

#Region "Eventos RegisterDeleteTasks"

    'Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
    '    MyBase.RegisterDeleteTasks(deleteProcess)
    '    deleteProcess.AddTask(Of DataRow)(AddressOf Reordenar)
    'End Sub

#End Region

#Region "Eventos GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim Obrl As New BusinessRules
        Obrl.Add("IDCliente", AddressOf CambioCliente)
        Obrl.Add("Orden", AddressOf CambioOrden)
        Return Obrl
    End Function

    <Task()> Public Shared Sub CambioCliente(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim Dr As DataRow = New Cliente().GetItemRow(data.Value)
            data.Current("DescCliente") = Dr("DescCliente")
        End If
    End Sub

    <Task()> Public Shared Sub CambioOrden(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            If Not IsNumeric(data.Value) Then
                ApplicationService.GenerateError("Campo no numérico.")
            End If
        End If
    End Sub

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarPrimaryKey)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarIDCliente)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarPromocion)
    End Sub

    <Task()> Public Shared Sub ValidarPrimaryKey(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDPromocion")) = 0 Then ApplicationService.GenerateError("La Promoción es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarIDCliente(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDCliente")) = 0 Then ApplicationService.GenerateError("El Cliente es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarPromocion(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            'Comprobación de la existencia de la Promoción
            Dim dt As DataTable = New ClientePromocion().SelOnPrimaryKey(data("IDCliente"), data("IDPromocion"))
            If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("Ya existe esta promoción para el cliente actual.")
            End If
        End If
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf ActualizarOrden)
        'updateProcess.AddTask(Of DataRow)(AddressOf ActualizarReOrden)
    End Sub

    <Task()> Public Shared Sub ActualizarOrden(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("Orden")) = 0 Then data("Orden") = 0
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarReOrden(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Modified Then
            If data("Orden", DataRowVersion.Original) & String.Empty <> data("Orden") & String.Empty Then
                ProcessServer.ExecuteTask(Of DataRow)(AddressOf ReOrdenar, data, services)
            End If
        End If

    End Sub

#End Region

#Region "Funciones Privadas"

    <Task()> Public Shared Sub ReOrdenar(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Not IsNothing(data) Then
            'Funcion que establece el orden de una determinada columna, haciendo que sus valores sean correlativos y que esten ordenados.
            'Cuando se le llama desde el DELETE: Se le pasa el rs con un 0 en la columna que hay que reordenar.
            'La llamada a esta funcion se hace despues de hacer una modificación o un borrado.

            'Hay que seleccionar solo las promociones generales.
            Dim f As New Filter
            f.Add(New StringFilterItem("IDCliente", data("IDCliente")))
            f.Add(New StringFilterItem("IDPromocion", FilterOperator.NotEqual, data("IDPromocion")))
            f.Add(New NumberFilterItem("Orden", FilterOperator.NotEqual, 0))
            Dim dtPromosEnCurso As DataTable = New ClientePromocion().Filter(f, "Orden")
            If Not dtPromosEnCurso Is Nothing AndAlso dtPromosEnCurso.Rows.Count > 0 Then
                If data("Orden") > dtPromosEnCurso.Rows.Count + 1 Then
                    Dim dr As DataRow = New ClientePromocion().GetItemRow(data("IDCliente"), data("IDPromocion"))
                    dr("Orden") = dtPromosEnCurso.Rows.Count + 1
                    BusinessHelper.UpdateTable(dr.Table)
                End If
                Dim intOrden As Integer = 1
                For Each drPromosEnCurso As DataRow In dtPromosEnCurso.Rows
                    Dim dr As DataRow = New ClientePromocion().GetItemRow(drPromosEnCurso("IDCliente"), drPromosEnCurso("IDPromocion"))
                    If intOrden = data("Orden") Then intOrden += 1
                    dr("Orden") = intOrden
                    BusinessHelper.UpdateTable(dr.Table)
                    intOrden += 1
                Next
            End If
        End If
    End Sub

#End Region

End Class