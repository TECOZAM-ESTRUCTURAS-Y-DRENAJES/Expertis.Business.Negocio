Public Class ClienteRepresentante

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper
    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub
    Private Const cnEntidad As String = "tbClienteRepresentante"

#End Region

#Region "Clases"
    <Serializable()> _
    Public Class DataComision
        Public Comision As String
        Public Porcentaje As Boolean
    End Class

#End Region

#Region "Eventos GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim Obrl As New BusinessRules
        Obrl.Add("IDArticulo", AddressOf CambioArticulo)
        Return Obrl
    End Function

    <Task()> Public Shared Sub CambioArticulo(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) = 0 Then
            ApplicationService.GenerateError("El artículo es obligatorio.")
        Else
            Dim DtArt As DataTable = New Articulo().SelOnPrimaryKey(data.Value)
            If DtArt Is Nothing OrElse DtArt.Rows.Count = 0 Then
                ApplicationService.GenerateError("El codigo del artículo no existe")
            End If
        End If
    End Sub

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarIDRepresentante)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarIDCliente)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarIDArticulo)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarComision)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarArticuloCliente)
    End Sub

    <Task()> Public Shared Sub ValidarIDRepresentante(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IdRepresentante")) = 0 Then
            ApplicationService.GenerateError("El representante es obligatorio.")
        Else
            Dim DtRepresen As DataTable = New Representante().SelOnPrimaryKey(data("IdRepresentante"))
            If DtRepresen Is Nothing OrElse DtRepresen.Rows.Count = 0 Then
                ApplicationService.GenerateError("El codigo del representante no existe")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarIDCliente(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDCliente")) = 0 Then
            ApplicationService.GenerateError("El Cliente es un dato obligatorio.")
        Else
            'Validamos que exista el cliente
            Dim DtCliente As DataTable = New Cliente().SelOnPrimaryKey(data("IDCliente"))
            If DtCliente Is Nothing OrElse DtCliente.Rows.Count = 0 Then
                ApplicationService.GenerateError("El cliente no existe en la Base de Datos")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarIDArticulo(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDArticulo")) > 0 Then
            'Validamos que exista el artículo
            Dim DtArt As DataTable = New Articulo().SelOnPrimaryKey(data("IDArticulo"))
            If DtArt Is Nothing OrElse DtArt.Rows.Count = 0 Then
                'El artículo introducido no existe en la BD
                ApplicationService.GenerateError("El Artículo introducido no existe")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarComision(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("Comision")) = 0 Then
            ApplicationService.GenerateError("El campo Comisión ha de ser numérico.")
        Else
            If Not data("Porcentaje") Is System.DBNull.Value AndAlso data("Porcentaje") = True Then
                If CDbl(data("Comision")) < 0 Or CDbl(data("Comision")) > 100 Then
                    ApplicationService.GenerateError("El campo Comisión es un porcentaje, introduzca valores entre 0 y 100.")
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarArticuloCliente(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim f As New Filter(FilterUnionOperator.And)
        f.Add(New StringFilterItem("IDCliente", FilterOperator.Equal, data("IDCliente")))
        f.Add(New StringFilterItem("IDRepresentante", FilterOperator.Equal, data("IDRepresentante")))
        If Length(data("IDArticulo")) = 0 Then
            f.Add(New IsNullFilterItem("IDArticulo", True))
        Else
            f.Add(New StringFilterItem("IDArticulo", FilterOperator.Equal, data("IDArticulo")))
        End If
        If Not data.RowState = DataRowState.Added Then
            f.Add(New NumberFilterItem("IDClienteRepresentante", FilterOperator.NotEqual, data("IDClienteRepresentante")))
        End If
        Dim DtDatos As DataTable = New ClienteRepresentante().Filter(f)
        If Not DtDatos Is Nothing AndAlso DtDatos.Rows.Count > 0 Then
            ' Hay repetición
            If Length(data("IDArticulo")) = 0 Then
                ApplicationService.GenerateError("Ya existe un registro con ese Cliente.")
            Else
                ApplicationService.GenerateError("Ya existe un registro con ese Cliente y ese Artículo.")
            End If
        End If
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarPrimaryKey)
    End Sub

    <Task()> Public Shared Sub AsignarPrimaryKey(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            data("IdClienteRepresentante") = AdminData.GetAutoNumeric
        End If
    End Sub

#End Region

End Class