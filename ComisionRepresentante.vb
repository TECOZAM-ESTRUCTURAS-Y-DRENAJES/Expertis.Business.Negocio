Public Class ComisionRepresentante

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbComisionRepresentante"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IdRepresentante")) = 0 Then
            ApplicationService.GenerateError("El representante es obligatorio.")
        Else
            Dim DtRepresen As DataTable = New Representante().SelOnPrimaryKey(data("IdRepresentante"))
            If DtRepresen Is Nothing OrElse DtRepresen.Rows.Count = 0 Then
                ApplicationService.GenerateError("El codigo del representante no existe")
            End If
        End If
        If data("IDTipo") Is Nothing OrElse Length(data("IDTipo")) = 0 Then
            ApplicationService.GenerateError("El Tipo de artículo es un dato obligatorio.")
        Else
            Dim DtTipo As DataTable = New TipoArticulo().SelOnPrimaryKey(data("IDTipo"))
            If DtTipo Is Nothing OrElse DtTipo.Rows.Count = 0 Then
                ApplicationService.GenerateError("El Tipo del Artículo no existe en la Base de datos.")
            End If
        End If
        If Not data("IDFamilia") Is Nothing AndAlso Length(data("IDFamilia")) > 0 Then
            Dim DtFamilia As DataTable = New Familia().SelOnPrimaryKey(data("IDTipo"), data("IDFamilia"))
            If DtFamilia Is Nothing OrElse DtFamilia.Rows.Count = 0 Then
                ApplicationService.GenerateError("La Familia no existe en la Base de datos.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim f As New Filter
        f.Add(New StringFilterItem("IDTipo", FilterOperator.Equal, data("IDTipo")))
        f.Add(New StringFilterItem("IDRepresentante", FilterOperator.Equal, data("IDRepresentante")))
        If Length(data("IDFamilia")) = 0 Then
            f.Add(New IsNullFilterItem("IDFamilia", True))
        Else : f.Add(New StringFilterItem("IDFamilia", FilterOperator.Equal, data("IDFamilia")))
        End If
        If Not data.RowState = DataRowState.Added Then
            f.Add(New NumberFilterItem("IDComisionRepresentante", FilterOperator.NotEqual, data("IDComisionRepresentante")))
        End If
        Dim DtDatos As DataTable = New ComisionRepresentante().Filter(f)
        If Not DtDatos Is Nothing AndAlso DtDatos.Rows.Count > 0 Then
            If data("IDFamilia") Is Nothing OrElse Length(data("IDFamilia")) = 0 Then
                ApplicationService.GenerateError("Ya existe un registro con ese Tipo y sin Familia.")
            Else : ApplicationService.GenerateError("Ya existe un registro con ese Tipo y Familia")
            End If
        End If
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClave)
    End Sub

    <Task()> Public Shared Sub AsignarClave(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then data("IDComisionRepresentante") = AdminData.GetAutoNumeric
    End Sub

#End Region

End Class