Public Class ZonaRepresentante

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbZonaRepresentante"

#End Region

#Region "Eventos RegisterValidateTask "

    ''' <summary>
    ''' Relación de tareas asociadas a la validación 
    ''' </summary>
    ''' <param name="validateProcess">Proceso en el que se registran las tareas de edición</param>
    ''' <remarks></remarks>
    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarRepresentante)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarZona)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarTipoArticulo)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarFamilia)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarComision)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarRepetidos)
    End Sub

    ''' <summary>
    ''' Comprobar que el representante exista
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarRepresentante(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IdRepresentante")) = 0 Then
            ApplicationService.GenerateError("El representante es obligatorio.")
        Else
            Dim DtAux As DataTable = New Representante().SelOnPrimaryKey(data("IdRepresentante"))
            If Not DtAux Is Nothing AndAlso DtAux.Rows.Count = 0 Then
                ApplicationService.GenerateError("El código del representante no existe")
            End If
        End If
    End Sub

    ''' <summary>
    ''' Comprobar que la zona exista
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarZona(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IdZona")) = 0 Then
            ApplicationService.GenerateError("La zona es obligatoria.")
        Else
            Dim DtAux As DataTable = New Zona().SelOnPrimaryKey(data("IDZona"))
            If DtAux Is Nothing OrElse DtAux.Rows.Count = 0 Then
                ApplicationService.GenerateError("El código de zona no existe")
            End If
        End If
    End Sub

    ''' <summary>
    ''' Comprobar que el tipo de artículo exista
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarTipoArticulo(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDTipo")) <> 0 Then
            Dim DtAux As DataTable = New TipoArticulo().SelOnPrimaryKey(data("IDTipo"))
            If Not DtAux Is Nothing AndAlso DtAux.Rows.Count = 0 Then
                ApplicationService.GenerateError("El Tipo del Artículo no existe en la Base de datos.")
            End If
        End If
    End Sub

    ''' <summary>
    ''' Comprobar que la familia de artículo exista
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarFamilia(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDFamilia")) <> 0 Then
            Dim DtAux As DataTable = New Familia().SelOnPrimaryKey(data("IDTipo"), data("IDFamilia"))
            If Not DtAux Is Nothing AndAlso DtAux.Rows.Count = 0 Then
                ApplicationService.GenerateError("La Familia del Artículo no existe en la Base de datos.")
            End If
        End If
    End Sub

    ''' <summary>
    ''' Comprobar que la comisión es correcta
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarComision(ByVal data As DataRow, ByVal services As ServiceProvider)
        ' Comisión
        If data("Comision") Is Nothing OrElse data("Comision").ToString.Trim.Length = 0 Then
            ApplicationService.GenerateError("El campo Comisión ha de ser numérico.")
        Else
            If Not data("Porcentaje") Is System.DBNull.Value AndAlso data("Porcentaje") = True Then
                If CDbl(data("Comision")) < 0 Or CDbl(data("Comision")) > 100 Then
                    ApplicationService.GenerateError("El campo Comisión es un porcentaje, introduzca valores entre 0 y 100.")
                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' Comprobar que no existe el registro
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarRepetidos(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim f As New Filter(FilterUnionOperator.And)
        f.Add(New StringFilterItem("IDZona", FilterOperator.Equal, data("IDZona")))
        f.Add(New StringFilterItem("IDRepresentante", FilterOperator.Equal, data("IDRepresentante")))
        If data("IDTipo").ToString.Trim.Length = 0 Then
            f.Add(New IsNullFilterItem("IDTipo", True))
        Else : f.Add(New StringFilterItem("IDTipo", FilterOperator.Equal, data("IDTipo")))
        End If
        If data("IDFamilia").ToString.Trim.Length = 0 Then
            f.Add(New IsNullFilterItem("IDFamilia", True))
        Else : f.Add(New StringFilterItem("IDFamilia", FilterOperator.Equal, data("IDFamilia")))
        End If
        If Not data.RowState = DataRowState.Added Then
            f.Add(New NumberFilterItem("IDZonaRepresentante", FilterOperator.NotEqual, data("IDZonaRepresentante")))
        End If

        Dim dtDatos As DataTable = New ZonaRepresentante().Filter(f)
        If Not dtDatos Is Nothing AndAlso dtDatos.Rows.Count > 0 Then
            ' Hay repetición
            If data("IDTipo").ToString.Trim.Length = 0 Then
                ApplicationService.GenerateError("Ya existe un registro con esa Zona.")
            Else
                If data("IDFamilia") Is Nothing OrElse data("IDFamilia").ToString.Trim.Length = 0 Then
                    ApplicationService.GenerateError("Ya existe un registro con esa Zona, Tipo y sin Familia.")
                Else : ApplicationService.GenerateError("Ya existe un registro con esa Zona, Tipo y Familia")
                End If
            End If
        End If
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    ''' <summary>
    ''' Relación de tareas asociadas a la modificación 
    ''' </summary>
    ''' <param name="updateProcess">Proceso en el que se registran las tareas de edición</param>
    ''' <remarks></remarks>
    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarPrimaryKey)
    End Sub

    ''' <summary>
    ''' Asignar la información por defecto
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub AsignarPrimaryKey(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Or data.RowState = DataRowState.Modified Then
            If Length(data("IdZonaRepresentante")) = 0 Then data("IdZonaRepresentante") = AdminData.GetAutoNumeric
        End If
    End Sub

#End Region

End Class