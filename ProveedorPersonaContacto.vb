Public Class ProveedorPersonaContacto

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbProveedorPersonaContacto"

#End Region

#Region "Eventos RegisterAddNewTasks"

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        data("IDPersona") = AdminData.GetAutoNumeric
    End Sub

#End Region

#Region "Eventos RegisterValidateTasks "

    ''' <summary>
    ''' Relación de tareas asociadas a la validación 
    ''' </summary>
    ''' <param name="validateProcess">Proceso en el que se registran las tareas de edición</param>
    ''' <remarks></remarks>
    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    ''' <summary>
    ''' Comprobar que el código y la descripción no sean nulos
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDProveedor")) = 0 Then ApplicationService.GenerateError("El Proveedor es un dato obligatorio.")
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    ''' <summary>
    ''' Relación de tareas asociadas a la modificación 
    ''' </summary>
    ''' <param name="updateProcess">Proceso en el que se registran las tareas de edición</param>
    ''' <remarks></remarks>
    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarPrimaryKey)
        updateProcess.AddTask(Of DataRow)(AddressOf TratarPersonaPredeterminada)
    End Sub

    ''' <summary>
    ''' Asignar la información por defecto
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub AsignarPrimaryKey(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDPersona")) = 0 Then data("IDPersona") = AdminData.GetAutoNumeric
        End If
    End Sub

#End Region

#Region "Funciones Privadas/Públicas"

    <Serializable()> _
    Public Class DatosPersonaNuevoContacto
        Public Dt As DataTable
        Public IDProveedor As String
        Public Nombre As String
    End Class

    ''' <summary>
    ''' Establecer la persona predeterminada
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub TratarPersonaPredeterminada(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim ofilter As New Filter
        ofilter.Add(New StringFilterItem("IDProveedor", data("IDProveedor")))
        ofilter.Add(New BooleanFilterItem("Predeterminada", True))
        Dim dtPersona As DataTable = New ProveedorPersonaContacto().Filter(ofilter)
        If dtPersona Is Nothing OrElse dtPersona.Rows.Count = 0 Then
            data("Predeterminada") = True
        Else
            If Length(data("Predeterminada")) = 0 Then data("Predeterminada") = False
            If data("Predeterminada") Then
                If data("IDPersona") <> dtPersona.Rows(0)("IDPersona") Then
                    dtPersona.Rows(0)("Predeterminada") = False
                    BusinessHelper.UpdateTable(dtPersona)
                End If
            ElseIf data.RowState = DataRowState.Modified AndAlso data("Predeterminada") <> data("Predeterminada", DataRowVersion.Original) Then
                data("Predeterminada") = True
            End If
        End If
    End Sub

    <Task()> Public Shared Sub NuevaPersonaContacto(ByVal data As DatosPersonaNuevoContacto, ByVal services As ServiceProvider)
        If Not data.Dt Is Nothing AndAlso data.Dt.Rows.Count > 0 Then
            Dim dtNewPC As DataTable = New ProveedorPersonaContacto().AddNewForm()
            dtNewPC.Rows(0)("IdProveedor") = data.IDProveedor
            dtNewPC.Rows(0)("Email") = data.Dt.Rows(0)("Email")
            dtNewPC.Rows(0)("Telefono1") = data.Dt.Rows(0)("TelefonoDirecto")
            dtNewPC.Rows(0)("Telefono2") = data.Dt.Rows(0)("TelefonoMovil")
            dtNewPC.Rows(0)("Fax") = data.Dt.Rows(0)("Fax")
            dtNewPC.Rows(0)("IDCargo") = data.Dt.Rows(0)("IDCargo")
            dtNewPC.Rows(0)("IDPersona") = data.Dt.Rows(0)("IDPersona")
            dtNewPC.Rows(0)("Nombre") = data.Nombre
            BusinessHelper.UpdateTable(dtNewPC)
        End If
    End Sub

#End Region

End Class