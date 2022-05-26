Public Class UsuarioCentroGestion

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbUsuarioCentroGestion"

#End Region

#Region "Eventos RegisterValidateTasks"

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
    ''' Comprobar que exista el ususario y que no esté asignado a otro centro de gestión
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = New UsuarioCentroGestion().SelOnPrimaryKey(data("IDUsuario"))
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("El Usuario ya está asignado a un Centro de Gestión")
            Else
                'Comprobamos que el usuario elegido exista
                Dim f As New Filter
                f.Add(New GuidFilterItem("IDUsuario", FilterOperator.Equal, data("IDUsuario")))
                Dim dtOperario As DataTable = ProcessServer.ExecuteTask(Of Filter, DataTable)(AddressOf DatosSistema.DevuelveUsuariosBD, f, services)
                If dtOperario Is Nothing OrElse dtOperario.Rows.Count = 0 Then
                    ApplicationService.GenerateError("Debe introducir uno de los usuarios de la lista o dejarlo vacío.")
                End If
            End If
        End If

    End Sub

#End Region

#Region "Funciones Publicas"

    <Serializable()> _
    Public Class UsuarioCentroGestionInfo
        Public IDOperario As String
        Public IDCentroGestion As String
        Public gIDUsuario As Guid
        Public NoAsignarPredeterminado As Boolean = False 'Para Alquiler
    End Class

    <Task()> Public Shared Function ObtenerUsuarioCentroGestion(ByVal ucg As UsuarioCentroGestionInfo, ByVal services As ServiceProvider) As UsuarioCentroGestionInfo
        If IsNothing(ucg) Then ucg = New UsuarioCentroGestionInfo()
        If ucg.gIDUsuario = New Guid AndAlso Length(ucg.IDOperario) = 0 Then ucg.gIDUsuario = NegocioGeneral.UserID()
        If Length(ucg.IDOperario) <> 0 Then
            Dim o As New Operario
            Dim dtO As DataTable = o.SelOnPrimaryKey(ucg.IDOperario)
            If Not IsNothing(dtO) AndAlso dtO.Rows.Count > 0 Then
                If Length(dtO.Rows(0)("IDUsuario")) > 0 Then
                    ucg.gIDUsuario = dtO.Rows(0)("IDUsuario")
                End If
            End If
        End If
        If Length(ucg.gIDUsuario) <> 0 Then
            'Retorna el Centro de Gestión del usuario que se envíe como parámetro real de la función.
            'En caso de que usuario no tenga un Centro de Gestión Asociado, retorna el Centro de Gestión
            'Predeterminado (de tbParametro)
            Dim dt As DataTable = New UsuarioCentroGestion().SelOnPrimaryKey(ucg.gIDUsuario)
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                ucg.IDCentroGestion = dt.Rows(0)("IdCentroGestion")
            End If
        End If
        If Not ucg.NoAsignarPredeterminado Then
            If Length(ucg.IDCentroGestion) = 0 Then
                ucg.IDCentroGestion = New Parametro().CGestionPredet
            End If
        End If
        Return ucg
    End Function

#End Region

    Public Function CentroGestionUsuario() As String
        CentroGestionUsuario = CentroGestionUsuario(NegocioGeneral.UserID)
    End Function

    Public Function CentroGestionUsuario(ByVal gIDUsuario As Guid) As String
        Dim strCGestion As String

        If Length(gIDUsuario.ToString) > 0 Then
            'Retorna el Centro de Gestión del usuario que se envíe como parámetro real de la función.
            'En caso de que usuario no tenga un Centro de Gestión Asociado, retorna el Centro de Gestión
            'Predeterminado (de tbParametro)
            Dim dt As DataTable = SelOnPrimaryKey(gIDUsuario)
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                strCGestion = dt.Rows(0)("IdCentroGestion")
            End If
        End If

        If Length(strCGestion) = 0 Then
            Dim p As New Parametro
            strCGestion = p.CGestionPredet
        End If

        Return strCGestion
    End Function

    Public Sub CentroGestionUsuario(ByVal strIDOperario As String, ByRef strCentroGestion As String)
        If Length(strIDOperario) > 0 Then
            Dim o As New Operario
            Dim dtO As DataTable = o.SelOnPrimaryKey(strIDOperario)
            If Not IsNothing(dtO) AndAlso dtO.Rows.Count > 0 Then
                If Length(dtO.Rows(0)("IDUsuario")) > 0 Then
                    Dim gIDUsuario As Guid = dtO.Rows(0)("IDUsuario")
                    strCentroGestion = CentroGestionUsuario(gIDUsuario)
                End If
            End If
        End If
    End Sub


End Class