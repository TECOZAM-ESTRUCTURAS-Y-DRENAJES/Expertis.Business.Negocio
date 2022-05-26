Public Class ActivoEstructura

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbActivoEstructura"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarComponenteObligatorio)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarComponenteExistenteEnSistema)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarComponenteEnArbol)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarComponenteExistenteEnActivo)
    End Sub

    ''' <summary>
    ''' Método que valida si se ha indicado el componente
    ''' </summary>
    ''' <param name="data">Registro de ActivoEstructura</param>
    ''' <param name="services">Objeto para compartir información</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarComponenteObligatorio(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDActivoComponente")) = 0 Then ApplicationService.GenerateError("El Componente es obligatorio.")
    End Sub

    ''' <summary>
    ''' Método que valida si el componente existe en el sistema o si es Padre.
    ''' </summary>
    ''' <param name="data">Registro de ActivoEstructura</param>
    ''' <param name="services">Objeto para compartir información</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarComponenteExistenteEnSistema(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim f As New Filter
        f.Add(New StringFilterItem("IDActivo", data("IDActivoComponente")))
        f.Add(New BooleanFilterItem("Padre", False))

        Dim dt As DataTable = New Activo().Filter(f)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then
            ApplicationService.GenerateError("El Activo Componente no existe o es Padre.")
        End If
    End Sub

    ''' <summary>
    ''' Método que valida si el Componente forma parte del árbol del Activo.
    ''' </summary>
    ''' <param name="data">Registro de ActivoEstructura</param>
    ''' <param name="services">Objeto para compartir información</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarComponenteEnArbol(ByVal data As DataRow, ByVal services As ServiceProvider)
        '//Comprobamos que no forme parte del árbol anteriormente. 
        Dim datos As New DataExisteActivoEnExplosion
        datos.IDActivoHijo = data("IDActivoComponente")
        datos.IDActivoBase = data("IDActivo")
        datos.ValidarPadre = False
        If ProcessServer.ExecuteTask(Of DataExisteActivoEnExplosion, Boolean)(AddressOf ExisteActivoEnExplosion, datos, services) Then
            ApplicationService.GenerateError("El Activo Componente {0} ya forma parte del árbol.", Quoted(data("IDActivoComponente")))
        End If
    End Sub

    ''' <summary>
    ''' Método que valida si el Componente ya existe en el Activo.
    ''' </summary>
    ''' <param name="data">Registro de ActivoEstructura</param>
    ''' <param name="services">Objeto para compartir información</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarComponenteExistenteEnActivo(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim f As New Filter
            f.Add(New StringFilterItem("IDActivo", data("IDActivo")))
            f.Add(New StringFilterItem("IDActivoComponente", data("IDActivoComponente")))
            Dim dt As DataTable = New ActivoEstructura().Filter(f)
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("Componente duplicado para el Activo actual")
            End If
        End If
    End Sub

#End Region

#Region "Eventos GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBRL As New BusinessRules
        oBRL.Add("IDActivoComponente", AddressOf CambioActivoComponente)
        Return oBRL
    End Function

    ''' <summary>
    ''' Método que realiza las operaciones derivadas del cambio del Componente de un Activo
    ''' </summary>
    ''' <param name="data">Objeto BusinessRuleData</param>
    ''' <param name="services">Objeto para compartir información</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub CambioActivoComponente(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDActivoComponente")) > 0 Then
            Dim objNegActivo As New Activo
            Dim dr As DataRow = objNegActivo.GetItemRow(data.Current("IDActivoComponente"))
            If Not IsNothing(dr) Then
                If data.Current.ContainsKey("DescActivoComponente") Then data.Current("DescActivoComponente") = dr("DescActivo")
            Else
                If data.Current.ContainsKey("DescActivoComponente") Then data.Current("DescActivoComponente") = System.DBNull.Value
            End If
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Serializable()> _
    Public Class DataExisteActivoEnExplosion
        Public IDActivoHijo As String
        Public IDActivoBase As String
        Public ValidarPadre As Boolean

        Public Sub New()
        End Sub
        Public Sub New(ByVal IDActivoHijo As String, ByVal IDActivoBase As String, ByVal ValidarPadre As Boolean)
            Me.IDActivoHijo = IDActivoHijo
            Me.IDActivoBase = IDActivoBase
            Me.ValidarPadre = ValidarPadre
        End Sub
    End Class

    ''' <summary>
    ''' Método que valida si el Activo puede ser padre, o forma parte de otro Activo
    ''' </summary>
    ''' <param name="data">Estructura que contiene el ActivoBase, ActivoHijo y si se debe Validar el Padre.</param>
    ''' <param name="services">Objeto para compartir información</param>
    ''' <returns>Retorna un objeto de tipo Boolean que indica si Existe el Activo (IDActivoHijo) en la Explosión del Activo IDActivoBase.</returns>
    ''' <remarks>Validar si el Activo puede ser padre, o forma parte de otro activo.</remarks>
    <Task()> Public Shared Function ExisteActivoEnExplosion(ByVal data As DataExisteActivoEnExplosion, ByVal services As ServiceProvider) As Boolean
        If data.ValidarPadre Then
            '//Cuando se marca como padre. Se comprueba si existe el componente previamente.
            '//(Llamada desde el formulario - Mnto.Activos)
            Dim f As New Filter
            f.Add(New StringFilterItem("IDActivoComponente", data.IDActivoBase))
            Dim dtAE As DataTable = New ActivoEstructura().Filter(f)
            If Not dtAE Is Nothing AndAlso dtAE.Rows.Count > 0 Then
                Return True
            Else
                Return False
            End If
        Else
            '//Comprobamos que IDActivoHijo(Componente) no forme parte del árbol anteriormente (IDActivoBase)(Activo)). 
            '//(Validación en el Update)
            Dim dtBases As New DataTable
            dtBases.Columns.Add("ID", GetType(String))
            dtBases.PrimaryKey = New DataColumn() {dtBases.Columns("ID")}

            Dim drNew As DataRow = dtBases.NewRow
            drNew("ID") = data.IDActivoBase
            dtBases.Rows.Add(drNew)
            drNew = dtBases.NewRow
            drNew("ID") = data.IDActivoHijo
            dtBases.Rows.Add(drNew)


            Dim f As New Filter
            f.Add(New StringFilterItem("IDActivo", data.IDActivoHijo))

            While True
                Dim dtAE As DataTable = New ActivoEstructura().Filter(f)
                If dtAE Is Nothing OrElse dtAE.Rows.Count = 0 Then
                    Return False
                Else
                    If dtBases.Rows.Find(dtAE.Rows(0)("IDActivoComponente")) Is Nothing Then
                        drNew = dtBases.NewRow
                        drNew("ID") = dtAE.Rows(0)("IDActivoComponente")
                        dtBases.Rows.Add(drNew)

                        f.Clear()
                        f.Add(New StringFilterItem("IDActivo", dtAE.Rows(0)("IDActivoComponente")))
                    Else
                        Return True
                    End If
                End If
            End While
            Return False
        End If
    End Function

#End Region

End Class