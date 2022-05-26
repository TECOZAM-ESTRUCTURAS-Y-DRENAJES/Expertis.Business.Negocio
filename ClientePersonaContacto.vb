Public Class ClientePersonaContacto

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper
    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub
    Private Const cnEntidad As String = "tbClientePersonaContacto"

#End Region

#Region "Clases"

    <Serializable()> _
    Public Class DataNuevoContacto
        Public IDCliente As String
        Public Nombre As String
        Public Telefono1 As String
        Public Telefono2 As String
        Public Fax As String
        Public EMail As String
        Public IDCargo As String
    End Class

    Public Class DataNuevaPersonaContacto
        Public DrCon As DataRow
        Public IDCliente As String
        Public NomComp As String
    End Class

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

#Region "Eventos RegisterBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim Obrl As New BusinessRules
        Obrl.Add("IDPersonaContacto", AddressOf CambioPersonaContacto)
        Return Obrl
    End Function

    <Task()> Public Shared Sub CambioPersonaContacto(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ClientePersonaContacto.InfoPersonaContacto, data.Current, services)
    End Sub

    <Task()> Public Shared Sub InfoPersonaContacto(ByVal data As IPropertyAccessor, ByVal services As ServiceProvider)
        If Length(data("IDPersonaContacto")) > 0 Then
            Dim objNegPC As New PersonaContacto
            Dim dtPC As DataTable = objNegPC.SelOnPrimaryKey(data("IDPersonaContacto"))
            If Not IsNothing(dtPC) AndAlso dtPC.Rows.Count > 0 Then
                data("Nombre") = dtPC.Rows(0)("Nombre") & Space(1) & dtPC.Rows(0)("Apellidos")
                data("IDCargo") = dtPC.Rows(0)("IDCargo")
                data("AltaAutomatica") = False
            Else
                ApplicationService.GenerateError("El Contacto introducido no existe.")
            End If
        Else
            data("Nombre") = System.DBNull.Value
            data("IDCargo") = System.DBNull.Value
        End If
    End Sub

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarIDCliente)
    End Sub

    <Task()> Public Shared Sub ValidarIDCliente(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDCliente")) = 0 Then ApplicationService.GenerateError("El Cliente es un dato obligatorio.")
        End If
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarPrimaryKey)
        updateProcess.AddTask(Of DataRow)(AddressOf TratarPersonaPredeterminada)
    End Sub

    <Task()> Public Shared Sub AsignarPrimaryKey(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDPersona")) = 0 Then data("IDPersona") = AdminData.GetAutoNumeric
        End If
    End Sub

    <Task()> Public Shared Sub TratarPersonaPredeterminada(ByVal dr As DataRow, ByVal services As ServiceProvider)
        Dim ofilter As New Filter
        ofilter.Add(New StringFilterItem("IDCliente", dr("IDCliente")))
        ofilter.Add(New BooleanFilterItem("Predeterminada", True))
        Dim dtPersona As DataTable = New ClientePersonaContacto().Filter(ofilter)
        If dtPersona Is Nothing OrElse dtPersona.Rows.Count = 0 Then
            dr("Predeterminada") = True
        Else
            If Length(dr("Predeterminada")) = 0 Then dr("Predeterminada") = False
            If dr("Predeterminada") Then
                If dr("IDPersona") <> dtPersona.Rows(0)("IDPersona") Then
                    dtPersona.Rows(0)("Predeterminada") = False
                    BusinessHelper.UpdateTable(dtPersona)
                End If
            ElseIf dr.RowState = DataRowState.Modified AndAlso dr("Predeterminada") <> dr("Predeterminada", DataRowVersion.Original) Then
                dr("Predeterminada") = True
            End If
        End If
    End Sub
#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Function AltaNuevoContacto(ByVal data As DataNuevoContacto, ByVal services As ServiceProvider) As Integer
        '//Dar de alta el contacto en tbMaestroPersonaContacto y traer su IDPersonaContacto. Sólo con los datos necesarios para el contacto.
        Dim StDatos As New PersonaContacto.DatosAltaNuevoContacto
        StDatos.Nombre = data.Nombre
        StDatos.IDCargo = data.IDCargo
        Dim intIDPersonaContacto As Integer = ProcessServer.ExecuteTask(Of PersonaContacto.DatosAltaNuevoContacto, Integer)(AddressOf PersonaContacto.AltaNuevoContacto, StDatos, services)
        Dim ClsCliePerCon As New ClientePersonaContacto
        '//Si estamos en Obras, tb hay que darlo de alta en ClientePersonaContacto, con todos los datos.
        Dim dtCltePC As DataTable = ClsCliePerCon.AddNewForm()
        If Not IsNothing(dtCltePC) AndAlso dtCltePC.Rows.Count > 0 Then
            dtCltePC.Rows(0)("IDPersonaContacto") = intIDPersonaContacto
            dtCltePC.Rows(0)("Nombre") = data.Nombre
            dtCltePC.Rows(0)("IDCliente") = data.IDCliente
            dtCltePC.Rows(0)("Telefono1") = data.Telefono1
            dtCltePC.Rows(0)("Telefono2") = data.Telefono2
            dtCltePC.Rows(0)("Fax") = data.Fax
            dtCltePC.Rows(0)("Email") = data.EMail
            dtCltePC.Rows(0)("IDCargo") = data.IDCargo
            ClsCliePerCon.Update(dtCltePC)
        End If
        Return intIDPersonaContacto
    End Function

    <Task()> Public Shared Sub NuevaPersonaContacto(ByVal data As DataNuevaPersonaContacto, ByVal services As ServiceProvider)
        Dim DtNewPC As DataTable = New ClientePersonaContacto().AddNewForm
        If Not DtNewPC Is Nothing Then
            DtNewPC.Rows(0)("IdCliente") = data.IDCliente
            DtNewPC.Rows(0)("Email") = data.DrCon("Email")
            DtNewPC.Rows(0)("Fax") = data.DrCon("Fax")
            DtNewPC.Rows(0)("Telefono1") = data.DrCon("TelefonoDirecto")
            DtNewPC.Rows(0)("Telefono2") = data.DrCon("TelefonoMovil")
            DtNewPC.Rows(0)("IDCargo") = data.DrCon("IDCargo")
            DtNewPC.Rows(0)("IDPersona") = data.DrCon("IDPersona")
            DtNewPC.Rows(0)("Nombre") = data.NomComp
            BusinessHelper.UpdateTable(DtNewPC)
        End If
    End Sub

#End Region

End Class