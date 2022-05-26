Public Class PersonaContacto

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroPersonaContacto"

#End Region

#Region "Eventos RegisterAddNewTasks"

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        data("IDPersonaContacto") = AdminData.GetAutoNumeric
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClave)
    End Sub

    <Task()> Public Shared Sub AsignarClave(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDPersonaContacto")) > 0 Then data("IDPersonaContacto") = AdminData.GetAutoNumeric
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Serializable()> _
    Public Class DatosAltaNuevoContacto
        Public Nombre As String
        Public IDCargo As String
    End Class

    <Task()> Public Shared Function AltaNuevoContacto(ByVal data As DatosAltaNuevoContacto, ByVal services As ServiceProvider) As Integer
        '//Dar de alta el contacto en tbMaestroPersonaContacto y traer su IDPersonaContacto. Sólo con los datos necesarios para el contacto.
        Dim dtContacto As DataTable = New PersonaContacto().AddNewForm()
        If Not IsNothing(dtContacto) AndAlso dtContacto.Rows.Count > 0 Then
            dtContacto.Rows(0)("Nombre") = data.Nombre
            dtContacto.Rows(0)("IDCargo") = data.IDCargo
            PersonaContacto.UpdateTable(dtContacto)

            Return dtContacto.Rows(0)("IDPersonaContacto")
        End If
        Return 0
    End Function

#End Region

End Class