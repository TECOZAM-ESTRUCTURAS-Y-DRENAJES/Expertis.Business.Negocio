Public Class DatosEmpresa

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbDatosEmpresa"

#End Region

#Region "Funciones Register / Validate / Update"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDescEmpresa)
    End Sub

    <Task()> Public Shared Sub ValidarDescEmpresa(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("DescEmpresa")) = 0 Then ApplicationService.GenerateError("La descripción es un dato obligatorio.")
        If Length(data("DatosRegistrales")) = 0 Then ApplicationService.GenerateError("Los datos registrales son obligatorios.")
    End Sub



#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Function ObtenerDatosEmpresa(ByVal data As Object, ByVal services As ServiceProvider) As DatosEmpresaInfo
        Dim e As New DatosEmpresaInfo
        Dim dt As DataTable = New DatosEmpresa().Filter()
        If Not dt Is Nothing AndAlso dt.Rows.Count Then
            Dim oRw As DataRow = dt.Rows(0)
            e.DescEmpresa = dt.Rows(0)("DescEmpresa") & String.Empty
            e.CIF = dt.Rows(0)("CIF") & String.Empty
            e.Direccion = dt.Rows(0)("Direccion") & String.Empty
            e.CodPostal = dt.Rows(0)("CodPostal") & String.Empty
            e.Poblacion = dt.Rows(0)("Poblacion") & String.Empty
            e.Provincia = dt.Rows(0)("Provincia") & String.Empty
            e.Telefono = dt.Rows(0)("Telefono") & String.Empty
            e.Fax = dt.Rows(0)("Fax") & String.Empty
            e.Email = dt.Rows(0)("Email") & String.Empty
            e.DatosRegistrales = dt.Rows(0)("DatosRegistrales") & String.Empty
            e.IDPais = dt.Rows(0)("IDPais") & String.Empty
        End If
        Return e
    End Function

    <Task()> Public Shared Function ObtenerFondoExp(ByVal data As Object, ByVal services As ServiceProvider) As String
        Dim DtDB As DataTable = New BE.DataEngine().Filter("xDataBase", New GuidFilterItem("IDBaseDatos", AdminData.GetConnectionInfo.IDDataBase), , , , True)
        If Not DtDB Is Nothing AndAlso DtDB.Rows.Count > 0 Then
            Return Nz(DtDB.Rows(0)("Imagen"), String.Empty)
        End If
    End Function

    <Task()> Public Shared Function GrabarFondoExp(ByVal data As String, ByVal services As ServiceProvider) As Boolean
        Dim StrUpdate As String = "UPDATE xDataBase "
        StrUpdate &= IIf(Length(data) > 0, "SET Imagen = '" & data & "' ", "SET Imagen = NULL ")
        StrUpdate &= ", ModoImagen = 1 "
        StrUpdate &= "WHERE IDBaseDatos = '" & AdminData.GetConnectionInfo.IDDataBase.ToString & "'"
        Try
            AdminData.Execute(StrUpdate, ExecuteCommand.ExecuteNonQuery, True)
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function

#End Region

End Class

<Serializable()> _
Public Class DatosEmpresaInfo
    Public ID As String
    Public DescEmpresa As String
    Public CIF As String
    Public Direccion As String
    Public CodPostal As String
    Public Poblacion As String
    Public Provincia As String
    Public IDPais As String
    Public Telefono As String
    Public Fax As String
    Public Email As String
    Public DatosRegistrales As String
End Class