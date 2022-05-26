Public Class UdMedida

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroUdMedida"

#End Region

#Region "Eventos RegisterValidateTasks"

    ''' <summary>
    ''' Relaci�n de tareas asociadas a la validaci�n 
    ''' </summary>
    ''' <param name="validateProcess">Proceso en el que se registran las tareas de edici�n</param>
    ''' <remarks></remarks>
    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClavePrimaria)
    End Sub

    ''' <summary>
    ''' Comprobar que el c�digo y la descripci�n no sean nulos
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Informaci�n compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDUdMedida")) = 0 Then ApplicationService.GenerateError("El c�digo de Medida es obligatorio.")
    End Sub

    ''' <summary>
    ''' Comprobar que no exista la clave
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Informaci�n compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = New UdMedida().SelOnPrimaryKey(data("IDUdMedida"))
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("El identificativo de medida ya existe en la Base de Datos")
            End If
        End If
    End Sub

#End Region

#Region "Funciones P�blicas"

    <Task()> Public Shared Function ValidaUDMedida(ByVal strIDUdMedida As String, ByVal services As ServiceProvider) As DataTable
        Dim dt As DataTable = New UdMedida().SelOnPrimaryKey(strIDUdMedida)
        If IsNothing(dt) OrElse dt.Rows.Count = 0 Then
            ApplicationService.GenerateError("La | no existe.", strIDUdMedida)
        End If

        Return dt
    End Function

    <Task()> Public Shared Function CargarMedidasAB(ByVal FilForm As Filter, ByVal services As ServiceProvider) As DataTable
        'Primero cargamos los datos de la vista cuyos articulos tengan configurado Articulos AB
        Dim DtArtAB As DataTable = New BE.DataEngine().Filter("vfrmCIStockUD", FilForm)
        If Not DtArtAB Is Nothing AndAlso DtArtAB.Rows.Count > 0 Then
            Dim StrArt(DtArtAB.Rows.Count - 1) As String
            Dim i As Integer = 0
            For Each Dr As DataRow In DtArtAB.Select("", "IDArticulo")
                StrArt(i) = Dr("IDArticulo")
                i += 1
            Next
            FilForm.Add(New InListFilterItem("IDArticulo", StrArt, FilterType.String, False))
            Dim DtArtConvAB As DataTable = New BE.DataEngine().Filter("vFrmCIStockConverUD", FilForm)
            If Not DtArtConvAB Is Nothing AndAlso DtArtConvAB.Rows.Count > 0 Then
                For Each DrConv As DataRow In DtArtConvAB.Select("", "IDArticulo")
                    DtArtAB.Rows.Add(DrConv.ItemArray)
                Next
                DtArtAB.AcceptChanges()
            End If
            Return DtArtAB
        Else
            Dim DtArtConvAB As DataTable = New BE.DataEngine().Filter("vFrmCIStockConverUD", FilForm)
            If Not DtArtConvAB Is Nothing AndAlso DtArtConvAB.Rows.Count > 0 Then
                Return DtArtConvAB
            End If
        End If
    End Function

#End Region

End Class