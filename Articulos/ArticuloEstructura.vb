Public Class ArticuloEstructura
#Region "Constructor"
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbArticuloEstructura"
#End Region
#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarRuta)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDEstructura")) = 0 Then ApplicationService.GenerateError("El Identificador de Estructura es un dato obligatorio.")
        If Length(data("DescEstructura")) = 0 Then ApplicationService.GenerateError("La descripción de la estructura es un dato obligatorio.")
        If Length(data("IDArticulo")) = 0 Then ApplicationService.GenerateError("El Artículo es obligatorio.")
    End Sub
    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim FilAP As New Filter
            FilAP.Add("IDEstructura", FilterOperator.Equal, data("IDEstructura"), FilterType.String)
            FilAP.Add("IDArticulo", FilterOperator.Equal, data("IDArticulo"), FilterType.String)
            Dim DtAP As DataTable = New ArticuloEstructura().Filter(FilAP)
            If Not DtAP Is Nothing AndAlso DtAP.Rows.Count > 0 Then
                ApplicationService.GenerateError("La estructura ya existe en la lista actual.", data("IDArticulo"), data("IDEstructura"))
            End If
        End If
    End Sub
    <Task()> Public Shared Sub ValidarRuta(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDRuta")) > 0 Then
            Dim ClsAR As New ArticuloRuta
            Dim FilAR As New Filter
            FilAR.Add("IDArticulo", FilterOperator.Equal, data("IDArticulo"), FilterType.String)
            FilAR.Add("IDRuta", FilterOperator.Equal, data("IDRuta"), FilterType.String)
            Dim DtAR As DataTable = ClsAR.Filter(FilAR)
            If DtAR.Rows.Count = 0 Then
                ApplicationService.GenerateError("El artículo | no tiene asociada la ruta |.", data("IDArticulo"), data("IDRuta"))

            End If
        End If
    End Sub

#End Region
#Region "Eventos RegisterUpdateTask"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf TratarPrincipal)
    End Sub

    <Task()> Public Shared Sub TratarPrincipal(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim FilArtEst As New Filter
        FilArtEst.Add(New StringFilterItem("IDArticulo", data("IDArticulo")))
        FilArtEst.Add(New BooleanFilterItem("Principal", True))
        Dim DtPrincipal As DataTable = New ArticuloEstructura().Filter(FilArtEst)
        If IsNothing(DtPrincipal) OrElse DtPrincipal.Rows.Count = 0 Then
            data("Principal") = True
        Else
            If Nz(data("Principal"), False) Then
                If data("IDEstructura") <> DtPrincipal.Rows(0)("IDEstructura") Then
                    DtPrincipal.Rows(0)("Principal") = False
                    BusinessHelper.UpdateTable(DtPrincipal)
                End If
            ElseIf data.RowState = DataRowState.Modified AndAlso data("Principal") <> data("Principal", DataRowVersion.Original) AndAlso DtPrincipal.Rows.Count = 1 Then
                data("Principal") = True
            End If
        End If
    End Sub

#End Region
#Region "Eventos RegisterDeleteTask "
    ''' <summary>
    ''' Relación de tareas asociadas al proceso de borrado
    ''' </summary>
    ''' <param name="deleteProcess">Proceso en el que se registran las tareas de borrado</param>
    ''' <remarks></remarks>
    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf EliminarComponente)
    End Sub
    ''' <summary>
    ''' Borrado de artículos
    ''' </summary>
    ''' <param name="data">Registro del artículo estructura a borrar</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub EliminarComponente(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Not IsNothing(data) Then
            Dim datos As New Estructura.DatosElimComp
            datos.IDArticulo = data("IDArticulo")
            datos.IDEstructura = data("IDEstructura")
            ' datos.IDComponente = data("IDComponente")
            ProcessServer.ExecuteTask(Of Estructura.DatosElimComp)(AddressOf Estructura.EliminarComponente, datos, services)
        End If
    End Sub

#End Region

#Region "Funciones Públicas"
    <Task()> Public Shared Function EstructuraPpal(ByVal StrIDArticulo As String, ByVal services As ServiceProvider) As String
        Dim f As New Filter
        f.Add(New StringFilterItem("IDArticulo", StrIDArticulo))
        f.Add(New BooleanFilterItem("Principal", True))
        Dim DtEstructura As DataTable = New ArticuloEstructura().Filter(f)
        If Not DtEstructura Is Nothing AndAlso DtEstructura.Rows.Count > 0 Then
            EstructuraPpal = DtEstructura.Rows(0)("IDEstructura")
        End If
    End Function

    <Task()> Public Shared Sub EstablecerEstructuraPpal(ByVal data As Estructura.DatosElimComp, ByVal services As ServiceProvider)
        If Length(data.IDArticulo) > 0 AndAlso Length(data.IDEstructura) > 0 Then
            Dim ae As New ArticuloEstructura()
            Dim f As New Filter
            f.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
            f.Add(New StringFilterItem("IDEstructura", data.IDEstructura))

            Dim dtArticuloEstructura As DataTable = ae.Filter(f)
            If Not dtArticuloEstructura Is Nothing AndAlso dtArticuloEstructura.Rows.Count > 0 Then
                'Quitar la estructura principal actual
                f.Clear()
                f.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
                f.Add(New BooleanFilterItem("Principal", True))
                Dim dtEstructuraPpal As DataTable = ae.Filter(f)
                If Not dtEstructuraPpal Is Nothing AndAlso dtEstructuraPpal.Rows.Count > 0 Then
                    dtEstructuraPpal.Rows(0)("Principal") = False
                    ae.Update(dtEstructuraPpal)
                End If
                'Establecer la estructura principal
                dtArticuloEstructura.Rows(0)("Principal") = True
                ae.Update(dtArticuloEstructura)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CopiarEstructura(ByVal info As CopiarComponentesInfo, ByVal services As ServiceProvider)
        Dim strIDArticuloOrigen, strIDEstructuraOrigen, strIDArticuloDestino As String
        If Not IsNothing(info) Then
            strIDArticuloOrigen = info.IDArticuloOrigen
            strIDEstructuraOrigen = info.IDEstructuraOrigen
            strIDArticuloDestino = info.IDArticuloDestino
        End If

        If Length(strIDArticuloOrigen) > 0 AndAlso Length(strIDEstructuraOrigen) > 0 AndAlso Length(strIDArticuloDestino) > 0 Then
            'Obtener datos de la estructura origen
            Dim ae As New ArticuloEstructura()
            Dim dtArtEstOrigen As DataTable = ae.SelOnPrimaryKey(strIDArticuloOrigen, strIDEstructuraOrigen)
            If Not dtArtEstOrigen Is Nothing AndAlso dtArtEstOrigen.Rows.Count > 0 Then
                'Agregar cabecera de estructura
                Dim dtArtEstrucDest As DataTable = ae.AddNewForm()
                dtArtEstrucDest.Rows(0)("IDArticulo") = strIDArticuloDestino
                dtArtEstrucDest.Rows(0)("IDEstructura") = strIDEstructuraOrigen
                dtArtEstrucDest.Rows(0)("DescEstructura") = dtArtEstOrigen.Rows(0)("DescEstructura")
                dtArtEstrucDest.Rows(0)("FechaVigencia") = dtArtEstOrigen.Rows(0)("FechaVigencia")
                dtArtEstrucDest.Rows(0)("IDRuta") = dtArtEstOrigen.Rows(0)("IDRuta")
                dtArtEstrucDest = ae.Update(dtArtEstrucDest)

                'Agregar componentes
                Dim ClsEstruc As New Estructura
                Dim f As New Filter
                f.Add(New StringFilterItem("IDArticulo", strIDArticuloOrigen))
                f.Add(New StringFilterItem("IDEstructura", strIDEstructuraOrigen))
                Dim dt As DataTable = ClsEstruc.Filter(f)
                If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                    Dim StDatos As New Estructura.DatosCopiaComp
                    StDatos.Dt = dt
                    StDatos.IDArticuloDestino = strIDArticuloDestino
                    StDatos.IDEstructuraDestino = strIDEstructuraOrigen
                    ProcessServer.ExecuteTask(Of Estructura.DatosCopiaComp)(AddressOf Estructura.CopiarComponente, StDatos, services)
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub EliminarEstructura(ByVal data As Estructura.DatosElimComp, ByVal services As ServiceProvider)
        If Length(data.IDArticulo) > 0 AndAlso Length(data.IDEstructura) > 0 Then
            Dim ae As New ArticuloEstructura()
            'Eliminar componentes(detalles)
            Dim StDatos As New Estructura.DatosElimComp
            StDatos.IDArticulo = data.IDArticulo
            StDatos.IDEstructura = data.IDEstructura
            ProcessServer.ExecuteTask(Of Estructura.DatosElimComp)(AddressOf Estructura.EliminarComponente, StDatos, services)
            'Eliminar cabecera de estructura
            Dim f As New Filter
            f.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
            f.Add(New StringFilterItem("IDEstructura", data.IDEstructura))
            Dim dt As DataTable = ae.Filter(f)
            ae.Delete(dt)
        End If
    End Sub

    <Task()> Public Shared Function CalcularEstructuraExplosion(ByVal StrIdArticulo As String, ByVal services As ServiceProvider) As DataTable
        Return AdminData.Execute("sp_EstructuraExplosion", False, StrIdArticulo)
    End Function

#End Region

End Class