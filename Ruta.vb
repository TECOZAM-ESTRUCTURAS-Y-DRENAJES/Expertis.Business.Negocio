Public Class Ruta

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbRuta"

#End Region

#Region "Eventos RegisterAddNewTasks"

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf AsignarValoresPredeterminados)
    End Sub
    
    ''' <summary>
    ''' Asignar la información por defecto
    ''' </summary>
    ''' <param name="data">Registro Nuevo</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub AsignarValoresPredeterminados(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim UdTiempo As Integer = New Parametro().UdTiempoPred()
        data("UdTiempoPrep") = UdTiempo
        data("UdTiempoEjec") = UdTiempo
        data("UdTiempo") = UdTiempo
        data("TipoOperacion") = CInt(enumtrTipoOperacion.trInterna)
        data("FactorProduccion") = True
        data("ControlProduccion") = CInt(enumrControlProduccion.rSi)
    End Sub

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
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarOperacion)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarCentro)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarSecuencia)
    End Sub

    ''' <summary>
    ''' Comprobar que el código y la descripción no sean nulos
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data("IDOperacion").ToString.Trim.Length = 0 Or data("IDCentro").ToString.Trim.Length = 0 Or data("IDUdProduccion").ToString.Trim.Length = 0 Then
            ApplicationService.GenerateError("La Operación, el Centro y la Unidad de Produccion son datos obligatorios.")
        End If
    End Sub

    ''' <summary>
    ''' Comprobar que la operación es válida 
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarOperacion(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim DtAux As DataTable = New Operacion().SelOnPrimaryKey(data("IDOperacion"))
        If DtAux Is Nothing OrElse DtAux.Rows.Count = 0 Then
            ApplicationService.GenerateError("La Operación | no existe.", data("IDOperacion"))
        End If
    End Sub

    ''' <summary>
    ''' Comprobar que el centro es válido 
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarCentro(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim DtAux As DataTable = New Centro().SelOnPrimaryKey(data("IDCentro"))
        If DtAux Is Nothing OrElse DtAux.Rows.Count = 0 Then
            ApplicationService.GenerateError("El centro | no existe.", data("IDOperacion"))
        End If
    End Sub

    ''' <summary>
    ''' Comprobar que la secuencia es válida 
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarSecuencia(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data("Secuencia") Is System.DBNull.Value OrElse data("Secuencia") = 0 Then
            ApplicationService.GenerateError("La Secuencia no puede tomar valor cero.")
        End If
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClavePrimaria)
        updateProcess.AddTask(Of DataRow)(AddressOf Business.General.Comunes.UpdateEntityRow)
        updateProcess.AddTask(Of DataRow)(AddressOf Business.General.Comunes.MarcarComoActualizado)
        updateProcess.AddTask(Of DataRow)(AddressOf EliminarRutaParametros)
        updateProcess.AddTask(Of DataRow)(AddressOf ActualizarRutaParametros)
        updateProcess.AddTask(Of DataRow, DataTable)(AddressOf Reordenar)
    End Sub

    <Task()> Public Shared Sub AsignarClavePrimaria(ByVal data As datarow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then data("IDRutaOp") = AdminData.GetAutoNumeric
    End Sub

    ''' <summary>
    ''' Eliminar los parámetros de la operación
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub EliminarRutaParametros(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf RutaParametro.EliminarRutaParametros, data, services)
        End If
    End Sub

    ''' <summary>
    ''' Añadir los parámetros de la operación
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ActualizarRutaParametros(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf RutaParametro.ActualizarRutaParametros, data, services)
        End If
    End Sub

    ''' <summary>
    ''' Ordenar las operaciones
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Function Reordenar(ByVal data As DataRow, ByVal services As ServiceProvider) As DataTable
        Dim blnModificado As Boolean
        Dim intSecuencia As Short
        Dim lngSkip As Integer

        Dim FilRuta As New Filter
        If Not data.IsNull("IDArticulo") Then
            FilRuta.Add("IDArticulo", FilterOperator.Equal, data("IDArticulo"), FilterType.String)
        End If
        If Not data.IsNull("IDRuta") Then
            FilRuta.Add("IDRuta", FilterOperator.Equal, data("IDRuta"), FilterType.String)
        ElseIf data.Table.Columns.Contains("IDTipoRuta") AndAlso Length(data("IDTipoRuta")) > 0 Then
            FilRuta.Add("IDTipoRuta", FilterOperator.Equal, data("IDTipoRuta"), FilterType.String)
        End If
        FilRuta.Add("IDRutaOp", FilterOperator.NotEqual, data("IDRutaOp"), FilterType.String)
        Dim DtRuta As DataTable = New Ruta().Filter(FilRuta, "Secuencia")
        'Si no entra en el If mas interno, la funcion devuelve nothing por defecto
        If Not DtRuta Is Nothing AndAlso DtRuta.Rows.Count > 0 Then
            If data("Secuencia") <= DtRuta.Rows(DtRuta.Rows.Count - 1)("Secuencia") Then
                intSecuencia = data("Secuencia")
                lngSkip = 0
                For Each Dr As DataRow In DtRuta.Select
                    Dim DrDatos() As DataRow = DtRuta.Select("Secuencia=" & intSecuencia)
                    If DrDatos.Length > 0 Then
                        lngSkip = CInt(DrDatos.Length)
                        DrDatos(0)("Secuencia") += 1
                        intSecuencia = DrDatos(0)("Secuencia")
                        blnModificado = True
                    Else
                        Exit For
                    End If
                Next
                'todo ¿Se actualiza?
                If blnModificado Then Return DtRuta
            End If
        End If
    End Function

#End Region

End Class