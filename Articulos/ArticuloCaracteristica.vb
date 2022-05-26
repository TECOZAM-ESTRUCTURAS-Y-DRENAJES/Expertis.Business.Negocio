Public Class ArticuloCaracteristica
#Region "Constructor"
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbArticuloCaracteristica"
#End Region
#Region "RegisterValidateTasks"
    ''' <summary>
    ''' Relación de tareas asociadas a la validación 
    ''' </summary>
    ''' <param name="validateProcess">Proceso en el que se registran las tareas de edición</param>
    ''' <remarks></remarks>
    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidaDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidaArticuloCaracteristica)
    End Sub
    ''' <summary>
    ''' Comprobar que existan el artículo y la característica
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidaDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDArticulo")) = 0 Then ApplicationService.GenerateError("El Artículo es obligatorio.")
        If Length(data("IDCaracteristica")) = 0 Then ApplicationService.GenerateError("El código de la Característica es obligatorio")
    End Sub
    ''' <summary>
    ''' Comprobar que el artículo no tenga ya esa caracterçistica asignada
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidaArticuloCaracteristica(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = New ArticuloCaracteristica().SelOnPrimaryKey(data("IDArticulo"), data("IDCaracteristica"))
            If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("La característica '|' ya está asignada en este artículo.", data("IDCaracteristica"))
            End If
        End If
    End Sub

#End Region
#Region "Funciones Públicas"
    ''' <summary>
    ''' Generar los características asociadas a un artículo
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks>Genera los características asociadas a ese tipo y familia</remarks>
    <Task()> Public Shared Sub AddArticuloCaracteristica(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDArticulo")) > 0 AndAlso Length(data("IDTipo")) > 0 AndAlso Length(data("IDFamilia")) > 0 Then
            Dim ac As New ArticuloCaracteristica
            Dim objFilter As New Filter
            objFilter.Add(New StringFilterItem("IDArticulo", data("IDArticulo")))

            Dim dt As DataTable = ac.Filter(objFilter)

            ac.Delete(dt)

            Dim fc As New FamiliaCaracteristica
            objFilter.Clear()
            objFilter.Add(New StringFilterItem("IDTipo", data("IDTipo")))
            objFilter.Add(New StringFilterItem("IDFamilia", data("IDFamilia")))
            Dim dtFC As DataTable = fc.Filter(objFilter)

            If Not dtFC Is Nothing AndAlso dtFC.Rows.Count > 0 Then
                dt = ac.AddNew
                For Each drFC As DataRow In dtFC.Rows
                    Dim drNewRow As DataRow = dt.NewRow
                    drNewRow("IDArticulo") = data("IDArticulo")
                    drNewRow("IDCaracteristica") = drFC("IDCaracteristica")
                    dt.Rows.Add(drNewRow)
                Next

                ac.Update(dt)
            End If
        End If
    End Sub
#End Region

End Class