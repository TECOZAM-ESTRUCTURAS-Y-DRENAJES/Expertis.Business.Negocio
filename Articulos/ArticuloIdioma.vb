Public Class ArticuloIdioma

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbArticuloIdioma"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDArticulo")) = 0 Then ApplicationService.GenerateError("El Artículo es obligatorio.")
        If Length(data("IDIdioma")) = 0 Then ApplicationService.GenerateError("El idioma es un dato obligatorio.") ' '
        If Length(data("DescArticuloIdioma")) = 0 Then ApplicationService.GenerateError("La descripción del idioma es un dato obligatorio.") ' '

    End Sub
    
    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As datarow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = New ArticuloIdioma().SelOnPrimaryKey(data("IDArticulo"), data("IDIdioma"))
            If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("El Idioma '|' ya está asignado en este artículo.", data("IDIdioma"))
            End If
        End If
    End Sub

#End Region

End Class