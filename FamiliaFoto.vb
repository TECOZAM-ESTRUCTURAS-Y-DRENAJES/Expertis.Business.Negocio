Public Class FamiliaFoto

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbFamiliaFoto"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDTipo")) = 0 Then ApplicationService.GenerateError("El Tipo Artículo no puede ser vacío.")
        If Length(data("IDFamilia")) = 0 Then ApplicationService.GenerateError("El campo Familia no puede ser vacío.")
        If IsDBNull(data("Foto")) OrElse data("Foto") Is Nothing OrElse CType(data("Foto"), Array).Length = 0 Then
            ApplicationService.GenerateError("Debe incluir una foto.")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = New FamiliaFoto().SelOnPrimaryKey(data("IDTipo"), data("IDFamilia"))
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("Ya existe una foto para la Familia {0}.", Quoted(data("IDFamilia")))
            End If
        End If
    End Sub

#End Region

End Class