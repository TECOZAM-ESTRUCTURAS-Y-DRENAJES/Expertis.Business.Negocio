Public Class Concepto

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroConcepto"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("DescConcepto")) = 0 Then ApplicationService.GenerateError("La Descripción del Concepto es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDConcepto")) > 0 Then
                Dim DtTemp As DataTable = New Concepto().SelOnPrimaryKey(data("IDConcepto"))
                If Not DtTemp Is Nothing AndAlso DtTemp.Rows.Count > 0 Then
                    ApplicationService.GenerateError("El Concepto está duplicado")
                End If
            Else : ApplicationService.GenerateError("El Concepto es un dato obligatorio")
            End If
        End If
    End Sub

#End Region

End Class