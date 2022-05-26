Public Class DiaPago

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroDiaPago"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("DescDiaPago")) = 0 Then ApplicationService.GenerateError("Introduzca la descripción del registro de días de pago")
        If (Length(data("Dia1")) = 0 OrElse data("Dia1") = 0) AndAlso (((Not Length(data("Dia2")) = 0 AndAlso data("Dia2") <> 0) OrElse (Not Length(data("Dia3")) = 0 AndAlso data("Dia3") <> 0))) Then
            ApplicationService.GenerateError("No se puede dar valor a días 2 y 3 sin tener día 1")
        End If
        If (Length(data("Dia2")) = 0 OrElse data("Dia2") = 0) AndAlso (Not Length(data("Dia3")) = 0 AndAlso data("Dia3") <> 0) Then
            ApplicationService.GenerateError("No se puede dar valor a día 3 sin tener día 2")
        End If
        If Not Length(data("Dia1")) = 0 AndAlso (data("Dia1") < 0 Or data("Dia1") > 31) Then
            ApplicationService.GenerateError("El día de pago 1 debe estar comprendido entre el 0 y el 31")
        ElseIf Not Length(data("Dia2")) = 0 AndAlso (data("Dia2") < 0 Or data("Dia2") > 31) Then
            ApplicationService.GenerateError("El día de pago 2 debe estar comprendido entre el 0 y el 31")
        ElseIf Not Length(data("Dia3")) = 0 AndAlso (data("Dia3") < 0 Or data("Dia3") > 31) Then
            ApplicationService.GenerateError("El día de pago 3 debe estar comprendido entre el 0 y el 31")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDDiaPago")) > 0 Then
                Dim DtTemp As DataTable = New DiaPago().SelOnPrimaryKey(data("IDDiaPago"))
                If Not DtTemp Is Nothing AndAlso DtTemp.Rows.Count > 0 Then
                    ApplicationService.GenerateError("El id. introducido ya existe.")
                End If
            Else : ApplicationService.GenerateError("Introduzca el código del día de pago.")
            End If
        End If
    End Sub

#End Region

End Class