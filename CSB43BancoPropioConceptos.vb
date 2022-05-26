Public Class CSB43BancoPropioConceptos

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbCSB43BancoPropioConceptos"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarCContableContrapartida)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarCentroGestion)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarConcepto)
    End Sub

    <Task()> Public Shared Sub ValidarCContableContrapartida(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("CContableContrapartida")) = 0 Then ApplicationService.GenerateError("Debe introducir la cuenta contable de cada concepto")
    End Sub

    <Task()> Public Shared Sub ValidarCentroGestion(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDCentroGestion")) > 0 Then
            Dim DtExternos As DataTable = New CentroGestion().SelOnPrimaryKey(data("IDCentroGestion"))
            If DtExternos Is Nothing OrElse DtExternos.Rows.Count = 0 Then
                ApplicationService.GenerateError("El centro de gestión elegido no existe en la base de datos")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As datarow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDConceptoComun")) = 0 OrElse Length(data("IDConceptoPropio")) = 0 Then
                ApplicationService.GenerateError("Los campos concepto común y concepto propio son obligatorios")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarConcepto(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim DtTemp As DataTable = New CSB43BancoPropioConceptos().SelOnPrimaryKey(data)
            If Not DtTemp Is Nothing AndAlso DtTemp.Rows.Count > 0 Then
                ApplicationService.GenerateError("Ya existe un concepto con las mismas claves de concepto común y concepto propio")
            End If
        End If
    End Sub

#End Region

End Class