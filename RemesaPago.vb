Public Class RemesaPago

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbRemesaPago"

#End Region

#Region "Eventos RegisterAddNewTasks"

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf GetIDRemesa)
    End Sub

    <Task()> Public Shared Sub GetIDRemesa(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim StrIDContador As String = New Parametro().RemesaPago
        Dim DtContador As DataTable = New Contador().SelOnPrimaryKey(StrIDContador)
        If Not DtContador Is Nothing AndAlso DtContador.Rows.Count > 0 Then
            If DtContador.Rows(0)("Numerico") Then
                data("IDRemesa") = DtContador.Rows(0)("Contador")
            End If
        End If
    End Sub

#End Region

#Region "Eventos RegisterDeleteTasks"

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf ValidarDeleteRemesa)
        deleteProcess.AddTask(Of DataRow)(AddressOf ActualizarPago)
    End Sub

    <Task()> Public Shared Sub ValidarDeleteRemesa(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim AppConta As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
        If Not AppConta.Contabilidad Then Exit Sub

        Dim f As New Filter
        f.Add(New NumberFilterItem("IDRemesa", data("IDRemesa")))
        Dim fFilterOr As New Filter(FilterUnionOperator.Or)
        fFilterOr.Add(New BooleanFilterItem("GeneradoAsientoRemesa", True))
        fFilterOr.Add(New BooleanFilterItem("Contabilizado", True))
        f.Add(fFilterOr)
        Dim dtPago As DataTable = New Pago().Filter(f)
        If Not dtPago Is Nothing AndAlso dtPago.Rows.Count > 0 Then
            ApplicationService.GenerateError("No se puede eliminar una Remesa contabilizada o con Pagos contabilizados.")
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarPago(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim f As New Filter
        f.Add(New NumberFilterItem("IDRemesa", data("IDRemesa")))
        f.Add(New BooleanFilterItem("GeneradoAsientoRemesa", False))
        Dim ClsPago As New Pago
        Dim dtPago As DataTable = ClsPago.Filter(f)
        If Not dtPago Is Nothing AndAlso dtPago.Rows.Count > 0 Then
            For Each dr As DataRow In dtPago.Rows
                dr("IdRemesa") = System.DBNull.Value
                dr("Situacion") = enumPagoSituacion.NoPagado
            Next
            ClsPago.Update(dtPago)
        End If
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClave)
    End Sub

    <Task()> Public Shared Sub AsignarClave(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            data("IdRemesa") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, New Parametro().RemesaPago, services)
        End If
        If data("IdRemesa") = -1 Then data("IdRemesa") = DBNull.Value
    End Sub

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Sub AddRemesa(ByVal data As Hashtable, ByVal services As ServiceProvider)
        If Not IsNothing(data) AndAlso data.Count > 0 Then
            ProcessServer.ExecuteTask(Of Object)(AddressOf Business.General.Comunes.BeginTransaction, Nothing, services)
            Dim fPagosRemesar As New Filter
            fPagosRemesar.Add(New GuidFilterItem("IDProcess", data("IdProcess")))
            fPagosRemesar.Add(New IsNullFilterItem("IDRemesa"))
            Dim dtPagosRemesar As DataTable = New BE.DataEngine().Filter("frmPagosNeg", fPagosRemesar)
            If dtPagosRemesar Is Nothing OrElse dtPagosRemesar.Rows.Count = 0 Then
                ApplicationService.GenerateError("No hay pagos a remesar. No se generará la remesa.")
            Else
                Dim dtNew As DataTable = New RemesaPago().AddNewForm
                dtNew.Rows(0)("FechaRemesa") = data("FechaRemesa")
                dtNew.Rows(0)("TipoFichero") = data("TipoFichero")
                dtNew.Rows(0)("Ruta") = data("Ruta")
                'Dim blnImpreso As Boolean = (Length(data("Ruta")) > 0)
                dtNew = New RemesaPago().Update(dtNew)
                Dim AsocPagoRem As New Pago.DataAsociarPagoARemesa
                AsocPagoRem.IDProcess = data("IdProcess")
                AsocPagoRem.IDRemesa = dtNew.Rows(0)("IDRemesa")
                AsocPagoRem.IDBancoPropio = data("IDBancoPropio")
                'AsocPagoRem.Impreso = blnImpreso
                AsocPagoRem.PagosRemesar = dtPagosRemesar
                If data.ContainsKey("NuevaSituacion") AndAlso Length(data("NuevaSituacion")) > 0 Then
                    AsocPagoRem.NuevaSituacion = data("NuevaSituacion")
                Else
                    AsocPagoRem.NuevaSituacion = -1
                End If
                ProcessServer.ExecuteTask(Of Pago.DataAsociarPagoARemesa)(AddressOf Pago.AsociarPagoARemesa, AsocPagoRem, services)
            End If
        End If

    End Sub

#End Region

End Class