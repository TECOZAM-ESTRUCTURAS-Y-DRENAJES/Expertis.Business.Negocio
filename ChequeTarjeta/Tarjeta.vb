Public Class Tarjeta
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

#Region " Constructor "

    Private Const cnEntidad As String = "tbTarjeta"

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

#End Region

#Region " RegisterAddnewTasks "

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf AsignarIdentificador)
    End Sub

#End Region

#Region " RegisterValidateTaks "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClavePrimaria)
    End Sub


    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("Fecha")) = 0 Then ApplicationService.GenerateError("La Fecha es un dato obligatorio.")
        If Length(data("NumeroTarjeta")) = 0 Then ApplicationService.GenerateError("El Número de Tarjeta es un dato obligatorio.")
        If Length(data("IDBancoPropio")) = 0 Then ApplicationService.GenerateError("La Caja es un dato obligatorio.")
    End Sub


    <Task()> Public Shared Sub ValidarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = New Tarjeta().SelOnPrimaryKey(data("IDTarjeta"))
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("El registro introducido ya existe.")
            End If
        End If
    End Sub

#End Region

#Region " RegisterUpdateTaks "

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarIdentificador)
    End Sub

    <Task()> Public Shared Sub AsignarIdentificador(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDTarjeta")) = 0 Then data("IDTarjeta") = AdminData.GetAutoNumeric
    End Sub

#End Region



#Region " Gestión de Tarjeta "

    <Serializable()> _
    Public Class dataAddInformacionTarjeta
        Public IDTarjeta As Integer?
        Public IDCobros(-1) As Integer
        Public NumeroTarjeta As String
        Public Voucher As String    ' Nº Transferencia
        Public IDBancoTarjeta As String
        Public Fecha As Date
        Public IDBancoPropio As String

        Public Sub New(ByVal IDCobros As Integer(), ByVal Fecha As Date, ByVal NumeroTarjeta As String, ByVal Voucher As String, ByVal IDBancoPropio As String, ByVal IDBancoTarjeta As String)
            Me.IDCobros = IDCobros
            Me.Fecha = Fecha
            Me.Voucher = Voucher
            Me.NumeroTarjeta = NumeroTarjeta
            Me.IDBancoPropio = IDBancoPropio
            Me.IDBancoTarjeta = IDBancoTarjeta
        End Sub

        Public Sub New(ByVal IDTarjeta As Integer, ByVal Fecha As Date, ByVal NumeroTarjeta As String, ByVal Voucher As String, ByVal IDBancoPropio As String, ByVal IDBancoTarjeta As String)
            Me.IDTarjeta = IDTarjeta
            Me.Fecha = Fecha
            Me.Voucher = Voucher
            Me.NumeroTarjeta = NumeroTarjeta
            Me.IDBancoPropio = IDBancoPropio
            Me.IDBancoTarjeta = IDBancoTarjeta
        End Sub
    End Class

    <Task()> Public Shared Function AddInformacionTarjeta(ByVal data As dataAddInformacionTarjeta, ByVal services As ServiceProvider) As Integer
        AddInformacionTarjeta = -1
        Dim blnNewTarjeta As Boolean
        Dim Tar As New Tarjeta
        Dim c As New Cobro
        Dim dtTarjeta As DataTable
        Dim dtCobrosTarjeta As DataTable
        If Not data.IDTarjeta Is Nothing Then
            '//Modificando Tarjeta
            dtTarjeta = Tar.SelOnPrimaryKey(data.IDTarjeta)
            dtCobrosTarjeta = c.Filter(New NumberFilterItem("IDTarjeta", data.IDTarjeta))
        ElseIf Not data.IDCobros Is Nothing AndAlso data.IDCobros.Count > 0 Then
            '//Creando Tarjeta
            Dim IdCobroObj(data.IDCobros.Length - 1) As Object
            data.IDCobros.CopyTo(IdCobroObj, 0)
            dtTarjeta = Tar.AddNewForm
            blnNewTarjeta = True
            dtCobrosTarjeta = c.Filter(New InListFilterItem("IDCobro", IdCobroObj, FilterType.Numeric))
        End If

        If dtTarjeta.Rows.Count > 0 Then
            dtTarjeta.Rows(0)("NumeroTarjeta") = data.NumeroTarjeta
            dtTarjeta.Rows(0)("Fecha") = data.Fecha
            If Length(data.Voucher) > 0 Then
                dtTarjeta.Rows(0)("Voucher") = data.Voucher
            Else
                dtTarjeta.Rows(0)("Voucher") = System.DBNull.Value
            End If
            dtTarjeta.Rows(0)("IDBancoPropio") = data.IDBancoPropio
            dtTarjeta.Rows(0)("IDBancoTarjeta") = data.IDBancoTarjeta

            Dim IDFras(-1) As Object
            Dim htFras As New Hashtable
            For Each drCobro As DataRow In dtCobrosTarjeta.Rows
                If blnNewTarjeta Then drCobro("IDTarjeta") = dtTarjeta.Rows(0)("IDTarjeta")

                drCobro("IDBancoPropio") = data.IDBancoPropio

                If Length(drCobro("IDFactura")) > 0 AndAlso Not htFras.ContainsKey(drCobro("IDFactura")) Then
                    ReDim Preserve IDFras(IDFras.Length)
                    IDFras(IDFras.Length - 1) = drCobro("IDFactura")

                    htFras(drCobro("IDFactura")) = drCobro("IDFactura")
                End If
            Next

            Dim fFacturas As New Filter
            fFacturas.Add(New InListFilterItem("IDFactura", IDFras, FilterType.Numeric))
            fFacturas.Add(New BooleanFilterItem("VencimientosManuales", False))
            Dim dtFras As DataTable = New FacturaVentaCabecera().Filter(fFacturas)
            If dtFras.Rows.Count > 0 Then
                For Each drFra As DataRow In dtFras.Rows
                    drFra("VencimientosManuales") = True
                Next
            End If


            AdminData.BeginTx()
            BusinessHelper.UpdateTable(dtTarjeta)
            BusinessHelper.UpdateTable(dtCobrosTarjeta)
            BusinessHelper.UpdateTable(dtFras)
            AdminData.CommitTx(True)

            AddInformacionTarjeta = dtTarjeta.Rows(0)("IDTarjeta")
        End If
    End Function

#End Region

End Class

