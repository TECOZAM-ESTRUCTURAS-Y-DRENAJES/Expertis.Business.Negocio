Public Class Cheque
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

#Region " Constructor "

    Private Const cnEntidad As String = "tbCheque"

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
        If Length(data("NumeroCheque")) = 0 Then ApplicationService.GenerateError("El Número de Cheque es un dato obligatorio.")
        If Length(data("IDBancoCheque")) = 0 Then ApplicationService.GenerateError("El Banco del Cheque es un dato obligatorio.")
        If Length(data("IDBancoPropio")) = 0 Then ApplicationService.GenerateError("La Caja es un dato obligatorio.")
    End Sub


    <Task()> Public Shared Sub ValidarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = New Cheque().SelOnPrimaryKey(data("IDCheque"))
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
        If Length(data("IDCheque")) = 0 Then data("IDCheque") = AdminData.GetAutoNumeric
    End Sub

#End Region


#Region " Gestión de Cheques "

    <Serializable()> _
    Public Class dataAddInformacionCheque
        Public IDCheque As Integer?
        Public IDCobros(-1) As Integer
        Public Fecha As Date
        Public NumeroCheque As String
        Public NombreCheque As String
        Public DireccionCheque As String
        Public IDBancoCheque As String
        Public TelefCheque As String
        Public NumCuentaCheque As String
        Public IDBancoPropio As String

        Public Sub New(ByVal IDCobros As Integer(), ByVal Fecha As Date, ByVal NumeroCheque As String, ByVal NombreCheque As String, ByVal DireccionCheque As String, ByVal TelefCheque As String, ByVal NumCuentaCheque As String, ByVal IDBancoCheque As String, ByVal IDBancoPropio As String)
            Me.IDCobros = IDCobros
            Me.NumeroCheque = NumeroCheque
            Me.NombreCheque = NombreCheque
            Me.DireccionCheque = DireccionCheque
            Me.TelefCheque = TelefCheque
            Me.NumCuentaCheque = NumCuentaCheque
            Me.IDBancoCheque = IDBancoCheque
            Me.IDBancoPropio = IDBancoPropio
            Me.Fecha = Fecha
        End Sub

        Public Sub New(ByVal IDCheque As Integer, ByVal Fecha As Date, ByVal NumeroCheque As String, ByVal NombreCheque As String, ByVal DireccionCheque As String, ByVal TelefCheque As String, ByVal NumCuentaCheque As String, ByVal IDBancoCheque As String, ByVal IDBancoPropio As String)
            Me.IDCheque = IDCheque
            Me.NumeroCheque = NumeroCheque
            Me.NombreCheque = NombreCheque
            Me.DireccionCheque = DireccionCheque
            Me.TelefCheque = TelefCheque
            Me.NumCuentaCheque = NumCuentaCheque
            Me.IDBancoCheque = IDBancoCheque
            Me.IDBancoPropio = IDBancoPropio
            Me.Fecha = Fecha
        End Sub
    End Class

    <Task()> Public Shared Function AddInformacionCheque(ByVal data As dataAddInformacionCheque, ByVal services As ServiceProvider) As Integer
        AddInformacionCheque = -1
        Dim blnNewCheque As Boolean
        Dim CHQ As New Cheque
        Dim c As New Cobro
        Dim dtCheque As DataTable
        Dim dtCobrosCheque As DataTable
        If Not data.IDCheque Is Nothing Then
            '//Modificando cheque
            dtCheque = CHQ.SelOnPrimaryKey(data.IDCheque)
            dtCobrosCheque = c.Filter(New NumberFilterItem("IDCheque", data.IDCheque))
        ElseIf Not data.IDCobros Is Nothing AndAlso data.IDCobros.Count > 0 Then
            '//Creando cheque
            Dim IdCobroObj(data.IDCobros.Length - 1) As Object
            data.IDCobros.CopyTo(IdCobroObj, 0)
            dtCheque = CHQ.AddNewForm
            blnNewCheque = True
            dtCobrosCheque = c.Filter(New InListFilterItem("IDCobro", IdCobroObj, FilterType.Numeric))
        End If

        If dtCheque.Rows.Count > 0 Then
            dtCheque.Rows(0)("NumeroCheque") = data.NumeroCheque
            dtCheque.Rows(0)("Fecha") = data.Fecha
            If Length(data.NombreCheque) > 0 Then
                dtCheque.Rows(0)("NombreCheque") = data.NombreCheque
            Else
                dtCheque.Rows(0)("NombreCheque") = System.DBNull.Value
            End If
            If Length(data.DireccionCheque) > 0 Then
                dtCheque.Rows(0)("DireccionCheque") = data.DireccionCheque
            Else
                dtCheque.Rows(0)("DireccionCheque") = System.DBNull.Value
            End If

            If Length(data.TelefCheque) > 0 Then
                dtCheque.Rows(0)("TelefCheque") = data.TelefCheque
            Else
                dtCheque.Rows(0)("TelefCheque") = System.DBNull.Value
            End If

            If Length(data.NumCuentaCheque) > 0 Then
                dtCheque.Rows(0)("NumCuentaCheque") = data.NumCuentaCheque
            Else
                dtCheque.Rows(0)("NumCuentaCheque") = System.DBNull.Value
            End If

            dtCheque.Rows(0)("IDBancoCheque") = data.IDBancoCheque
            dtCheque.Rows(0)("IDBancoPropio") = data.IDBancoPropio

            Dim IDFras(-1) As Object
            Dim htFras As New Hashtable
            For Each drCobro As DataRow In dtCobrosCheque.Rows
                If blnNewCheque Then drCobro("IDCheque") = dtCheque.Rows(0)("IDCheque")
                drCobro("IDBancoPropio") = data.IDBancoPropio

                If Length(drCobro("IDFactura")) > 0 AndAlso Not htFras.ContainsKey(drCobro("IDFactura")) Then
                    ReDim Preserve IDFras(IDFras.Length)
                    IDFras(IDFras.Length - 1) = drCobro("IDFactura")

                    htFras(drCobro("IDFactura")) = drCobro("IDFactura")
                End If
            Next

            Dim dtFras As DataTable
            If htFras.Count > 0 Then
                Dim fFacturas As New Filter
                fFacturas.Add(New InListFilterItem("IDFactura", IDFras, FilterType.Numeric))
                fFacturas.Add(New BooleanFilterItem("VencimientosManuales", False))
                dtFras = New FacturaVentaCabecera().Filter(fFacturas)
                If dtFras.Rows.Count > 0 Then
                    For Each drFra As DataRow In dtFras.Rows
                        drFra("VencimientosManuales") = True
                    Next
                End If
            End If

            AdminData.BeginTx()
            BusinessHelper.UpdateTable(dtCheque)
            BusinessHelper.UpdateTable(dtCobrosCheque)
            If Not dtFras Is Nothing AndAlso dtFras.Rows.Count > 0 Then BusinessHelper.UpdateTable(dtFras)
            AdminData.CommitTx(True)

            AddInformacionCheque = dtCheque.Rows(0)("IDCheque")
        End If
    End Function

#End Region


End Class
