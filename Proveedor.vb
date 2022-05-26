Public Class ProveedorInfo
    Inherits ClassEntityInfo

    Private mIDProveedor As String
    Public Property IDProveedor() As String
        Get
            Return mIDProveedor
        End Get
        Set(ByVal Value As String)
            mIDProveedor = Value

            Dim objFilter As New Filter
            objFilter.Add(New StringFilterItem("IDProveedor", Value))
            objFilter.Add(New BooleanFilterItem("Predeterminado", True))
            Dim dtPB As DataTable = New ProveedorBanco().Filter(objFilter)
            If Not IsNothing(dtPB) AndAlso dtPB.Rows.Count > 0 Then
                If Length(dtPB.Rows(0)("IdProveedorBanco")) > 0 Then
                    IDProveedorBanco = dtPB.Rows(0)("IdProveedorBanco")
                End If
            End If
        End Set
    End Property
    Public IDPais As String
    Public IDTipoIVA As String
    Public EmpresaGrupo As Boolean
    Public BaseDatos As Guid   '//Base Datos Multiempresa
    Public RazonSocial As String
    Public Extranjero As Boolean
    Public CanariasCeutaMelilla As Boolean
    Public CEE As Boolean
    Public IDMoneda As String
    Public IDCentroGestion As String
    Public IDFormaPago As String
    Public IDCondicionPago As String
    Public IDDiaPago As String
    Public IDFormaEnvio As String
    Public IDCondicionEnvio As String
    Public DtoComercial As Double
    Public IDModoTransporte As String
    Public GrupoProveedor As String
    Public GrupoFactura As Boolean
    Public IDContadorCargo As String
    Public CCProveedor As String
    Public CCRetencion As String
    Public CCAnticipo As String
    Public CCEfectos As String
    Public CCFianza As String
    Public CCInMovilizadoCortoPlazo As String
    Public DescProveedor As String
    Public TipoFactura As Integer
    Public DtoProntoPago As Double
    Public RecFinan As Double
    Public IDBancoPropio As String
    Public CifProveedor As String
    Public Direccion As String
    Public Provincia As String
    Public Poblacion As String
    Public CodPostal As String
    Public Telefono As String
    Public Fax As String
    Public IDProveedorBanco As Integer
    Public IDTipoAsiento As Integer
    Public RetencionIRPF As Double
    Public TipoRetencionIRPF As Integer
    Public RegimenEspecial As Boolean
    Public IDCalificacion As String     '//Calidad
    Public Homologable As Boolean       '//Calidad
    Public CalidadConcertada As Boolean '//Calidad
    Public PorcentajeTolCierre As Double
    Public IDIdioma As String
    Public IDTipoClasif As String
    Public TipoDocIdentidad As enumTipoDocIdent
    Public IDAlmacenProveedor As String
    Public IVACaja As Boolean
    Public IDCNAE As String

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal data As DataRow)
        MyBase.New(data)
    End Sub
    Public Sub New(ByVal IDProveedor As String)
        MyBase.New()
        Me.Fill(IDProveedor)
    End Sub
    Public Overloads Overrides Sub Fill(ByVal ParamArray PrimaryKey() As Object)
        Dim dtProvInfo As DataTable = New BE.DataEngine().Filter("vNegProveedorInfo", New StringFilterItem("IDProveedor", PrimaryKey(0)))
        If dtProvInfo.Rows.Count > 0 Then
            Me.Fill(dtProvInfo.Rows(0))
        Else
            ApplicationService.GenerateError("El Proveedor | no existe.", Quoted(PrimaryKey(0)))
        End If
    End Sub

End Class

Public Class Proveedor

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper
    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroProveedor"

#End Region

#Region "Variables / Interfaces"

    Private mblnADD As Boolean

    Public Interface IProveedor
        Function ComprobarCategoriaEmpresaCRMProveedor(ByVal IDEmpresa As String, ByVal Entidad As String, ByVal IDValor As String) As DataTable
    End Interface

#End Region

#Region "Eventos RegisterAddNewTasks"

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim StDatos As New Contador.DatosDefaultCounterValue(data, GetType(Proveedor).Name, "IDProveedor")
        ProcessServer.ExecuteTask(Of Contador.DatosDefaultCounterValue)(AddressOf Contador.LoadDefaultCounterValue, StDatos, services)
        Dim P As New Parametro
        data("FechaAlta") = Date.Now
        data("IDFormaPago") = P.FormaPago
        data("IDCondicionPago") = P.CondicionPago
        data("IDPais") = P.Pais
        data("IDMoneda") = P.MonedaPred
        data("IDTipoIVA") = P.TipoIva
        data("EmpresaGrupo") = 0
        'data("IDTipoClasif") = P.TipoClasificacionProveedor
        data("TipoDocIdentidad") = P.TipoDocIdentificativo
    End Sub


#End Region

#Region "Eventos RegisterDeleteTasks"

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf AccionCheckCRM)
    End Sub

    <Task()> Public Shared Sub AccionCheckCRM(ByVal data As DataRow, ByVal services As ServiceProvider)
        'Control del CRM: si es un proveedor que ha venido convertido desde empresa, para borrarlo de los datos de Empresa del CRM
        Dim ClsCRMEmp As BusinessHelper = BusinessHelper.CreateBusinessObject("EmpresaCategoria")
        Dim DtCRM As DataTable = CType(ClsCRMEmp, IProveedor).ComprobarCategoriaEmpresaCRMProveedor(data("IDEmpresa") & String.Empty, "Proveedor", data("IDProveedor"))
        If Not DtCRM Is Nothing Then ClsCRMEmp.Delete(DtCRM)
    End Sub

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ComprobarPais)
        validateProcess.AddTask(Of DataRow)(AddressOf ComprobarMoneda)
        validateProcess.AddTask(Of DataRow)(AddressOf ComprobarCPostal)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarCifProveedor)
        validateProcess.AddTask(Of DataRow)(AddressOf ComprobarCContable)
        validateProcess.AddTask(Of DataRow)(AddressOf ComprobarDescripcion)
        'validateProcess.AddTask(Of DataRow)(AddressOf ComprobarEmpresaGrupo)
        validateProcess.AddTask(Of DataRow)(AddressOf ComprobarDocumento)
        validateProcess.AddTask(Of DataRow)(AddressOf ComprobarCondicionEnvio)
        validateProcess.AddTask(Of DataRow)(AddressOf ComprobarFormaEnvio)
    End Sub

    <Task()> Public Shared Sub ComprobarPais(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDPais")) = 0 Then ApplicationService.GenerateError("El País es un campo obligatorio.")
    End Sub

    <Task()> Public Shared Sub ComprobarMoneda(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDMoneda")) = 0 Then ApplicationService.GenerateError("La moneda es una dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ComprobarCPostal(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("CodPostal")) = 0 Then ApplicationService.GenerateError("El Código Postal es un datos obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarCifProveedor(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("TipoDocIdentidad")) = 0 Then ApplicationService.GenerateError("Es necesario especificar el tipo de documento.")
        If Length(data("CifProveedor")) = 0 Then ApplicationService.GenerateError("El documento de Identificación es un dato obligatorio.")
        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ValidaDocumentoIdentificativo, New DataRowPropertyAccessor(data), services)
    End Sub

    <Task()> Public Shared Sub ComprobarCContable(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim ClsParam As New Parametro
        If ClsParam.Contabilidad Then
            If Length(data("CCProveedor")) = 0 Then ApplicationService.GenerateError("La Cuenta Contable es un dato obligatorio.")
        End If
    End Sub

    <Task()> Public Shared Sub ComprobarDescripcion(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("DescProveedor")) = 0 Then ApplicationService.GenerateError("La descripción del Proveeedor es un dato obligatorio.")
    End Sub

    '<Task()> Public Shared Sub ComprobarEmpresaGrupo(ByVal data As DataRow, ByVal services As ServiceProvider)
    '    If data("EmpresaGrupo") <> 0 Then
    '        If IsDBNull(data("BaseDatos")) Then ApplicationService.GenerateError("El Nombre de la Empresa del Grupo es Obligatorio.")
    '    End If
    'End Sub

    <Task()> Public Shared Sub ComprobarDocumento(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("CifProveedor")) = 0 Then
            Dim ClsPais As New Pais
            Dim DtPais As DataTable = ClsPais.Filter(New FilterItem("IDPais", FilterOperator.Equal, data("IDPais"), FilterType.String))
            If Not DtPais Is Nothing AndAlso DtPais.Rows.Count > 0 Then
                If DtPais.Rows(0)("Extranjero") = 0 Then
                    ApplicationService.GenerateError("El Documento de Identificación es un dato obligatorio.")
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ComprobarCondicionEnvio(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDCondicionEnvio")) = 0 Then data("IDCondicionEnvio") = New Parametro().CondicionEnvio
    End Sub

    <Task()> Public Shared Sub ComprobarFormaEnvio(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDFormaEnvio")) = 0 Then data("IDFormaEnvio") = New Parametro().FormaEnvio
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClave)
        updateProcess.AddTask(Of DataRow)(AddressOf ActualizarRegimenEspecial)
        updateProcess.AddTask(Of DataRow)(AddressOf ValidaGrupoProveedor)
        updateProcess.AddTask(Of DataRow)(AddressOf Business.General.Comunes.UpdateEntityRow)
        updateProcess.AddTask(Of DataRow)(AddressOf Business.General.Comunes.MarcarComoActualizado)
        updateProcess.AddTask(Of DataRow)(AddressOf CrearProveedorDireccion)
        updateProcess.AddTask(Of DataRow)(AddressOf CrearProveedorBodegaMercado)
    End Sub

    <Task()> Public Shared Sub AsignarClave(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IdProveedor") & String.Empty) = 0 Then
                ' Caso en el que no esté introducido un código de proveedor
                If Length(data("IdContador")) > 0 Then
                    data("IdProveedor") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, data("IDContador"), services)
                Else
                    Dim StDatos As New Contador.DatosDefaultCounterValue
                    StDatos.Row = data
                    StDatos.EntityName = "Proveedor"
                    StDatos.FieldName = "IDProveedor"
                    ProcessServer.ExecuteTask(Of Contador.DatosDefaultCounterValue)(AddressOf Contador.LoadDefaultCounterValue, StDatos, services)
                    If Length(data("IDContador")) > 0 Then
                        data("IdProveedor") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, data("IDContador"), services)
                    Else
                        ApplicationService.GenerateError("No ha configurado un contador predeterminado para los proveedores.")
                    End If
                End If
            Else
                ' Caso en el que el código de proveedor proviene de la asignación de contador,
                '  en caso contrario sería código manual por lo que no hay que mover el contador
                If Length(data("IdContador")) > 0 Then data("IdProveedor") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, data("IDContador"), services)
                Dim dtProveedor As DataTable = New Proveedor().SelOnPrimaryKey(data("IdProveedor"))
                If Not IsNothing(dtProveedor) AndAlso dtProveedor.Rows.Count > 0 Then
                    dtProveedor.Rows.Clear()
                    ApplicationService.GenerateError("El Proveedor introducido ya existe.")
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarRegimenEspecial(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Modified Then
            If data("RegimenEspecial") <> data("RegimenEspecial", DataRowVersion.Original) Then
                Dim ClsFactCC As New FacturaCompraCabecera
                ClsFactCC.ActualizarRegimenEspecial(data)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidaGrupoProveedor(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IdGrupoProveedor")) > 0 Then
            If data("IDProveedor") = data("IdGrupoProveedor") Then
                ApplicationService.GenerateError("El Proveedor no puede ser cabecera y miembro de un grupo al mismo tiempo")
            Else
                Dim dtGrupo As DataTable = New Proveedor().SelOnPrimaryKey(data("IdGrupoProveedor"))
                If Not dtGrupo Is Nothing AndAlso dtGrupo.Rows.Count > 0 Then
                    If Length(dtGrupo.Rows(0)("IdGrupoProveedor")) > 0 Then
                        ApplicationService.GenerateError("El Proveedor no puede ser cabecera de grupo porque pertenece a otro grupo.")
                    Else
                        Dim dtHijos As DataTable = New Proveedor().Filter(New FilterItem("IdGrupoProveedor", FilterOperator.Equal, data("IDProveedor")))
                        If Not dtHijos Is Nothing AndAlso dtHijos.Rows.Count > 0 Then
                            ApplicationService.GenerateError("El cliente no puede pertenecer al grupo porque es cabecera de otro grupo.")
                        End If
                    End If
                End If
            End If
        Else
            Dim DtProv As DataTable = New Proveedor().Filter(New FilterItem("IDGrupoProveedor", FilterOperator.Equal, data("IDProveedor")))
            If data.RowState = DataRowState.Added OrElse (data.RowState = DataRowState.Modified AndAlso DtProv.Rows.Count = 0) Then
                data("GrupoFactura") = False
                data("GrupoArticulo") = False
                If data("EmpresaGrupo") = 0 And (data("TipoDocIdentidad") = enumTipoDocIdent.NIF Or data("TipoDocIdentidad") = enumTipoDocIdent.CertificiadoResiFiscal) Then
                    Dim StDatos As New DatosCifRepetido
                    StDatos.Documento = data("CifProveedor") & String.Empty
                    StDatos.TipoDocumento = Nz(data("TipoDocIdentidad"), 0)
                    StDatos.IDProveedor = data("IDProveedor")
                    StDatos.IDGrupoProveedor = data("IDGrupoProveedor") & String.Empty
                    If ProcessServer.ExecuteTask(Of DatosCifRepetido, Boolean)(AddressOf CIFRepetido, StDatos, services) Then
                        ApplicationService.GenerateError("Ya existe un Proveedor con el Documento |. | Sólo se permiten proveedores con el mismo Documento si pertenecen al mismo grupo ", data("CifProveedor"), vbNewLine)
                    End If
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CrearProveedorDireccion(ByVal data As DataRow, ByVal services As ServiceProvider)
        'Controlar que no me llega una de presentación
        Dim upPresenta As UpdatePackage = services.GetService(Of UpdatePackage)()
        Dim dt As DataTable = upPresenta.Item(GetType(ProveedorDireccion).Name).First
        Dim adr() As DataRow
        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            adr = dt.Select("IDProveedor='" & data("IDProveedor") & "'")
        End If
        If adr Is Nothing OrElse adr.Length = 0 Then
            If data.RowState = DataRowState.Added Then
                'Hay que guardar 1 registro con 3 checks:   .- Dirección de Envio
                'Todos marcados y predeterminados           .- Dirección de Facturación
                '                                           .- Dirección de Giro
                dt = New ProveedorDireccion().AddNewForm
                dt.Rows(0)("IdProveedor") = data("IdProveedor")
                dt.Rows(0)("RazonSocial") = data("RazonSocial")
                dt.Rows(0)("Direccion") = data("Direccion")
                dt.Rows(0)("CodPostal") = data("CodPostal")
                dt.Rows(0)("Poblacion") = data("Poblacion")
                dt.Rows(0)("Provincia") = data("Provincia")
                dt.Rows(0)("IDPais") = data("IDPais")
                dt.Rows(0)("Telefono") = data("Telefono")
                dt.Rows(0)("Fax") = data("Fax")
                dt.Rows(0)("Email") = data("Email")
                dt.Rows(0)("Envio") = 1
                dt.Rows(0)("Factura") = 1
                dt.Rows(0)("Giro") = 1
                dt.Rows(0)("PredeterminadaEnvio") = 1
                dt.Rows(0)("PredeterminadaFactura") = 1
                dt.Rows(0)("PredeterminadaGiro") = 1
                BusinessHelper.UpdateTable(dt)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CrearProveedorBodegaMercado(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added OrElse (data.RowState = DataRowState.Modified AndAlso data("IDMercado") & String.Empty <> data("IDMercado", DataRowVersion.Original) & String.Empty) Then
            If Length(data("IDMercado")) > 0 Then
                Dim dt As DataTable = New Parametro().SelOnPrimaryKey("BDGMERCADO")
                If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                    If Length(Nz(dt.Rows(0)("Valor"), String.Empty)) > 0 Then
                        If data("IDMercado") = dt.Rows(0)("Valor") Then
                            Dim ClsBdgProv As BusinessHelper = BusinessHelper.CreateBusinessObject("BdgProveedor")
                            Dim dtProveedorBdg As DataTable = ClsBdgProv.SelOnPrimaryKey(data("IDProveedor"))
                            If dtProveedorBdg.Rows.Count = 0 Then
                                Dim DtNewProv As DataTable = ClsBdgProv.AddNewForm
                                DtNewProv.Rows(0)("IDProveedor") = data("IDProveedor")
                                'DtNewProv.Rows(0)("IDGrupo") = 
                                'DtNewProv.Rows(0)("IDTarifaT") = 
                                'DtNewProv.Rows(0)("IDTarifaB") = 
                                DtNewProv.Rows(0)("PrecioOrigenT") = 0
                                DtNewProv.Rows(0)("PrecioOrigenB") = 0
                                DtNewProv.Rows(0)("PrecioExcedenteT") = 0
                                DtNewProv.Rows(0)("PrecioExcedenteB") = 0
                                ClsBdgProv.Update(DtNewProv)
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End Sub

#End Region

#Region "Eventos GetBusinessRules"

    <Serializable()> _
    Public Class DatosCifRepetido
        Public Documento As String
        Public TipoDocumento As enumTipoDocIdent
        Public IDProveedor As String
        Public IDGrupoProveedor As String
    End Class

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBrl As New BusinessRules
        oBrl.Add("FechaAlta", AddressOf ComprobarDato)
        oBrl.Add("FechaValidezHomologacion", AddressOf ComprobarDato)
        oBrl.Add("FechaHomologacion", AddressOf ComprobarDato)
        oBrl.Add("FechaUltimaCalificacion", AddressOf ComprobarDato)

        oBrl.Add("PorcentajeTolCierre", AddressOf ComprobarDato)
        oBrl.Add("DtoComercial", AddressOf ComprobarDato)
        oBrl.Add("RetencionIRPF", AddressOf ComprobarDato)
        oBrl.Add("Resultado", AddressOf ComprobarDato)
        oBrl.Add("ResultadoCC", AddressOf ComprobarDato)

        oBrl.Add("CifProveedor", AddressOf TratarCif)
        oBrl.Add("IDPais", AddressOf TratarCif)
        oBrl.Add("TipoDocIdentidad", AddressOf TratarCif)

        oBrl.Add("CodPostal", AddressOf TratarCPostal)
        Return oBrl
    End Function

    <Task()> Public Shared Sub ComprobarDato(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) = 0 Then ApplicationService.GenerateError("El campo: |, no puede ir vacio.", data.ColumnName)
    End Sub

    <Task()> Public Shared Sub TratarCif(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 AndAlso Length(data.Current("CifProveedor")) > 0 Then
            data.Current(data.ColumnName) = data.Value
            ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ValidaDocumentoIdentificativo, data.Current, services)
        End If
    End Sub

    <Task()> Public Shared Sub ValidaDocumentoIdentificativo(ByVal data As IPropertyAccessor, ByVal services As ServiceProvider)
        If data("TipoDocIdentidad") = enumTipoDocIdent.NIF Or data("TipoDocIdentidad") = enumTipoDocIdent.CertificiadoResiFiscal Then ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf TratarCifProveedor, data, services)
        Dim info As New DataDocIdentificacion(data("CifProveedor"), data("IDPais"), data("TipoDocIdentidad"))
        info = ProcessServer.ExecuteTask(Of DataDocIdentificacion, DataDocIdentificacion)(AddressOf Comunes.ValidarDocumentoIdentificativo, info, services)
        If Not info.EsCorrecto Then
            ApplicationService.GenerateError("El Documento introducido no es un '|'. Intoduzca uno correcto o cambie de tipo de documento", info.TipoDocumento)
        End If
    End Sub

    <Task()> Public Shared Sub TratarCifProveedor(ByVal data As IPropertyAccessor, ByVal services As ServiceProvider)
        If Length(data("IDPais")) > 0 And ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf Pais.Nacional, data("IDPais"), services) Then
            If Length(data("CifProveedor")) > 0 Then
                Dim blnCancel As Boolean
                'ValidaCIF(data, blnCancel)
                If data("EmpresaGrupo") = 0 Then
                    Dim StDatos As New DatosCifRepetido
                    StDatos.Documento = data("CifProveedor")
                    StDatos.TipoDocumento = Nz(data("TipoDocIdentidad"), 0)
                    StDatos.IDProveedor = Nz(data("IDProveedor"), String.Empty)
                    StDatos.IDGrupoProveedor = data("IDGrupoProveedor") & String.Empty
                    Dim BlnResul As Boolean = ProcessServer.ExecuteTask(Of DatosCifRepetido, Boolean)(AddressOf CIFRepetido, StDatos, services)
                    If BlnResul Then
                        ApplicationService.GenerateError("Ya existe un Proveedor con el Documento |. | Sólo se permiten proveedores con el mismo Documento si pertenecen al mismo grupo ", data("CifProveedor"), vbNewLine)
                    End If
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Function CIFRepetido(ByVal data As DatosCifRepetido, ByVal services As ServiceProvider) As Boolean
        If Length(data.Documento) > 0 Then
            Dim objFilter As New Filter
            objFilter.Add(New StringFilterItem("CifProveedor", data.Documento))
            objFilter.Add(New NumberFilterItem("TipoDocIdentidad", data.TipoDocumento))
            objFilter.Add(New StringFilterItem("IDProveedor", FilterOperator.NotEqual, data.IDProveedor))
            Dim dt As DataTable = New Proveedor().Filter(objFilter)
            If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
                objFilter.Clear()
                objFilter.UnionOperator = FilterUnionOperator.Or
                If Length(data.IDGrupoProveedor) > 0 Then
                    objFilter.Add(New StringFilterItem("IDGrupoProveedor", data.IDGrupoProveedor))
                    objFilter.Add(New StringFilterItem("IDProveedor", data.IDGrupoProveedor))
                ElseIf Length(data.IDProveedor) > 0 Then
                    objFilter.Add(New StringFilterItem("IDGrupoProveedor", data.IDProveedor))
                Else
                    objFilter.Add(New NoRowsFilterItem)
                End If

                Dim WhereGrupoProveedor As String = objFilter.Compose(New AdoFilterComposer)
                Dim adr() As DataRow = dt.Select(WhereGrupoProveedor)
                If Not IsNothing(adr) AndAlso adr.Length > 0 Then
                    CIFRepetido = False
                Else
                    CIFRepetido = True
                End If
            End If
        End If
    End Function

    <Task()> Public Shared Sub TratarCPostal(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim infoCP As New CodPostalInfo(CStr(data.Value), data.Current("IDPais") & String.Empty)
            If Length(infoCP.DescPoblacion) > 0 Then data.Current("Poblacion") = infoCP.DescPoblacion
            If Length(infoCP.DescProvincia) > 0 Then data.Current("Provincia") = infoCP.DescProvincia
            If Length(infoCP.IDPais) > 0 Then data.Current("IDPais") = infoCP.IDPais
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Serializable()> _
    Public Class DatosEstRetraso
        Public IDProveedor As String
        Public FechaDesde As Date
        Public FechaHasta As Date
    End Class

    <Task()> Public Shared Function Nacional(ByVal data As String, ByVal services As ServiceProvider) As Boolean
        If Length(data) > 0 Then
            Dim dtProveedor As DataTable = New Proveedor().SelOnPrimaryKey(data)
            If Not IsNothing(dtProveedor) AndAlso dtProveedor.Rows.Count > 0 Then
                'Si el Proveedor no tiene País se le considerará como Nacional.
                If Length(dtProveedor.Rows(0)("IDPais")) > 0 Then
                    Return ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf Pais.Nacional, dtProveedor.Rows(0)("IDPais"), services)
                Else : Return True
                End If
            End If
        End If
        Return False
    End Function

    <Task()> Public Shared Function ValidaProveedor(ByVal data As String, ByVal services As ServiceProvider) As DataTable
        If Length(data) > 0 Then
            Dim dt As DataTable = New Proveedor().SelOnPrimaryKey(data)
            If IsNothing(dt) OrElse dt.Rows.Count = 0 Then
                ApplicationService.GenerateError("El Proveedor | no existe.", Quoted(data))
            End If
            Return dt
        End If
    End Function

    <Task()> Public Shared Function ObtenerDatosEstadisticaRetraso(ByVal data As DatosEstRetraso, ByVal services As ServiceProvider) As DataTable
        Dim selectSQL As New System.Text.StringBuilder
        selectSQL.Append("SET DATEFORMAT DMY; SELECT tbPedidoCompraCabecera.IDProveedor, tbMaestroProveedor.DescProveedor ,tbMaestroProveedor.IDTipoProveedor, AVG(DATEDIFF(dd, tbPedidoCompraCabecera.FechaPedido, " & _
        " tbAlbaranCompraCabecera.FechaAlbaran)) AS PlazoMedio, COUNT(tbAlbaranCompraLinea.IDAlbaran) AS Entregas," & _
        " SUM(CASE WHEN (tbAlbaranCompraCabecera.FechaAlbaran > tbPedidoCompraLinea.FechaEntrega) THEN 1 ELSE 0 END) AS Retrasos," & _
        " AVG(CASE WHEN (tbAlbaranCompraCabecera.FechaAlbaran > tbPedidoCompraLinea.FechaEntrega) THEN DATEDIFF(DD," & _
        " tbPedidoCompraLinea.FechaEntrega, tbAlbaranCompraCabecera.FechaAlbaran) ELSE 0 END) AS RetrasoMedio," & _
        " ROUND(AVG(tbPedidoCompraLinea.QPedida), 0) AS TamañoMedioPedido, round(AVG(tbAlbaranCompraLinea.QServida), 0)" & _
        " AS TamañoMedioEntrega FROM tbAlbaranCompraLinea INNER JOIN" & _
        " tbAlbaranCompraCabecera ON tbAlbaranCompraLinea.IDAlbaran = tbAlbaranCompraCabecera.IDAlbaran INNER JOIN" & _
        " tbMaestroProveedor INNER JOIN" & _
        " tbPedidoCompraCabecera ON tbMaestroProveedor.IDProveedor = tbPedidoCompraCabecera.IDProveedor INNER JOIN" & _
        " tbPedidoCompraLinea ON tbPedidoCompraCabecera.IDPedido = tbPedidoCompraLinea.IDPedido ON" & _
        " tbAlbaranCompraLinea.IDLineaPedido = tbPedidoCompraLinea.IDLineaPedido")

        Dim whereSQL As New Text.StringBuilder
        If data.IDProveedor.Length > 0 Then
            whereSQL.Append("tbPedidoCompraCabecera.IdProveedor = '" & data.IDProveedor & "' AND ")
        End If
        If Not data.FechaDesde = Nothing And Not data.FechaHasta = Nothing Then
            whereSQL.Append("FechaPedido BETWEEN '" & Convert.ToDateTime(data.FechaDesde) & "' AND '" & _
                Convert.ToDateTime(data.FechaHasta) & "' AND ")
        Else
            If Not data.FechaDesde = Nothing Then
                whereSQL.Append("FechaPedido >= '" & Convert.ToDateTime(data.FechaDesde) & "' AND ")
            End If
            If Not data.FechaHasta = Nothing Then
                whereSQL.Append("FechaPedido <= '" & Convert.ToDateTime(data.FechaHasta) & "' AND ")
            End If
        End If

        If whereSQL.Length > 0 Then
            selectSQL.Append(" WHERE ")
            selectSQL.Append(whereSQL.ToString.Substring(0, whereSQL.Length - 4))
        End If

        selectSQL.Append(" GROUP BY tbPedidoCompraCabecera.IDProveedor, tbMaestroProveedor.DescProveedor,tbMaestroProveedor.IDTipoProveedor")

        Dim ad As New AdminData
        Dim cmdEstadisticas As Common.DbCommand = AdminData.GetCommand
        cmdEstadisticas.CommandType = CommandType.Text
        cmdEstadisticas.CommandText = selectSQL.ToString()
        Return ad.Execute(cmdEstadisticas, ExecuteCommand.ExecuteReader)
    End Function

    <Task()> Public Shared Function ObtenerXDataBase(ByVal data As Object, ByVal services As ServiceProvider) As DataTable
        Return AdminData.Filter("xDataBase", , , , , True)
    End Function

    <Task()> Public Shared Function InformacionProveedor(ByVal data As String, ByVal services As ServiceProvider) As ProveedorInfo
        Dim ProvInfo As New ProveedorInfo
        ProvInfo.Fill(data)
        Return ProvInfo
    End Function

    <Serializable()> _
    Public Class DataCopiaProv
        Public IDProveedorNew As String
        Public Errores As String

        Public Sub New()
        End Sub
        Public Sub New(ByVal IDProveedorNew As String, ByVal Errores As String)
            Me.IDProveedorNew = IDProveedorNew
            Me.Errores = Errores
        End Sub
    End Class

    <Task()> Public Shared Function CopiarProveedor(ByVal IDProveedorOrigen As String, ByVal services As ServiceProvider) As DataCopiaProv
        Dim StDataReturn As New DataCopiaProv
        Try
            AdminData.BeginTx()
            'Cabecera Proveedor
            Dim ClsProv As New Proveedor
            Dim DtProvOrigen As DataTable = ClsProv.SelOnPrimaryKey(IDProveedorOrigen)
            Dim DtProvDestino As DataTable = ClsProv.AddNew
            DtProvDestino.Rows.Add(DtProvOrigen.Rows(0).ItemArray)
            If Length(DtProvDestino.Rows(0)("IDContador")) = 0 Then
                Dim DataCont As Contador.DefaultCounter = ProcessServer.ExecuteTask(Of String, Contador.DefaultCounter)(AddressOf Contador.GetDefaultCounterValue, "Proveedor", services)
                DtProvDestino.Rows(0)("IDContador") = DataCont.CounterID
            End If
            StDataReturn.IDProveedorNew = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, DtProvDestino.Rows(0)("IDContador"), services)
            DtProvDestino.Rows(0)("IDProveedor") = StDataReturn.IDProveedorNew
            DtProvDestino.Rows(0)("IDGrupoProveedor") = IDProveedorOrigen

            'Proveedor Banco
            Dim ClsProvBanco As New ProveedorBanco
            Dim DtProvBancoOrigen As DataTable = ClsProvBanco.Filter(New FilterItem("IDProveedor", FilterOperator.Equal, IDProveedorOrigen))
            Dim DtProvBancoDestino As DataTable = ClsProvBanco.AddNew
            For Each Dr As DataRow In DtProvBancoOrigen.Select
                DtProvBancoDestino.Rows.Add(Dr.ItemArray)
                DtProvBancoDestino.Rows(DtProvBancoDestino.Rows.Count - 1)("IDProveedor") = StDataReturn.IDProveedorNew
                DtProvBancoDestino.Rows(DtProvBancoDestino.Rows.Count - 1)("IDProveedorBanco") = AdminData.GetAutoNumeric
            Next

            'Proveedor Articulo
            Dim ClsProvArt As New ArticuloProveedor
            Dim DtProvArtOrigen As DataTable = ClsProvArt.Filter(New FilterItem("IDProveedor", FilterOperator.Equal, IDProveedorOrigen))
            Dim DtProvArtDestino As DataTable = ClsProvArt.AddNew
            For Each Dr As DataRow In DtProvArtOrigen.Select
                DtProvArtDestino.Rows.Add(Dr.ItemArray)
                DtProvArtDestino.Rows(DtProvArtDestino.Rows.Count - 1)("IDProveedor") = StDataReturn.IDProveedorNew
            Next

            'Proveedor Articulo Linea
            Dim ClsProvArtLinea As New ArticuloProveedorLinea
            Dim DtProvArtLinOrigen As DataTable = ClsProvArtLinea.Filter(New FilterItem("IDProveedor", FilterOperator.Equal, IDProveedorOrigen))
            Dim DtProvArtLinDestino As DataTable = ClsProvArtLinea.AddNew
            For Each Dr As DataRow In DtProvArtLinOrigen.Select
                DtProvArtLinDestino.Rows.Add(Dr.ItemArray)
                DtProvArtLinDestino.Rows(DtProvArtLinDestino.Rows.Count - 1)("IDProveedor") = StDataReturn.IDProveedorNew
            Next

            'Proveedor Direccion
            Dim ClsProvDirec As New ProveedorDireccion
            Dim DtProvDirecOrigen As DataTable = ClsProvDirec.Filter(New FilterItem("IDProveedor", FilterOperator.Equal, IDProveedorOrigen))
            Dim DtProvDirecDestino As DataTable = ClsProvDirec.AddNew
            For Each Dr As DataRow In DtProvDirecOrigen.Select
                DtProvDirecDestino.Rows.Add(Dr.ItemArray)
                DtProvDirecDestino.Rows(DtProvDirecDestino.Rows.Count - 1)("IDProveedor") = StDataReturn.IDProveedorNew
                DtProvDirecDestino.Rows(DtProvDirecDestino.Rows.Count - 1)("IDDireccion") = AdminData.GetAutoNumeric
            Next

            'Proveedor Familia Descuento
            Dim ClsProvFamDto As New ProveedorDescuentoFamilia
            Dim DtProvFamDtoOrigen As DataTable = ClsProvFamDto.Filter(New FilterItem("IDProveedor", FilterOperator.Equal, IDProveedorOrigen))
            Dim DtProvFamDtoDestino As DataTable = ClsProvFamDto.AddNew
            For Each Dr As DataRow In DtProvFamDtoOrigen.Select
                DtProvFamDtoDestino.Rows.Add(Dr.ItemArray)
                DtProvFamDtoDestino.Rows(DtProvFamDtoDestino.Rows.Count - 1)("IDProveedor") = StDataReturn.IDProveedorNew
                DtProvFamDtoDestino.Rows(DtProvFamDtoDestino.Rows.Count - 1)("IDProveedorFamilia") = AdminData.GetAutoNumeric
            Next

            'Proveedor Vacaciones
            Dim ClsProvVac As New ProveedorVacaciones
            Dim DtProvVacOrigen As DataTable = ClsProvVac.Filter(New FilterItem("IDProveedor", FilterOperator.Equal, IDProveedorOrigen))
            Dim DtProvVacDestino As DataTable = ClsProvVac.AddNew
            For Each Dr As DataRow In DtProvVacOrigen.Select
                DtProvVacDestino.Rows.Add(Dr.ItemArray)
                DtProvVacDestino.Rows(DtProvVacDestino.Rows.Count - 1)("IDProveedor") = StDataReturn.IDProveedorNew
                DtProvVacDestino.Rows(DtProvVacDestino.Rows.Count - 1)("IDVacacion") = AdminData.GetAutoNumeric
            Next

            'Proveedor Observacion
            Dim ClsProvObv As New ProveedorObservacion
            Dim DtProvObvOrigen As DataTable = ClsProvObv.Filter(New FilterItem("IDProveedor", FilterOperator.Equal, IDProveedorOrigen))
            Dim DtProvObvDestino As DataTable = ClsProvObv.AddNew
            For Each Dr As DataRow In DtProvObvOrigen.Select
                DtProvObvDestino.Rows.Add(Dr.ItemArray)
                DtProvObvDestino.Rows(DtProvObvDestino.Rows.Count - 1)("IDProveedor") = StDataReturn.IDProveedorNew
                DtProvObvDestino.Rows(DtProvObvDestino.Rows.Count - 1)("IDProveedorObservacion") = AdminData.GetAutoNumeric
            Next

            BusinessHelper.UpdateTable(DtProvDestino) : BusinessHelper.UpdateTable(DtProvBancoDestino)
            BusinessHelper.UpdateTable(DtProvArtDestino) : BusinessHelper.UpdateTable(DtProvArtLinDestino)
            BusinessHelper.UpdateTable(DtProvDirecDestino) : BusinessHelper.UpdateTable(DtProvFamDtoDestino)
            BusinessHelper.UpdateTable(DtProvVacDestino) : BusinessHelper.UpdateTable(DtProvObvDestino)

            AdminData.CommitTx(True)
        Catch ex As Exception
            AdminData.RollBackTx()
            StDataReturn.Errores = ex.Message
        End Try

        Return StDataReturn
    End Function

#End Region

End Class