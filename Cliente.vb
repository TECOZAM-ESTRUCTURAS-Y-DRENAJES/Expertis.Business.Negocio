Public Class Cliente

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper
    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub
    Private Const cnEntidad As String = "tbMaestroCliente"

#End Region

#Region "Interfaces"

    Public Interface ICliente
        Function ComprobarCategoriaEmpresaCRMCliente(ByVal IDEmpresa As String, ByVal Entidad As String, ByVal IDValor As String) As DataTable
    End Interface

#End Region

#Region "Clases"

    <Serializable()> _
    Public Class DataCIFRepetido
        Public Documento As String
        Public TipoDocumento As enumTipoDocIdent
        Public IDCliente As String
        Public IDGrupoCliente As String
    End Class

    <Serializable()> _
    Public Class DataBloqArtClie
        Public IDCliente As String
        Public IDArticulo As String
        Public RefCliente As String
    End Class

    <Serializable()> _
    Public Class DataRiesgoCliente
        Public IDCliente As String
        Public ImporteASumar As Double
    End Class

    <Serializable()> _
   Public Class DataGrupoCliente
        Public IDCliente As String
        Public TipoGrupo As String
    End Class
#End Region

#Region "Función RegisterAddNewTasks"

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim StDatos As New Contador.DatosDefaultCounterValue
        StDatos.Row = data
        StDatos.EntityName = "Cliente"
        StDatos.FieldName = "IDCliente"
        ProcessServer.ExecuteTask(Of Contador.DatosDefaultCounterValue)(AddressOf Contador.LoadDefaultCounterValue, StDatos, services)

        Dim ClsParam As New Parametro
        data("IDFormaPago") = ClsParam.FormaPago
        data("IDCondicionPago") = ClsParam.CondicionPago
        data("IDPais") = ClsParam.Pais
        data("IDMoneda") = ClsParam.MonedaPred
        data("IDTipoIVA") = ClsParam.TipoIva
        data("EmpresaGrupo") = 0
        data("TipoDocIdentidad") = ClsParam.TipoDocIdentificativo
    End Sub

#End Region

#Region "Funciones BusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBRL As New BusinessRules
        oBRL.Add("IDDiaPago", AddressOf CambioDiapago)
        oBRL.Add("CifCliente", AddressOf CambioCifPaisCliente)
        oBRL.Add("IDPais", AddressOf CambioCifPaisCliente)
        oBRL.Add("TipoDocIdentidad", AddressOf CambioCifPaisCliente)
        oBRL.Add("IDOperarioBloqueo", AddressOf CambioOperarioBloqueo)
        oBRL.Add("DescOperarioBloqueo", AddressOf CambioDescOperarioBloqueo)
        oBRL.Add("Bloqueado", AddressOf CambioBloqueado)
        oBRL.Add("DiaFacturacion", AddressOf CambioDiaFacturacion)
        oBRL.Add("CodPostal", AddressOf CambioCodPostal)
        oBRL.Add("IDTipoCliente", AddressOf CambioIDTipoCliente)
        oBRL.Add("IDCondicionPago", AddressOf CambioIDCondicionPago)
        oBRL.Add("IDFormaPago", AddressOf CambioIDFormaPago)
        oBRL.Add("IDAseguradora", AddressOf CambioAseguradora)
        oBRL.Add("RiesgoInterno", AddressOf CalculoRiesgoConcedido)
        oBRL.Add("LimiteCapitalAsegurado", AddressOf CalculoRiesgoConcedido)
        Return oBRL
    End Function

    <Task()> Public Shared Sub CambioIDTipoCliente(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf AsignarMotivoNoAsegurado, data, services)
    End Sub

    <Task()> Public Shared Sub CambioIDCondicionPago(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf AsignarMotivoNoAsegurado, data, services)
    End Sub

    <Task()> Public Shared Sub CambioIDFormaPago(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf AsignarMotivoNoAsegurado, data, services)
    End Sub

    <Task()> Public Shared Sub AsignarMotivoNoAsegurado(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf AsignarMotivoNoAseguradoIProp, data.Current, services)
    End Sub


    <Task()> Public Shared Sub AsignarMotivoNoAseguradoIProp(ByVal data As IPropertyAccessor, ByVal services As ServiceProvider)
        Dim IDMotivoNoAsegurado As String
        ' data.Current("IDMotivoNoAsegurado") = System.DBNull.Value

        Dim fCliente As New Filter
        fCliente.Add(New StringFilterItem("IDCliente", data("IDCliente")))
        Dim dtHistorico As DataTable = New HistoricoRiesgoCliente().Filter(fCliente, "FechaConcedido DESC", "TOP 2 *")
        If dtHistorico.Rows.Count > 0 Then
            Dim ImporteConcedido As Double = dtHistorico.Rows(0)("ImporteConcedido")
            IDMotivoNoAsegurado = dtHistorico.Rows(0)("IDMotivoNoAsegurado") & String.Empty
            If ImporteConcedido = 0 AndAlso Nz(dtHistorico.Rows(0)("FechaConcedido"), cnMinDate) = cnMinDate AndAlso dtHistorico.Rows.Count > 1 Then
                ImporteConcedido = dtHistorico.Rows(1)("ImporteConcedido")
                IDMotivoNoAsegurado = dtHistorico.Rows(1)("IDMotivoNoAsegurado") & String.Empty
            End If
        End If

        If Length(IDMotivoNoAsegurado) = 0 Then
            If Length(data("IDPais")) > 0 Then
                Dim dtPais As DataTable = New Pais().Filter(New StringFilterItem("IDPais", data("IDPais")))
                If dtPais.Rows.Count > 0 Then
                    IDMotivoNoAsegurado = dtPais.Rows(0)("IDMotivoNoAsegurado") & String.Empty
                End If
            End If
        End If


        If Length(IDMotivoNoAsegurado) = 0 Then
            If Length(data("IDTipoCliente")) > 0 Then
                Dim dtTipoClte As DataTable = New TipoCliente().Filter(New StringFilterItem("IDTipoCliente", data("IDTipoCliente")))
                If dtTipoClte.Rows.Count > 0 Then
                    IDMotivoNoAsegurado = dtTipoClte.Rows(0)("IDMotivoNoAsegurado") & String.Empty
                End If
            End If
        End If

        If Length(IDMotivoNoAsegurado) = 0 Then
            If Length(data("IDCondicionPago")) > 0 Then
                Dim dtCondPago As DataTable = New CondicionPago().Filter(New StringFilterItem("IDCondicionPago", data("IDCondicionPago")))
                If dtCondPago.Rows.Count > 0 Then
                    IDMotivoNoAsegurado = dtCondPago.Rows(0)("IDMotivoNoAsegurado") & String.Empty
                End If
            End If
        End If

        If Length(IDMotivoNoAsegurado) = 0 Then
            If Length(data("IDFormaPago")) > 0 Then
                Dim dtFormaPago As DataTable = New FormaPago().Filter(New StringFilterItem("IDFormaPago", data("IDFormaPago")))
                If dtFormaPago.Rows.Count > 0 Then
                    IDMotivoNoAsegurado = dtFormaPago.Rows(0)("IDMotivoNoAsegurado") & String.Empty
                End If
            End If
        End If

        data("IDMotivoNoAsegurado") = IDMotivoNoAsegurado
    End Sub



    <Task()> Public Shared Sub CambioAseguradora(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDAseguradora")) > 0 Then
            Dim dtAseg As DataTable = New Aseguradora().SelOnPrimaryKey(data.Current("IDAseguradora"))
            If dtAseg.Rows.Count > 0 Then
                data.Current("NPolizaAseg") = dtAseg.Rows(0)("NumPoliza")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CalculoRiesgoConcedido(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        data.Current("RiesgoConcedido") = Nz(data.Current("RiesgoInterno"), 0) + Nz(data.Current("LimiteCapitalAsegurado"), 0)
    End Sub

    <Task()> Public Shared Sub CambioDiapago(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        Dim DtDiaPago As DataTable = New DiaPago().SelOnPrimaryKey(data.Value)
        If Not DtDiaPago Is Nothing AndAlso DtDiaPago.Rows.Count = 0 Then
            ApplicationService.GenerateError("El Día de Pago introducido no existe.")
        End If
    End Sub

    <Task()> Public Shared Sub CambioCifPaisCliente(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Value) > 0 AndAlso Length(data.Current("CifCliente")) > 0 Then
            ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ValidaDocumentoIdentificativo, data.Current, services)
        End If
        If data.ColumnName = "IDPais" Then
            ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf AsignarMotivoNoAsegurado, data, services)
        End If
    End Sub

    <Task()> Public Shared Sub ValidaDocumentoIdentificativo(ByVal data As IPropertyAccessor, ByVal services As ServiceProvider)
        If data("TipoDocIdentidad") = enumTipoDocIdent.NIF Or data("TipoDocIdentidad") = enumTipoDocIdent.CertificiadoResiFiscal Then ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf TratarCIFCliente, data, services)
        Dim info As New DataDocIdentificacion(data("CifCliente"), data("IDPais"), data("TipoDocIdentidad"))
        info = ProcessServer.ExecuteTask(Of DataDocIdentificacion, DataDocIdentificacion)(AddressOf Comunes.ValidarDocumentoIdentificativo, info, services)
        If Not info.EsCorrecto Then
            ApplicationService.GenerateError("El Documento introducido no es un '|'. Intoduzca uno correcto o cambie de tipo de documento", info.TipoDocumento)
        End If
    End Sub

    <Task()> Public Shared Sub CambioOperarioBloqueo(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim drOperario As DataRow = New Operario().GetItemRow(data.Value)
            If Not drOperario Is Nothing Then
                data.Current("DescOperarioBloqueo") = drOperario("DescOperario") & String.Empty
                data.Current("Bloqueado") = True
            End If
        Else
            If Length(data.Current("DescOperarioBloqueo")) = 0 Then
                data.Current("Bloqueado") = False
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioDescOperarioBloqueo(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            If data.Current("Bloqueado") = False Then
                data.Current("Bloqueado") = True
            End If
        Else
            If Length(data.Current("IDOperarioBloqueo")) = 0 Then
                data.Current("Bloqueado") = False
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioBloqueado(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Not data.Value Then
            If Length(data.Current("DescOperarioBloqueo")) > 0 Then data.Current("DescOperarioBloqueo") = String.Empty
            If Length(data.Current("IDOperarioBloqueo")) > 0 Then data.Current("IDOperarioBloqueo") = String.Empty
        End If
    End Sub

    <Task()> Public Shared Sub CambioDiaFacturacion(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            If data.Value > 31 OrElse data.Value < 0 Then
                ApplicationService.GenerateError("El día introducido no es válido.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioCodPostal(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim infoCP As New CodPostalInfo(CStr(data.Value), data.Current("IDPais") & String.Empty)
            If Length(infoCP.DescPoblacion) > 0 Then data.Current("Poblacion") = infoCP.DescPoblacion
            If Length(infoCP.DescProvincia) > 0 Then data.Current("Provincia") = infoCP.DescProvincia
            If Length(infoCP.IDPais) > 0 Then data.Current("IDPais") = infoCP.IDPais
        End If
    End Sub

#End Region

#Region "Funciones ValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDescCliente)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarCodPostal)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarCifCliente)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarPais)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarCContable)
        ' validateProcess.AddTask(Of DataRow)(AddressOf ValidarGrupo)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarBloqueos)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarFechaAlta)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDiaFacturacion)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarGrupoCliente)
    End Sub

    <Task()> Public Shared Sub ValidarDescCliente(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("DescCliente")) = 0 Then ApplicationService.GenerateError("La Descripción del Cliente es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarCodPostal(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("CodPostal")) = 0 Then ApplicationService.GenerateError("Código Postal es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarCifCliente(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("TipoDocIdentidad")) = 0 Then ApplicationService.GenerateError("Es necesario especificar el tipo de documento.")
        If Length(data("CifCliente")) = 0 Then ApplicationService.GenerateError("El documento de Identificación es un dato obligatorio.")
        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ValidaDocumentoIdentificativo, New DataRowPropertyAccessor(data), services)
    End Sub

    <Task()> Public Shared Sub ValidarPais(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IdPais")) = 0 Then ApplicationService.GenerateError("País es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarMoneda(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDMoneda")) = 0 Then ApplicationService.GenerateError("Moneda es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarCContable(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim ClsParam As New Parametro
        If ClsParam.Contabilidad Then
            If Length(data("CCCliente")) = 0 Then ApplicationService.GenerateError("CCCliente es un dato obligatorio.")
        End If
    End Sub

    '<Task()> Public Shared Sub ValidarGrupo(ByVal data As DataRow, ByVal services As ServiceProvider)
    '    If data("EmpresaGrupo") <> 0 Then
    '        If IsDBNull(data("BaseDatos")) Then ApplicationService.GenerateError("El Nombre de Empresa del Grupo es Obligatorio.")
    '    End If
    'End Sub

    <Task()> Public Shared Sub ValidarBloqueos(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Not IsDBNull(data("IDOperarioBloqueo")) Then
            If Length(data("IDOperarioBloqueo")) > 0 Then
                If Not Nz(data("Bloqueado"), False) Then ApplicationService.GenerateError("Hay seleccionado un responsable de bloqueo pero el cliente no está bloqueado.")
            ElseIf Nz(data("Bloqueado"), False) Then
                ApplicationService.GenerateError("No ha seleccionado ningún Responsable del Bloqueo de los Clientes.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarFechaAlta(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("FechaAlta")) = 0 Then data("FechaAlta") = Today.Date
    End Sub

    <Task()> Public Shared Sub ValidarDiaFacturacion(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("DiaFacturacion")) > 0 AndAlso (data("DiaFacturacion") > 31 OrElse data("DiaFacturacion") < 0) Then
            ApplicationService.GenerateError("El día introducido no es válido.")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarGrupoCliente(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IdGrupoCliente")) > 0 AndAlso Length(data("IdCliente")) > 0 Then
            If data("IDCliente") = data("IdGrupoCliente") Then
                ApplicationService.GenerateError("Un cliente no puede ser cabecera y miembro de un grupo al mismo tiempo.")  'El cliente no puede ser cabecera y miembro de un grupo al mismo tiempo
            Else
                Dim dtGrupo As DataTable = New Cliente().Filter(New FilterItem("IdCliente", FilterOperator.Equal, data("IdGrupoCliente")))
                If Not dtGrupo Is Nothing AndAlso dtGrupo.Rows.Count > 0 Then
                    If Length(dtGrupo.Rows(0)("IdGrupoCliente")) > 0 Then
                        ApplicationService.GenerateError("El cliente no puede ser cabecera de grupo porque pertenece a otro grupo.")  'Si el IDGrupoCliente de la cabecera es <> vbnullstring , el cliente ya pertenece a un grupo
                    Else
                        Dim dtHijos As DataTable = New Cliente().Filter(New FilterItem("IDGrupoCliente", FilterOperator.Equal, data("IDCliente")))
                        If Not dtHijos Is Nothing AndAlso dtHijos.Rows.Count > 0 Then
                            ApplicationService.GenerateError("El cliente no puede pertenecer al grupo porque es cabecera de otro grupo.")  'Si ese cliente ya es una cabecera de grupo
                        End If
                    End If
                End If
            End If
        Else
            data("GrupoDireccion") = False
            data("GrupoTarifa") = False
            data("GrupoFactura") = False
            data("GrupoArticulo") = False
        End If
        Dim StCIFRepetido As New DataCIFRepetido
        StCIFRepetido.Documento = data("CifCliente") & String.Empty
        StCIFRepetido.TipoDocumento = Nz(data("TipoDocIdentidad"), 0)
        StCIFRepetido.IDCliente = data("IDCliente") & String.Empty
        StCIFRepetido.IDGrupoCliente = data("IDGrupoCliente") & String.Empty
        If ProcessServer.ExecuteTask(Of DataCIFRepetido, Boolean)(AddressOf Cliente.CIFRepetido, StCIFRepetido, services) Then
            ApplicationService.GenerateError("Ya existe un Cliente con el Documento '|'.|Sólo se permiten clientes con el mismo Documento si pertenecen al mismo grupo.", data("CifCliente"), vbCrLf)
        End If
    End Sub

#End Region

#Region "Funciones UpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClavePrimaria)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarContenedoresCajas)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarCondicionEnvio)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarFormaEnvio)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarRiesgoPorCIF)
        updateProcess.AddTask(Of DataRow)(AddressOf Business.General.Comunes.UpdateEntityRow)
        updateProcess.AddTask(Of DataRow)(AddressOf Business.General.Comunes.MarcarComoActualizado)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClienteDireccion)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClienteComprador)
    End Sub

    <Task()> Public Shared Sub AsignarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDCliente")) = 0 Then
                ' Caso en el que no esté introducido un código de cliente
                If Length(data("IdContador")) > 0 Then
                    data("IDCliente") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, data("IDContador"), services)
                Else
                    Dim StDatos As New Contador.DatosDefaultCounterValue
                    StDatos.row = data
                    StDatos.EntityName = "Cliente"
                    StDatos.FieldName = "IDCliente"
                    ProcessServer.ExecuteTask(Of Contador.DatosDefaultCounterValue)(AddressOf Contador.LoadDefaultCounterValue, StDatos, services)
                    If Length(data("IDContador")) > 0 Then
                        data("IDCliente") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, data("IDContador"), services)
                    Else
                        ApplicationService.GenerateError("No se ha configurado contador predeterminado para Clientes.")
                    End If
                End If
            Else
                ' Caso en el que el código de cliente proviene de la asignación de contador,
                '  en caso contrario sería código manual por lo que no hay que mover el contador
                If Length(data("IdContador")) > 0 Then data("IDCliente") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, data("IDContador"), services)
                Dim dtCliente As DataTable = New Cliente().SelOnPrimaryKey(data("IDCliente"))
                If Not IsNothing(dtCliente) AndAlso dtCliente.Rows.Count > 0 Then
                    ApplicationService.GenerateError("Ese cliente ya existe en la Base de Datos")
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarContenedoresCajas(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDTipoEtiquetaContenedor")) = 0 Or Length(data("IDTipoEtiquetaCaja")) = 0 Then
                Dim E As New TipoEtiqueta
                Dim FilEti As New Filter(FilterUnionOperator.Or)
                FilEti.Add("PredeterminadaContenedor", FilterOperator.Equal, 1)
                FilEti.Add("PredeterminadaCaja", FilterOperator.Equal, 1)
                Dim dtE As DataTable = E.Filter(FilEti)
                If Not dtE Is Nothing AndAlso dtE.Rows.Count > 0 Then
                    If Length(data("IDTipoEtiquetaContenedor")) = 0 Then
                        Dim drE As DataRow() = dtE.Select("PredeterminadaContenedor=1")
                        If drE.Length > 0 Then
                            data("IDTipoEtiquetaContenedor") = drE(0)("IDTipoEtiqueta")
                        End If
                    End If
                    If Length(data("IDTipoEtiquetaCaja")) = 0 Then
                        Dim drE As DataRow() = dtE.Select("PredeterminadaCaja=1")
                        If drE.Length > 0 Then
                            data("IDTipoEtiquetaCaja") = drE(0)("IDTipoEtiqueta")
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarCondicionEnvio(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDCondicionEnvio")) = 0 Then data("IDCondicionEnvio") = New Parametro().CondicionEnvio
    End Sub

    <Task()> Public Shared Sub AsignarFormaEnvio(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDFormaEnvio")) = 0 Then data("IDFormaEnvio") = New Parametro().FormaEnvio
    End Sub

    <Task()> Public Shared Sub AsignarRiesgoPorCIF(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added OrElse (data.RowState = DataRowState.Modified AndAlso data("CIFCliente") & String.Empty <> data("CIFCliente", DataRowVersion.Original) & String.Empty) Then
            Dim c As New Cliente
            Dim IDAseguradora As String = String.Empty
            Dim ImporteConcedido As Double = 0
            Dim NumPoliza As String = String.Empty
            Dim IDMotivoNoAsegurado As String '= cboMotivoNoAsegurado.Value & String.Empty

            If Length(data("CIFCliente")) > 0 Then
                Dim fCliente As New Filter
                fCliente.Add(New StringFilterItem("CIFCliente", data("CIFCliente")))
                Dim dtHistorico As DataTable = New HistoricoRiesgoCliente().Filter(fCliente, "FechaConcedido DESC", "TOP 2 *")
                If dtHistorico.Rows.Count > 0 Then

                    ImporteConcedido = dtHistorico.Rows(0)("ImporteConcedido")
                    IDMotivoNoAsegurado = dtHistorico.Rows(0)("IDMotivoNoAsegurado") & String.Empty
                    IDAseguradora = dtHistorico.Rows(0)("IDAseguradora") & String.Empty
                    NumPoliza = dtHistorico.Rows(0)("NumPoliza") & String.Empty
                    If ImporteConcedido = 0 AndAlso Nz(dtHistorico.Rows(0)("FechaConcedido"), cnMinDate) = cnMinDate AndAlso dtHistorico.Rows.Count > 1 Then
                        ImporteConcedido = dtHistorico.Rows(1)("ImporteConcedido")
                        IDMotivoNoAsegurado = dtHistorico.Rows(1)("IDMotivoNoAsegurado") & String.Empty
                        IDAseguradora = dtHistorico.Rows(1)("IDAseguradora") & String.Empty
                        NumPoliza = dtHistorico.Rows(1)("NumPoliza") & String.Empty
                    End If
                    If Length(IDAseguradora) > 0 Then
                        data("IDAseguradora") = IDAseguradora
                    Else
                        data("IDAseguradora") = System.DBNull.Value
                    End If
                    If Length(NumPoliza) > 0 Then
                        data("NPolizaAseg") = NumPoliza
                    Else
                        data("NPolizaAseg") = System.DBNull.Value
                    End If
                    data = c.ApplyBusinessRule("LimiteCapitalAsegurado", ImporteConcedido, data, Nothing)
                    If Length(IDMotivoNoAsegurado) > 0 Then
                        data("IDMotivoNoAsegurado") = IDMotivoNoAsegurado
                    Else
                        data("IDMotivoNoAsegurado") = System.DBNull.Value
                    End If
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarClienteDireccion(ByVal data As DataRow, ByVal services As ServiceProvider)
        'Controlar que no me llega una de presentación
        Dim upPresenta As UpdatePackage = services.GetService(Of UpdatePackage)()
        Dim dt As DataTable = upPresenta.Item(GetType(ClienteDireccion).Name).First
        Dim adr() As DataRow
        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            adr = dt.Select("IDCliente='" & data("IDCliente") & "'")
        End If
        If adr Is Nothing OrElse adr.Length = 0 Then
            If data.RowState = DataRowState.Added Then
                'Hay que guardar 1 registro con 3 checks:   .- Dirección de Envio
                'Todos marcados y predeterminados           .- Dirección de Facturación
                '                                           .- Dirección de Giro
                dt = New ClienteDireccion().AddNewForm
                dt.Rows(0)("IDCliente") = data("IDCliente")
                dt.Rows(0)("IDCAE") = data("IDCAE")
                dt.Rows(0)("RazonSocial") = data("RazonSocial")
                dt.Rows(0)("Direccion") = data("Direccion")
                dt.Rows(0)("CodPostal") = data("CodPostal")
                dt.Rows(0)("Poblacion") = data("Poblacion")
                dt.Rows(0)("Provincia") = data("Provincia")
                dt.Rows(0)("IDPais") = data("IDPais")
                dt.Rows(0)("CIfCliente") = data("CIfCliente")
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

    <Task()> Public Shared Sub AsignarClienteComprador(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IdCliente")) > 0 Then
                ' Se crea o modifica el Contratante
                Dim c As BusinessHelper
                Try
                    c = BusinessHelper.CreateBusinessObject("ObraPromoClienteComprador")
                Catch ex As Exception
                    Exit Sub
                End Try
                Dim dtNew As DataTable = c.AddNewForm
                dtNew.Rows(0)("IDCliente") = data("IdCliente")
                dtNew.Rows(0)("DescComprador") = data("DescCliente")
                dtNew.Rows(0)("DniContacto") = data("CifCliente")
                dtNew.Rows(0)("Direccion") = data("Direccion")
                dtNew.Rows(0)("CodPostal") = data("CodPostal")
                dtNew.Rows(0)("Poblacion") = data("Poblacion")
                dtNew.Rows(0)("Provincia") = data("Provincia")
                dtNew.Rows(0)("IdPais") = data("IdPais")
                dtNew.Rows(0)("Telefono") = data("Telefono")
                dtNew.Rows(0)("Fax") = data("Fax")
                dtNew.Rows(0)("eMail") = data("eMail")
                c.Update(dtNew)
            End If
        End If
    End Sub

#End Region

#Region "Funciones DeleteTasks"

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf ComprobarCabeceraGrupo)
        deleteProcess.AddTask(Of DataRow)(AddressOf ComprobarCRM)
    End Sub

    <Task()> Public Shared Sub ComprobarCabeceraGrupo(ByVal data As DataRow, ByVal services As ServiceProvider)
        'Si el cliente es cabecera de grupo, no se puede eliminar
        Dim dt As DataTable = New Cliente().Filter(New StringFilterItem("IDGrupoCliente", data("IDCliente")))
        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            ApplicationService.GenerateError("Este cliente no se puede eliminar porque es cabecera de grupo. ")
        End If
    End Sub

    <Task()> Public Shared Sub ComprobarCRM(ByVal data As DataRow, ByVal services As ServiceProvider)
        'Control del CRM: si es un cliente que ha venido convertido desde empresa, para borrarlo de los datos de Empresa del CRM
        Dim ClsCRMEmp As BusinessHelper = BusinessHelper.CreateBusinessObject("EmpresaCategoria")
        Dim DtCRM As DataTable = CType(ClsCRMEmp, ICliente).ComprobarCategoriaEmpresaCRMCliente(data("IDEmpresa") & String.Empty, "Cliente", data("IDCliente"))
        If Not DtCRM Is Nothing Then ClsCRMEmp.Delete(DtCRM)
    End Sub

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Function Idioma(ByVal strID As String, ByVal services As ServiceProvider) As String
        If Length(strID & String.Empty) > 0 Then
            Dim dtMe As DataTable = New Cliente().SelOnPrimaryKey(strID)
            If Not dtMe Is Nothing AndAlso dtMe.Rows.Count > 0 Then
                Return dtMe.Rows(0)("IDIdioma") & String.Empty
            End If
        End If
        Return String.Empty
    End Function

    <Task()> Public Shared Function Nacional(ByVal strIDCliente As String, ByVal services As ServiceProvider) As Boolean
        Nacional = False
        If Length(strIDCliente) > 0 Then
            Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
            Dim Clteinfo As ClienteInfo = Clientes.GetEntity(strIDCliente)
            If Not Clteinfo Is Nothing AndAlso Length(Clteinfo.IDCliente) > 0 Then
                'Si el Cliente no tiene País se le considerará como Nacional.
                If Length(Clteinfo.Pais) > 0 Then
                    Nacional = Not Clteinfo.Extranjero
                Else
                    Nacional = True
                End If
            End If
        End If
    End Function

    <Task()> Public Shared Function Grupo(ByVal DatosGrupoCliente As DataGrupoCliente, ByVal services As ServiceProvider) As String
        Dim objFilter As New Filter
        objFilter.Add(New StringFilterItem("IDCliente", DatosGrupoCliente.IDCliente))
        Select Case DatosGrupoCliente.TipoGrupo
            Case "Direccion"
                objFilter.Add(New BooleanFilterItem("GrupoDireccion", True))
            Case "Factura"
                objFilter.Add(New BooleanFilterItem("GrupoFactura", True))
        End Select

        Dim dtGrupo As DataTable = New Cliente().Filter(objFilter)
        If Not IsNothing(dtGrupo) AndAlso dtGrupo.Rows.Count > 0 Then
            Return dtGrupo.Rows(0)("IdGrupoCliente") & String.Empty
        End If
        Return String.Empty
    End Function

    <Task()> Public Shared Sub TratarCIFCliente(ByVal data As IPropertyAccessor, ByVal services As ServiceProvider)
        If Length(data("IDPais")) > 0 AndAlso ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf Pais.Nacional, data("IDPais"), services) Then
            If Length(data("CifCliente")) > 0 Then
                Dim blnCancel As Boolean
                'ValidaCIF(data, blnCancel)
                If Not blnCancel Then
                    Dim StCifRepetido As New DataCIFRepetido
                    StCifRepetido.Documento = data("CifCliente") & String.Empty
                    StCifRepetido.TipoDocumento = Nz(data("TipoDocIdentidad"), 0)
                    StCifRepetido.IDCliente = data("IDCliente")
                    StCifRepetido.IDGrupoCliente = data("IdGrupoCliente") & String.Empty
                    If ProcessServer.ExecuteTask(Of DataCIFRepetido, Boolean)(AddressOf Cliente.CIFRepetido, StCifRepetido, services) Then
                        ApplicationService.GenerateError("Ya existe un Cliente con el Documento '|'.|Sólo se permiten clientes con el mismo Documento si pertenecen al mismo grupo.", data("CifCliente"), vbCrLf)
                    End If
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Function CIFRepetido(ByVal data As DataCIFRepetido, ByVal services As ServiceProvider) As Boolean
        If Length(data.Documento) > 0 Then
            Dim f As New Filter
            f.Add(New StringFilterItem("CifCliente", data.Documento))
            f.Add(New NumberFilterItem("TipoDocIdentidad", data.TipoDocumento))
            f.Add(New StringFilterItem("IDCliente", FilterOperator.NotEqual, data.IDCliente))
            Dim dt As DataTable = New Cliente().Filter(f)
            If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
                f.Clear()
                f.UnionOperator = FilterUnionOperator.Or
                If Length(data.IDGrupoCliente) > 0 Then
                    f.Add(New StringFilterItem("IDGrupoCliente", data.IDGrupoCliente))
                    f.Add(New StringFilterItem("IDCliente", data.IDGrupoCliente))
                Else
                    f.Add(New StringFilterItem("IDGrupoCliente", data.IDCliente))
                End If
                Dim WhereClteGrupo As String = f.Compose(New AdoFilterComposer)
                Dim adr() As DataRow = dt.Select(WhereClteGrupo)
                If Not IsNothing(adr) AndAlso adr.Length > 0 Then
                    CIFRepetido = False
                Else
                    CIFRepetido = True
                End If
            End If
        End If
    End Function

    <Task()> Public Shared Function ObtenerXDataBase(ByVal data As Object, ByVal services As ServiceProvider) As DataTable
        Return New BE.DataEngine().Filter("xDataBase", "", "", , , True)
    End Function

    <Task()> Public Shared Function ComprobarBloqueoCliente(ByVal data As String, ByVal services As ServiceProvider) As Boolean
        Dim DtCliente As DataTable = New Cliente().SelOnPrimaryKey(data)
        If Not DtCliente Is Nothing AndAlso DtCliente.Rows.Count > 0 Then
            Return DtCliente.Rows(0)("Bloqueado")
        End If
    End Function

    <Task()> Public Shared Function ComprobarBloqueoArticuloCliente(ByVal data As DataBloqArtClie, ByVal services As ServiceProvider) As Boolean
        Dim ClsArtClie As New ArticuloCliente
        Dim FilArtClie As New Filter
        FilArtClie.Add("IDCliente", FilterOperator.Equal, data.IDCliente, FilterType.String)
        If Length(data.IDArticulo) > 0 Then
            FilArtClie.Add("IDArticulo", FilterOperator.Equal, data.IDArticulo, FilterType.String)
        ElseIf Length(data.RefCliente) > 0 Then
            FilArtClie.Add("RefCliente", FilterOperator.Equal, data.RefCliente, FilterType.String)
        End If
        Dim DtArtClie As DataTable = ClsArtClie.Filter(FilArtClie)
        If Not DtArtClie Is Nothing AndAlso DtArtClie.Rows.Count > 0 Then
            Return DtArtClie.Rows(0)("Bloqueado")
        End If
    End Function

    <Task()> Public Shared Function GetParamsRiesgoCliente(ByVal IDCliente As String, ByVal services As ServiceProvider) As DataParamRiesgoCliente
        Dim p As New Parametro
        Dim datParams As New DataParamRiesgoCliente
        datParams.TipoAlbaranDeDeposito = p.TipoAlbaranDeDeposito
        datParams.TipoAlbaranRetornoAlquiler = p.TipoAlbaranRetornoAlquiler
        datParams.GestionAlquiler = p.AplicacionGestionAlquiler


        If Length(IDCliente) > 0 Then
            datParams.IDCliente = IDCliente

            Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
            Dim ClteInfo As ClienteInfo = Clientes.GetEntity(IDCliente)
            datParams.RiesgoInterno = ClteInfo.RiesgoInterno
            datParams.RiesgoConcedido = ClteInfo.RiesgoConcedido
            datParams.RiesgoGrupo = ClteInfo.GrupoRiesgo
            datParams.IDClienteMatriz = ClteInfo.GrupoCliente
            datParams.IDProveedorAsociado = ClteInfo.IDProveedorAsociado
            datParams.LimiteCapitalAsegurado = ClteInfo.LimiteCapitalAsegurado

            If datParams.RiesgoGrupo Then
                ReDim Preserve datParams.IDClientesGrupo(datParams.IDClientesGrupo.Length)
                datParams.IDClientesGrupo(datParams.IDClientesGrupo.Length - 1) = datParams.IDCliente

                If Length(datParams.IDClienteMatriz) > 0 Then
                    ReDim Preserve datParams.IDClientesGrupo(datParams.IDClientesGrupo.Length)
                    datParams.IDClientesGrupo(datParams.IDClientesGrupo.Length - 1) = datParams.IDClienteMatriz

                    Dim ClteGrupoInfo As ClienteInfo = Clientes.GetEntity(datParams.IDClienteMatriz)
                    datParams.RiesgoConcedido = ClteGrupoInfo.RiesgoConcedido
                    datParams.RiesgoInterno = ClteGrupoInfo.RiesgoInterno
                    datParams.IDProveedorAsociado = ClteGrupoInfo.IDProveedorAsociado
                    datParams.LimiteCapitalAsegurado = ClteGrupoInfo.LimiteCapitalAsegurado

                Else
                    Dim ClteGrupoInfo As ClienteInfo = Clientes.GetEntity(datParams.IDCliente)
                    datParams.RiesgoConcedido = ClteGrupoInfo.RiesgoConcedido
                    datParams.RiesgoInterno = ClteGrupoInfo.RiesgoInterno
                    datParams.LimiteCapitalAsegurado = ClteGrupoInfo.LimiteCapitalAsegurado
                End If

                Dim f As New Filter
                If Length(datParams.IDClienteMatriz) > 0 Then
                    f.Add(New StringFilterItem("IDGrupoCliente", datParams.IDClienteMatriz))
                    f.Add(New StringFilterItem("IDCliente", FilterOperator.NotEqual, datParams.IDCliente))
                    f.Add(New BooleanFilterItem("GrupoRiesgo", True))
                Else
                    f.Add(New StringFilterItem("IDGrupoCliente", datParams.IDCliente))
                    f.Add(New BooleanFilterItem("GrupoRiesgo", True))
                End If

                Dim dtClientesGrupo As DataTable = New Cliente().Filter(f)
                If dtClientesGrupo.Rows.Count > 0 Then
                    For Each drClte As DataRow In dtClientesGrupo.Rows
                        ReDim Preserve datParams.IDClientesGrupo(datParams.IDClientesGrupo.Length)
                        datParams.IDClientesGrupo(datParams.IDClientesGrupo.Length - 1) = drClte("IDCliente")
                    Next
                End If
            Else
                '//Si no tiene Riesgo grupo, puede ser el cliente matriz.

                Dim f As New Filter

                f.Add(New StringFilterItem("IDGrupoCliente", datParams.IDCliente))
                f.Add(New BooleanFilterItem("GrupoRiesgo", True))

                Dim dtClientesGrupo As DataTable = New Cliente().Filter(f)
                If dtClientesGrupo.Rows.Count > 0 Then

                    ReDim Preserve datParams.IDClientesGrupo(datParams.IDClientesGrupo.Length)
                    datParams.IDClientesGrupo(datParams.IDClientesGrupo.Length - 1) = datParams.IDCliente

                    datParams.RiesgoConcedido = ClteInfo.RiesgoConcedido
                    datParams.RiesgoInterno = ClteInfo.RiesgoInterno
                    datParams.LimiteCapitalAsegurado = ClteInfo.LimiteCapitalAsegurado

                    For Each drClte As DataRow In dtClientesGrupo.Rows
                        ReDim Preserve datParams.IDClientesGrupo(datParams.IDClientesGrupo.Length)
                        datParams.IDClientesGrupo(datParams.IDClientesGrupo.Length - 1) = drClte("IDCliente")
                    Next
                End If
            End If
        End If

        If Length(datParams.IDProveedorAsociado) > 0 Then
            Dim dtProv As DataTable = New Proveedor().SelOnPrimaryKey(datParams.IDProveedorAsociado)
            If dtProv.Rows.Count > 0 Then
                datParams.DescProveedorAsociado = dtProv.Rows(0)("DescProveedor") & String.Empty

                Dim datRiesgoProv As RiesgoProveedor = ProcessServer.ExecuteTask(Of String, RiesgoProveedor)(AddressOf GetRiesgoProveedorAsociado, datParams.IDProveedorAsociado, services)
                datParams.PdteFacturar = datRiesgoProv.PdteFacturar
                datParams.PagosNoPagados = datRiesgoProv.PagosNoPagados
            End If
        End If

        Return datParams
    End Function

    <Task()> Public Shared Function ObtenerRiesgoCliente(ByVal data As DataRiesgoCliente, ByVal services As ServiceProvider) As RiesgoCliente
        Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
        Dim Riesgo As New RiesgoCliente
        If AppParams.RiesgoCliente Then
            Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
            Dim ClteInfo As ClienteInfo = Clientes.GetEntity(data.IDCliente)
            Riesgo.IDCliente = data.IDCliente
            Riesgo.DescCliente = ClteInfo.DescCliente
            Riesgo.RiesgoConcedido = ClteInfo.RiesgoConcedido
            Riesgo.RiesgoGrupo = ClteInfo.GrupoRiesgo
            Riesgo.RiesgoInterno = ClteInfo.RiesgoInterno
            If Riesgo.RiesgoGrupo AndAlso Length(ClteInfo.GrupoCliente) > 0 Then
                Dim ClteGrupoInfo As ClienteInfo = Clientes.GetEntity(ClteInfo.GrupoCliente)
                Riesgo.RiesgoConcedido = ClteGrupoInfo.RiesgoConcedido
            End If
            Riesgo.SuperaRiesgo = False
            If ClteInfo.Riesgo Then
                Riesgo.Riesgo = True

                '//Cálculo riesgo
                If Not ClteInfo.RiesgoConsumidoCalculado Then
                    Dim ViewName As String

                    ClteInfo.RiesgoConsumido = 0
                    Dim fCliente As New Filter
                    If Riesgo.RiesgoGrupo Then
                        If Length(ClteInfo.GrupoCliente) > 0 Then
                            Dim dtClientesGrupo As DataTable = New Cliente().Filter(New StringFilterItem("IDGrupoCliente", ClteInfo.GrupoCliente))
                            If dtClientesGrupo.Rows.Count > 0 Then
                                Dim fClteGrupo As New Filter(FilterUnionOperator.Or)
                                fClteGrupo.Add(New StringFilterItem("IDCliente", ClteInfo.GrupoCliente))
                                For Each drClte As DataRow In dtClientesGrupo.Rows
                                    fClteGrupo.Add(New StringFilterItem("IDCliente", drClte("IDCliente")))
                                Next
                                fCliente.Add(fClteGrupo)
                            Else
                                fCliente.Add(New StringFilterItem("IDCliente", data.IDCliente))
                            End If
                        Else
                            fCliente.Add(New StringFilterItem("IDCliente", data.IDCliente))
                        End If
                    Else
                        '//Si el cliente es de grupo, buscamos el riesgo de sus hijos que tengan la marca de Riesgo.
                        Dim fCltesRiesgoGrupo As New Filter
                        fCltesRiesgoGrupo.Add(New StringFilterItem("IDGrupoCliente", data.IDCliente))
                        fCltesRiesgoGrupo.Add(New BooleanFilterItem("GrupoRiesgo", True))

                        Dim dtClientesGrupo As DataTable = New Cliente().Filter(fCltesRiesgoGrupo)
                        If dtClientesGrupo.Rows.Count > 0 Then
                            Dim fClteGrupo As New Filter(FilterUnionOperator.Or)
                            fClteGrupo.Add(New StringFilterItem("IDCliente", data.IDCliente))
                            For Each drClte As DataRow In dtClientesGrupo.Rows
                                fClteGrupo.Add(New StringFilterItem("IDCliente", drClte("IDCliente")))
                            Next
                            fCliente.Add(fClteGrupo)
                        Else
                            fCliente.Add(New StringFilterItem("IDCliente", data.IDCliente))
                        End If
                    End If

                    ViewName = "VFrmMntoRiesgoClientePedidosDeposito"
                    Dim dtRiesgoPedDep As DataTable = AdminData.GetData(ViewName, fCliente)
                    If dtRiesgoPedDep.Rows.Count > 0 Then
                        ClteInfo.RiesgoConsumido += dtRiesgoPedDep.Compute("SUM(Importe)", Nothing)
                    End If

                    ViewName = "VFrmMntoRiesgoClientePedidosSinDeposito"
                    Dim dtRiesgoPedSinDep As DataTable = AdminData.GetData(ViewName, fCliente)
                    If dtRiesgoPedSinDep.Rows.Count > 0 Then
                        ClteInfo.RiesgoConsumido += dtRiesgoPedSinDep.Compute("SUM(Importe)", Nothing)
                    End If

                    ViewName = "VFrmMntoRiesgoClienteAlbaranes"
                    Dim dtRiesgoAlbaranes As DataTable = AdminData.GetData(ViewName, fCliente)
                    If dtRiesgoAlbaranes.Rows.Count > 0 Then
                        ClteInfo.RiesgoConsumido += dtRiesgoAlbaranes.Compute("SUM(Importe)", Nothing)
                    End If

                    ViewName = "VFrmMntoRiesgoClienteFacturas"
                    Dim dtRiesgoFacturas As DataTable = AdminData.GetData(ViewName, fCliente)
                    If dtRiesgoFacturas.Rows.Count > 0 Then
                        ClteInfo.RiesgoConsumido += dtRiesgoFacturas.Compute("SUM(Importe)", Nothing)
                    End If

                    ViewName = "VFrmMntoRiesgoClienteCobrosFactura"
                    Dim dtRiesgoCobroFras As DataTable = AdminData.GetData(ViewName, fCliente)
                    If dtRiesgoCobroFras.Rows.Count > 0 Then
                        ClteInfo.RiesgoConsumido += dtRiesgoCobroFras.Compute("SUM(Importe)", Nothing)
                    End If

                    ViewName = "VFrmMntoRiesgoClienteCobrosManuales"
                    Dim dtRiesgoCobrosManuales As DataTable = AdminData.GetData(ViewName, fCliente)
                    If dtRiesgoCobrosManuales.Rows.Count > 0 Then
                        ClteInfo.RiesgoConsumido += dtRiesgoCobrosManuales.Compute("SUM(Importe)", Nothing)
                    End If

                    ViewName = "vCIClientesRiesgoObrasCertificaciones"
                    Dim dtRiesgoObrasCertif As DataTable = AdminData.GetData(ViewName, fCliente)
                    If dtRiesgoObrasCertif.Rows.Count > 0 Then
                        ClteInfo.RiesgoConsumido += dtRiesgoObrasCertif.Compute("SUM(ImpVencimientoA)", Nothing)  'dtRiesgoObrasCertif.Rows(0)("ImpVencimientoA")
                    End If

                    ViewName = "vCIClientesRiesgoObrasHitos"
                    Dim dtRiesgoObrasHitos As DataTable = AdminData.GetData(ViewName, fCliente)
                    If dtRiesgoObrasHitos.Rows.Count > 0 Then
                        ClteInfo.RiesgoConsumido += dtRiesgoObrasHitos.Compute("SUM(ImpVencimientoA)", Nothing)  'dtRiesgoObrasHitos.Rows(0)("ImpVencimientoA")
                    End If

                    ViewName = "vCIClientesRiesgoObrasMateriales"
                    Dim dtRiesgoObrasMateriales As DataTable = AdminData.GetData(ViewName, fCliente)
                    If dtRiesgoObrasMateriales.Rows.Count > 0 Then
                        ClteInfo.RiesgoConsumido += dtRiesgoObrasMateriales.Compute("SUM(ImpVencimientoA)", Nothing)  'dtRiesgoObrasMateriales.Rows(0)("ImpVencimientoA")
                    End If

                    ViewName = "vCIClientesRiesgoObrasMOD"
                    Dim dtRiesgoObrasMOD As DataTable = AdminData.GetData(ViewName, fCliente)
                    If dtRiesgoObrasMOD.Rows.Count > 0 Then
                        ClteInfo.RiesgoConsumido += dtRiesgoObrasMOD.Compute("SUM(ImpVencimientoA)", Nothing)  'dtRiesgoObrasMOD.Rows(0)("ImpVencimientoA")
                    End If

                    ViewName = "vCIClientesRiesgoObrasGasto"
                    Dim dtRiesgoObrasGastos As DataTable = AdminData.GetData(ViewName, fCliente)
                    If dtRiesgoObrasGastos.Rows.Count > 0 Then
                        ClteInfo.RiesgoConsumido += dtRiesgoObrasGastos.Compute("SUM(ImpVencimientoA)", Nothing)  'dtRiesgoObrasGastos.Rows(0)("ImpVencimientoA")
                    End If

                    ClteInfo.RiesgoConsumidoCalculado = True
                End If

                Dim datRiesgoProv As RiesgoProveedor = ProcessServer.ExecuteTask(Of String, RiesgoProveedor)(AddressOf GetRiesgoProveedorAsociado, Riesgo.IDProveedorAsociado, services)
                Riesgo.PdteFacturar = datRiesgoProv.PdteFacturar
                Riesgo.PagosNoPagados = datRiesgoProv.PagosNoPagados

                Riesgo.RiesgoConsumido = ClteInfo.RiesgoConsumido + data.ImporteASumar - Riesgo.PdteFacturar - Riesgo.PagosNoPagados

                Riesgo.SuperaRiesgo = (Riesgo.RiesgoConsumido >= Riesgo.RiesgoConcedido)
            End If
        End If


        Return Riesgo
    End Function

    <Serializable()> _
    Public Class RiesgoProveedor
        Public PdteFacturar As Double
        Public PagosNoPagados As Double
    End Class

    <Task()> Public Shared Function GetRiesgoProveedorAsociado(ByVal IDProveedorAsociado As String, ByVal services As ServiceProvider) As RiesgoProveedor
        Dim datRiesgo As New RiesgoProveedor
        If Length(IDProveedorAsociado) > 0 Then
            Dim fProveedor As New Filter
            fProveedor.Add(New StringFilterItem("IDProveedor", IDProveedorAsociado))
            Dim ViewName As String = "vFrmMntoRiesgoClienteProvFrasPdtes"
            Dim dtRiesgoFrasPdtesProv As DataTable = AdminData.GetData(ViewName, fProveedor)
            If dtRiesgoFrasPdtesProv.Rows.Count > 0 Then
                datRiesgo.PdteFacturar = dtRiesgoFrasPdtesProv.Compute("SUM(Importe)", Nothing)
            End If

            ViewName = "VFrmMntoRiesgoClienteProvPagos"
            Dim dtRiesgoPagosPdtesProv As DataTable = AdminData.GetData(ViewName, fProveedor)
            If dtRiesgoPagosPdtesProv.Rows.Count > 0 Then
                datRiesgo.PagosNoPagados = dtRiesgoPagosPdtesProv.Compute("SUM(Importe)", Nothing)
            End If
        End If
        Return datRiesgo

    End Function



    '<Task()> Public Shared Function ObtenerRiesgoCliente(ByVal data As DataRiesgoCliente, ByVal services As ServiceProvider) As RiesgoCliente
    '    Dim Riesgo As New RiesgoCliente
    '    Dim dtParam As DataTable = New Parametro().SelOnPrimaryKey("RIESGOCLTE")
    '    Dim blnComprobar As Boolean = True
    '    If Not IsNothing(dtParam) AndAlso dtParam.Rows.Count > 0 Then
    '        blnComprobar = (dtParam.Rows(0)("Valor") = True)
    '    End If
    '    If blnComprobar Then
    '        Dim dt As DataTable = AdminData.GetData("VFrmMntoRiesgoConsumidoClientePV", New StringFilterItem("IDCliente", FilterOperator.Equal, data.IDCliente & String.Empty))
    '        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
    '            Riesgo.IDCliente = data.IDCliente
    '            Riesgo.DescCliente = dt.Rows(0)("DescCliente")
    '            Riesgo.RiesgoConcedido = dt.Rows(0)("RiesgoConcedido")
    '            Riesgo.SuperaRiesgo = False
    '            If dt.Rows(0)("Riesgo") = True Then
    '                Riesgo.Riesgo = True
    '                Riesgo.RiesgoConsumido = dt.Rows(0)("RiesgoConsumido") + data.ImporteASumar
    '                Riesgo.SuperaRiesgo = (Riesgo.RiesgoConsumido >= Riesgo.RiesgoConcedido)
    '            End If

    '        End If
    '    End If
    '    Return Riesgo
    'End Function

    <Serializable()> _
    Public Class StCambioPermisos
        Public DtClientes As DataTable
        Public Permiso As enumPermisoFactElec

        Public Sub New()
        End Sub
        Public Sub New(ByVal DtClientes As DataTable, ByVal Permiso As enumPermisoFactElec)
            Me.DtClientes = DtClientes
            Me.Permiso = Permiso
        End Sub
    End Class

    <Task()> Public Shared Sub CambiarPermisos(ByVal data As StCambioPermisos, ByVal services As ServiceProvider)
        If Not data.DtClientes Is Nothing AndAlso data.DtClientes.Rows.Count > 0 Then
            For Each Dr As DataRow In data.DtClientes.Select
                Dr("FacturaElectronica") = data.Permiso
            Next
            BusinessHelper.UpdateTable(data.DtClientes)
        End If
    End Sub

#Region " Bloquear / Descbloquear Cliente"

    <Serializable()> _
    Public Class dataBloqueoCliente
        Public Bloqueado As Boolean
        Public IDOperarioBloqueo As String = String.Empty
        Public DescOperarioBloqueo As String = String.Empty
        Public IDCliente() As String

        Public Sub New(ByVal IDCliente() As String, ByVal Bloqueado As Boolean, ByVal IDOperarioBloqueo As String, ByVal DescOperarioBloqueo As String)
            Me.IDCliente = IDCliente
            Me.Bloqueado = Bloqueado
            Me.IDOperarioBloqueo = IDOperarioBloqueo
            Me.DescOperarioBloqueo = DescOperarioBloqueo
        End Sub

        Public Sub New(ByVal IDCliente() As String, ByVal Bloqueado As Boolean)
            Me.IDCliente = IDCliente
            Me.Bloqueado = Bloqueado
        End Sub
    End Class

    <Task()> Public Shared Sub BloqueoCliente(ByVal data As dataBloqueoCliente, ByVal services As ServiceProvider)
        If data.IDCliente.Length > 0 Then
            Dim dtClientes As DataTable = New Cliente().Filter(New InListFilterItem("IDCliente", data.IDCliente, FilterType.String))
            If Not dtClientes Is Nothing AndAlso dtClientes.Rows.Count > 0 Then
                For Each drCliente As DataRow In dtClientes.Rows
                    drCliente("Bloqueado") = data.Bloqueado
                    drCliente("IDOperarioBloqueo") = IIf(Len(data.IDOperarioBloqueo) > 0, data.IDOperarioBloqueo, DBNull.Value)
                    drCliente("DescOperarioBloqueo") = IIf(Len(data.DescOperarioBloqueo) > 0, data.DescOperarioBloqueo, DBNull.Value)

                    BusinessHelper.UpdateTable(dtClientes)
                Next
            End If
        End If
    End Sub

#End Region

    <Serializable()> _
    Public Class DataCopiaCliente
        Public IDClienteNew As String
        Public Errores As String

        Public Sub New()
        End Sub
        Public Sub New(ByVal IDClienteNew As String, ByVal Errores As String)
            Me.IDClienteNew = IDClienteNew
            Me.Errores = Errores
        End Sub
    End Class

    <Task()> Public Shared Function CopiarCliente(ByVal IDClienteOrigen As String, ByVal services As ServiceProvider) As DataCopiaCliente
        Dim StDataReturn As New DataCopiaCliente
        Try
            AdminData.BeginTx()
            'Cabecera Cliente
            Dim ClsClie As New Cliente
            Dim DtClieOrigen As DataTable = ClsClie.SelOnPrimaryKey(IDClienteOrigen)
            Dim DtClieDestino As DataTable = ClsClie.AddNew
            DtClieDestino.Rows.Add(DtClieOrigen.Rows(0).ItemArray)
            If Length(DtClieDestino.Rows(0)("IDContador")) = 0 Then
                Dim DataCont As Contador.DefaultCounter = ProcessServer.ExecuteTask(Of String, Contador.DefaultCounter)(AddressOf Contador.GetDefaultCounterValue, "Cliente", services)
                DtClieDestino.Rows(0)("IDContador") = DataCont.CounterID
            End If
            StDataReturn.IDClienteNew = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, DtClieDestino.Rows(0)("IDContador"), services)
            DtClieDestino.Rows(0)("IDCliente") = StDataReturn.IDClienteNew
            DtClieDestino.Rows(0)("IDGrupoCliente") = IDClienteOrigen

            'Cliente Banco
            Dim ClsClieBanco As New ClienteBanco
            Dim DtClieBancoOrigen As DataTable = ClsClieBanco.Filter(New FilterItem("IDCliente", FilterOperator.Equal, IDClienteOrigen))
            Dim DtClieBancoDestino As DataTable = ClsClieBanco.AddNew
            For Each Dr As DataRow In DtClieBancoOrigen.Select
                DtClieBancoDestino.Rows.Add(Dr.ItemArray)
                DtClieBancoDestino.Rows(DtClieBancoDestino.Rows.Count - 1)("IDCliente") = StDataReturn.IDClienteNew
                DtClieBancoDestino.Rows(DtClieBancoDestino.Rows.Count - 1)("IDClienteBanco") = AdminData.GetAutoNumeric
            Next

            'Cliente Articulo
            Dim ClsClieArt As New ArticuloCliente
            Dim DtClieArtOrigen As DataTable = ClsClieArt.Filter(New FilterItem("IDCliente", FilterOperator.Equal, IDClienteOrigen))
            Dim DtClieArtDestino As DataTable = ClsClieArt.AddNew
            For Each Dr As DataRow In DtClieArtOrigen.Select
                DtClieArtDestino.Rows.Add(Dr.ItemArray)
                DtClieArtDestino.Rows(DtClieArtDestino.Rows.Count - 1)("IDCliente") = StDataReturn.IDClienteNew
            Next

            'Cliente Articulo Linea
            Dim ClsClieArtLinea As New ArticuloClienteLinea
            Dim DtClieArtLinOrigen As DataTable = ClsClieArtLinea.Filter(New FilterItem("IDCliente", FilterOperator.Equal, IDClienteOrigen))
            Dim DtClieArtLinDestino As DataTable = ClsClieArtLinea.AddNew
            For Each Dr As DataRow In DtClieArtLinOrigen.Select
                DtClieArtLinDestino.Rows.Add(Dr.ItemArray)
                DtClieArtLinDestino.Rows(DtClieArtLinDestino.Rows.Count - 1)("IDCliente") = StDataReturn.IDClienteNew
            Next

            'Cliente Direccion
            Dim ClsClieDirec As New ClienteDireccion
            Dim DtClieDirecOrigen As DataTable = ClsClieDirec.Filter(New FilterItem("IDCliente", FilterOperator.Equal, IDClienteOrigen))
            Dim DtClieDirecDestino As DataTable = ClsClieDirec.AddNew
            For Each Dr As DataRow In DtClieDirecOrigen.Select
                DtClieDirecDestino.Rows.Add(Dr.ItemArray)
                DtClieDirecDestino.Rows(DtClieDirecDestino.Rows.Count - 1)("IDCliente") = StDataReturn.IDClienteNew
                DtClieDirecDestino.Rows(DtClieDirecDestino.Rows.Count - 1)("IDDireccion") = AdminData.GetAutoNumeric
            Next

            'Cliente Representante
            Dim ClsClieRep As New ClienteRepresentante
            Dim DtClieRepOrigen As DataTable = ClsClieRep.Filter(New FilterItem("IDCliente", FilterOperator.Equal, IDClienteOrigen))
            Dim DtClieRepDestino As DataTable = ClsClieRep.AddNew
            For Each Dr As DataRow In DtClieRepOrigen.Select
                DtClieRepDestino.Rows.Add(Dr.ItemArray)
                DtClieRepDestino.Rows(DtClieRepDestino.Rows.Count - 1)("IDCliente") = StDataReturn.IDClienteNew
                DtClieRepDestino.Rows(DtClieRepDestino.Rows.Count - 1)("IDClienteRepresentante") = AdminData.GetAutoNumeric
            Next

            'Cliente Tarifa
            Dim ClsClieTar As New ClienteTarifa
            Dim DtClieTarOrigen As DataTable = ClsClieTar.Filter(New FilterItem("IDCliente", FilterOperator.Equal, IDClienteOrigen))
            Dim DtClieTarDestino As DataTable = ClsClieTar.AddNew
            For Each Dr As DataRow In DtClieTarOrigen.Select
                DtClieTarDestino.Rows.Add(Dr.ItemArray)
                DtClieTarDestino.Rows(DtClieTarDestino.Rows.Count - 1)("IDCliente") = StDataReturn.IDClienteNew
            Next

            'Cliente Descuento Familia
            Dim ClsClieDtoFam As New ClienteDescuentoFamilia
            Dim DtClieDtoFamOrigen As DataTable = ClsClieDtoFam.Filter(New FilterItem("IDCliente", FilterOperator.Equal, IDClienteOrigen))
            Dim DtClieDtoFamDestino As DataTable = ClsClieDtoFam.AddNew
            For Each Dr As DataRow In DtClieDtoFamOrigen.Select
                DtClieDtoFamDestino.Rows.Add(Dr.ItemArray)
                DtClieDtoFamDestino.Rows(DtClieDtoFamDestino.Rows.Count - 1)("IDCliente") = StDataReturn.IDClienteNew
                DtClieDtoFamDestino.Rows(DtClieDtoFamDestino.Rows.Count - 1)("IDClienteFamilia") = AdminData.GetAutoNumeric
            Next

            'Cliente Promoción
            Dim ClsCliePromo As New ClientePromocion
            Dim DtCliePromoOrigen As DataTable = ClsCliePromo.Filter(New FilterItem("IDCliente", FilterOperator.Equal, IDClienteOrigen))
            Dim DtCliePromoDestino As DataTable = ClsCliePromo.AddNew
            For Each Dr As DataRow In DtCliePromoOrigen.Select
                DtCliePromoDestino.Rows.Add(Dr.ItemArray)
                DtCliePromoDestino.Rows(DtCliePromoDestino.Rows.Count - 1)("IDCliente") = StDataReturn.IDClienteNew
            Next

            'Cliente Vacación
            Dim ClsClieVaca As New ClienteVacaciones
            Dim DtClieVacaOrigen As DataTable = ClsClieVaca.Filter(New FilterItem("IDCliente", FilterOperator.Equal, IDClienteOrigen))
            Dim DtClieVacaDestino As DataTable = ClsClieVaca.AddNew
            For Each Dr As DataRow In DtClieVacaOrigen.Select
                DtClieVacaDestino.Rows.Add(Dr.ItemArray)
                DtClieVacaDestino.Rows(DtClieVacaDestino.Rows.Count - 1)("IDCliente") = StDataReturn.IDClienteNew
                DtClieVacaDestino.Rows(DtClieVacaDestino.Rows.Count - 1)("IDVacacion") = AdminData.GetAutoNumeric
            Next

            'Cliente Observación
            Dim ClsClieObs As New ClienteObservacion
            Dim DtClieObsOrigen As DataTable = ClsClieObs.Filter(New FilterItem("IDCliente", FilterOperator.Equal, IDClienteOrigen))
            Dim DtClieObsDestino As DataTable = ClsClieObs.AddNew
            For Each Dr As DataRow In DtClieObsOrigen.Select
                DtClieObsDestino.Rows.Add(Dr.ItemArray)
                DtClieObsDestino.Rows(DtClieObsDestino.Rows.Count - 1)("IDCliente") = StDataReturn.IDClienteNew
                DtClieObsDestino.Rows(DtClieObsDestino.Rows.Count - 1)("IDClienteObservacion") = AdminData.GetAutoNumeric
            Next

            BusinessHelper.UpdateTable(DtClieDestino) : BusinessHelper.UpdateTable(DtClieBancoDestino)
            BusinessHelper.UpdateTable(DtClieArtDestino) : BusinessHelper.UpdateTable(DtClieArtLinDestino)
            BusinessHelper.UpdateTable(DtClieDirecDestino) : BusinessHelper.UpdateTable(DtClieRepDestino)
            BusinessHelper.UpdateTable(DtClieTarDestino) : BusinessHelper.UpdateTable(DtClieDtoFamDestino)
            BusinessHelper.UpdateTable(DtCliePromoDestino) : BusinessHelper.UpdateTable(DtClieVacaDestino)
            BusinessHelper.UpdateTable(DtClieObsDestino)

            AdminData.CommitTx(True)
        Catch ex As Exception
            AdminData.RollBackTx()
            StDataReturn.Errores = ex.Message
        End Try

        Return StDataReturn
    End Function

#End Region

End Class


Public Class ClienteInfo
    Inherits ClassEntityInfo

    Public IDCliente As String
    Public DescCliente As String
    Public CifCliente As String
    Public RazonSocial As String
    Public Direccion As String
    Public CodPostal As String
    Public Poblacion As String
    Public Provincia As String
    Public Telefono As String
    Public Fax As String
    Public Pais As String
    Public TipoIVA As String
    Public EmpresaGrupo As Boolean
    Public Idioma As String
    Public Extranjero As Boolean
    Public CanariasCeutaMelilla As Boolean
    Public CEE As Boolean
    Public Moneda As String
    Public CentroGestion As String
    Public AgrupacionPedido As enummcAgrupPedido
    Public Bloqueado As Boolean
    Public FormaPago As String
    Public CondicionPago As String
    Public DiaPago As String
    Public FormaEnvio As String
    Public CondicionEnvio As String
    Public DtoComercial As Double
    Public Prioridad As Integer
    Public GrupoCliente As String
    Public GrupoDireccion As Boolean
    Public GrupoFactura As Boolean
    Public GrupoTarifa As Boolean
    Public GrupoRiesgo As Boolean
    Public GrupoArticulo As Boolean
    Public Zona As String
    Public IDTipoAsiento As enumTipoAsiento
    Public RetencionIRPF As Double
    Public IDBancoPropio As String
    Public CCCliente As String
    Public CCRetencion As String
    Public CCEfectosCliente As String
    Public CCEfectosGestionCobros As String
    Public CCAnticipo As String
    Public CCFianza As String
    Public IDContadorCargo As String
    Public IDModoTransporte As String
    Public Riesgo As Boolean
    Public IDAlmacenContenedor As String
    Public FacturarTasaResiduos As Boolean
    Public AlbaranValorado As Boolean
    Public TipoGeneracionSeguros As Integer
    Public TieneRE As Boolean
    Public IDEmpresa As String
    Public PortesEspSalidas As Boolean
    Public PortesEspRetornos As Boolean
    Public CondicionesEspPortes As String
    Public DiaFacturacion As Integer
    Public FianzaObligatoria As Boolean
    Public IDClasificacionObra As String
    Public Email As String
    Public IDConsignatario As String
    Public IDEDIFormato As String
    Public IVAReducido As Boolean
    Public DtoComercialLinea As Double
    Public RiesgoConcedido As Double
    Public RiesgoConsumido As Double
    Public RiesgoConsumidoCalculado As Boolean
    Public RiesgoInterno As Double
    Public IDProveedor As String
    Public IDOperario As String
    Public IDCCEfectosCartera As String
    Public IDCNAE As String
    Public IDProveedorAsociado As String
    Public LimiteCapitalAsegurado As Double
    Public IDMotivoNoAsegurado As String
    Public IDTarifaAbono As String
    'Public TipoDocIdentidad As enumTipoDocIdent


    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal data As DataRow)
        MyBase.New(data)
    End Sub

    Public Sub New(ByVal IDCliente As String)
        MyBase.New()
        Me.Fill(IDCliente)
    End Sub

    Public Overloads Overrides Sub Fill(ByVal ParamArray PrimaryKey() As Object)
        Dim dtClteInfo As DataTable = New BE.DataEngine().Filter("vNegClienteInfo", New StringFilterItem("IDCliente", PrimaryKey(0)))
        If dtClteInfo.Rows.Count > 0 Then
            Me.Fill(dtClteInfo.Rows(0))
        Else
            ApplicationService.GenerateError("El Cliente | no existe.", Quoted(PrimaryKey(0)))
        End If
    End Sub

End Class

<Serializable()> _
Public Class RiesgoCliente

    Public IDCliente As String
    Public DescCliente As String
    Public Riesgo As Boolean
    Public RiesgoGrupo As Boolean
    Public RiesgoConcedido As Double
    Public RiesgoConsumido As Double
    Public RiesgoInterno As Double
    Public SuperaRiesgo As Boolean


    Public IDProveedorAsociado As String
    Public DescProveedorAsociado As String
    Public PdteFacturar As Double
    Public PagosNoPagados As Double

    Public Sub New()
        IDCliente = String.Empty
        Riesgo = False
        RiesgoGrupo = False
        RiesgoConcedido = 0
        RiesgoConsumido = 0
        SuperaRiesgo = False
    End Sub

End Class