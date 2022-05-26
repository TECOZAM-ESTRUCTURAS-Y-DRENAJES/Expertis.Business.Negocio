<Serializable()> _
Public Class CopiarComponentesInfo
    Public IDArticuloOrigen, IDArticuloDestino As String
    Public IDEstructuraOrigen, IDEstructuraDestino As String
    Public IDComponente As String
    Public IDArticuloPadreComponente, IDEstructuraPadreComponente As String
    Public Cantidad, Merma, Factor, CantidadProduccion As Double
    Public IDUdMedidaProduccion As String
    Public EsComponente As Boolean
End Class

Public Class Estructura

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbEstructura"

#End Region

#Region "Eventos GetBusinessRules"

    <Serializable()> _
    Public Class DatosCalcFactor
        Public IDComponente As String
        Public IDUDInterna As String
        Public IDUDMedidaProduccion As String
    End Class

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBrl As New BusinessRules
        oBrl.Add("Cantidad", AddressOf CambioCantidad)
        oBrl.Add("CantidadProduccion", AddressOf CambioCantidadProd)
        oBrl.Add("Factor", AddressOf CambioFactor)
        oBrl.Add("IDComponente", AddressOf CambioComponente)
        oBrl.Add("IDUDInterna", AddressOf CambioUnidad)
        oBrl.Add("IDUDMedidaProduccion", AddressOf CambioUnidad)
        Return oBrl
    End Function

    <Task()> Public Shared Sub CambioCantidad(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            data.Current("CantidadProduccion") = data.Value * Nz(data.Current("Factor"), 1)
        Else
            data.Current("CantidadProduccion") = data.Value
            data.Current("Factor") = 1
        End If
    End Sub

    <Task()> Public Shared Sub CambioCantidadProd(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            If Nz(data.Current("Factor")) > 0 Then
                data.Current("Cantidad") = data.Value / data.Current("Factor")
            Else
                data.Current("Cantidad") = data.Value
                data.Current("Factor") = 1
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioFactor(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current("Factor") = Nz(data.Value, 1)
        If Length(data.Current("CantidadProduccion")) > 0 Then
            data.Current("Cantidad") = data.Current("CantidadProduccion") / data.Current("Factor")
        ElseIf Length(data.Current("Cantidad")) > 0 Then
            data.Current("CantidadProduccion") = data.Current("Cantidad") * data.Current("Factor")
        End If
    End Sub

    <Task()> Public Shared Sub CambioComponente(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim dt As DataTable = New Articulo().SelOnPrimaryKey(data.Value)
            If IsNothing(dt) OrElse dt.Rows.Count = 0 Then
                ApplicationService.GenerateError("El Componente | no existe.", Quoted(data.Value))
            End If
            If Length(data.Current("IDArticulo")) > 0 And Length(data.Current("IDComponente")) > 0 Then
                Dim datCambioComponente As New DataCambioComponente(data.Current("IDArticulo"), data.Current("IDComponente"), data.Value)
                ProcessServer.ExecuteTask(Of DataCambioComponente)(AddressOf ValidarItemsEstructuraConfigurable, datCambioComponente, services)
            End If
            If Length(data.Current("IDUDInterna")) > 0 And Length(data.Current("IDUDMedidaProduccion")) > 0 Then
                Dim StDatos As New DatosCalcFactor
                StDatos.IDComponente = data.Value
                StDatos.IDUDInterna = data.Current("IDUDInterna")
                StDatos.IDUDMedidaProduccion = data.Current("IDUDMedidaProduccion")
                data.Current("Factor") = ProcessServer.ExecuteTask(Of DatosCalcFactor, Double)(AddressOf CalcularFactor, StDatos, services)
            ElseIf Length(data.Value) > 0 And Length(data.Current("IDUDInterna")) > 0 And Length(data.Current("IDUDMedidaProduccion")) = 0 Then
                data.Current("IDUDMedidaProduccion") = data.Current("IDUdInterna")
                data.Current("Factor") = 1
            Else : data.Current("Factor") = 1
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarDelItemsEstructuraConfigurable(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDArticulo")) > 0 And Length(data("IDComponente")) > 0 Then
            Dim datCambioComponente As New DataCambioComponente(data("IDArticulo"), data("IDComponente"))
            ProcessServer.ExecuteTask(Of DataCambioComponente)(AddressOf ValidarItemsEstructuraConfigurable, datCambioComponente, services)
        End If
    End Sub

    <Serializable()> _
    Public Class DataCambioComponente
        Public IDArticulo As String
        Public IDComponenteOld As String
        Public IDComponenteNew As String

        Public Sub New(ByVal IDArticulo As String, ByVal IDComponenteOld As String, Optional ByVal IDComponenteNew As String = Nothing)
            Me.IDArticulo = IDArticulo
            Me.IDComponenteOld = IDComponenteOld
            Me.IDComponenteNew = IDComponenteNew
        End Sub
    End Class
    <Task()> Public Shared Sub ValidarItemsEstructuraConfigurable(ByVal data As DataCambioComponente, ByVal services As ServiceProvider)
        If Length(data.IDArticulo) > 0 AndAlso Length(data.IDComponenteOld) > 0 Then
            '//Si cambiamos el componente o lo quitamos, validamos si tiene asociado algo del configurador.
            If data.IDComponenteOld <> data.IDComponenteNew & String.Empty Then
                Dim f As New Filter
                f.Add(New StringFilterItem("IDPadre", data.IDArticulo))
                f.Add(New StringFilterItem("IDComponente", data.IDComponenteOld))

                Dim Caract As BusinessHelper = BusinessHelper.CreateBusinessObject("CfgEstructuraCaracteristica")
                Dim dtCaracteristicas As DataTable = Caract.Filter(f)
                If dtCaracteristicas.Rows.Count > 0 Then
                    ApplicationService.GenerateError("Existen Características asociadas al componente {0} en la estructura del artículo {1}.", Quoted(data.IDComponenteOld), Quoted(data.IDArticulo))
                Else
                    Dim Condiciones As BusinessHelper = BusinessHelper.CreateBusinessObject("CfgEstructuraCondicion")
                    Dim dtCondiciones As DataTable = Condiciones.Filter(f)
                    If dtCondiciones.Rows.Count > 0 Then
                        ApplicationService.GenerateError("Existen Condiciones asociadas al componente {0} en la estructura del artículo {1}.", Quoted(data.IDComponenteOld), Quoted(data.IDArticulo))
                    Else
                        Dim Acciones As BusinessHelper = BusinessHelper.CreateBusinessObject("CfgEstructuraAccion")
                        Dim dtAcciones As DataTable = Acciones.Filter(f)
                        If dtAcciones.Rows.Count > 0 Then
                            ApplicationService.GenerateError("Existen Acciones asociadas al componente {0} en la estructura del artículo {1}.", Quoted(data.IDComponenteOld), Quoted(data.IDArticulo))
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioUnidad(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 And data.ColumnName = "IDUDMedidaProduccion" Then
            Dim u As New UdMedida
            u.GetItemRow(data.Value)
        End If
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDUDInterna")) > 0 AndAlso Length(data.Current("IDUDMedidaProduccion")) > 0 AndAlso Length(data.Current("IDComponente")) > 0 Then
            Dim StDatos As New DatosCalcFactor
            StDatos.IDComponente = data.Current("IDComponente")
            StDatos.IDUDInterna = data.Current("IDUDInterna")
            StDatos.IDUDMedidaProduccion = data.Current("IDUDMedidaProduccion")
            data.Current("Factor") = ProcessServer.ExecuteTask(Of DatosCalcFactor, Double)(AddressOf CalcularFactor, StDatos, services)
        Else : data.Current("Factor") = 1
        End If
        data.Current = New Estructura().ApplyBusinessRule("Factor", data.Current("Factor"), data.Current, data.Context)
    End Sub

    <Task()> Public Shared Function CalcularFactor(ByVal data As DatosCalcFactor, ByVal services As ServiceProvider) As Double
        Dim StDatos As New ArticuloUnidadAB.DatosFactorConversion
        StDatos.IDArticulo = data.IDComponente
        StDatos.IDUdMedidaA = data.IDUDInterna
        StDatos.IDUdMedidaB = data.IDUDMedidaProduccion
        StDatos.UnoSiNoExiste = True
        Return ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, StDatos, services)
    End Function

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarComponente)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosEstructura)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDComponente")) = 0 Then ApplicationService.GenerateError("El Componente es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarComponente(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Modified Then
            If data("IDComponente") <> data("IDComponente", DataRowVersion.Original) & String.Empty Or Nz(data("Secuencia")) <> Nz(data("Secuencia", DataRowVersion.Original)) Then
                Dim f As New Filter
                f.Add(New StringFilterItem("IDArticulo", data("IDArticulo")))
                f.Add(New StringFilterItem("IDComponente", data("IDComponente")))
                f.Add(New StringFilterItem("IDEstructura", data("IDEstructura")))
                If Length(data("Secuencia")) > 0 Then
                    f.Add(New NumberFilterItem("Secuencia", data("Secuencia")))
                Else : f.Add(New IsNullFilterItem("Secuencia", True))
                End If
                Dim dt As DataTable = New Estructura().Filter(f)
                If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                    ApplicationService.GenerateError("El Componente '|' no puede tener Secuencias repetidas.", data("IDComponente"))
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarDatosEstructura(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDArticulo")) > 0 Then
            data("IDTipoEstructura") = System.DBNull.Value
            Dim StDatos As New DatosExistPadre
            StDatos.ArticuloPadre = data("IDArticulo")
            StDatos.ArticuloHijo = data("IDComponente")
            If ProcessServer.ExecuteTask(Of DatosExistPadre, Boolean)(AddressOf ExisteComoPadreFn, StDatos, services) Then
                ApplicationService.GenerateError("Este componente no puede formar parte de su propia estructura.")
            End If
            If Length(data("Secuencia")) > 0 Then
                Dim f As New Filter
                f.Clear()
                f.Add(New StringFilterItem("IDArticulo", data("IDArticulo")))
                f.Add(New StringFilterItem("IDEstructura", data("IDEstructura")))

                Dim dtEstructura As DataTable = New ArticuloEstructura().Filter(f)
                If Not dtEstructura Is Nothing AndAlso dtEstructura.Rows.Count > 0 Then
                    If Length(dtEstructura.Rows(0)("IDRuta")) > 0 Then
                        f.Clear()
                        f.Add(New StringFilterItem("IDArticulo", data("IDArticulo")))
                        f.Add(New NumberFilterItem("Secuencia", data("Secuencia")))
                        f.Add(New StringFilterItem("IDRuta", dtEstructura.Rows(0)("IDRuta")))
                        Dim dtRuta As DataTable = New Ruta().Filter(f)
                        If dtRuta Is Nothing OrElse dtRuta.Rows.Count = 0 Then
                            ApplicationService.GenerateError("La Secuencia no existe en la Ruta asociada a la Estructura activa.")
                        End If
                    End If
                End If
            End If

            If data.RowState = DataRowState.Modified Then
                Dim datCambioComponente As New DataCambioComponente(data("IDArticulo"), data("IDComponente", DataRowVersion.Original), data("IDComponente"))
                ProcessServer.ExecuteTask(Of DataCambioComponente)(AddressOf ValidarItemsEstructuraConfigurable, datCambioComponente, services)
            End If
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
            If Length(data("IDEstrComp")) = 0 Then data("IDEstrComp") = AdminData.GetAutoNumeric
        End If
    End Sub

#End Region


#Region "Eventos RegisterDeleteTasks"

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf ValidarDelItemsEstructuraConfigurable)
    End Sub

#End Region

#Region "Funciones Públicas"

    <Serializable()> _
    Public Class DatosEstrucPrincipal
        Public ArticuloPadre As String
        Public Estructura As String
        Public TipoEstructura As String
    End Class

    <Serializable()> _
    Public Class DatosExistPadre
        Public ArticuloPadre As String
        Public ArticuloHijo As String
    End Class

    <Serializable()> _
    Public Class DatosCopiaComp
        Public Dt As DataTable
        Public IDArticuloDestino As String
        Public IDEstructuraDestino As String
    End Class

    <Serializable()> _
    Public Class DatosElimComp
        Public IDArticulo As String
        Public IDEstructura As String
        Public IDComponente As String
    End Class

    <Serializable()> _
    Public Class DatosObtEstConfig
        Public IDArticuloPadre As String
        Public IDEstructura As String
    End Class

    <Serializable()> _
   Public Class DatosSustComp
        Public IDComponenteAnterior As String
        Public IDComponenteNuevo As String
        Public DtMarcados As DataTable
    End Class

    <Task()> Public Shared Function ObtenerEstructuraPrincipal(ByVal data As DatosEstrucPrincipal, ByVal services As ServiceProvider) As DataTable
        'Function a la cual le pasamos una cadena formada por la llamada a un procedimiento almacenado
        'y nos devuelve el datatable resultante de hacer la select en la tabla "tbEstructura" con todos los componentes de clave principal
        If Length(data.ArticuloPadre) > 0 Then
            If Length(data.Estructura) = 0 Then
                Dim ae As New ArticuloEstructura
                data.Estructura = ProcessServer.ExecuteTask(Of String, String)(AddressOf ae.EstructuraPpal, data.ArticuloPadre, services)
            End If
            Dim strSQL As String
            If Len(data.Estructura) Then
                Return AdminData.Execute("sp_EstructuraPrincipalMultinivel", False, data.ArticuloPadre, data.Estructura, 0)
            Else : Return AdminData.Execute("sp_EstructuraPrincipalMultinivel", False, data.ArticuloPadre, Nothing, 0)
            End If
        End If
    End Function

    <Task()> Public Shared Function ObtenerImplosion(ByVal data As String, ByVal services As ServiceProvider) As DataTable
        If Length(data) > 0 Then
            Return AdminData.Execute("sp_EstructuraPrincipalMultinivel", False, data, Nothing, 1)
        End If
    End Function

    <Task()> Public Shared Sub ExisteComoPadre(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim dt As DataTable = data.Table
        If dt.Columns.Contains("IDTipoEstructura") AndAlso dt.Columns.Contains("IDArticulo") Then
            Dim dtEstructura As DataTable = New Estructura().Filter("IDComponente", "IDTipoEstructura= '" & data("IDTipoEstructura") & "'")
            If Not IsNothing(dtEstructura) AndAlso dtEstructura.Rows.Count > 0 Then
                For Each drEstructura As DataRow In dtEstructura.Rows
                    Dim StDatos As New DatosExistPadre
                    StDatos.ArticuloPadre = data("IDArticulo")
                    StDatos.ArticuloHijo = data("IDComponente")
                    If ProcessServer.ExecuteTask(Of DatosExistPadre, Boolean)(AddressOf ExisteComoPadreFn, StDatos, services) Then
                        ApplicationService.GenerateError("El artículo no puede formar parte de la propia estructura.")
                    End If
                Next
            End If
        End If
    End Sub

    <Task()> Public Shared Function ExisteComoPadreFn(ByVal data As DatosExistPadre, ByVal services As ServiceProvider) As Boolean
        Dim cmm As Common.DbCommand = AdminData.GetCommand
        cmm.CommandType = CommandType.StoredProcedure
        cmm.CommandText = "sp_ExisteComoPadre"

        Dim parameter1 As Common.DbParameter = cmm.CreateParameter
        cmm.Parameters.Add(parameter1)
        parameter1.ParameterName = "@pArticuloPadre"
        parameter1.Value = data.ArticuloPadre
        Dim Parameter2 As Common.DbParameter = cmm.CreateParameter
        cmm.Parameters.Add(Parameter2)
        Parameter2.ParameterName = "@pArticuloHijo"
        Parameter2.Value = data.ArticuloHijo
        Dim Parameter3 As Common.DbParameter = cmm.CreateParameter
        cmm.Parameters.Add(Parameter3)
        Parameter3.DbType = DbType.Int32
        Parameter3.ParameterName = "@pEncontrado"
        Parameter3.Direction = ParameterDirection.Output

        AdminData.Execute(cmm)

        Return Parameter3.Value
    End Function

    <Task()> Public Shared Sub CopiarComponente(ByVal data As DatosCopiaComp, ByVal services As ServiceProvider)
        If Not data.Dt Is Nothing AndAlso data.Dt.Rows.Count > 0 Then
            Dim info As New CopiarComponentesInfo
            info.IDArticuloOrigen = data.Dt.Rows(0)("IDArticulo")
            info.IDEstructuraOrigen = data.Dt.Rows(0)("IDEstructura")
            info.IDArticuloDestino = data.IDArticuloDestino
            info.IDEstructuraDestino = data.IDEstructuraDestino

            For Each dr As DataRow In data.Dt.Rows
                info.IDComponente = dr("IDComponente") & String.Empty
                info.Cantidad = dr("Cantidad")
                info.Merma = dr("Merma")
                info.IDUdMedidaProduccion = dr("IDUdMedidaProduccion") & String.Empty
                info.Factor = dr("Factor")
                info.CantidadProduccion = dr("CantidadProduccion")
                ProcessServer.ExecuteTask(Of CopiarComponentesInfo)(AddressOf CopiarComponenteInfo, info, services)
            Next
        End If
    End Sub

    <Task()> Public Shared Sub CopiarComponenteInfo(ByVal data As CopiarComponentesInfo, ByVal services As ServiceProvider)
        Dim strIDArticuloOrigen, strIDEstructuraOrigen, strIDArticuloDestino, strIDEstructuraDestino, strIDComponenteOrigen As String
        If Not IsNothing(data) Then
            strIDArticuloOrigen = data.IDArticuloOrigen
            strIDEstructuraOrigen = data.IDEstructuraOrigen
            strIDArticuloDestino = data.IDArticuloDestino
            strIDEstructuraDestino = data.IDEstructuraDestino
            strIDComponenteOrigen = data.IDComponente
        End If
        If Len(strIDArticuloOrigen) > 0 And Len(strIDEstructuraOrigen) > 0 And Len(strIDArticuloDestino) > 0 And Len(strIDEstructuraDestino) > 0 Then
            Dim dtNew As DataTable = New Estructura().AddNewForm
            dtNew.Rows(0)("IDEstrComp") = AdminData.GetAutoNumeric
            dtNew.Rows(0)("IDArticulo") = strIDArticuloDestino
            dtNew.Rows(0)("IDEstructura") = strIDEstructuraDestino
            If Length(strIDComponenteOrigen) > 0 Then
                dtNew.Rows(0)("IDComponente") = strIDComponenteOrigen
                dtNew.Rows(0)("Cantidad") = data.Cantidad
                dtNew.Rows(0)("Merma") = data.Merma
                If Len(data.IDUdMedidaProduccion) > 0 Then
                    dtNew.Rows(0)("IDUdMedidaProduccion") = data.IDUdMedidaProduccion
                End If
                dtNew.Rows(0)("Factor") = data.Factor
                dtNew.Rows(0)("CantidadProduccion") = data.CantidadProduccion
            End If
            Dim ClsEst As New Estructura
            ClsEst.Update(dtNew)
        End If
    End Sub

    <Task()> Public Shared Sub EliminarComponente(ByVal data As DatosElimComp, ByVal services As ServiceProvider)
        If Length(data.IDArticulo) > 0 And Length(data.IDEstructura) > 0 Then
            Dim f As New Filter
            f.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
            f.Add(New StringFilterItem("IDEstructura", data.IDEstructura))
            If Length(data.IDComponente) > 0 Then f.Add(New StringFilterItem("IDComponente", data.IDComponente))
            Dim ClsEst As New Estructura
            Dim dt As DataTable = ClsEst.Filter(f)
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                ClsEst.Delete(dt)
            End If
        End If
    End Sub

    <Task()> Public Shared Function ObtenerEstructuraConfigurable(ByVal data As DatosObtEstConfig, ByVal service As ServiceProvider) As DataTable
        Return AdminData.Execute("sp_EstructuraConfigurable", False, data.IDArticuloPadre, data.IDEstructura)
    End Function

    <Task()> Public Shared Function SustituirComponente(ByVal data As DatosSustComp, ByVal services As ServiceProvider) As DataTable
        '//El nuevo componente no puede estar por arriba, ni por abajo en las estructuras en las que se encuentre el Componente antiguo
        data.DtMarcados.Columns("MensajeError").ReadOnly = False
        data.DtMarcados.Columns("Sustituido").ReadOnly = False

        Dim objFilterExisteComp As New Filter
        Dim dt As DataTable
        Dim objNegArticulo As New Articulo
        Dim dtArticulo As DataTable

        Dim objFilter As New Filter
        objFilter.Add(New BooleanFilterItem("Expertis.CheckValue", False))
        Dim WhereNoMarcados As String = objFilter.Compose(New AdoFilterComposer)
        For Each drNoMarcado As DataRow In data.DtMarcados.Select(WhereNoMarcados)
            drNoMarcado("Sustituido") = System.DBNull.Value
            drNoMarcado("MensajeError") = String.Empty
        Next

        objFilter.Clear()
        objFilter.Add(New BooleanFilterItem("Expertis.CheckValue", True))
        Dim WhereMarcados As String = objFilter.Compose(New AdoFilterComposer)
        For Each drMarcado As DataRow In data.DtMarcados.Select(WhereMarcados)
            Dim StDatos As New DatosExistPadre
            StDatos.ArticuloPadre = drMarcado("IDArticulo")
            StDatos.ArticuloHijo = data.IDComponenteNuevo
            If data.IDComponenteAnterior = drMarcado("IDComponente") AndAlso Not ProcessServer.ExecuteTask(Of DatosExistPadre, Boolean)(AddressOf ExisteComoPadreFn, StDatos, services) Then
                objFilterExisteComp.Clear()
                objFilterExisteComp.Add(New StringFilterItem("IDArticulo", drMarcado("IDArticulo")))
                objFilterExisteComp.Add(New StringFilterItem("IDComponente", data.IDComponenteNuevo))
                objFilterExisteComp.Add(New StringFilterItem("IDEstructura", drMarcado("IDEstructura")))
                If Length(drMarcado("IDRuta")) > 0 Then objFilterExisteComp.Add(New StringFilterItem("IDRuta", drMarcado("IDRuta")))
                If Length(drMarcado("Secuencia")) > 0 Then objFilterExisteComp.Add(New NumberFilterItem("Secuencia", drMarcado("Secuencia")))
                dt = New BE.DataEngine().Filter("vFrmMntoSustituirComponente", objFilterExisteComp)
                If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
                    drMarcado("Sustituido") = False
                    drMarcado("MensajeError") = "El componente ya existe para la estructura indicada."
                Else
                    drMarcado("IDComponente") = data.IDComponenteNuevo
                    dtArticulo = objNegArticulo.SelOnPrimaryKey(data.IDComponenteNuevo)
                    If Not IsNothing(dtArticulo) AndAlso dtArticulo.Rows.Count > 0 Then
                        drMarcado("DescComponente") = dtArticulo.Rows(0)("DescArticulo")
                    End If
                    drMarcado("Sustituido") = True
                    drMarcado("MensajeError") = "Sustitución realizada con éxito."
                End If
            Else
                drMarcado("Sustituido") = False
                If data.IDComponenteAnterior <> drMarcado("IDComponente") Then
                    drMarcado("MensajeError") = "El componente había sido sustituido anteriormente. Asegúrese de filtrar correctamente antes del proceso de Sustitución."
                Else
                    drMarcado("MensajeError") = "El nuevo componente no puede formar parte de su propia estructura."
                End If
            End If
        Next

        data.DtMarcados.TableName = "Estructura"
        BusinessHelper.UpdateTable(data.DtMarcados)

        Return data.DtMarcados
    End Function

#End Region

End Class