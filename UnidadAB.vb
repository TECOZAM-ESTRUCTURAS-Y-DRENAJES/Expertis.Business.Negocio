Public Class UnidadAB

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Private Const cnEntidad As String = "tbUnidadAB"

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

#End Region

#Region "Eventos RegisterValidateTasks"

    ''' <summary>
    ''' Relación de tareas asociadas a la validación 
    ''' </summary>
    ''' <param name="validateProcess">Proceso en el que se registran las tareas de edición</param>
    ''' <remarks></remarks>
    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClavePrimaria)
    End Sub

    ''' <summary>
    ''' Comprobar que el código y la descripción no sean nulos
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDUdMedidaA")) = 0 Then ApplicationService.GenerateError("El campo Unidad A es un dato obligatorio.")
        If Length(data("IDUdMedidaB")) = 0 Then ApplicationService.GenerateError("El campo Unidad B es un dato obligatorio.")
        If Nz(data("Factor"), 0) <= 0 Then
            ApplicationService.GenerateError("El factor de conversión debe ser mayor de 0.")
        End If
    End Sub

    ''' <summary>
    ''' Comprobar que no exista la clave
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = New UnidadAB().SelOnPrimaryKey(data("IDUdMedidaA"), data("IDUdMedidaB"))
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("Ya está definida una conversión entre las unidades indicadas. -")
            End If
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Serializable()> _
      Public Class UnidadMedidaInfo
        Public IDUdMedidaA As String
        Public IDUdMedidaB As String
        Public Cantidad As Double
        Public UnoSiNoExiste As Boolean = True
    End Class

    ''' <summary>
    ''' Buscar el factor de conversión entre 2 unidades de medida
    ''' </summary>
    ''' <param name="data">Unidades de medida</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Function FactorDeConversion(ByVal UDMedida As UnidadMedidaInfo, ByVal services As ServiceProvider) As Double
        If Not IsNothing(UDMedida) Then
            Dim oFltr As Filter = New Filter(FilterUnionOperator.Or)
            Dim dblFactor As Double
            Dim blnDividir As Boolean

            Dim oFltrA As Filter = New Filter
            oFltrA.Add("IDUdMedidaA", FilterOperator.Equal, UDMedida.IDUdMedidaA & String.Empty)
            oFltrA.Add("IDUdMedidaB", FilterOperator.Equal, UDMedida.IDUdMedidaB & String.Empty)
            oFltr.Add(oFltrA)

            Dim oFltrB As Filter = New Filter
            oFltrB.Add("IDUdMedidaA", FilterOperator.Equal, UDMedida.IDUdMedidaB & String.Empty)
            oFltrB.Add("IDUdMedidaB", FilterOperator.Equal, UDMedida.IDUdMedidaA & String.Empty)
            oFltr.Add(oFltrB)

            Dim dt As DataTable = New UnidadAB().Filter(oFltr)

            Select Case dt.Rows.Count
                Case 0
                    If UDMedida.UnoSiNoExiste Then
                        dblFactor = 1
                    Else
                        dblFactor = 0
                    End If
                Case 1
                    Dim oRw As DataRow = dt.Rows(0)
                    dblFactor = oRw("Factor")
                    blnDividir = (CStr(oRw("IDUdMedidaA")) = UDMedida.IDUdMedidaB)
                Case 2
                    If CStr(dt.Rows(0)("IDUdMedidaA")) = UDMedida.IDUdMedidaA Then
                        dblFactor = dt.Rows(0)("Factor")
                    Else
                        dblFactor = dt.Rows(1)("Factor")
                    End If
            End Select

            If blnDividir Then
                If dblFactor <> 0 Then
                    Return 1 / dblFactor
                Else
                    Return 0
                End If
            Else
                Return dblFactor
            End If
        End If
    End Function

    ''' <summary>
    ''' Buscar la cantidad equivalente
    ''' </summary>
    ''' <param name="data">Unidades de medida</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Function ConvertirA(ByVal UDMedida As UnidadMedidaInfo, ByVal services As ServiceProvider) As Double
        If Not IsNothing(UDMedida) Then
            Return UDMedida.Cantidad * ProcessServer.ExecuteTask(Of UnidadMedidaInfo, Double)(AddressOf FactorDeConversion, UDMedida, services)
        End If
    End Function

#End Region

End Class