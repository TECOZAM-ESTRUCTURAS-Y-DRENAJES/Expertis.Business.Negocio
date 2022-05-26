Public Class TipoAmortizacionCabecera

#Region "Construtor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbTipoAmortizacionCabecera"

#End Region

#Region "Eventos RegisterDeleteTasks"

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf ComprobarRelaciones)
    End Sub

    <Task()> Public Shared Sub ComprobarRelaciones(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim FilOR As New Filter(FilterUnionOperator.Or)
        FilOR.Add(New StringFilterItem("IDCodigoAmortizacionContable", data("IDTipoAmortizacion")))
        FilOR.Add(New StringFilterItem("IDCodigoAmortizacionTecnica", data("IDTipoAmortizacion")))
        FilOR.Add(New StringFilterItem("IDCodigoAmortizacionFiscal", data("IDTipoAmortizacion")))
        Dim dt As DataTable = New ElementoAmortizable().Filter(FilOR)
        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            ApplicationService.GenerateError("Hay elementos amortizables con este tipo de amortización")
        End If
    End Sub

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ComprobarPorcentaje)
    End Sub

    <Task()> Public Shared Sub ComprobarPorcentaje(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Not data.IsNull("PorcentajeReposicion") AndAlso (data("PorcentajeReposicion") > 100 OrElse data("PorcentajeReposicion") < 0) Then
            ApplicationService.GenerateError("El porcentaje de reposición debe estar entre 0 y 100")
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Function AsignadoTipo(ByVal data As String, ByVal services As ServiceProvider) As Boolean
        'Comprueba si el tipo de Amortizacion esta relacionado con algun ElementoAmortizable.
        'Mirar que el IdTipoAmortizacion exista en GrupoAmortizacion 
        'y a su vez IDGrupoAmortizacion esté en ElementoAmortizable.

        'Si tiene elementos relacionados pero el usuario quiere hacer la modificacion 
        'de todas formas, en presentación puede elegir continuar y dar de alta de todos modos.
        Dim oGrupo As New GrupoAmortizacion
        Dim dtGrupo As DataTable
        dtGrupo = oGrupo.Filter(New StringFilterItem("IdTipoAmortiz", data))
        If Not dtGrupo Is Nothing AndAlso dtGrupo.Rows.Count > 0 Then
            Dim oElem As New ElementoAmortizable
            Dim dtElem As DataTable
            dtElem = oElem.Filter(New StringFilterItem("IdGrupoAmortizacion", dtGrupo.Rows(0)("IdGrupoAmortiz")))
            If Not dtElem Is Nothing AndAlso dtElem.Rows.Count > 0 Then
                'El TipoAmortiz tiene Elementos Amortizables relacionados
                Return True
            Else : Return False
            End If
        Else : Return False
        End If
    End Function

    <Task()> Public Shared Function ObtenerAmortizacionLineal(ByVal data As String, ByVal services As ServiceProvider) As Double
        Dim DtAmortizacionLineal As DataTable = New BE.DataEngine().Filter("vAmortizacionLineal", New FilterItem("IDTipoAmortizacion", FilterOperator.Equal, data))
        If Not DtAmortizacionLineal Is Nothing AndAlso DtAmortizacionLineal.Rows.Count = 1 Then
            Return DtAmortizacionLineal.Rows(0)("PorcentajeAmortizar")
        Else : Return 0
        End If
    End Function

#End Region

End Class