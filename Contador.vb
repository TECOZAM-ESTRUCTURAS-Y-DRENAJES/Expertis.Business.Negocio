Option Strict Off
Option Explicit On

Public Class Contador2
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Private Const cnEntidad As String = "tbMaestroContador2"

    Private Const cCounterIDFieldName As String = "IDContador2"

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Public Overloads Function Update(ByVal dttSource As System.Data.DataTable) As System.Data.DataTable
        ' Validación de datos
        If Not dttSource Is Nothing AndAlso dttSource.Rows.Count > 0 Then
            For Each dr As DataRow In dttSource.Rows

                ' 1.- Que haya introducido la descripción del Contador2.
                If dr("DescContador2").ToString.Trim.Length = 0 Then
                    ApplicationService.GenerateError("Introduzca la descripción del Contador2")
                End If
                ' 2.- Datos obligatorios
                If dr("Longitud").ToString.Trim.Length = 0 Then
                    ApplicationService.GenerateError("Debe establecer valor a Longitud. -")
                End If

                If dr("Contador2Ini").ToString.Trim.Length = 0 Then
                    ApplicationService.GenerateError("Mensaje 5568 - Introduzca el valor inicial del Contador2. -")
                    'ApplicationService.GenerateError(5568, vbCritical, ExpertisApp.Title) 'Introduzca el valor inicial del Contador2
                End If
                If dr("Contador2Fin").ToString.Trim.Length = 0 Then
                    ApplicationService.GenerateError("Mensaje 5567 - Introduzca el valor final del Contador2. -")
                    'ApplicationService.GenerateError(5567, vbCritical, ExpertisApp.Title)  'Introduzca el valor final del Contador2.
                End If

                If dr("Contador2Ini") > dr("Contador2Fin") Then
                    ApplicationService.GenerateError("Mensaje 5566 - El valor del Contador2 inicial es mayor que el valor del Contador2 final. -")
                    'ApplicationService.GenerateError(5566, vbInformation, ExpertisApp.Title)   'El valor del Contador2 inicial es mayor que el valor del Contador2 final.
                End If



                ' Si es un registro nuevo, miramos que la clave no esté vacía ni duplicada
                If dr.RowState = DataRowState.Added Then
                    If dr("IDContador2").ToString.Trim.Length > 0 Then
                        Dim rcsTemp As DataTable
                        rcsTemp = Me.SelOnPrimaryKey(dr("IDContador2"))
                        If Not rcsTemp Is Nothing AndAlso rcsTemp.Rows.Count > 0 Then
                            ApplicationService.GenerateError("Ya existe un Contador2 con esa clave. -")
                        End If
                        rcsTemp = Nothing
                    Else
                        ApplicationService.GenerateError("Introduzca el código del Contador2")
                    End If
                End If

            Next

            AdminData.SetData(dttSource)

        End If

        Return dttSource

    End Function

    Public Function CounterRs(ByVal strEntityName As String) As DataTable

        Dim strProc As String

        strProc = "FwObtenerContador2es " & Quoted(strEntityName)
        Return AdminData.GetData(strProc, False)

    End Function

    Public Shared Function CounterValueTx(ByVal strIDContador2 As String) As CounterTx
        Dim oRslt As New CounterTx

        Dim data As DataTable
        Dim oRw As DataRow

        Dim mMe As New Contador2     'debido a que esta función es shared
        data = mMe.SelOnPrimaryKey(strIDContador2)
        If data Is Nothing OrElse data.Rows.Count = 0 Then
            ApplicationService.GenerateError("No se encontró el Contador2 | ", strIDContador2)
        Else
            oRw = data.Rows(0)
        End If

        oRslt.rcsCounter = data
        oRslt.strCounterValue = FormatCounter(oRw)
        oRw("Contador2") += 1
        Return oRslt
    End Function

    Public Shared Function CounterValue(ByVal strIdCounter As String) As String
        Dim dttCounter As DataTable
        Dim dtrCounter As DataRow
        Dim strCounterValue As String
        Dim mMe As New Contador2     'debido a que esta función es shared
        dttCounter = mMe.SelOnPrimaryKey(strIdCounter)

        If Not dttCounter Is Nothing Then
            If dttCounter.Rows.Count > 0 Then
                dtrCounter = dttCounter.Rows(0)
                strCounterValue = FormatCounter(dtrCounter)

                dtrCounter("Contador2") = CInt(dtrCounter("Contador2")) + 1

                mMe.Update(dttCounter)
                Return strCounterValue
            End If
        End If
    End Function

    Public Shared Function CounterValue(ByVal IDCounter As String, ByVal targetClass As BusinessHelper, ByVal targetField As String, ByVal dateField As String, ByVal dateValue As Date) As String
        Dim dtCounter As DataTable
        Dim Counter As DataRow
        Dim formattedValue As String

        Dim mMe As New Contador2
        dtCounter = mMe.SelOnPrimaryKey(IDCounter)
        If dtCounter.Rows.Count > 0 Then
            Counter = dtCounter.Rows(0)

            formattedValue = FormatCounter(Counter)

            Dim f As New Filter
            f.Add(New StringFilterItem("IDContador2", Counter("IDContador2")))
            f.Add(New StringFilterItem(targetField, formattedValue))
            f.Add(New NumberFilterItem("YEAR(" & dateField & ")", dateValue.Year))
            Dim control As DataTable = targetClass.Filter(f)
            If control.Rows.Count > 0 Then
                ApplicationService.GenerateError("El " & targetField & " ya existe. Modifique el Contador2 correspondiente.")
            End If

            Counter("Contador2") += 1
            mMe.Update(dtCounter)

            Return formattedValue
        End If
    End Function

    Public Shared Function CounterDefault(ByVal strIdEntity As String) As DataTable
        Dim strSql As String

        strSql = "Exec GetCounterDefault '" & strIdEntity & "'"
        Return AdminData.GetData(strSql, False)

    End Function

    Private Shared Function FormatCounter(ByVal Numeric As Boolean, _
                                      ByVal Counter As Integer, _
                                      ByVal mLen As Integer, _
                                      ByVal strText As String) As String

        Dim strCounter As String = CStr(Counter)
        If Numeric Then

        Else
            Dim intPad As Integer = mLen - Len(strCounter) - Len(strText)
            If intPad > 0 Then
                strCounter = strText & New String("0", intPad) & strCounter
            Else
                strCounter = strText & strCounter
            End If
        End If
        Return strCounter
    End Function

    Friend Shared Function FormatCounter(ByVal rwCounter As DataRow) As String
        Dim EsNumerico As Boolean
        Dim texto As String
        Dim longitud As Integer
        Dim ValorContador2 As Integer

        EsNumerico = rwCounter("Numerico")
        ValorContador2 = rwCounter("Contador2")

        If Not rwCounter.IsNull("Longitud") Then
            longitud = rwCounter("Longitud")
        End If
        If rwCounter.IsNull("Texto") Then
            texto = String.Empty
        Else
            texto = rwCounter("Texto")
        End If

        Return FormatCounter(EsNumerico, ValorContador2, longitud, texto)
    End Function

    Friend Shared Function FormattedValueToValue(ByVal formattedValue As String) As Integer
        Dim valor As Integer
        If IsNumeric(formattedValue) Then
            valor = CInt(formattedValue)
        Else
            Dim pos As Integer = 1
            For Each c As Char In formattedValue
                If Char.IsNumber(c) Then
                    valor = CInt(formattedValue.Substring(pos))
                    Exit For
                Else
                    pos += 1
                End If
            Next
        End If

        Return valor
    End Function

    Public Function GetDefaultCounterValue(ByVal EntityName As String) As DefaultCounter
        Dim dtCont As DataTable = CounterDefault(EntityName)
        If Not dtCont Is Nothing AndAlso dtCont.Rows.Count <> 0 Then
            Dim rwC As DataRow = dtCont.Rows(0)
            Dim oRslt As New DefaultCounter
            oRslt.CounterID = dtCont.Rows(0)("IDContador2")
            oRslt.CounterValue = FormatCounter(rwC("Numerico"), rwC("Contador2"), rwC("Longitud"), rwC("Texto") & String.Empty)
            Return oRslt
        End If
    End Function

    Public Shared Sub LoadDefaultCounterValue(ByVal Row As DataRow, _
                                            ByVal EntityName As String, _
                                            ByVal FiledName As String, _
                                            Optional ByVal CounterIDFieldName As String = cCounterIDFieldName)
        Dim ClsCont As New Contador2
        Dim oDC As DefaultCounter = ClsCont.GetDefaultCounterValue(EntityName)
        If Not oDC Is Nothing Then
            Row(FiledName) = oDC.CounterValue
            Row(CounterIDFieldName) = oDC.CounterID
        End If
    End Sub

    Public Shared Sub DecrementCounter(ByVal strIdCounter As String, ByVal Value As String)
        Dim oMe As New Contador2
        Dim oRw As DataRow = oMe.GetItemRow(strIdCounter)
        oRw("Contador2") -= 1
        If FormatCounter(oRw) = Value Then oMe.Update(oRw.Table)
    End Sub

    Public Structure CounterTx
        Public strCounterValue As String
        Public rcsCounter As DataTable
    End Structure

    <Serializable()> _
    Public Class DefaultCounter
        Public CounterValue As String
        Public CounterID As String
    End Class

    Public Overrides Function AddNewForm() As DataTable
        Dim dt As DataTable = MyBase.AddNewForm
        dt.Rows(0)("Contador2Ini") = 1
        dt.Rows(0)("Contador2Fin") = 0
        dt.Rows(0)("Contador2") = 1
        Return dt
    End Function

    Public Function Contador2Numerico(ByVal idContador2 As String) As Boolean
        Dim numerico As Boolean
        Dim dt As DataTable
        dt = SelOnPrimaryKey(idContador2)
        Return dt.Rows(0)("Numerico")
    End Function

End Class

