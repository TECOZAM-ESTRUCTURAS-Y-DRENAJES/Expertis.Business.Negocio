Public Class ProvisionalCounter
    Private ht As Hashtable = New Hashtable
    Public Sub AddCounter(ByVal strIDCounter As String, ByVal rwCounter As DataRow)
        ht.Add(strIDCounter, rwCounter)
    End Sub
    Public Function GetCounter(ByVal strIDCounter As String) As DataRow
        Return ht.Item(strIDCounter)
    End Function
    Public Function GetCounterValue(ByVal strIDCounter As String) As String
        Dim rwCounter As DataRow = ht.Item(strIDCounter)
        If rwCounter Is Nothing Then
            Dim oCntr As Contador.CounterTx = ProcessServer.ExecuteTask(Of String, Contador.CounterTx)(AddressOf Contador.CounterValueTx, strIDCounter, New ServiceProvider)
            ht.Add(strIDCounter, oCntr.DtCounter.Rows(0))
            Return oCntr.strCounterValue
        Else
            Dim strCounter As String = ProcessServer.ExecuteTask(Of DataRow, String)(AddressOf Contador.FormatCounterDr, rwCounter, New ServiceProvider)
            rwCounter("Contador") += 1
            Return strCounter
        End If
    End Function
End Class
