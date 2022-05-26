Public Class ValoresAyB
    Private mLinea As IPropertyAccessor
    Private mImporte As Double
    Private mIDMoneda As String
    Private mCambioA As Double
    Private mCambioB As Double

    Public Sub New()

    End Sub

    '//Sobrecarga utilizada para actualizar importes de Monedas en un registro dato
    Public Sub New(ByVal linea As IPropertyAccessor, ByVal IDMoneda As String, ByVal CambioA As Double, ByVal CambioB As Double)
        mLinea = linea
        mIDMoneda = IDMoneda
        mCambioA = CambioA
        mCambioB = CambioB
    End Sub

    '//Sobrecarga utilizada para calcular importes de Monedas de un importe dado
    Public Sub New(ByVal Importe As Double, ByVal IDMoneda As String, ByVal CambioA As Double, ByVal CambioB As Double)
        mImporte = Importe
        mIDMoneda = IDMoneda
        mCambioA = CambioA
        mCambioB = CambioB
    End Sub

    Public Property Linea() As IPropertyAccessor
        Get
            Return mLinea
        End Get
        Set(ByVal value As IPropertyAccessor)
            mLinea = value
        End Set
    End Property

    Public Property Importe() As Double
        Get
            Return mImporte
        End Get
        Set(ByVal value As Double)
            mImporte = value
        End Set
    End Property

    Public Property IDMoneda() As String
        Get
            Return mIDMoneda
        End Get
        Set(ByVal value As String)
            mIDMoneda = value
        End Set
    End Property

    Public Property CambioA() As Double
        Get
            Return mCambioA
        End Get
        Set(ByVal value As Double)
            mCambioA = value
        End Set
    End Property

    Public Property CambioB() As Double
        Get
            Return mCambioB
        End Get
        Set(ByVal value As Double)
            mCambioB = value
        End Set
    End Property

End Class