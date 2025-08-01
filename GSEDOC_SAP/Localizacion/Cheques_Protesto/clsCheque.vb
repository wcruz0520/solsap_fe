Public Class clsCheque
    Private _NumPago As Integer
    Public Property NumPago() As Integer
        Get
            Return _NumPago
        End Get
        Set(ByVal value As Integer)
            _NumPago = value
        End Set
    End Property

    Private _Cheque_Num As String
    Public Property Cheque_Num() As String
        Get
            Return _Cheque_Num
        End Get
        Set(ByVal value As String)
            _Cheque_Num = value
        End Set
    End Property

    Private _Cheque_Valor As Decimal
    Public Property Cheque_Valor() As Decimal
        Get
            Return _Cheque_Valor
        End Get
        Set(ByVal value As Decimal)
            _Cheque_Valor = value
        End Set
    End Property

    Private _Banco As String
    Public Property Banco() As String
        Get
            Return _Banco
        End Get
        Set(ByVal value As String)
            _Banco = value
        End Set
    End Property

    Private _Cliente_Codigo As String
    Public Property Cliente_Codigo() As String
        Get
            Return _Cliente_Codigo
        End Get
        Set(ByVal value As String)
            _Cliente_Codigo = value
        End Set
    End Property

    Private _Cliente As String
    Public Property Cliente() As String
        Get
            Return _Cliente
        End Get
        Set(ByVal value As String)
            _Cliente = value
        End Set
    End Property



    Private _Doc_Protesto As String
    Public Property Doc_Protesto() As String
        Get
            Return _Doc_Protesto
        End Get
        Set(ByVal value As String)
            _Doc_Protesto = value
        End Set
    End Property

    Private _Pago_Coments As String
    Public Property Pago_Coments() As String
        Get
            Return _Pago_Coments
        End Get
        Set(ByVal value As String)
            _Pago_Coments = value
        End Set
    End Property

    Private _Doc_Coments As String
    Public Property Doc_Coments() As String
        Get
            Return _Doc_Coments
        End Get
        Set(ByVal value As String)
            _Doc_Coments = value
        End Set
    End Property

    Private _CuentaContableDeposito As String
    Public Property CuentaContableDeposito() As String
        Get
            Return _CuentaContableDeposito
        End Get
        Set(ByVal value As String)
            _CuentaContableDeposito = value
        End Set
    End Property

    Private _NombreCuentaContableDeposito As String
    Public Property NombreCuentaContableDeposito() As String
        Get
            Return _NombreCuentaContableDeposito
        End Get
        Set(ByVal value As String)
            _NombreCuentaContableDeposito = value
        End Set
    End Property

    Private _NumeroDeposito As String
    Public Property NumeroDeposito() As String
        Get
            Return _NumeroDeposito
        End Get
        Set(ByVal value As String)
            _NumeroDeposito = value
        End Set
    End Property

    Sub New(NumPago As Integer, Cheque_Num As String, Cheque_Valor As String, Banco As String, Cliente_Codigo As String, Cliente As String,
            Doc_Protesto As String, Pago_Coments As String, CuentaContableDeposito As String, NombreCuentaContableDeposito As String, NumeroDeposito As String)

        _NumPago = NumPago
        _Cheque_Num = Cheque_Num
        _Cheque_Valor = Cheque_Valor
        _Banco = Banco
        _Cliente_Codigo = Cliente_Codigo
        _Cliente = Cliente
        _Doc_Protesto = Doc_Protesto
        _Pago_Coments = Pago_Coments
        _CuentaContableDeposito = CuentaContableDeposito
        _NombreCuentaContableDeposito = NombreCuentaContableDeposito
        _NumeroDeposito = NumeroDeposito

    End Sub

    Sub New()

    End Sub
End Class
