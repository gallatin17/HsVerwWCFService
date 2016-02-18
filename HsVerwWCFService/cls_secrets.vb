Public Class cls_secrets

    Public Sub New()
        MyBase.New
        SmtpServer = "mail.ralfabels.de"
        SmtpServerPwd = "68IbeMP19"
        SmtpServerUser = "gallatin17@ralfabels.de"
    End Sub

    Private _smtpserver As String
    Public Property SmtpServer As String
        Get
            Return _smtpserver
        End Get
        Set(value As String)
            _smtpserver = value
        End Set
    End Property

    Private _smtpserverpwd As String
    Public Property SmtpServerPwd As String
        Get
            Return _smtpserverpwd
        End Get
        Set(value As String)
            _smtpserverpwd = value
        End Set
    End Property

    Private _smtpserveruser As String
    Public Property SmtpServerUser As String
        Get
            Return _smtpserveruser
        End Get
        Set(value As String)
            _smtpserveruser = value
        End Set
    End Property

End Class
