Public Class Mail

    Private _credentialName As String = "no.reply@polestar.jp"
    Private _credentialDisplayName As String = "Polestar No Reply"
    Private _credentialPassword As String = "3Edc4Rfv,"
    Private _host As String = "smtp.gmail.com"
    Private _port As Long = 587

    Private _subject As String
    Public Property Subject As String
        Get
            Return Me._subject
        End Get
        Set(value As String)
            Me._subject = Me.EncodeMailHeader(value)
        End Set
    End Property

    Private _body As String
    Public Property Body As String
        Get
            Return Me._body
        End Get
        Set(value As String)
            Me._body = value
        End Set
    End Property

    Private _addresses() As String
    Public Property Addresses As String()
        Get
            Return Me._addresses
        End Get
        Set(value As String())
            Me._addresses = value
        End Set
    End Property

    Public Function Send() As Long

        Dim _smtp As System.Net.Mail.SmtpClient
        Dim _message As System.Net.Mail.MailMessage

        '-------------------------
        '   Create Mail Body
        _message = New System.Net.Mail.MailMessage(
            New System.Net.Mail.MailAddress(Me._credentialName, Me._credentialDisplayName),
            New System.Net.Mail.MailAddress(Me._credentialName))

        _message.Subject = Me.Subject
        _message.Body = Me.Body
        _message.BodyEncoding = System.Text.Encoding.GetEncoding("iso-2022-jp")

        For Each _to As String In Me.Addresses
            Dim _a As New System.Net.Mail.MailAddress(_to)
            _message.To.Add(_a)

        Next

        '-------------------------
        '   Send Email
        _smtp = New System.Net.Mail.SmtpClient(Me._host, Me._port)
        _smtp.Credentials = New System.Net.NetworkCredential(Me._credentialName, Me._credentialPassword)
        _smtp.EnableSsl = True
        _smtp.Send(_message)
        _message.Dispose()
        _smtp.Dispose()

        Return 0
    End Function

    Private Function EncodeMailHeader(subject As String) As String

        Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("iso-2022-jp")
        Dim base64 As String = Convert.ToBase64String((enc.GetBytes(subject)))
        Return String.Format("=?{0}?B?{1}?=", "iso-2022-jp", base64)

    End Function

End Class
