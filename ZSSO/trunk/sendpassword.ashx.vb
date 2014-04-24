Imports System.Web
Imports System.Web.Services
Imports System.Text.RegularExpressions.Regex
Imports System.Web.Script.Serialization
Imports System.Data.SqlClient
Imports System.Runtime.Caching
Imports System.IO
Imports System.Security.Cryptography
Imports System.Net.Mail
Imports BCrypt.Net.BCrypt

Public Class sendpassword
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim Email As String
        Dim NewPassword As String
        Dim cachedCounterByIp As Int32
        Dim EmailFound As Boolean = False
        Dim HttpCache As Caching.Cache = HttpRuntime.Cache

        Try
            cachedCounterByIp = CInt(HttpCache("request_newpassword_" + context.Request.UserHostAddress))
        Catch
            cachedCounterByIp = 0
        End Try
        If cachedCounterByIp > 3 Then
            context.Response.ContentType = "text/plain"
            context.Response.StatusCode = 435
            context.Response.Write("Too many requests")
            ZSSOUtilities.WriteLog("SendPassword : Too many requests")
            Return
        Else
            cachedCounterByIp = cachedCounterByIp + 1
            HttpCache.Insert("request_newpassword_" + context.Request.UserHostAddress, cachedCounterByIp, Nothing, DateTime.Now.AddMinutes(10.0), TimeSpan.Zero)
            If context.Request.HttpMethod = "GET" Then
                Email = HttpUtility.UrlDecode(context.Request.QueryString("email"))
                ZSSOUtilities.WriteLog("SendPassword : " + context.Request.QueryString("email"))
                If Not testaccount.SearchEmail(Email) Then
                    context.Response.ContentType = "text/plain"
                    context.Response.StatusCode = 433
                    context.Response.Write("Incorrect Parameter")
                    ZSSOUtilities.WriteLog("SendPassword : Incorrect parameter")
                    Return
                End If

                NewPassword = System.Web.Security.Membership.GeneratePassword(8, 0)
                Dim Salt = BCrypt.Net.BCrypt.GenerateSalt()
                Dim PasswordHash As String = BCrypt.Net.BCrypt.HashPassword(NewPassword, Salt)

                Using oConnexion As New SqlConnection(ZSSOUtilities.getConnectionString())
                    oConnexion.Open()

                    Dim QueryString = "UPDATE Account SET Password = @new_password WHERE email=@email"

                    Using oSqlCmdUpdate As New SqlCommand(QueryString, oConnexion)

                        oSqlCmdUpdate.Parameters.AddWithValue("@email", Email)
                        oSqlCmdUpdate.Parameters.AddWithValue("@new_password", PasswordHash)

                        Try
                            If oSqlCmdUpdate.ExecuteNonQuery() < 1 Then
                                context.Response.ContentType = "text/plain"
                                context.Response.StatusCode = 434
                                context.Response.Write("Email not found")
                                ZSSOUtilities.WriteLog("SendPassword : Email not found")
                                Return
                            End If
                        Catch ex As Exception
                            Return
                        End Try

                    End Using

                    QueryString = "SELECT TOP 1 HtmlTemplate FROM EmailTemplate WHERE Name = @template_name"
                    Using oSqlSelectTemplace As New SqlCommand(QueryString, oConnexion)
                        oSqlSelectTemplace.Parameters.AddWithValue("@template_name", "sendpassword")

                        Using QueryResult As SqlDataReader = oSqlSelectTemplace.ExecuteReader()
                            Dim Template As String = ""

                            If QueryResult.Read() Then
                                Template = QueryResult(QueryResult.GetOrdinal("HtmlTemplate"))
                            End If

                            If Template.Length > 0 Then
                                Dim HtmlTemplate = String.Format(Template, NewPassword)
                                Dim HtmlEmail As New Mail
                                HtmlEmail.receiver = Email
                                HtmlEmail.subject = "Demande de nouveau mot de passe"
                                HtmlEmail.body = HtmlTemplate
                                HtmlEmail.send()
                            End If
                        End Using
                    End Using
                End Using
            End If
        End If
        ZSSOUtilities.WriteLog("SendPassword : OK")
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class

Public NotInheritable Class Mail
    Public Property subject As String
    Public Property body As String
    Public Property receiver As String

    Public Sub send()
        Try
            Dim smtpServer As New SmtpClient()
            Dim mail As New MailMessage()
            smtpServer.UseDefaultCredentials = False
            smtpServer.Credentials = New Net.NetworkCredential("service-informatique@zee3dcompany.com", "uBXf9JhuFAg7FfeJVAVvkA")
            smtpServer.Port = 587
            smtpServer.EnableSsl = True
            smtpServer.Host = "smtp.mandrillapp.com"

            mail = New MailMessage()
            mail.From = New MailAddress("service-informatique@zee3dcompany.com")
            mail.To.Add(receiver)
            mail.Subject = subject
            mail.IsBodyHtml = True
            mail.Body = body
            smtpServer.Send(mail)
        Catch ex As Exception
            'MsgBox(ex.Message & vbNewLine & ex.StackTrace)
        End Try

    End Sub
End Class