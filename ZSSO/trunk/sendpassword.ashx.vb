Imports System.Web
Imports System.Web.Services
Imports System.Text.RegularExpressions.Regex
Imports System.Web.Script.Serialization
Imports System.Data.SqlClient
Imports System.Runtime.Caching
Imports System.IO
Imports System.Security.Cryptography
Imports System.Net.Mail

Public Class sendpassword
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim Email As String
        Dim NewPassword As String
        Dim NewSalt As String
        Dim cachedCounterByIp As Int32
        Dim EmailFound As Boolean = False
        Dim cacheMemory As ObjectCache = MemoryCache.Default

        Dim cachedRequestNewPassword = TryCast(cacheMemory("request_newpassword"), Dictionary(Of String, Integer))
        If IsNothing(cachedRequestNewPassword) Then
            cachedRequestNewPassword = New Dictionary(Of String, Integer)
        End If
        If cachedRequestNewPassword.ContainsKey(context.Request.UserHostAddress) Then
            cachedCounterByIp = CInt(cachedRequestNewPassword(context.Request.UserHostAddress))
            If IsNothing(cachedCounterByIp) Then
                cachedCounterByIp = 0
            End If
        End If

        If cachedCounterByIp > 3 Then
            context.Response.ContentType = "text/plain"
            context.Response.StatusCode = 435
            context.Response.Write("Too many requests")
            ZSSOUtilities.WriteLog("SendPassword : Too many requests")
            Return
        Else
            cachedCounterByIp = cachedCounterByIp + 1
            cachedRequestNewPassword(context.Request.UserHostAddress) = cachedCounterByIp
            cacheMemory.Set("request_newpassword", cachedRequestNewPassword, DateTime.Now.AddMinutes(10.0), Nothing)
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
                NewSalt = System.Web.Security.Membership.GeneratePassword(5, 0)

                Using oConnexion As New SqlConnection("Data Source=(LocalDB)\v11.0;AttachDbFilename=C:\Users\ZPFr1\Desktop\zsso\ZSSO\trunk\App_Data\Database1.mdf;Integrated Security=True;MultipleActiveResultSets=True")
                    oConnexion.Open()

                    Dim QueryString = "UPDATE Account SET Password = @new_password, Salt = @new_salt WHERE email=@email"

                    Using oSqlCmdDelete As New SqlCommand(QueryString, oConnexion)

                        Using md5Hash As MD5 = MD5.Create()

                            oSqlCmdDelete.Parameters.AddWithValue("@email", Email)
                            oSqlCmdDelete.Parameters.AddWithValue("@new_password", ZSSOUtilities.GetMd5Hash(md5Hash, NewPassword + NewSalt))
                            oSqlCmdDelete.Parameters.AddWithValue("@new_salt", NewSalt)

                            Try
                                oSqlCmdDelete.ExecuteNonQuery()
                            Catch ex As Exception
                                context.Response.Write("Error : " + " commande " + ex.Message)
                                Return
                            End Try

                        End Using

                    End Using

                    QueryString = "SELECT HtmlTemplate FROM EmailTemplate WHERE Name = @template_name"
                    Using oSqlSelectTemplace As New SqlCommand(QueryString, oConnexion)
                        oSqlSelectTemplace.Parameters.AddWithValue("@template_name", "sendpassword")

                        Dim QueryResult As SqlDataReader = oSqlSelectTemplace.ExecuteReader()
                        Dim Template As String = ""

                        While QueryResult.Read()
                            Template = QueryResult(QueryResult.GetOrdinal("HtmlTemplate"))
                        End While

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
            MsgBox(ex.Message & vbNewLine & ex.StackTrace)
        End Try

    End Sub
End Class