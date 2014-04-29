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

    Sub ProcessRequest(ByVal oContext As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim sEmail As String
        Dim sNewPassword As String
        Dim iCachedCounterByIp As Int32
        Dim bEmailFound As Boolean = False
        Dim oHttpCache As Caching.Cache = HttpRuntime.Cache

        Try
            iCachedCounterByIp = CInt(oHttpCache("request_newpassword_" & oContext.Request.UserHostAddress))
        Catch
            iCachedCounterByIp = 0
        End Try
        If iCachedCounterByIp > 3 Then
            oContext.Response.ContentType = "text/plain"
            oContext.Response.StatusCode = 435
            oContext.Response.Write("Too many requests")
            ZSSOUtilities.WriteLog("SendPassword : Too many requests")
            Return
        Else
            iCachedCounterByIp = iCachedCounterByIp + 1
            oHttpCache.Insert("request_newpassword_" & oContext.Request.UserHostAddress, iCachedCounterByIp, Nothing, DateTime.Now.AddMinutes(10.0), TimeSpan.Zero)
            If oContext.Request.HttpMethod = "GET" Then
                sEmail = HttpUtility.UrlDecode(oContext.Request.QueryString("email"))
                ZSSOUtilities.WriteLog("SendPassword : " & ZSSOUtilities.oSerializer.Serialize({sEmail}))
                If Not testaccount.SearchEmail(sEmail) Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 433
                    oContext.Response.Write("Incorrect Parameter")
                    ZSSOUtilities.WriteLog("SendPassword : Incorrect parameter")
                    Return
                End If

                sNewPassword = System.Web.Security.Membership.GeneratePassword(8, 0)
                Dim sPasswordHash As String = BCrypt.Net.BCrypt.HashPassword(sNewPassword, BCrypt.Net.BCrypt.GenerateSalt())

                Using oConnexion As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
                    oConnexion.Open()

                    Dim sQuery = "UPDATE Account SET Password = @new_password WHERE email=@email AND Deleted IS NULL"

                    Using oSqlCmdUpdate As New SqlCommand(sQuery, oConnexion)

                        oSqlCmdUpdate.Parameters.AddWithValue("@email", sEmail)
                        oSqlCmdUpdate.Parameters.AddWithValue("@new_password", sPasswordHash)

                        Try
                            If oSqlCmdUpdate.ExecuteNonQuery() < 1 Then
                                oContext.Response.ContentType = "text/plain"
                                oContext.Response.StatusCode = 434
                                oContext.Response.Write("Email not found")
                                ZSSOUtilities.WriteLog("SendPassword : Email not found")
                                Return
                            End If
                        Catch ex As Exception
                            Return
                        End Try

                    End Using

                    sQuery = "SELECT TOP 1 HtmlTemplate, Subject FROM EmailTemplate WHERE Name = @template_name"
                    Using oSqlCmdSelect As New SqlCommand(sQuery, oConnexion)
                        oSqlCmdSelect.Parameters.AddWithValue("@template_name", "sendpassword")

                        Using oQueryResult As SqlDataReader = oSqlCmdSelect.ExecuteReader()
                            Dim sTemplate As String = ""
                            Dim sSubject As String = ""

                            If oQueryResult.Read() Then
                                sTemplate = oQueryResult(oQueryResult.GetOrdinal("HtmlTemplate"))
                                sSubject = oQueryResult(oQueryResult.GetOrdinal("Subject"))
                            End If

                            If sTemplate.Length > 0 And sSubject.Length > 0 Then
                                Dim sHtmlTemplate = String.Format(sTemplate, sNewPassword)
                                Dim oHtmlEmail As New Mail
                                oHtmlEmail.sReceiver = sEmail
                                oHtmlEmail.sSubject = sSubject
                                oHtmlEmail.sBody = sHtmlTemplate
                                oHtmlEmail.Send()
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
    Public Property sSubject As String
    Public Property sBody As String
    Public Property sReceiver As String

    Public Sub Send()
        Try
            Dim oSmtpServer As New SmtpClient()
            Dim oMail As New MailMessage()
            oSmtpServer.UseDefaultCredentials = False
            oSmtpServer.Credentials = New Net.NetworkCredential("service-informatique@zee3dcompany.com", "uBXf9JhuFAg7FfeJVAVvkA")
            oSmtpServer.Port = 587
            oSmtpServer.EnableSsl = True
            oSmtpServer.Host = "smtp.mandrillapp.com"

            oMail = New MailMessage()
            oMail.From = New MailAddress("service-informatique@zee3dcompany.com")
            oMail.To.Add(sReceiver)
            oMail.Subject = sSubject
            oMail.IsBodyHtml = True
            oMail.Body = sBody
            oSmtpServer.Send(oMail)
        Catch ex As Exception
            'MsgBox(ex.Message & vbNewLine & ex.StackTrace)
        End Try

    End Sub
End Class