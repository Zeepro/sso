Imports System.Web
Imports System.Web.Services
Imports System.Text.RegularExpressions.Regex
Imports System.Data.SqlClient
Imports System.Security.Cryptography
Imports System.Runtime.Caching
Imports System.IO
Imports BCrypt.Net.BCrypt

Public Class createaccount
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal oContext As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim sEmail As String
        Dim sPassword As String

        If ZSSOUtilities.CheckRequests(oContext.Request.UserHostAddress, "createaccount") > 5 Then
            oContext.Response.ContentType = "text/plain"
            oContext.Response.StatusCode = 435
            oContext.Response.Write("Too many requests")
            ZSSOUtilities.WriteLog("CreateAccount : Too many requests")
            Return
        Else
            If oContext.Request.HttpMethod = "GET" Then
                oContext.Response.ContentType = "text/html"
                oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title></head><body><form  method=""post"" action=""/createaccount.ashx"" accept-charset=""utf-8"">login <input id=""email"" name=""email"" type=""text"" /><br />password <input id=""password"" name=""password"" type=""text"" /><br /><input id=""Submit1"" type=""submit"" value=""Ok"" /></form></body></html>")
            Else
                sEmail = HttpUtility.UrlDecode(oContext.Request.Form("email"))
                sPassword = HttpUtility.UrlDecode(oContext.Request.Form("password"))

                ZSSOUtilities.WriteLog("CreateAccount : " & ZSSOUtilities.oSerializer.Serialize({sEmail}))
                If String.IsNullOrEmpty(sEmail) Or String.IsNullOrEmpty(sPassword) Then
                    oContext.Response.StatusCode = 432
                    oContext.Response.Write("Missing parameter")
                    ZSSOUtilities.WriteLog("CreateAccount : Missing parameter")
                    Return
                End If

                If Not (ZSSOUtilities.oEmailRegex.IsMatch(sEmail)) Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 433
                    oContext.Response.Write("Incorrect Parameter")
                    ZSSOUtilities.WriteLog("CreateAccount : Incorrect parameter")
                    Return
                End If

                Using oConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
                    oConnection.Open()

                    Dim oSqlCmd As SqlCommand = oConnection.CreateCommand()
                    Dim oTransaction As SqlTransaction = oConnection.BeginTransaction("createaccount")

                    oSqlCmd.Connection = oConnection
                    oSqlCmd.Transaction = oTransaction

                    Dim oRandom As Random = New Random()
                    Dim iCode As Integer = oRandom.Next(1, 9999)
                    Dim sPasswordHash As String = BCrypt.Net.BCrypt.HashPassword(sPassword, BCrypt.Net.BCrypt.GenerateSalt())

                    Try
                        'Insert new account
                        oSqlCmd.CommandText = "UPDATE Account SET Password=@password, Code=@code WHERE Email=@email IF @@ROWCOUNT=0 INSERT INTO Account (Email, Password, Code) VALUES (@email, @password, @code)"
                        oSqlCmd.Parameters.AddWithValue("@email", sEmail)
                        oSqlCmd.Parameters.AddWithValue("@password", sPasswordHash)
                        oSqlCmd.Parameters.AddWithValue("@code", iCode)
                        oSqlCmd.ExecuteNonQuery()

                        'Select email template
                        oSqlCmd.CommandText = "SELECT TOP 1 HtmlTemplate, Subject FROM EmailTemplate WHERE Name = @template_name"
                        oSqlCmd.Parameters.AddWithValue("@template_name", "createaccount")

                        oTransaction.Commit()

                        'Send email with confirm code
                        Using oQueryResult As SqlDataReader = oSqlCmd.ExecuteReader()

                            Dim sTemplate As String = ""
                            Dim sSubject As String = ""

                            If oQueryResult.Read() Then
                                sTemplate = oQueryResult(oQueryResult.GetOrdinal("HtmlTemplate"))
                                sSubject = oQueryResult(oQueryResult.GetOrdinal("Subject"))
                            End If

                            If sTemplate.Length > 0 And sSubject.Length > 0 Then
                                Dim sHtmlTemplate = String.Format(sTemplate, String.Format("{0:0000}", iCode))
                                Dim oHtmlEmail As New Mail
                                oHtmlEmail.sReceiver = sEmail
                                oHtmlEmail.sSubject = sSubject
                                oHtmlEmail.sBody = sHtmlTemplate
                                oHtmlEmail.Send()
                            End If
                        End Using

                    Catch ex As Exception
                        If ex.HResult = -2146232060 Then
                            oContext.Response.ContentType = "text/plain"
                            oContext.Response.StatusCode = 437
                            oContext.Response.Write("Already Exist")
                            ZSSOUtilities.WriteLog("CreateAccount : Already exist")
                        End If
                        ZSSOUtilities.WriteLog("CreateAccount : NOK : " & ex.Message)
                        Return
                    End Try

                End Using
                ZSSOUtilities.WriteLog("CreateAccount : OK")
            End If
        End If
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property
End Class