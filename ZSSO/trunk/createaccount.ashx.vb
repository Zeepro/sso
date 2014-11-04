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
        Dim sLanguage As String = "en"

        If ZSSOUtilities.CheckRequests(oContext.Request.UserHostAddress, "createaccount") > 5 Then
            oContext.Response.ContentType = "text/plain"
            oContext.Response.StatusCode = 435
            oContext.Response.Write("Too many requests")
            ZSSOUtilities.WriteLog("CreateAccount : Too many requests")
            Return
        Else
            If oContext.Request.HttpMethod = "GET" Then
                oContext.Response.ContentType = "text/html"
                oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title>" & _
                                        "<script src=""https://code.jquery.com/jquery-1.10.2.js""></script><script type=""text/javascript"">function load_wait() { $(""#overlay"").addClass(""gray-overlay""); $("".ui-loader"").css(""display"", ""block""); }</script>" & _
                                        "<link rel=""stylesheet"" type=""text/css"" href=""style.css"">" & _
                                        "</head><body><div id=""overlay""></div><div class=""ui-loader ui-corner-all ui-body-a ui-loader-default""><span class=""ui-icon-loading""></span><h1>loading</h1></div><form  method=""post"" action=""/createaccount.ashx"" accept-charset=""utf-8"">login <input id=""email"" name=""email"" type=""text"" /><br />password <input id=""password"" name=""password"" type=""text"" /><br />language <input id=""language"" name=""language"" type=""text"" /><input id=""Submit1"" type=""submit"" value=""Ok""  onclick=""javascript: load_wait();"" /></form></body></html>")
            Else
                sEmail = HttpUtility.UrlDecode(oContext.Request.Form("email"))
                sPassword = HttpUtility.UrlDecode(oContext.Request.Form("password"))
                If String.IsNullOrEmpty(oContext.Request.Form("language")) = False Then
                    sLanguage = HttpUtility.UrlDecode(oContext.Request.Form("language"))
                End If

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

                If testaccount.SearchEmail(sEmail) Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 437
                    oContext.Response.Write("Already Exist")
                    ZSSOUtilities.WriteLog("CreateAccount : Already exist")
                    Return
                End If

                Using oConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
                    oConnection.Open()

                    'Check email template language
                    Dim sQuery = "SELECT TOP 1 * FROM EmailTemplate WHERE Name = @template_name AND Language = @language"

                    Using oSqlCmdSelect As New SqlCommand(sQuery, oConnection)
                        oSqlCmdSelect.Parameters.AddWithValue("@template_name", "createaccount")
                        oSqlCmdSelect.Parameters.AddWithValue("@language", sLanguage.ToLower)
                        Try
                            Using oQueryResult As SqlDataReader = oSqlCmdSelect.ExecuteReader()
                                If oQueryResult.HasRows = False Then
                                    sLanguage = "en"
                                End If
                            End Using
                        Catch
                        End Try
                    End Using

                    Dim oSqlCmd As SqlCommand = oConnection.CreateCommand()
                    Dim oTransaction As SqlTransaction = oConnection.BeginTransaction("createaccount")

                    oSqlCmd.Connection = oConnection
                    oSqlCmd.Transaction = oTransaction

                    Dim oRandom As Random = New Random()
                    Dim iCode As Integer = oRandom.Next(1, 9999)
                    Dim sPasswordHash As String = BCrypt.Net.BCrypt.HashPassword(sPassword, BCrypt.Net.BCrypt.GenerateSalt())

                    Try
                        'Insert new account
                        oSqlCmd.CommandText = "UPDATE Account SET Password=@password, Code=@code WHERE Email=@email IF @@ROWCOUNT=0 INSERT INTO Account (Email, Password, Code, Language) VALUES (@email, @password, @code, @language)"
                        oSqlCmd.Parameters.AddWithValue("@email", sEmail.ToLower)
                        oSqlCmd.Parameters.AddWithValue("@password", sPasswordHash)
                        oSqlCmd.Parameters.AddWithValue("@code", iCode)
                        oSqlCmd.Parameters.AddWithValue("@language", sLanguage.ToLower)
                        oSqlCmd.ExecuteNonQuery()

                        'Select email template
                        oSqlCmd.CommandText = "SELECT TOP 1 HtmlTemplate, Subject FROM EmailTemplate WHERE Name = @template_name AND Language = @language"
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
                            Return
                        End If
                        oSqlCmd.Transaction.Rollback()
                        Throw ex
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