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

                    Dim sQuery As String = "INSERT INTO Account (Email, Password) VALUES (@email, @password)"

                    Using oSqlCmdInsert As New SqlCommand(sQuery, oConnection)

                        Dim sPasswordHash As String = BCrypt.Net.BCrypt.HashPassword(sPassword, BCrypt.Net.BCrypt.GenerateSalt())

                        oSqlCmdInsert.Parameters.AddWithValue("email", sEmail)
                        oSqlCmdInsert.Parameters.AddWithValue("password", sPasswordHash)

                        Try
                            oSqlCmdInsert.ExecuteNonQuery()
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