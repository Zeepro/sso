Imports System.Web
Imports System.Web.Services
Imports System.Text.RegularExpressions.Regex
Imports System.Data.SqlClient
Imports System.Security.Cryptography
Imports System.Runtime.Caching
Imports System.IO

Public Class deleteaccount
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim Email As String
        Dim Password As String
        Dim cacheMemory As ObjectCache = MemoryCache.Default

        If ZSSOUtilities.CheckRequests(context.Request.UserHostAddress, "deleteaccount") > 5 Then
            context.Response.ContentType = "text/plain"
            context.Response.StatusCode = 435
            context.Response.Write("Too many requests")
            ZSSOUtilities.WriteLog("DeleteAccount : Too many requests")
            Return
        Else
            If context.Request.HttpMethod = "GET" Then
                context.Response.ContentType = "text/html"
                context.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title></head><body><form  method=""post"" action=""/deleteaccount.ashx"" accept-charset=""utf-8"">login <input id=""email"" name=""email"" type=""text"" /><br />password <input id=""password"" name=""password"" type=""text"" /><br /><input id=""Submit1"" type=""submit"" value=""Ok"" /></form></body></html>")
            Else
                Email = HttpUtility.UrlDecode(context.Request.Form("email"))
                Password = HttpUtility.UrlDecode(context.Request.Form("password"))

                ZSSOUtilities.WriteLog("DeleteAccount : " + ZSSOUtilities.oSerializer.Serialize(context.Request.Form))
                If String.IsNullOrEmpty(Email) Or String.IsNullOrEmpty(Password) Then
                    context.Response.StatusCode = 432
                    context.Response.Write("Missing parameter")
                    ZSSOUtilities.WriteLog("DeleteAccount : Missing parameter")
                    Return
                End If

                'check required password pattern
                If Not (ZSSOUtilities.emailExpression.IsMatch(Email)) Then
                    context.Response.ContentType = "text/plain"
                    context.Response.StatusCode = 433
                    context.Response.Write("Incorrect Parameter")
                    ZSSOUtilities.WriteLog("Delete : Incorrect parameter")
                    Return
                End If

                Using oConnexion As New SqlConnection(ZSSOUtilities.getConnectionString())
                    oConnexion.Open()

                    If Not ZSSOUtilities.Login(oConnexion, Email, Password) Then
                        context.Response.ContentType = "text/plain"
                        context.Response.StatusCode = 434
                        context.Response.Write("Login failed")
                        ZSSOUtilities.WriteLog("DeleteAccount : Login failed")
                        Return
                    End If

                    Dim QueryString = "DELETE FROM Account WHERE email=@email"

                    Using oSqlCmdDelete As New SqlCommand(QueryString, oConnexion)
                        oSqlCmdDelete.Parameters.AddWithValue("@email", Email)

                        Try
                            oSqlCmdDelete.ExecuteNonQuery()
                        Catch ex As Exception
                            'context.Response.Write("Error : " + " commande " + ex.Message)
                            Return
                        End Try

                    End Using

                End Using
            End If
        End If
        ZSSOUtilities.WriteLog("DeleteAccount : OK")
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class