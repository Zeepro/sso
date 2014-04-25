Imports System.Web
Imports System.Web.Services
Imports System.Text.RegularExpressions.Regex
Imports System.Data.SqlClient
Imports System.Security.Cryptography
Imports System.Runtime.Caching
Imports System.IO

Public Class deleteaccount
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal oContext As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim sEmail As String
        Dim sPassword As String

        If ZSSOUtilities.CheckRequests(oContext.Request.UserHostAddress, "deleteaccount") > 5 Then
            oContext.Response.ContentType = "text/plain"
            oContext.Response.StatusCode = 435
            oContext.Response.Write("Too many requests")
            ZSSOUtilities.WriteLog("DeleteAccount : Too many requests")
            Return
        Else
            If oContext.Request.HttpMethod = "GET" Then
                oContext.Response.ContentType = "text/html"
                oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title></head><body><form  method=""post"" action=""/deleteaccount.ashx"" accept-charset=""utf-8"">login <input id=""email"" name=""email"" type=""text"" /><br />password <input id=""password"" name=""password"" type=""text"" /><br /><input id=""Submit1"" type=""submit"" value=""Ok"" /></form></body></html>")
            Else
                sEmail = HttpUtility.UrlDecode(oContext.Request.Form("email"))
                sPassword = HttpUtility.UrlDecode(oContext.Request.Form("password"))

                ZSSOUtilities.WriteLog("DeleteAccount : " & ZSSOUtilities.oSerializer.Serialize(oContext.Request.Form))
                If String.IsNullOrEmpty(sEmail) Or String.IsNullOrEmpty(sPassword) Then
                    oContext.Response.StatusCode = 432
                    oContext.Response.Write("Missing parameter")
                    ZSSOUtilities.WriteLog("DeleteAccount : Missing parameter")
                    Return
                End If

                If Not (ZSSOUtilities.oEmailRegex.IsMatch(sEmail)) Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 433
                    oContext.Response.Write("Incorrect Parameter")
                    ZSSOUtilities.WriteLog("Delete : Incorrect parameter")
                    Return
                End If

                Using oConnexion As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
                    oConnexion.Open()

                    If Not ZSSOUtilities.Login(oConnexion, sEmail, sPassword) Then
                        oContext.Response.ContentType = "text/plain"
                        oContext.Response.StatusCode = 434
                        oContext.Response.Write("Login failed")
                        ZSSOUtilities.WriteLog("DeleteAccount : Login failed")
                        Return
                    End If

                    Dim sQuery = "UPDATE Account SET [Delete] = GETDATE() WHERE email=@email"

                    Using oSqlCmdUpdate As New SqlCommand(sQuery, oConnexion)
                        oSqlCmdUpdate.Parameters.AddWithValue("@email", sEmail)

                        Try
                            oSqlCmdUpdate.ExecuteNonQuery()
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