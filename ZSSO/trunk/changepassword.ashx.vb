Imports System.Web
Imports System.Web.Services
Imports System.Text.RegularExpressions.Regex
Imports System.Data.SqlClient
Imports System.Security.Cryptography
Imports System.Runtime.Caching
Imports System.IO
Imports BCrypt.Net.BCrypt

Public Class changepassword
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal oContext As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim sEmail As String
        Dim sOldPassword As String
        Dim sNewPassword As String

        If ZSSOUtilities.CheckRequests(oContext.Request.UserHostAddress, "changepassword") > 5 Then
            oContext.Response.ContentType = "text/plain"
            oContext.Response.StatusCode = 435
            oContext.Response.Write("Too many requests")
            ZSSOUtilities.WriteLog("ChangePassword : Too many requests")
            Return
        Else
            If oContext.Request.HttpMethod = "GET" Then
                oContext.Response.ContentType = "text/html"
                oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title></head><body><form  method=""post"" action=""/changepassword.ashx"" accept-charset=""utf-8"">login <input id=""email"" name=""email"" type=""text"" /><br />old password <input id=""old_password"" name=""old_password"" type=""text"" /><br />new password <input id=""new_password"" name=""new_password"" type=""text"" /><br /><input id=""Submit1"" type=""submit"" value=""Ok"" /></form></body></html>")
            Else
                sEmail = HttpUtility.UrlDecode(oContext.Request.Form("email"))
                sOldPassword = HttpUtility.UrlDecode(oContext.Request.Form("old_password"))
                sNewPassword = HttpUtility.UrlDecode(oContext.Request.Form("new_password"))

                ZSSOUtilities.WriteLog("ChangePassword : " & ZSSOUtilities.oSerializer.Serialize(oContext.Request.Form))
                If String.IsNullOrEmpty(sEmail) Or String.IsNullOrEmpty(sOldPassword) Or String.IsNullOrEmpty(sNewPassword) Then
                    oContext.Response.StatusCode = 432
                    oContext.Response.Write("Missing parameter")
                    ZSSOUtilities.WriteLog("ChangePassword : Missing parameter")
                    Return
                End If

                If Not (ZSSOUtilities.oEmailRegex.IsMatch(sEmail)) Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 433
                    oContext.Response.Write("Incorrect Parameter")
                    ZSSOUtilities.WriteLog("ChangePassword : Incorrect parameter")
                    Return
                End If

                Using oConnexion As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
                    oConnexion.Open()

                    If Not ZSSOUtilities.Login(oConnexion, sEmail, sOldPassword) Then
                        oContext.Response.ContentType = "text/plain"
                        oContext.Response.StatusCode = 434
                        oContext.Response.Write("Login failed")
                        ZSSOUtilities.WriteLog("ChangePassword : Login failed")
                        Return
                    End If

                    Dim sQuery = "UPDATE Account SET Password = @new_password WHERE email=@email"

                    Using oSqlCmdUpdate As New SqlCommand(sQuery, oConnexion)

                        Dim sNewPasswordHash As String = BCrypt.Net.BCrypt.HashPassword(sNewPassword, BCrypt.Net.BCrypt.GenerateSalt())

                        oSqlCmdUpdate.Parameters.AddWithValue("@email", sEmail)
                        oSqlCmdUpdate.Parameters.AddWithValue("@new_password", sNewPasswordHash)

                        Try
                            oSqlCmdUpdate.ExecuteNonQuery()
                        Catch ex As Exception
                            Return
                        End Try

                    End Using
                End Using
            End If
        End If
        ZSSOUtilities.WriteLog("ChangePassword : OK")
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class