Imports System.Web
Imports System.Web.Services
Imports System.Text.RegularExpressions.Regex
Imports System.Data.SqlClient
Imports System.Security.Cryptography
Imports System.Runtime.Caching
Imports System.IO

Public Class changepassword
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim Email As String
        Dim OldPassword As String
        Dim NewPassword As String
        Dim AccountSalt = ""
        Dim cacheMemory As ObjectCache = MemoryCache.Default

        If ZSSOUtilities.CheckRequests(context.Request.UserHostAddress) > 5 Then
            context.Response.ContentType = "text/plain"
            context.Response.StatusCode = 435
            context.Response.Write("Too many requests")
            ZSSOUtilities.WriteLog("ChangePassword : Too many requests")
            Return
        Else
            If context.Request.HttpMethod = "GET" Then
                context.Response.ContentType = "text/html"
                context.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title></head><body><form  method=""post"" action=""/changepassword.ashx"" accept-charset=""utf-8"">login <input id=""email"" name=""email"" type=""text"" /><br />old password <input id=""old_password"" name=""old_password"" type=""text"" /><br />new password <input id=""new_password"" name=""new_password"" type=""text"" /><br /><input id=""Submit1"" type=""submit"" value=""Ok"" /></form></body></html>")
            Else
                Email = HttpUtility.UrlDecode(context.Request.Form("email"))
                OldPassword = HttpUtility.UrlDecode(context.Request.Form("old_password"))
                NewPassword = HttpUtility.UrlDecode(context.Request.Form("new_password"))

                ZSSOUtilities.WriteLog("ChangePassword : " + ZSSOUtilities.oSerializer.Serialize(context.Request.Form))
                If String.IsNullOrEmpty(Email) Or String.IsNullOrEmpty(OldPassword) Or String.IsNullOrEmpty(NewPassword) Then
                    context.Response.StatusCode = 432
                    context.Response.Write("Missing parameter")
                    ZSSOUtilities.WriteLog("ChangePassword : Missing parameter")
                    Return
                End If

                'check required password pattern
                If Not (ZSSOUtilities.emailExpression.IsMatch(Email)) Then
                    context.Response.ContentType = "text/plain"
                    context.Response.StatusCode = 433
                    context.Response.Write("Incorrect Parameter")
                    ZSSOUtilities.WriteLog("ChangePassword : Incorrect parameter")
                    Return
                End If

                Using oConnexion As New SqlConnection("Data Source=(LocalDB)\v11.0;AttachDbFilename=C:\Users\ZPFr1\Desktop\zsso\ZSSO\trunk\App_Data\Database1.mdf;Integrated Security=True;MultipleActiveResultSets=True")
                    oConnexion.Open()

                    If Not ZSSOUtilities.Login(oConnexion, Email, OldPassword) Then
                        context.Response.ContentType = "text/plain"
                        context.Response.StatusCode = 434
                        context.Response.Write("Login failed")
                        ZSSOUtilities.WriteLog("ChangePassword : Login failed")
                        Return
                    End If

                    Dim NewSalt As String = System.Web.Security.Membership.GeneratePassword(5, 0)
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