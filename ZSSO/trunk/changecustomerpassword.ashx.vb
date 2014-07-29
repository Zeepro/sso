Imports System.Web
Imports System.Web.Services
Imports System.Text.RegularExpressions.Regex
Imports System.Data.SqlClient
Imports System.Security.Cryptography
Imports System.Runtime.Caching
Imports System.IO
Imports BCrypt.Net.BCrypt

Public Class changecustomerpassword
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal oContext As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim sAdminEmail As String
        Dim sAdminPassword As String
        Dim sCustomerEmail As String
        Dim sCustomerPassword As String


        If ZSSOUtilities.CheckRequests(oContext.Request.UserHostAddress, "changecustomerpassword") > 5 Then
            oContext.Response.ContentType = "text/plain"
            oContext.Response.StatusCode = 435
            oContext.Response.Write("Too many requests")
            ZSSOUtilities.WriteLog("ChangeCustomerPassword : Too many requests")
            Return
        Else
            If oContext.Request.HttpMethod = "GET" Then
                oContext.Response.ContentType = "text/html"
                oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title></head><body><form  method=""post"" action=""/changecustomerpassword.ashx"" accept-charset=""utf-8"">Admin Email <input id=""admin_email"" name=""admin_email"" type=""text"" /><br />Admin Password <input id=""admin_password"" name=""admin_password"" type=""text"" /><br />Customer Email <input id=""customer_email"" name=""customer_email"" type=""text"" /><br />Customer Password <input id=""customer_password"" name=""customer_password"" type=""text"" /><br /><input id=""Submit1"" type=""submit"" value=""Ok"" /></form></body></html>")
            Else
                sAdminEmail = HttpUtility.UrlDecode(oContext.Request.Form("admin_email"))
                sAdminPassword = HttpUtility.UrlDecode(oContext.Request.Form("admin_password"))
                sCustomerEmail = HttpUtility.UrlDecode(oContext.Request.Form("customer_email"))
                sCustomerPassword = HttpUtility.UrlDecode(oContext.Request.Form("customer_password"))

                ZSSOUtilities.WriteLog("ChangeCustomerPassword : " & ZSSOUtilities.oSerializer.Serialize({sCustomerEmail}))
                If String.IsNullOrEmpty(sAdminEmail) Or String.IsNullOrEmpty(sAdminPassword) Or String.IsNullOrEmpty(sCustomerEmail) Or String.IsNullOrEmpty(sCustomerPassword) Then
                    oContext.Response.StatusCode = 432
                    oContext.Response.Write("Missing parameter")
                    ZSSOUtilities.WriteLog("ChangeCustomerPassword : Missing parameter")
                    Return
                End If

                If Not (ZSSOUtilities.oEmailRegex.IsMatch(sAdminEmail)) Or Not (ZSSOUtilities.oEmailRegex.IsMatch(sCustomerEmail)) Or Not testaccount.SearchEmail(sCustomerEmail) Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 433
                    oContext.Response.Write("Incorrect Parameter")
                    ZSSOUtilities.WriteLog("ChangeCustomerPassword : Incorrect parameter")
                    Return
                End If

                Using oConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
                    oConnection.Open()

                    If String.Compare(sAdminEmail, ZSSOUtilities.oZoombaiAdminEmail) <> 0 Or Not ZSSOUtilities.Login(oConnection, sAdminEmail, sAdminPassword) Then
                        oContext.Response.ContentType = "text/plain"
                        oContext.Response.StatusCode = 434
                        oContext.Response.Write("Login failed")
                        ZSSOUtilities.WriteLog("ChangeCustomerPassword : Login failed")
                        Return
                    End If

                    Dim sQuery = "UPDATE Account SET Password = @new_password WHERE email=@email"

                    Using oSqlCmdUpdate As New SqlCommand(sQuery, oConnection)

                        Dim sNewPasswordHash As String = BCrypt.Net.BCrypt.HashPassword(sCustomerPassword, BCrypt.Net.BCrypt.GenerateSalt())

                        oSqlCmdUpdate.Parameters.AddWithValue("@email", sCustomerEmail)
                        oSqlCmdUpdate.Parameters.AddWithValue("@new_password", sNewPasswordHash)

                        Try
                            oSqlCmdUpdate.ExecuteNonQuery()
                        Catch ex As Exception
                            ZSSOUtilities.WriteLog("ChangeCustomerPassword : NOK : " & ex.Message)
                            Return
                        End Try

                    End Using
                End Using
                ZSSOUtilities.WriteLog("ChangeCustomerPassword : OK")
            End If
        End If
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class