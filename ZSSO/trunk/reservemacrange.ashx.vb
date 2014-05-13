Imports System.Web
Imports System.Web.Services
Imports System.Text.RegularExpressions.Regex
Imports System.Data.SqlClient
Imports System.Security.Cryptography
Imports System.Runtime.Caching
Imports System.IO

Public Class reservemacrange
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal oContext As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim sEmail As String
        Dim sPassword As String
        Dim sRangeCode As String

        If ZSSOUtilities.CheckRequests(oContext.Request.UserHostAddress, "reservemacrange") > 5 Then
            oContext.Response.ContentType = "text/plain"
            oContext.Response.StatusCode = 435
            oContext.Response.Write("Too many requests")
            ZSSOUtilities.WriteLog("ReserveMacRange : Too many requests")
            Return
        Else
            If oContext.Request.HttpMethod = "GET" Then
                oContext.Response.ContentType = "text/html"
                oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title></head><body><form  method=""post"" action=""/reservemacrange.ashx"" accept-charset=""utf-8"">login <input id=""email"" name=""email"" type=""text"" /><br />password <input id=""password"" name=""password"" type=""text"" /><br />range code <input id=""rangecode"" name=""rangecode"" type=""text"" /><br /><input id=""Submit1"" type=""submit"" value=""Ok"" /></form></body></html>")
            Else
                sEmail = HttpUtility.UrlDecode(oContext.Request.Form("email"))
                sPassword = HttpUtility.UrlDecode(oContext.Request.Form("password"))
                sRangeCode = HttpUtility.UrlDecode(oContext.Request.Form("rangecode"))

                ZSSOUtilities.WriteLog("ReserveMacRange : " & ZSSOUtilities.oSerializer.Serialize({sEmail, sRangeCode}))
                If String.IsNullOrEmpty(sEmail) Or String.IsNullOrEmpty(sPassword) Or String.IsNullOrEmpty(sRangeCode) Then
                    oContext.Response.StatusCode = 432
                    oContext.Response.Write("Missing parameter")
                    ZSSOUtilities.WriteLog("ReserveMacRange : Missing parameter")
                    Return
                End If

                If Not (ZSSOUtilities.oEmailRegex.IsMatch(sEmail)) Or Not (ZSSOUtilities.CheckRangeCode(sRangeCode)) Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 433
                    oContext.Response.Write("Incorrect Parameter")
                    ZSSOUtilities.WriteLog("ReserveMacRange : Incorrect parameter")
                    Return
                End If

                Using oConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
                    oConnection.Open()

                    If String.Compare(sEmail, ZSSOUtilities.oAdminEmail) <> 0 Or Not ZSSOUtilities.Login(oConnection, sEmail, sPassword) Then
                        oContext.Response.ContentType = "text/plain"
                        oContext.Response.StatusCode = 434
                        oContext.Response.Write("Login failed")
                        ZSSOUtilities.WriteLog("ReserveMacRange : Login failed")
                        Return
                    End If

                    If (ZSSOUtilities.CheckRangeStart(oConnection, sRangeCode.Substring(0, 12)) = False) Then
                        oContext.Response.ContentType = "text/plain"
                        oContext.Response.StatusCode = 438
                        oContext.Response.Write("Unavailable range")
                        ZSSOUtilities.WriteLog("ReserveMacRange : Unavailable range")
                        Return
                    End If

                    Dim sQuery = "UPDATE LastAddress SET Value = @value"

                    Using oSqlCmdUpdate As New SqlCommand(sQuery, oConnection)
                        oSqlCmdUpdate.Parameters.AddWithValue("@value", CLng("&H" & sRangeCode.Substring(12, 12)) + 1)

                        Try
                            oSqlCmdUpdate.ExecuteNonQuery()
                        Catch ex As Exception
                            ZSSOUtilities.WriteLog("ReserveMacRange : NOK : " & ex.Message)
                            Return
                        End Try

                    End Using
                End Using
                oContext.Response.ContentType = "text/plain"
                oContext.Response.Write("ok")
                ZSSOUtilities.WriteLog("ReserveMacRange : OK")
            End If
        End If
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class