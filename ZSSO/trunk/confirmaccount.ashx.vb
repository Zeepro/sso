Imports System.Web
Imports System.Web.Services
Imports System.Text.RegularExpressions.Regex
Imports System.Data.SqlClient
Imports System.Security.Cryptography
Imports System.Runtime.Caching
Imports System.IO
Imports BCrypt.Net.BCrypt

Public Class confirmaccount
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal oContext As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim sEmail As String
        Dim sCode As String

        If ZSSOUtilities.CheckRequests(oContext.Request.UserHostAddress, "confirmaccount") > 5 Then
            oContext.Response.ContentType = "text/plain"
            oContext.Response.StatusCode = 435
            oContext.Response.Write("Too many requests")
            ZSSOUtilities.WriteLog("ConfirmAccount : Too many requests")
            Return
        Else
            If oContext.Request.HttpMethod = "GET" Then
                oContext.Response.ContentType = "text/html"
                oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title>" & _
                                        "<script src=""https://code.jquery.com/jquery-1.10.2.js""></script><script type=""text/javascript"">function load_wait() { $(""#overlay"").addClass(""gray-overlay""); $("".ui-loader"").css(""display"", ""block""); }</script>" & _
                                        "<link rel=""stylesheet"" type=""text/css"" href=""style.css"">" & _
                                        "</head><body><div id=""overlay""></div><div class=""ui-loader ui-corner-all ui-body-a ui-loader-default""><span class=""ui-icon-loading""></span><h1>loading</h1></div><form  method=""post"" action=""/confirmaccount.ashx"" accept-charset=""utf-8"">login <input id=""email"" name=""email"" type=""text"" /><br />code <input id=""code"" name=""code"" type=""text"" /><br /><input id=""Submit1"" type=""submit"" value=""Ok""  onclick=""javascript: load_wait();"" /></form></body></html>")
            Else
                sEmail = HttpUtility.UrlDecode(oContext.Request.Form("email"))
                sCode = HttpUtility.UrlDecode(oContext.Request.Form("code"))

                ZSSOUtilities.WriteLog("ConfirmAccount : " & ZSSOUtilities.oSerializer.Serialize({sEmail}))
                If String.IsNullOrEmpty(sEmail) Or String.IsNullOrEmpty(sCode) Then
                    oContext.Response.StatusCode = 432
                    oContext.Response.Write("Missing parameter")
                    ZSSOUtilities.WriteLog("ConfirmAccount : Missing parameter")
                    Return
                End If

                If Not (ZSSOUtilities.oEmailRegex.IsMatch(sEmail)) Or sCode.Count <> 4 Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 433
                    oContext.Response.Write("Incorrect Parameter")
                    ZSSOUtilities.WriteLog("ConfirmAccount : Incorrect parameter")
                    Return
                End If


                Using oConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
                    oConnection.Open()

                    Dim sQuery = "SELECT TOP 1 * " & _
                        "FROM Account " & _
                        "WHERE Email=@email AND Deleted IS NULL"
                    Dim iCodeAccount As Integer = 0

                    'Check account status and get confirm code
                    Using oSqlCmdSelect As New SqlCommand(sQuery, oConnection)
                        oSqlCmdSelect.Parameters.AddWithValue("@email", sEmail.ToLower)

                        Using oQueryResult As SqlDataReader = oSqlCmdSelect.ExecuteReader()
                            If oQueryResult.Read() Then
                                Dim bConfirmed As Boolean = oQueryResult(oQueryResult.GetOrdinal("Confirmed"))
                                If bConfirmed Then
                                    oContext.Response.ContentType = "text/plain"
                                    oContext.Response.StatusCode = 437
                                    oContext.Response.Write("Already exist")
                                    ZSSOUtilities.WriteLog("ConfirmAccount : Already exist (confirmed)")
                                    Return
                                Else
                                    iCodeAccount = oQueryResult(oQueryResult.GetOrdinal("Code"))
                                End If
                            End If
                        End Using

                    End Using

                    'Check confirm code
                    Dim iCode As Integer = 0
                    Try
                        iCode = Convert.ToInt32(sCode)
                    Catch ex As Exception
                        oContext.Response.ContentType = "text/plain"
                        oContext.Response.StatusCode = 433
                        oContext.Response.Write("Incorrect Parameter")
                        ZSSOUtilities.WriteLog("ConfirmAccount : Incorrect parameter (code)")
                        Return
                    End Try

                    If iCode <> iCodeAccount Then
                        oContext.Response.ContentType = "text/plain"
                        oContext.Response.StatusCode = 433
                        oContext.Response.Write("Incorrect Parameter")
                        ZSSOUtilities.WriteLog("ConfirmAccount : Incorrect parameter (bad email or code)")
                        Return
                    End If

                    'Update account status
                    sQuery = "UPDATE Account SET Confirmed=1 WHERE Email=@email"

                    Using oSqlCmUpdate As New SqlCommand(sQuery, oConnection)
                        oSqlCmUpdate.Parameters.AddWithValue("@email", sEmail.ToLower)
                            oSqlCmUpdate.ExecuteNonQuery()

                    End Using
                End Using

                ZSSOUtilities.WriteLog("ConfirmAccount : OK")
            End If
        End If
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property
End Class