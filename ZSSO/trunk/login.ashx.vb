Imports System.Web
Imports System.Web.Services
Imports System.Text.RegularExpressions.Regex
Imports System.Data.SqlClient
Imports System.Security.Cryptography
Imports System.Runtime.Caching
Imports System.IO

Public Class login
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal oContext As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim sEmail As String
        Dim sPassword As String
        Dim sOptin As String
        Dim lOptin As Nullable(Of Boolean)
        Dim iStatusCode As Integer = 202

        If oContext.Request.HttpMethod = "GET" Then
            oContext.Response.ContentType = "text/html"
            oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title>" & _
                                    "<script src=""https://code.jquery.com/jquery-1.10.2.js""></script><script type=""text/javascript"">function load_wait() { $(""#overlay"").addClass(""gray-overlay""); $("".ui-loader"").css(""display"", ""block""); }</script>" & _
                                    "<link rel=""stylesheet"" type=""text/css"" href=""style.css"">" & _
                                    "</head><body><div id=""overlay""></div><div class=""ui-loader ui-corner-all ui-body-a ui-loader-default""><span class=""ui-icon-loading""></span><h1>loading</h1></div>" & _
                                    "<form  method=""post"" action=""/login.ashx"" accept-charset=""utf-8"">" & _
                                    "login <input id=""email"" name=""email"" type=""text"" /><br />" & _
                                    "password <input id=""password"" name=""password"" type=""text"" /><br />" & _
                                    "optin (on/off) <input id=""optin"" name=""optin"" type=""text"" /><br />" & _
                                    "<input id=""Submit1"" type=""submit"" value=""Ok""  onclick=""javascript: load_wait();"" /></form></body></html>")
        Else
            sEmail = HttpUtility.UrlDecode(oContext.Request.Form("email"))
            sPassword = HttpUtility.UrlDecode(oContext.Request.Form("password"))
            sOptin = HttpUtility.UrlDecode(oContext.Request.Form("optin")).ToLower
            If sOptin <> "" Then
                If sOptin = "on" Then
                    lOptin = True
                Else
                    lOptin = False
                End If
            End If

            ZSSOUtilities.WriteLog("Login : " & ZSSOUtilities.oSerializer.Serialize({sEmail}))
            If String.IsNullOrEmpty(sEmail) Or String.IsNullOrEmpty(sPassword) Then
                oContext.Response.StatusCode = 432
                oContext.Response.Write("Missing parameter")
                ZSSOUtilities.WriteLog("Login : Missing parameter")
                Return
            End If

            If Not (ZSSOUtilities.oEmailRegex.IsMatch(sEmail)) Then
                oContext.Response.ContentType = "text/plain"
                oContext.Response.StatusCode = 433
                oContext.Response.Write("Incorrect Parameter")
                ZSSOUtilities.WriteLog("Login : Incorrect parameter")
                Return
            End If

            Using oConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
                oConnection.Open()

                If Not testaccount.SearchEmail(sEmail) Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 439
                    oContext.Response.Write("Unknown email")
                    ZSSOUtilities.WriteLog("Login : Unknown email")
                    Return
                End If


                If Not ZSSOUtilities.Login(oConnection, sEmail, sPassword) Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 434
                    oContext.Response.Write("Login failed")
                    ZSSOUtilities.WriteLog("Login : Login failed")
                    Return
                End If
                oContext.Response.StatusCode = 202

                ' Optin update
                Try
                    If Not lOptin Is Nothing Then
                        Using oSqlCmd As New SqlCommand("UPDATE Account SET optin = @optin WHERE email = @email", oConnection)
                            oSqlCmd.Parameters.AddWithValue("@email", sEmail)
                            If lOptin Then
                                oSqlCmd.Parameters.AddWithValue("@optin", 1)
                            Else
                                oSqlCmd.Parameters.AddWithValue("@optin", 0)
                            End If
                            oSqlCmd.ExecuteNonQuery()
                        End Using
                    End If
                Catch ex As Exception
                    ZSSOUtilities.WriteLog("Login: " & ex.Message)
                End Try

            End Using
            ZSSOUtilities.WriteLog("Login : OK")
        End If
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class