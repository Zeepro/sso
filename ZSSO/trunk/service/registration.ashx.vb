'TODO: Sort token by date to avoid/manage duplicate (if duplicate can't be avoid)...

Imports System.Web
Imports System.Web.Services
Imports System.Data.SqlClient

Public Class registration
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal oContext As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim sToken As String
        Dim sService As String
        Dim sURL As String
        Dim sState As String

        If oContext.Request.HttpMethod = "GET" Then
            oContext.Response.ContentType = "text/html"
            oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title>" & _
                                    "<script src=""https://code.jquery.com/jquery-1.10.2.js""></script><script type=""text/javascript"">function load_wait() { $(""#overlay"").addClass(""gray-overlay""); $("".ui-loader"").css(""display"", ""block""); }</script>" & _
                                    "<link rel=""stylesheet"" type=""text/css"" href=""style.css"">" & _
                                    "</head><body><div id=""overlay""></div><div class=""ui-loader ui-corner-all ui-body-a ui-loader-default""><span class=""ui-icon-loading""></span><h1>Registration</h1></div>" & _
                                    "<form  method=""post"" action=""registration.ashx"" accept-charset=""utf-8"">" & _
                                    "Token <input id=""token"" name=""token"" type=""text"" /><br />" & _
                                    "Service <input id=""service"" name=""service"" type=""text"" /><br />" & _
                                    "URL <input id=""url"" name=""url"" type=""text"" /><br />" & _
                                    "State <input id=""state"" name=""state"" type=""text"" /><br />" & _
                                    "<input id=""Submit1"" type=""submit"" value=""Ok""  onclick=""javascript: load_wait();"" /></form></body></html>")
        Else
            sToken = HttpUtility.UrlDecode(oContext.Request.Form("token"))
            sService = HttpUtility.UrlDecode(oContext.Request.Form("service"))
            sURL = HttpUtility.UrlDecode(oContext.Request.Form("url"))
            sState = HttpUtility.UrlDecode(oContext.Request.Form("state"))

            If sToken = "" OrElse sService = "" OrElse sState = "" Then
                oContext.Response.ContentType = "text/plain"
                oContext.Response.StatusCode = 432
                oContext.Response.Write("Missing parameter")
                ZSSOUtilities.WriteLog("Registration: Missing parameter")
                Return
            End If

            If sToken.Length <> 40 Then
                oContext.Response.ContentType = "text/plain"
                oContext.Response.StatusCode = 433
                oContext.Response.Write("Incorrect parameter")
                ZSSOUtilities.WriteLog("Registration: Incorrect parameter")
                Return
            End If

            Select Case sState
                Case "available", "allocated", "loading", "working"
                    Exit Select
                Case Else
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 433
                    oContext.Response.Write("Incorrect parameter")
                    ZSSOUtilities.WriteLog("Registration: Incorrect parameter")
                    Return
            End Select

            Try
                Using oConnexion As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
                    oConnexion.Open()
                    Using oSqlCommande As New SqlCommand("register", oConnexion)
                        oSqlCommande.CommandType = CommandType.StoredProcedure
                        oSqlCommande.Parameters.AddWithValue("token", sToken)
                        oSqlCommande.Parameters.AddWithValue("service", sService)
                        oSqlCommande.Parameters.AddWithValue("url", sURL)
                        oSqlCommande.Parameters.AddWithValue("state", sState)
                        oSqlCommande.ExecuteNonQuery()
                    End Using
                End Using
            Catch ex As Exception
                ZSSOUtilities.WriteLog("Registration: " & ex.Message)
            End Try
        End If
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class