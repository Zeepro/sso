Imports System.Web
Imports System.Web.Services
Imports System.Data.SqlClient
Imports System.Runtime.Caching

Public Class url
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal oContext As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim sSerial As String
        Dim sURL As String
        Dim oHttpCache As Caching.Cache = HttpRuntime.Cache

        If ZSSOUtilities.CheckRequests(oContext.Request.UserHostAddress, "url") > 5 Then
            oContext.Response.ContentType = "text/plain"
            oContext.Response.StatusCode = 435
            oContext.Response.Write("Too many requests")
            ZSSOUtilities.WriteLog("Url: Too many requests")
            Return
        Else
            If oContext.Request.HttpMethod = "GET" Then
                oContext.Response.ContentType = "text/html"
                oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title>" & _
                                        "<script src=""https://code.jquery.com/jquery-1.10.2.js""></script><script type=""text/javascript"">function load_wait() { $(""#overlay"").addClass(""gray-overlay""); $("".ui-loader"").css(""display"", ""block""); }</script>" & _
                                        "<link rel=""stylesheet"" type=""text/css"" href=""style.css"">" & _
                                        "</head><body><div id=""overlay""></div><div class=""ui-loader ui-corner-all ui-body-a ui-loader-default""><span class=""ui-icon-loading""></span><h1>loading</h1></div><form  method=""post"" action=""/url.ashx"" accept-charset=""utf-8"">serial <input id=""printersn"" name=""printersn"" type=""text"" /><br />URL <input id=""URL"" name=""URL"" type=""text"" /><br /><input id=""Submit1"" type=""submit"" value=""Ok""  onclick=""javascript: load_wait();"" /></form></body></html>")
            Else
                sSerial = HttpUtility.UrlDecode(oContext.Request.Form("printersn"))
                sURL = HttpUtility.UrlDecode(oContext.Request.Form("URL"))

                ZSSOUtilities.WriteLog("Url: " & ZSSOUtilities.oSerializer.Serialize({sSerial, sURL}))
                If String.IsNullOrEmpty(sSerial) Or String.IsNullOrEmpty(sURL) Then
                    oContext.Response.StatusCode = 432
                    oContext.Response.Write("Missing parameter")
                    ZSSOUtilities.WriteLog("Url: Missing parameter")
                    Return
                End If

                If ZSSOUtilities.SearchSerial(sSerial) = False Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 436
                    oContext.Response.Write("Unknown printer")
                    ZSSOUtilities.WriteLog("Url: Unknown printer")
                    Return
                End If

                oHttpCache.Insert(sSerial, sURL, Nothing, DateTime.Now.AddMinutes(20), TimeSpan.Zero)

                ZSSOUtilities.WriteLog("Url: OK")
            End If
        End If
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class