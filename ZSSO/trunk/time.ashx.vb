Imports System.Web
Imports System.Web.Services
Imports System.Text.RegularExpressions.Regex
Imports System.Web.Script.Serialization
Imports System.Data.SqlClient
Imports System.Runtime.Caching
Imports System.IO

Public Class time
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal oContext As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim sSerial As String
        Dim oHttpCache As Caching.Cache = HttpRuntime.Cache

        If oContext.Request.HttpMethod = "GET" Then
            oContext.Response.ContentType = "text/html"
            oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title></head><body><form  method=""post"" action=""/time.ashx"" accept-charset=""utf-8"">serial <input id=""printersn"" name=""printersn"" type=""text"" /><br /><input id=""Submit1"" type=""submit"" value=""Ok"" /></form></body></html>")
        Else
            Try
                Dim iCachedCounterByIp = CInt(oHttpCache("time_" & oContext.Request.UserHostAddress))
            Catch
                oContext.Response.ContentType = "text/plain"
                oContext.Response.StatusCode = 435
                oContext.Response.Write("Too many requests")
                ZSSOUtilities.WriteLog("Time : Too many requests")
                Return
            End Try

            oHttpCache.Insert("time_" & oContext.Request.UserHostAddress, 1, Nothing, DateTime.Now.AddMinutes(1.0), TimeSpan.Zero)

            sSerial = HttpUtility.UrlDecode(oContext.Request.Form("printersn"))

            ZSSOUtilities.WriteLog("Time : " & ZSSOUtilities.oSerializer.Serialize({sSerial}))

            If String.IsNullOrEmpty(sSerial) Then
                oContext.Response.ContentType = "text/plain"
                oContext.Response.StatusCode = 432
                oContext.Response.Write("Missing parameter")
                ZSSOUtilities.WriteLog("Time : Missing parameter")
                Return
            End If

            If Not ZSSOUtilities.SearchSerial(sSerial) Then
                oContext.Response.ContentType = "text/plain"
                oContext.Response.StatusCode = 433
                oContext.Response.Write("Incorrect Parameter")
                ZSSOUtilities.WriteLog("Time : Incorrect parameter")
                Return
            End If

            oContext.Response.ContentType = "text/plain"
            oContext.Response.Write(DateTime.Now.ToString("s", System.Globalization.CultureInfo.InvariantCulture))
            ZSSOUtilities.WriteLog("Time : OK")
        End If
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class