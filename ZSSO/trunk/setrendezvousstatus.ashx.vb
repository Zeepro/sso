Imports System.Web
Imports System.Web.Services
Imports System.Text.RegularExpressions.Regex
Imports System.Web.Script.Serialization
Imports System.Data.SqlClient
Imports System.Net
Imports System.Runtime.Caching
Imports System.IO

Public Class setrendezvousstatus
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal oContext As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim sBandwidth As String = ""
        Dim sPercentage As String = ""
        Dim oHttpCache As Caching.Cache = HttpRuntime.Cache

        If oContext.Request.HttpMethod = "GET" Then
            oContext.Response.ContentType = "text/html"
            oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title></head><body><form  method=""post"" action=""/setrendezvousstatus.ashx"" accept-charset=""utf-8"">Bandwith <input id=""bandwidth"" name=""bandwidth"" type=""text"" /><br />Percentage <input id=""percentage"" name=""percentage"" type=""text"" /><br />Ip <input id=""ip"" name=""ip"" type=""text"" /><br />Hostname <input id=""hostname"" name=""hostname"" type=""text"" /><br /><input id=""Submit1"" type=""submit"" value=""Ok"" /></form></body></html>")
        Else
            sBandwidth = HttpUtility.UrlDecode(oContext.Request.Form("bandwidth"))
            sPercentage = HttpUtility.UrlDecode(oContext.Request.Form("percentage"))

            ZSSOUtilities.WriteLog("SetRDVStatus : " & ZSSOUtilities.oSerializer.Serialize({sBandwidth, sPercentage}))

            If String.IsNullOrEmpty(sBandwidth) Or String.IsNullOrEmpty(sPercentage) Then
                oContext.Response.StatusCode = 432
                oContext.Response.Write("Missing parameter")
                ZSSOUtilities.WriteLog("SetRDVStatus : Missing parameter")
                Return
            End If

            Dim arLocationData As Dictionary(Of String, String) = ZSSOUtilities.GetLocation(HttpUtility.UrlDecode(oContext.Request.Form("ip")))
            'Dim arLocationData As Dictionary(Of String, String) = ZSSOUtilities.GetLocation(oContext.Request.UserHostAddress)
            Dim arServerData = New Dictionary(Of String, String)
            arServerData("bandwidth") = sBandwidth
            arServerData("percentage") = sPercentage
            arServerData("latitude") = arLocationData("latitude")
            arServerData("longitude") = arLocationData("longitude")
            ' arServerData("hostname") = Dns.GetHostEntry(oContext.Request.UserHostAddress).HostName
            arServerData("hostname") = HttpUtility.UrlDecode(oContext.Request.Form("hostname"))

            'oHttpCache.Insert("rendezvous_server_" & oContext.Request.UserHostAddress, arServerData, Nothing, DateTime.Now.AddMinutes(20.0), TimeSpan.Zero)
            oHttpCache.Insert("rendezvous_server_" & HttpUtility.UrlDecode(oContext.Request.Form("ip")), arServerData, Nothing, DateTime.Now.AddMinutes(20.0), TimeSpan.Zero)

        End If
        ZSSOUtilities.WriteLog("SetRDVStatus : OK")
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property
End Class