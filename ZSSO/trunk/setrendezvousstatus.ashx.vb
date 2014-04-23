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

    Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim Bandwidth As String = ""
        Dim Percentage As String = ""
        Dim Token As String = ""
        Dim cacheMemory As ObjectCache = MemoryCache.Default

        If context.Request.HttpMethod = "GET" Then
            context.Response.ContentType = "text/html"
            context.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title></head><body><form  method=""post"" action=""/setrendezvousstatus.ashx"" accept-charset=""utf-8"">Bandwith <input id=""bandwidth"" name=""bandwidth"" type=""text"" /><br />Percentage <input id=""percentage"" name=""percentage"" type=""text"" /><br /><input id=""Submit1"" type=""submit"" value=""Ok"" /></form></body></html>")
        Else
            Dim cachedRendezvousServers = TryCast(cacheMemory("rendezvous_servers"), Dictionary(Of String, Dictionary(Of String, String)))
            If IsNothing(cachedRendezvousServers) Then
                cachedRendezvousServers = New Dictionary(Of String, Dictionary(Of String, String))
            End If

            Bandwidth = HttpUtility.UrlDecode(context.Request.Form("bandwidth"))
            Percentage = HttpUtility.UrlDecode(context.Request.Form("percentage"))
            Dim ServerData = New Dictionary(Of String, String)

            ZSSOUtilities.WriteLog("SetRDVStatus : " + ZSSOUtilities.oSerializer.Serialize(context.Request.Form))

            If String.IsNullOrEmpty(Bandwidth) Or String.IsNullOrEmpty(Percentage) Then
                context.Response.StatusCode = 432
                context.Response.Write("Missing parameter")
                ZSSOUtilities.WriteLog("SetRDVStatus : Missing parameter")
                Return
            End If

            ServerData("bandwidth") = Bandwidth
            ServerData("percentage") = Percentage
            Dim LocationData As Dictionary(Of String, String) = ZSSOUtilities.GetLocation(context.Request.UserHostAddress)
            ServerData("latitude") = LocationData("latitude")
            ServerData("longitude") = LocationData("longitude")

            cachedRendezvousServers(context.Request.UserHostAddress) = ServerData
            cacheMemory.Set("rendezvous_servers", cachedRendezvousServers, DateTime.Now.AddMinutes(20.0), Nothing)

        End If
        ZSSOUtilities.WriteLog("SetRDVStatus : OK")
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property
End Class