Imports System.Web
Imports System.Web.Services
Imports System.Text.RegularExpressions.Regex
Imports System.Web.Script.Serialization
Imports System.Data.SqlClient
Imports System.Net
Imports System.Runtime.Caching
Imports System.IO

Public Class listrendezvous
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim Serial As String = ""
        Dim Token As String = ""
        Dim cacheMemory As ObjectCache = MemoryCache.Default
        Dim ListRdv = New Dictionary(Of String, Double)
        Dim PrinterLocationData As Dictionary(Of String, String) = ZSSOUtilities.GetLocation(context.Request.UserHostAddress)

        If context.Request.HttpMethod = "GET" Then
            context.Response.ContentType = "text/html"
            context.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title></head><body><form  method=""post"" action=""/listrendezvous.ashx"" accept-charset=""utf-8"">Serial <input id=""serial"" name=""serial"" type=""text"" /><br /><input id=""Submit1"" type=""submit"" value=""Ok"" /></form></body></html>")
        Else
            Dim cachedRendezvousServers = TryCast(cacheMemory("rendezvous_servers"), Dictionary(Of String, Dictionary(Of String, String)))
            If IsNothing(cachedRendezvousServers) Then
                cachedRendezvousServers = New Dictionary(Of String, Dictionary(Of String, String))
            End If

            Serial = HttpUtility.UrlDecode(context.Request.Form("serial"))
            Dim ServerData = New Dictionary(Of String, String)

            ZSSOUtilities.WriteLog("ListRDV : " + ZSSOUtilities.oSerializer.Serialize(context.Request.Form))
            If String.IsNullOrEmpty(Serial) Then
                context.Response.StatusCode = 432
                context.Response.Write("Missing parameter")
                ZSSOUtilities.WriteLog("ListRDV : Missing parameter")
                Return
            End If

            Dim pair As KeyValuePair(Of String, Dictionary(Of String, String))
            For Each pair In cachedRendezvousServers
                ListRdv(pair.Value("hostname") + ".zeepro.com") = ZSSOUtilities.CalculateDistanceBetweenCoordinates(pair.Value("latitude"), pair.Value("longitude"), PrinterLocationData("latitude"), PrinterLocationData("longitude"))
            Next

            Dim sorted = From RdvServer In ListRdv
                         Order By RdvServer.Value
            Dim SortedListRdv = sorted.ToDictionary(Function(p) p.Key, Function(p) p.Value)

            context.Response.ContentType = "text/plain"
            context.Response.Write(ZSSOUtilities.oSerializer.Serialize(SortedListRdv.Keys))
            ZSSOUtilities.WriteLog("ListRDV : OK : " + ZSSOUtilities.oSerializer.Serialize(SortedListRdv.Keys))
        End If

    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property
End Class