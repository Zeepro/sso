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

    Sub ProcessRequest(ByVal oContext As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim sSerial As String = ""
        Dim sToken As String = ""
        Dim oHttpCache As Caching.Cache = HttpRuntime.Cache
        Dim arPrinterLocationData As Dictionary(Of String, String) = ZSSOUtilities.GetLocation(oContext.Request.UserHostAddress)

        If IsNothing(arPrinterLocationData) Then
            arPrinterLocationData = New Dictionary(Of String, String)
            arPrinterLocationData("latitude") = "0"
            arPrinterLocationData("longitude") = "0"
        End If

        If oContext.Request.HttpMethod = "GET" Then
            oContext.Response.ContentType = "text/html"
            oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title>" & _
                                    "<script src=""https://code.jquery.com/jquery-1.10.2.js""></script><script type=""text/javascript"">function load_wait() { $(""#overlay"").addClass(""gray-overlay""); $("".ui-loader"").css(""display"", ""block""); }</script>" & _
                                    "<link rel=""stylesheet"" type=""text/css"" href=""style.css"">" & _
                                    "</head><body><div id=""overlay""></div><div class=""ui-loader ui-corner-all ui-body-a ui-loader-default""><span class=""ui-icon-loading""></span><h1>loading</h1></div><form  method=""post"" action=""/listrendezvous.ashx"" accept-charset=""utf-8"">Serial <input id=""serial"" name=""serial"" type=""text"" /><br /><input id=""Submit1"" type=""submit"" value=""Ok""  onclick=""javascript: load_wait();"" /></form></body></html>")
        Else
            'Dim arPrinterLocationData As Dictionary(Of String, String) = ZSSOUtilities.GetLocation("199.188.203.167")
            'If IsNothing(arPrinterLocationData) Then
            '    arPrinterLocationData = New Dictionary(Of String, String)
            '    arPrinterLocationData("latitude") = "0"
            '    arPrinterLocationData("longitude") = "0"
            'End If

            sSerial = HttpUtility.UrlDecode(oContext.Request.Form("serial"))

            '' Code a supprimer lors de la mise en prod (champs a aussi supprimer dans l'html)
            'Dim sIp As String = HttpUtility.UrlDecode(oContext.Request.Form("ip"))
            'arPrinterLocationData = ZSSOUtilities.GetLocation(sIp)
            '' Fin du code a supprimer

            ZSSOUtilities.WriteLog("ListRDV : " & ZSSOUtilities.oSerializer.Serialize({sSerial}))
            If String.IsNullOrEmpty(sSerial) Then
                oContext.Response.StatusCode = 432
                oContext.Response.Write("Missing parameter")
                ZSSOUtilities.WriteLog("ListRDV : Missing parameter")
                Return
            End If

            If ZSSOUtilities.SearchSerial(sSerial) = False Then
                oContext.Response.ContentType = "text/plain"
                oContext.Response.StatusCode = 436
                oContext.Response.Write("Unknown printer")
                ZSSOUtilities.WriteLog("ListRDV : Unknown printer")
                Return
            End If

            Dim arListRdv = New Dictionary(Of String, Double)
            Dim oCacheEnum As IDictionaryEnumerator = oHttpCache.GetEnumerator()
            While oCacheEnum.MoveNext()
                If oCacheEnum.Current.Key.ToString.StartsWith("rendezvous_server_") Then
                    arListRdv(oCacheEnum.Current.Value("hostname")) = ZSSOUtilities.CalculateDistanceBetweenCoordinates(oCacheEnum.Current.Value("latitude"), oCacheEnum.Current.Value("longitude"), arPrinterLocationData("latitude"), arPrinterLocationData("longitude"))
                End If
            End While

            Dim lnkSorted = From RdvServer In arListRdv
                         Order By RdvServer.Value
            Dim arSortedListRdv = lnkSorted.ToDictionary(Function(p) p.Key, Function(p) p.Value)

            oContext.Response.ContentType = "text/plain"
            oContext.Response.Write(ZSSOUtilities.oSerializer.Serialize(arSortedListRdv.Keys))
            ZSSOUtilities.WriteLog("ListRDV : OK : " + ZSSOUtilities.oSerializer.Serialize(arSortedListRdv.Keys))
        End If

    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property
End Class