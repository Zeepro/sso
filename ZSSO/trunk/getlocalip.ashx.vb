Imports System.Web
Imports System.Web.Services
Imports System.Text.RegularExpressions.Regex
Imports System.Data.SqlClient
Imports System.Security.Cryptography
Imports System.Runtime.Caching
Imports System.IO

Public Class getlocalip
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal oContext As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim sSerial As String
        Dim oHttpCache As Caching.Cache = HttpRuntime.Cache

        If oContext.Request.HttpMethod = "GET" Then
            oContext.Response.ContentType = "text/html"
            oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title>" & _
                                    "<script src=""https://code.jquery.com/jquery-1.10.2.js""></script><script type=""text/javascript"">function load_wait() { $(""#overlay"").addClass(""gray-overlay""); $("".ui-loader"").css(""display"", ""block""); }</script>" & _
                                    "<link rel=""stylesheet"" type=""text/css"" href=""style.css"">" & _
                                    "</head><body><div id=""overlay""></div><div class=""ui-loader ui-corner-all ui-body-a ui-loader-default""><span class=""ui-icon-loading""></span><h1>loading</h1></div><form  method=""post"" action=""/getlocalip.ashx"" accept-charset=""utf-8"">Serial <input id=""printersn"" name=""printersn"" type=""text"" /><br /><input id=""Submit1"" type=""submit"" value=""Ok""  onclick=""javascript: load_wait();"" /></form></body></html>")
        Else
            sSerial = HttpUtility.UrlDecode(oContext.Request.Form("printersn"))

            ZSSOUtilities.WriteLog("GetLocalIp : " & ZSSOUtilities.oSerializer.Serialize({sSerial}))
            If String.IsNullOrEmpty(sSerial) Then
                oContext.Response.StatusCode = 432
                oContext.Response.Write("Missing parameter")
                ZSSOUtilities.WriteLog("GetLocalIp : Missing parameter")
                Return
            End If

            If sSerial.Length <> 12 Then
                oContext.Response.ContentType = "text/plain"
                oContext.Response.StatusCode = 433
                oContext.Response.Write("Incorrect Parameter")
                ZSSOUtilities.WriteLog("GetLocalIp : Incorrect parameter")
                Return
            End If

            Dim arReturn As Dictionary(Of String, String) = New Dictionary(Of String, String)

            'Dim arCachedPrinter = TryCast(oHttpCache("printer_" & sSerial.ToUpper), Dictionary(Of String, String))
            'If Not IsNothing(arCachedPrinter) Then
            '    arReturn("localIP") = arCachedPrinter("local_ip")
            '    arReturn("state") = "ok"
            'Else
            '    arReturn("localIP") = ""
            '    arReturn("state") = "unknown"
            'End If

            Using oConnexion As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
                oConnexion.Open()
                Using oSqlCmdSelect As New SqlCommand("DELETE ActivePrinter WHERE date < DATEADD(minute, -20, GETDATE());" & _
                                                      "SELECT local_ip FROM ActivePrinter WHERE serial = @serial", oConnexion)
                    oSqlCmdSelect.Parameters.AddWithValue("@serial", sSerial.ToUpper)
                    Dim oTmp As Object = oSqlCmdSelect.ExecuteScalar()
                    If oSqlCmdSelect.ExecuteScalar() Is Nothing Then
                        arReturn("localIP") = ""
                        arReturn("state") = "unknown"
                    Else
                        arReturn("localIP") = DirectCast(oTmp, String)
                        arReturn("state") = "ok"
                    End If
                End Using
            End Using


            oContext.Response.AddHeader("Access-Control-Allow-Origin", "*")
            oContext.Response.ContentType = "application/json"
            oContext.Response.Write(ZSSOUtilities.oSerializer.Serialize(arReturn))
            ZSSOUtilities.WriteLog("GetLocalIp : OK : " + ZSSOUtilities.oSerializer.Serialize(arReturn))
        End If
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class