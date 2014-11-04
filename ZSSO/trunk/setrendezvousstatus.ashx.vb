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
        Dim sFreeMemory As String = ""
        Dim oHttpCache As Caching.Cache = HttpRuntime.Cache

        If oContext.Request.HttpMethod = "GET" Then
            oContext.Response.ContentType = "text/html"
            oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title>" & _
                                    "<script src=""https://code.jquery.com/jquery-1.10.2.js""></script><script type=""text/javascript"">function load_wait() { $(""#overlay"").addClass(""gray-overlay""); $("".ui-loader"").css(""display"", ""block""); }</script>" & _
                                    "<link rel=""stylesheet"" type=""text/css"" href=""style.css"">" & _
                                    "</head><body><div id=""overlay""></div><div class=""ui-loader ui-corner-all ui-body-a ui-loader-default""><span class=""ui-icon-loading""></span><h1>loading</h1></div><form  method=""post"" action=""/setrendezvousstatus.ashx"" accept-charset=""utf-8"">Bandwith <input id=""bandwidth"" name=""bandwidth"" type=""text"" /><br />Percentage <input id=""percentage"" name=""percentage"" type=""text"" /><br />FreeMemory <input id=""freememory"" name=""freememory"" type=""text"" /><br /><input id=""Submit1"" type=""submit"" value=""Ok""  onclick=""javascript: load_wait();"" /></form></body></html>")
        Else
            sBandwidth = HttpUtility.UrlDecode(oContext.Request.Form("bandwidth"))
            sPercentage = HttpUtility.UrlDecode(oContext.Request.Form("percentage"))
            sFreeMemory = HttpUtility.UrlDecode(oContext.Request.Form("freememory"))
            Try
                ZSSOUtilities.WriteLog("SetRDVStatus : " & ZSSOUtilities.oSerializer.Serialize({"from : " & Dns.GetHostEntry(oContext.Request.UserHostAddress).HostName, sBandwidth, sPercentage}))
            Catch ex As Exception
                ZSSOUtilities.WriteLog("SetRDVStatus : " & ZSSOUtilities.oSerializer.Serialize({"from : " & Dns.GetHostEntry(oContext.Request.UserHostAddress).HostName, sBandwidth, sPercentage}))
            End Try

            If String.IsNullOrEmpty(sBandwidth) Or String.IsNullOrEmpty(sPercentage) Then
                oContext.Response.StatusCode = 432
                oContext.Response.Write("Missing parameter")
                ZSSOUtilities.WriteLog("SetRDVStatus : Missing parameter")
                Return
            End If

            'Dim arLocationData As Dictionary(Of String, String) = ZSSOUtilities.GetLocation(HttpUtility.UrlDecode(oContext.Request.Form("ip")))
            Dim arLocationData As Dictionary(Of String, String) = ZSSOUtilities.GetLocation(oContext.Request.UserHostAddress)
            Dim arServerData = New Dictionary(Of String, String)
            arServerData("bandwidth") = sBandwidth
            arServerData("percentage") = sPercentage
            arServerData("freememory") = sFreeMemory

            If IsNothing(arLocationData) Then
                arLocationData = New Dictionary(Of String, String)
                arLocationData("latitude") = "0"
                arLocationData("longitude") = "0"
            End If
            arServerData("latitude") = arLocationData("latitude")
            arServerData("longitude") = arLocationData("longitude")

            Try
                arServerData("hostname") = Dns.GetHostEntry(oContext.Request.UserHostAddress).HostName
            Catch ex As Exception
                arServerData("hostname") = Dns.GetHostEntry(oContext.Request.UserHostAddress).HostName
            End Try
            'arServerData("hostname") = HttpUtility.UrlDecode(oContext.Request.Form("hostname"))

                oHttpCache.Insert("rendezvous_server_" & oContext.Request.UserHostAddress, arServerData, Nothing, DateTime.Now.AddMinutes(20.0), TimeSpan.Zero)
                'oHttpCache.Insert("rendezvous_server_" & HttpUtility.UrlDecode(oContext.Request.Form("ip")), arServerData, Nothing, DateTime.Now.AddMinutes(20.0), TimeSpan.Zero)

            ' add data into DB
            Using oConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
                oConnection.Open()

                Dim sQuery = "DELETE FROM RdvStats WHERE ServerTime <= @threemonth"

                Using oSqlCmdDelete As New SqlCommand(sQuery, oConnection)
                    oSqlCmdDelete.Parameters.AddWithValue("@threemonth", Date.Today.AddMonths(-3))

                    oSqlCmdDelete.ExecuteNonQuery()
                End Using

                sQuery = "INSERT INTO [RdvStats] ([Hostname], [Bandwidth], [Percentage], [Freememory]) VALUES (@hostname, @bandwidth, @percentage, @freememory)"

                Using oSqlCmdUpdate As New SqlCommand(sQuery, oConnection)
                    oSqlCmdUpdate.Parameters.AddWithValue("@hostname", arServerData("hostname"))
                    oSqlCmdUpdate.Parameters.AddWithValue("@bandwidth", Convert.ToInt32(sBandwidth))
                    oSqlCmdUpdate.Parameters.AddWithValue("@percentage", Convert.ToInt32(sPercentage))
                    oSqlCmdUpdate.Parameters.AddWithValue("@freememory", Convert.ToInt32(sFreeMemory))

                    oSqlCmdUpdate.ExecuteNonQuery()

                End Using

            End Using

            ZSSOUtilities.WriteLog("SetRDVStatus : OK : " & ZSSOUtilities.oSerializer.Serialize(arServerData))
        End If
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property
End Class