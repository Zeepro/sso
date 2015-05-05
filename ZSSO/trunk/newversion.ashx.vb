Imports System.Web
Imports System.Web.Services
Imports System.Data.SqlClient

Public Class newversion
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal oContext As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim sSerial As String = ""
        Dim sNextVersion As String

        If oContext.Request.HttpMethod = "GET" Then
            oContext.Response.ContentType = "text/html"
            oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title>" & _
                                    "<script src=""https://code.jquery.com/jquery-1.10.2.js""></script><script type=""text/javascript"">function load_wait() { $(""#overlay"").addClass(""gray-overlay""); $("".ui-loader"").css(""display"", ""block""); }</script>" & _
                                    "<link rel=""stylesheet"" type=""text/css"" href=""style.css"">" & _
                                    "</head><body><div id=""overlay""></div><div class=""ui-loader ui-corner-all ui-body-a ui-loader-default""><span class=""ui-icon-loading""></span><h1>loading</h1></div>" & _
                                    "<form  method=""post"" action=""/newversion.ashx"" accept-charset=""utf-8"">" & _
                                    "Serial <input id=""serial"" name=""serial"" type=""text"" /><br />" & _
                                    "Next software version <input id=""nextversion"" name=""nextversion"" type=""text"" /><br />" & _
                                    "<input id=""Submit1"" type=""submit"" value=""Ok""  onclick=""javascript: load_wait();"" /></form></body></html>")
        Else
            sSerial = oContext.Request.Form("serial")
            sNextVersion = oContext.Request.Form("nextversion")


            If String.IsNullOrEmpty(sSerial) OrElse String.IsNullOrEmpty(sNextVersion) Then
                oContext.Response.StatusCode = 432
                oContext.Response.Write("Missing parameter")
                ZSSOUtilities.WriteLog("NewVersion: Missing parameter")
                Return
            End If

            If ZSSOUtilities.SearchSerial(sSerial) = False Then
                oContext.Response.ContentType = "text/plain"
                oContext.Response.StatusCode = 436
                oContext.Response.Write("Unknown printer")
                ZSSOUtilities.WriteLog("NewVersion: Unknown printer")
                Return
            End If

            Try
                Using oConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
                    oConnection.Open()
                    Using oSqlCmdSelect As New SqlCommand("UPDATE Printer SET next_software = @next_software WHERE serial = @serial", oConnection)
                        oSqlCmdSelect.Parameters.AddWithValue("@serial", sSerial)
                        oSqlCmdSelect.Parameters.AddWithValue("@next_software", sNextVersion)
                        oSqlCmdSelect.ExecuteNonQuery()
                    End Using
                End Using
            Catch ex As Exception
                ZSSOUtilities.WriteLog("NewVersion: " & ex.Message)
                Return
            End Try
            ZSSOUtilities.WriteLog("NewVersion: OK")
        End If
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class