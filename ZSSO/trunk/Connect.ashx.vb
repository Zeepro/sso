Imports System.Web
Imports System.Web.Services
Imports System.Data.SqlClient

Public Class Connect
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal oContext As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim sToken, sSerial, sCalibrationWarningMessage, sAccountEmail As String

        If ZSSOUtilities.CheckRequests(oContext.Request.UserHostAddress, "connect") > 5 Then
            oContext.Response.ContentType = "text/plain"
            oContext.Response.StatusCode = 435
            oContext.Response.Write("Too many requests")
            ZSSOUtilities.WriteLog("Connect: Too many requests")
            Return
        Else
            If oContext.Request.HttpMethod = "GET" Then
                oContext.Response.ContentType = "text/html"
                oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title>" & _
                                        "<script src=""https://code.jquery.com/jquery-1.10.2.js""></script><script type=""text/javascript"">function load_wait() { $(""#overlay"").addClass(""gray-overlay""); $("".ui-loader"").css(""display"", ""block""); }</script>" & _
                                        "<link rel=""stylesheet"" type=""text/css"" href=""style.css"">" & _
                                        "</head><body><div id=""overlay""></div><div class=""ui-loader ui-corner-all ui-body-a ui-loader-default""><span class=""ui-icon-loading""></span><h1>connect</h1></div>" & _
                                        "<form  method=""post"" action=""/connect.ashx"" accept-charset=""utf-8"">" & _
                                        "token <input id=""token"" name=""token"" type=""text"" /><br />" & _
                                        "serial <input id=""printersn"" name=""printersn"" type=""text"" /><br />" & _
                                        "calibrationwarningmessage <input id=""calibrationwarningmessage"" name=""calibrationwarningmessage"" type=""text"" /><br />" & _
                                        "<input id=""Submit1"" type=""submit"" value=""Ok""  onclick=""javascript: load_wait();"" /></form></body></html>")
            Else
                sToken = oContext.Request.Form("token")
                sSerial = oContext.Request.Form("printersn")
                sCalibrationWarningMessage = oContext.Request.Form("calibrationwarningmessage")

                If String.IsNullOrEmpty(sToken) OrElse _
                    String.IsNullOrEmpty(sSerial) OrElse _
                    String.IsNullOrEmpty(sCalibrationWarningMessage) Then
                    oContext.Response.StatusCode = 432
                    oContext.Response.Write("Missing parameter")
                    ZSSOUtilities.WriteLog("Connect: Missing parameter")
                    Return
                End If

                If sToken.Length <> 40 OrElse _
                    sSerial.Length <> 12 OrElse _
                    (sCalibrationWarningMessage <> "yes" AndAlso sCalibrationWarningMessage <> "no") Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 433
                    oContext.Response.Write("Incorrect Parameter")
                    ZSSOUtilities.WriteLog("Connect: Incorrect parameter")
                    Return
                End If

                sAccountEmail = ZSSOUtilities.SearchAccountEmail(sToken)

                If sAccountEmail Is Nothing Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 442
                    oContext.Response.Write("Unauthorized user")
                    ZSSOUtilities.WriteLog("Connect: Unauthorized user")
                    Return
                End If

                Using oConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
                    oConnection.Open()

                    Dim oSqlCmd As New SqlCommand("UPDATE AccountPrinterAssociation " & _
                            "SET lastaccess = GETDATE(), " & _
                            "calibrationwarningmessage = @calibrationwarningmessage " & _
                            "WHERE serial = @serial AND deleted IS NULL", _
                            oConnection)

                    Try
                        If sCalibrationWarningMessage = "yes" Then
                            oSqlCmd.Parameters.AddWithValue("@calibrationwarningmessage", 1)
                        Else
                            oSqlCmd.Parameters.AddWithValue("@calibrationwarningmessage", 0)
                        End If
                        oSqlCmd.Parameters.AddWithValue("@serial", sSerial)
                        oSqlCmd.ExecuteNonQuery()
                    Catch ex As Exception
                        ZSSOUtilities.WriteLog("Connect: " & ex.Message)
                    End Try
                End Using
                ZSSOUtilities.WriteLog("Connect: OK")
            End If
        End If
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property
End Class