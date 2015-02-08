'TODO: manage geolocal

Imports System.Web
Imports System.Web.Services
Imports System.Data.SqlClient

Public Class request
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal oContext As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim sSerial As String
        Dim sService As String
        Dim sTicket As String
        Static oLock As New Object

        If oContext.Request.HttpMethod = "GET" Then
            oContext.Response.ContentType = "text/html"
            oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title>" & _
                                    "<script src=""https://code.jquery.com/jquery-1.10.2.js""></script><script type=""text/javascript"">function load_wait() { $(""#overlay"").addClass(""gray-overlay""); $("".ui-loader"").css(""display"", ""block""); }</script>" & _
                                    "<link rel=""stylesheet"" type=""text/css"" href=""style.css"">" & _
                                    "</head><body><div id=""overlay""></div><div class=""ui-loader ui-corner-all ui-body-a ui-loader-default""><span class=""ui-icon-loading""></span><h1>Request</h1></div>" & _
                                    "<form  method=""post"" action=""request.ashx"" accept-charset=""utf-8"">" & _
                                    "Serial <input id=""serial"" name=""serial"" type=""text"" /><br />" & _
                                    "Service <input id=""service"" name=""service"" type=""text"" /><br />" & _
                                    "Ticket <input id=""ticket"" name=""ticket"" type=""text"" /><br />" & _
                                    "<input id=""Submit1"" type=""submit"" value=""Ok""  onclick=""javascript: load_wait();"" /></form></body></html>")
        Else
            sSerial = oContext.Request.Form("serial")
            sService = oContext.Request.Form("service")
            sTicket = oContext.Request.Form("ticket")

            If sSerial = "" OrElse sService = "" Then
                oContext.Response.ContentType = "text/plain"
                oContext.Response.StatusCode = 432
                oContext.Response.Write("Missing parameter")
                ZSSOUtilities.WriteLog("Registration: Missing parameter")
                Return
            End If

            Try
                Using oConnexion As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
                    oConnexion.Open()

                    Using oSqlCmdSelect As New SqlCommand("SELECT TOP 1 Serial FROM Printer WHERE Serial = @serial", oConnexion)
                        oSqlCmdSelect.Parameters.AddWithValue("@serial", sSerial.ToUpper)
                        If oSqlCmdSelect.ExecuteScalar() Is Nothing Then
                            oContext.Response.ContentType = "text/plain"
                            oContext.Response.StatusCode = 436
                            oContext.Response.Write("Unknown printer")
                            ZSSOUtilities.WriteLog("Registration: Missing parameter")
                            Return
                        End If
                    End Using

                    Using oSqlCmdSelect As New SqlCommand("SELECT TOP 1 service FROM Service WHERE service = @service", oConnexion)
                        oSqlCmdSelect.Parameters.AddWithValue("@service", sService)
                        If oSqlCmdSelect.ExecuteScalar() Is Nothing Then
                            oContext.Response.ContentType = "text/plain"
                            oContext.Response.StatusCode = 441
                            oContext.Response.Write("Unknown service")
                            ZSSOUtilities.WriteLog("Registration: Service unavailable")
                            Return
                        End If
                    End Using

                    SyncLock oLock
                        Try
                            Dim oResponse As New Dictionary(Of String, String)

                            Using oSqlCmd As New SqlCommand()
                                oSqlCmd.Connection = oConnexion
                                oSqlCmd.Parameters.AddWithValue("@serial", sSerial)
                                oSqlCmd.Parameters.AddWithValue("@service", sService)
                                oSqlCmd.Parameters.AddWithValue("@state", "waiting")
                                If sTicket = "" Then
                                    sTicket = System.Web.Security.Membership.GeneratePassword(40, 0)
                                    oSqlCmd.Parameters.AddWithValue("@ticket", sTicket)
                                    oSqlCmd.CommandText = "DELETE Ticket WHERE [update] < DATEADD(hour, -1, GETDATE());DELETE Ticket WHERE [update] < DATEADD(second, -10, GETDATE()) AND state = 'waiting';INSERT Ticket (ticket, service, state) VALUES (@ticket, @service, @state)"
                                    oSqlCmd.ExecuteNonQuery()
                                Else
                                    oSqlCmd.CommandText = "DELETE Ticket WHERE [update] < DATEADD(hour, -1, GETDATE());DELETE Ticket WHERE [update] < DATEADD(second, -10, GETDATE()) AND state = 'waiting';UPDATE Ticket SET [update] = GETDATE(), state = 'waiting' WHERE ticket = @ticket;SELECT @@ROWCOUNT"
                                    oSqlCmd.Parameters.AddWithValue("@ticket", sTicket)
                                    If oSqlCmd.ExecuteScalar() <> 1 Then
                                        ' Unknown ticket
                                        sTicket = System.Web.Security.Membership.GeneratePassword(40, 0)
                                        oSqlCmd.Parameters("@ticket").Value = sTicket
                                        oSqlCmd.CommandText = "INSERT Ticket (ticket, service, state) VALUES (@ticket, @service, @state)"
                                        oSqlCmd.ExecuteNonQuery()
                                    End If
                                End If
                                oSqlCmd.CommandText = "SELECT TOP (1) 1 FROM PrinterService WHERE serial = @serial"
                                If Not oSqlCmd.ExecuteScalar() Is Nothing Then
                                    ' Association rule
                                    oSqlCmd.CommandText = "SELECT TOP 1 ticket " & _
                                        "FROM Service INNER JOIN Ticket " & _
                                        "ON Service.service = Ticket.service " & _
                                        "WHERE Service.service = @service AND Service.state = 'available' AND EXISTS (SELECT 1 FROM PrinterService WHERE Service.url = PrinterService.url) AND Ticket.state = 'waiting' " & _
                                        "ORDER BY creation"
                                    If oSqlCmd.ExecuteScalar() = sTicket Then
                                        oSqlCmd.CommandText = "UPDATE Ticket SET state = 'processed' WHERE ticket = @ticket;" & _
                                            "UPDATE TOP (1) Service SET state = 'allocated', date = GETDATE() OUTPUT INSERTED.url, INSERTED.token WHERE service = @service AND Service.state = 'available' AND EXISTS (SELECT 1 FROM PrinterService WHERE Service.url = PrinterService.url)"
                                        Using oQueryResult As SqlDataReader = oSqlCmd.ExecuteReader()
                                            If oQueryResult.Read() Then
                                                oResponse.Add("URL", oQueryResult("url"))
                                                oResponse.Add("token", oQueryResult("token"))
                                            End If
                                        End Using
                                    End If
                                Else
                                    ' Common rule
                                    oSqlCmd.CommandText = "SELECT TOP 1 ticket " & _
                                        "FROM Service INNER JOIN Ticket " & _
                                        "ON Service.service = Ticket.service " & _
                                        "WHERE Service.service = @service AND Service.state = 'available' AND NOT EXISTS (SELECT 1 FROM PrinterService WHERE Service.url = PrinterService.url) AND Ticket.state = 'waiting' " & _
                                        "ORDER BY creation"
                                    If oSqlCmd.ExecuteScalar() = sTicket Then
                                        oSqlCmd.CommandText = "UPDATE Ticket SET state = 'processed' WHERE ticket = @ticket;" & _
                                            "UPDATE TOP (1) Service SET state = 'allocated', date = GETDATE() OUTPUT INSERTED.url, INSERTED.token WHERE service = @service AND Service.state = 'available' AND NOT EXISTS (SELECT 1 FROM PrinterService WHERE Service.url = PrinterService.url)"
                                        Using oQueryResult As SqlDataReader = oSqlCmd.ExecuteReader()
                                            If oQueryResult.Read() Then
                                                oResponse.Add("URL", oQueryResult("url"))
                                                oResponse.Add("token", oQueryResult("token"))
                                            End If
                                        End Using
                                    End If
                                End If
                                oContext.Response.ContentType = "text/plain"
                                oResponse.Add("ticket", sTicket)
                                oContext.Response.Write(ZSSOUtilities.oSerializer.Serialize(oResponse))
                            End Using
                        Catch ex As Exception
                            ZSSOUtilities.WriteLog("Registration: " & ex.Message)
                        End Try
                    End SyncLock
                End Using

            Catch ex As Exception
                ZSSOUtilities.WriteLog("Registration: " & ex.Message)
            End Try
        End If
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class