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
        Static oLock As New Object  ' Application level semaphore

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
                            oContext.Response.Write("Service unavailable")
                            ZSSOUtilities.WriteLog("Registration: Service unavailable")
                            Return
                        End If
                    End Using

                    SyncLock oLock
                        Try
                            Dim oDataTable As DataTable
                            Dim oResponse As New Dictionary(Of String, String)
                            Dim oTmp As Object

                            Using oSqlCmd As New SqlCommand()
                                oSqlCmd.Connection = oConnexion
                                oSqlCmd.Parameters.AddWithValue("@serial", sSerial)
                                oSqlCmd.Parameters.AddWithValue("@service", sService)

                                ' Ticket management
                                If sTicket = "" Then
                                    ' Creation
                                    sTicket = System.Web.Security.Membership.GeneratePassword(40, 0)
                                    oSqlCmd.Parameters.AddWithValue("@ticket", sTicket)
                                    oSqlCmd.CommandText = "DELETE Ticket WHERE [update] < DATEADD(hour, -1, GETDATE());DELETE Ticket WHERE [update] < DATEADD(second, -10, GETDATE()) AND state = 'waiting';INSERT Ticket (ticket, service, state) VALUES (@ticket, @service, 'waiting')"
                                    oSqlCmd.ExecuteNonQuery()
                                Else
                                    ' Update / recreation
                                    oSqlCmd.CommandText = "DELETE Ticket WHERE [update] < DATEADD(hour, -1, GETDATE());DELETE Ticket WHERE [update] < DATEADD(second, -100, GETDATE()) AND state = 'waiting';UPDATE Ticket SET [update] = GETDATE(), state = 'waiting' OUTPUT inserted.creation WHERE ticket = @ticket"
                                    oSqlCmd.Parameters.AddWithValue("@ticket", sTicket)

                                    Using oQueryResult As SqlDataReader = oSqlCmd.ExecuteReader()
                                        If oQueryResult.HasRows Then
                                            oQueryResult.Read()
                                            If DateDiff(DateInterval.Minute, DirectCast(oQueryResult("creation"), DateTime), Now) > 2 Then
                                                ' Timeout
                                                oContext.Response.ContentType = "text/plain"
                                                oContext.Response.StatusCode = 441
                                                oContext.Response.Write("Service unavailable")
                                                ZSSOUtilities.WriteLog("Registration: Service unavailable")
                                                Return
                                            End If
                                        Else
                                            oQueryResult.Close()
                                            ' Unknown ticket
                                            sTicket = System.Web.Security.Membership.GeneratePassword(40, 0)
                                            oSqlCmd.Parameters("@ticket").Value = sTicket
                                            oSqlCmd.CommandText = "INSERT Ticket (ticket, service, state) VALUES (@ticket, @service, 'waiting')"
                                            oSqlCmd.ExecuteNonQuery()
                                        End If
                                    End Using
                                End If

                                oSqlCmd.CommandText = "SELECT queue FROM PrinterQueue WHERE serial = @serial"
                                oTmp = oSqlCmd.ExecuteScalar()
                                If oTmp Is Nothing Then
                                    ' Common queue

                                    ' Ticket selection
                                    oSqlCmd.CommandText = "DELETE Service WHERE date < DATEADD(minute, -2, GETDATE());" & _
                                        "SELECT TOP 1 ticket " & _
                                        "FROM Service " & _
                                        "INNER JOIN Ticket ON Service.service = Ticket.service " & _
                                        "WHERE Service.service = @service " & _
                                        "AND Service.state = 'available' " & _
                                        "AND NOT EXISTS (SELECT 1 FROM ServiceQueue WHERE Service.url = ServiceQueue.url) " & _
                                        "AND Ticket.state = 'waiting' " & _
                                        "ORDER BY creation"
                                    If oSqlCmd.ExecuteScalar() = sTicket Then
                                        ' Server list
                                        oSqlCmd.CommandText = "SELECT token, url, lat = ISNULL(latitude, 0), long = ISNULL(longitude, 0), distance = 0.0 " & _
                                            "FROM Service " & _
                                            "WHERE Service.service = @service " & _
                                            "AND Service.state = 'available' " & _
                                            "AND NOT EXISTS (SELECT 1 FROM ServiceQueue WHERE Service.url = ServiceQueue.url)"
                                        Using oQueryResult As SqlDataReader = oSqlCmd.ExecuteReader()
                                            oDataTable = New DataTable
                                            oDataTable.Load(oQueryResult)
                                            oDataTable.Columns("distance").ReadOnly = False
                                        End Using
                                    End If
                                Else
                                    ' Dedicated queue
                                    oSqlCmd.Parameters.AddWithValue("@queue", DirectCast(oTmp, Integer))

                                    ' Ticket selection
                                    oSqlCmd.CommandText = "DELETE Service WHERE date < DATEADD(minute, -2, GETDATE());" & _
                                        "SELECT TOP 1 ticket " & _
                                        "FROM Service " & _
                                        "INNER JOIN Ticket ON Service.service = Ticket.service " & _
                                        "INNER JOIN ServiceQueue ON Servicequeue.url = Service.url " & _
                                        "WHERE Service.service = @service " & _
                                        "AND ServiceQueue.queue = @queue " & _
                                        "AND Service.state = 'available' " & _
                                        "AND Ticket.state = 'waiting' " & _
                                        "ORDER BY creation"
                                    If oSqlCmd.ExecuteScalar() = sTicket Then
                                        ' Server list
                                        oSqlCmd.CommandText = "SELECT token, Service.url, lat = ISNULL(latitude, 0), long = ISNULL(longitude, 0), distance = 0.0 " & _
                                            "FROM Service " & _
                                            "INNER JOIN ServiceQueue ON Servicequeue.url = Service.url " & _
                                            "WHERE Service.service = @service " & _
                                            "AND ServiceQueue.queue = @queue " & _
                                            "AND Service.state = 'available'"
                                        Using oQueryResult As SqlDataReader = oSqlCmd.ExecuteReader()
                                            oDataTable = New DataTable
                                            oDataTable.Load(oQueryResult)
                                            oDataTable.Columns("distance").ReadOnly = False
                                        End Using
                                    End If
                                End If

                                If Not oDataTable Is Nothing Then
                                    If oDataTable.Rows.Count > 0 Then
                                        ' Sort by distance
                                        If oDataTable.Rows.Count > 1 Then
                                            Dim arPrinterLocationData As Dictionary(Of String, String) = ZSSOUtilities.GetLocation(oContext.Request.UserHostAddress)
                                            Dim nLatitude As Double = CDbl(arPrinterLocationData("latitude"))
                                            Dim nLongitude As Double = CDbl(arPrinterLocationData("longitude"))

                                            For Each oRow As DataRow In oDataTable.Rows
                                                oRow("distance") = ZSSOUtilities.CalculateDistanceBetweenCoordinates(oRow("lat"), oRow("long"), nLatitude, nLongitude)
                                            Next
                                            Dim oDataView As New DataView(oDataTable)
                                            oDataView.Sort = "distance"
                                            oResponse.Add("URL", oDataView(0)("url"))
                                            oResponse.Add("token", oDataView(0)("token"))
                                        Else
                                            oResponse.Add("URL", oDataTable(0)("url"))
                                            oResponse.Add("token", oDataTable(0)("token"))
                                        End If
                                        ' Set service and ticket
                                        oSqlCmd.CommandText = "UPDATE Ticket SET state = 'processed' WHERE ticket = @ticket;" & _
                                            "UPDATE Service SET state = 'allocated', date = GETDATE() WHERE token = @token"
                                        oSqlCmd.Parameters.AddWithValue("@token", oResponse("token"))
                                        oSqlCmd.ExecuteNonQuery()
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