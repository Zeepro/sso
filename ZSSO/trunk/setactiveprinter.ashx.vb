Imports System.Web
Imports System.Web.Services
Imports System.Text.RegularExpressions.Regex
Imports System.Web.Script.Serialization
Imports System.Data.SqlClient
Imports System.Net
Imports System.Runtime.Caching
Imports System.IO

Public Class setactiveprinter
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal oContext As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim sSerial As String = ""
        Dim sIp As String = ""
        Dim sToken As String = ""
        Dim sPort As String = ""
        Dim bSerialFound As Boolean = False
        Dim oHttpCache As Caching.Cache = HttpRuntime.Cache

        If oContext.Request.HttpMethod = "GET" Then
            oContext.Response.ContentType = "text/html"
            oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title>" & _
                                    "<script src=""https://code.jquery.com/jquery-1.10.2.js""></script><script type=""text/javascript"">function load_wait() { $(""#overlay"").addClass(""gray-overlay""); $("".ui-loader"").css(""display"", ""block""); }</script>" & _
                                    "<link rel=""stylesheet"" type=""text/css"" href=""style.css"">" & _
                                    "</head><body><div id=""overlay""></div><div class=""ui-loader ui-corner-all ui-body-a ui-loader-default""><span class=""ui-icon-loading""></span><h1>loading</h1></div><form  method=""post"" action=""/setactiveprinter.ashx"" accept-charset=""utf-8"">Serial <input id=""printersn"" name=""printersn"" type=""text"" /><br />Local IP <input id=""localIPaddress"" name=""localIPaddress"" type=""text"" /><br />Token <input id=""token"" name=""token"" type=""text"" /><br />Port <input id=""port"" name=""port"" type=""text"" /><input id=""Submit1"" type=""submit"" value=""Ok""  onclick=""javascript: load_wait();"" /></form></body></html>")
        Else
            sSerial = HttpUtility.UrlDecode(oContext.Request.Form("printersn"))
            sIp = HttpUtility.UrlDecode(oContext.Request.Form("localIPaddress"))
            sToken = HttpUtility.UrlDecode(oContext.Request.Form("token"))
            sPort = HttpUtility.UrlDecode(oContext.Request.Form("port"))
            ZSSOUtilities.WriteLog("SetActivePrinter : " & ZSSOUtilities.oSerializer.Serialize({sSerial, sIp, sToken, sPort}))

            If String.IsNullOrEmpty(sSerial) Or String.IsNullOrEmpty(sIp) Or String.IsNullOrEmpty(sToken) Then
                oContext.Response.StatusCode = 432
                oContext.Response.Write("Missing parameter")
                ZSSOUtilities.WriteLog("SetActivePrinter : Missing parameter")
                Return
            End If

            If String.IsNullOrEmpty(sPort) Then
                sPort = 443
            End If

            Dim oIpa As IPAddress = Nothing
            Dim iPort As Integer
            If Not (IPAddress.TryParse(sIp, oIpa)) Or Not Integer.TryParse(sPort, iPort) Then
                oContext.Response.StatusCode = 433
                oContext.Response.Write("Incorrect Parameter")
                ZSSOUtilities.WriteLog("SetActivePrinter : Incorrect parameter")
                Return
            End If

            Using oConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
                oConnection.Open()

                Dim sQuery = "SELECT TOP 1 Serial " & _
                    "FROM Printer " & _
                    "WHERE Serial=@serial"

                Using oSqlCmdSelect As New SqlCommand(sQuery, oConnection)

                    oSqlCmdSelect.Parameters.AddWithValue("@serial", sSerial.ToUpper)

                    Try
                        Using oQueryResult As SqlDataReader = oSqlCmdSelect.ExecuteReader()
                            If oQueryResult.HasRows Then
                                bSerialFound = True
                            End If
                        End Using
                    Catch ex As Exception
                        ZSSOUtilities.WriteLog("SetActivePrinter : NOK : " & ex.Message)
                    End Try
                End Using

            End Using
            Dim arSerialData = New Dictionary(Of String, String)
            If bSerialFound Then
                arSerialData("local_ip") = sIp
                arSerialData("token") = sToken
                arSerialData("port") = iPort
                Try
                    arSerialData("server_hostname") = Dns.GetHostEntry(oContext.Request.UserHostAddress).HostName
                Catch ex As Exception
                    arSerialData("server_hostname") = Dns.GetHostEntry(oContext.Request.UserHostAddress).HostName
                End Try

                Try
                    Using oConnexion As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
                        oConnexion.Open()
                        Using oSqlCmdSelect As New SqlCommand("DELETE ActivePrinter WHERE date < DATEADD(minute, -20, GETDATE());" & _
                                                              "UPDATE ActivePrinter SET date = GETDATE(), local_ip = @local_ip, token = @token, port = @port, server_hostname = @server_hostname WHERE serial = @serial;" & _
                                                              "SELECT @@ROWCOUNT", oConnexion)
                            oSqlCmdSelect.Parameters.AddWithValue("@serial", sSerial.ToUpper)
                            oSqlCmdSelect.Parameters.AddWithValue("@local_ip", sIp)
                            oSqlCmdSelect.Parameters.AddWithValue("@token", sToken)
                            oSqlCmdSelect.Parameters.AddWithValue("@port", iPort)
                            oSqlCmdSelect.Parameters.AddWithValue("@server_hostname", Dns.GetHostEntry(oContext.Request.UserHostAddress).HostName)
                            If oSqlCmdSelect.ExecuteScalar() = 0 Then
                                oSqlCmdSelect.CommandText = "INSERT ActivePrinter (serial, local_ip, token, port, server_hostname) VALUES (@serial, @local_ip, @token, @port, @server_hostname)"
                                oSqlCmdSelect.ExecuteNonQuery()
                            End If
                        End Using
                    End Using
                Catch ex As Exception
                    ZSSOUtilities.WriteLog("SetActivePrinter : Err : " & ex.Message)
                End Try

                'oHttpCache.Insert("printer_" & sSerial.ToUpper, arSerialData, Nothing, DateTime.Now.AddMinutes(20.0), TimeSpan.Zero)
            Else
                oContext.Response.StatusCode = 436
                oContext.Response.Write("Unknown printer")
                ZSSOUtilities.WriteLog("SetActivePrinter : Unknown printer")
                Return
            End If

            ZSSOUtilities.WriteLog("SetActivePrinter : OK : " & ZSSOUtilities.oSerializer.Serialize(arSerialData))
        End If
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class