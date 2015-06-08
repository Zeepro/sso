Imports System.Web
Imports System.Web.Services
Imports System.Text.RegularExpressions.Regex
Imports System.Data.SqlClient
Imports System.Security.Cryptography
Imports System.Runtime.Caching
Imports System.IO

Public Class listprinter
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal oContext As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim sEmail As String
        Dim sPassword As String
        Dim sToken As String
        Dim oHttpCache As Caching.Cache = HttpRuntime.Cache

        If ZSSOUtilities.CheckRequests(oContext.Request.UserHostAddress, "listprinter") > 5 Then
            oContext.Response.ContentType = "text/plain"
            oContext.Response.StatusCode = 435
            oContext.Response.Write("Too many requests")
            ZSSOUtilities.WriteLog("ListPrinter : Too many requests")
            Return
        Else
            If oContext.Request.HttpMethod = "GET" Then
                oContext.Response.ContentType = "text/html"
                oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title>" & _
                                        "<script src=""https://code.jquery.com/jquery-1.10.2.js""></script><script type=""text/javascript"">function load_wait() { $(""#overlay"").addClass(""gray-overlay""); $("".ui-loader"").css(""display"", ""block""); }</script>" & _
                                        "<link rel=""stylesheet"" type=""text/css"" href=""style.css"">" & _
                                        "</head><body><div id=""overlay""></div><div class=""ui-loader ui-corner-all ui-body-a ui-loader-default""><span class=""ui-icon-loading""></span><h1>loading</h1></div>" & _
                                        "<form  method=""post"" action=""/listprinter.ashx"" accept-charset=""utf-8"">" & _
                                        "login <input id=""email"" name=""email"" type=""text"" /><br />" & _
                                        "password <input id=""password"" name=""password"" type=""text"" /><br />" & _
                                        "<br />" & _
                                        "(user) token <input id=""token"" name=""token"" type=""text"" /><br />" & _
                                        "<input id=""Submit1"" type=""submit"" value=""Ok""  onclick=""javascript: load_wait();"" /></form></body></html>")
            Else
                sEmail = HttpUtility.UrlDecode(oContext.Request.Form("email"))
                sPassword = HttpUtility.UrlDecode(oContext.Request.Form("password"))

                sToken = oContext.Request.Form("token")

                ZSSOUtilities.WriteLog("ListPrinter : " & ZSSOUtilities.oSerializer.Serialize({sEmail}))
                If (String.IsNullOrEmpty(sEmail) OrElse String.IsNullOrEmpty(sPassword)) AndAlso _
                    String.IsNullOrEmpty(sToken) Then
                    oContext.Response.StatusCode = 432
                    oContext.Response.Write("Missing parameter")
                    ZSSOUtilities.WriteLog("ListPrinter : Missing parameter")
                    Return
                End If

                If (Not String.IsNullOrEmpty(sEmail) AndAlso Not String.IsNullOrEmpty(sToken)) OrElse _
                    (Not String.IsNullOrEmpty(sEmail) AndAlso Not ZSSOUtilities.oEmailRegex.IsMatch(sEmail)) OrElse _
                    (Not String.IsNullOrEmpty(sToken) AndAlso sToken.Length <> 40) Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 433
                    oContext.Response.Write("Incorrect Parameter")
                    ZSSOUtilities.WriteLog("ListPrinter : Incorrect parameter")
                    Return
                End If

                Using oConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
                    oConnection.Open()

                    If Not String.IsNullOrEmpty(sEmail) AndAlso _
                        Not String.IsNullOrEmpty(sPassword) AndAlso _
                        Not ZSSOUtilities.Login(oConnection, sEmail, sPassword) Then
                        oContext.Response.ContentType = "text/plain"
                        oContext.Response.StatusCode = 434
                        oContext.Response.Write("Login failed")
                        ZSSOUtilities.WriteLog("ListPrinter: Login failed")
                        Return
                    End If

                    If Not String.IsNullOrEmpty(sToken) Then
                        sEmail = ZSSOUtilities.SearchAccountEmail(sToken)
                        If sEmail Is Nothing Then
                            oContext.Response.ContentType = "text/plain"
                            oContext.Response.StatusCode = 442
                            oContext.Response.Write("Unknow user")
                            ZSSOUtilities.WriteLog("ListPrinter: Unknow user")
                            Return
                        End If
                    End If

                    Dim sQuery = "DELETE ActivePrinter WHERE date < DATEADD(minute, -20, GETDATE());" & _
                        "SELECT [AccountPrinterAssociation].Email, [AccountPrinterAssociation].Deleted, [AccountPrinterAssociation].Serial, [Printer].Name, [ActivePrinter].local_ip, [ActivePrinter].server_hostname, [ActivePrinter].port, [ActivePrinter].token, [Printer].current_software, [Printer].next_software, [AccountPrinterAssociation].LastAccess, [AccountPrinterAssociation].CalibrationWarningMessage " & _
                        "FROM [AccountPrinterAssociation] " & _
                        "INNER JOIN Printer ON [AccountPrinterAssociation].Serial = [Printer].Serial " & _
                        "INNER JOIN ActivePrinter ON [ActivePrinter].Serial = [Printer].Serial " & _
                        "WHERE [AccountPrinterAssociation].Email = @email AND [AccountPrinterAssociation].Deleted IS NULL " & _
                        "ORDER BY [Printer].Name"

                    Using oSqlCmdSelect As New SqlCommand(sQuery, oConnection)
                        oSqlCmdSelect.Parameters.AddWithValue("@email", sEmail)

                        Dim arAccountPrinters = New Dictionary(Of String, Dictionary(Of String, String))
                        Using oQueryResult As SqlDataReader = oSqlCmdSelect.ExecuteReader()

                            While oQueryResult.Read()
                                Dim sSerial = oQueryResult(oQueryResult.GetOrdinal("Serial"))

                                Dim arPrinterData = New Dictionary(Of String, String)
                                arPrinterData("printername") = oQueryResult(oQueryResult.GetOrdinal("Name"))
                                arPrinterData("localIP") = oQueryResult("local_ip")
                                arPrinterData("URL") = sSerial & "." & oQueryResult("server_hostname") & ":" & oQueryResult("port")
                                arPrinterData("token") = oQueryResult("token")
                                If Not oQueryResult("current_software") Is System.DBNull.Value Then
                                    arPrinterData("current_software") = oQueryResult("current_software")
                                End If
                                If Not oQueryResult("next_software") Is System.DBNull.Value Then
                                    arPrinterData("next_software") = oQueryResult("next_software")
                                End If
                                If Not oQueryResult("lastaccess") Is System.DBNull.Value Then
                                    arPrinterData("lastaccess") = oQueryResult("lastaccess")
                                End If
                                If oQueryResult("calibrationwarningmessage") Then
                                    arPrinterData("calibrationwarningmessage") = "yes"
                                Else
                                    arPrinterData("calibrationwarningmessage") = "no"
                                End If
                                arAccountPrinters(sSerial) = arPrinterData
                            End While
                        End Using
                        oContext.Response.ContentType = "text/plain"
                        oContext.Response.Write(ZSSOUtilities.oSerializer.Serialize(arAccountPrinters.Values))
                        ZSSOUtilities.WriteLog("ListPrinter : OK : " & ZSSOUtilities.oSerializer.Serialize(arAccountPrinters.Values))
                    End Using
                End Using
            End If
        End If
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class