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
                oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title></head><body><form  method=""post"" action=""/listprinter.ashx"" accept-charset=""utf-8"">login <input id=""email"" name=""email"" type=""text"" /><br />password <input id=""password"" name=""password"" type=""text"" /><br /><input id=""Submit1"" type=""submit"" value=""Ok"" /></form></body></html>")
            Else
                sEmail = HttpUtility.UrlDecode(oContext.Request.Form("email"))
                sPassword = HttpUtility.UrlDecode(oContext.Request.Form("password"))

                ZSSOUtilities.WriteLog("ListPrinter : " & ZSSOUtilities.oSerializer.Serialize(oContext.Request.Form))
                If String.IsNullOrEmpty(sEmail) Or String.IsNullOrEmpty(sPassword) Then
                    oContext.Response.StatusCode = 432
                    oContext.Response.Write("Missing parameter")
                    ZSSOUtilities.WriteLog("ListPrinter : Missing parameter")
                    Return
                End If

                If Not (ZSSOUtilities.oEmailRegex.IsMatch(sEmail)) Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 433
                    oContext.Response.Write("Incorrect Parameter")
                    ZSSOUtilities.WriteLog("ListPrinter : Incorrect parameter")
                    Return
                End If

                Using oConnexion As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
                    oConnexion.Open()

                    If Not ZSSOUtilities.Login(oConnexion, sEmail, sPassword) Then
                        oContext.Response.ContentType = "text/plain"
                        oContext.Response.StatusCode = 434
                        oContext.Response.Write("Login failed")
                        ZSSOUtilities.WriteLog("ListPrinter : Login failed")
                        Return
                    End If

                    Dim sQuery = "SELECT [AccountPrinterAssociation].Email, [AccountPrinterAssociation].[Create], [AccountPrinterAssociation].Serial, [Printer].Name " & _
                        "FROM [AccountPrinterAssociation] " & _
                        "INNER JOIN " & _
                        "(SELECT max([Create]) AS LatestDate, Serial " & _
                        "FROM [AccountPrinterAssociation ]" & _
                        "GROUP BY Serial) AS LastAssociation " & _
                        "ON [AccountPrinterAssociation].[Create] = LastAssociation.LatestDate AND [AccountPrinterAssociation].Serial = LastAssociation.Serial " & _
                        "INNER JOIN Printer ON LastAssociation.Serial = [Printer].Serial " & _
                        "WHERE [AccountPrinterAssociation].Email = @email"

                    Using oSqlCmdSelect As New SqlCommand(sQuery, oConnexion)
                        oSqlCmdSelect.Parameters.AddWithValue("@email", sEmail)

                        Dim arAccountPrinters = New Dictionary(Of String, Dictionary(Of String, String))
                        Using oQueryResult As SqlDataReader = oSqlCmdSelect.ExecuteReader()

                            While oQueryResult.Read()
                                Dim sSerial = oQueryResult(oQueryResult.GetOrdinal("Serial"))

                                Dim arCachedPrinter = TryCast(oHttpCache("printer_" & sSerial), Dictionary(Of String, String))
                                If Not IsNothing(arCachedPrinter) Then
                                    Dim arPrinterData = New Dictionary(Of String, String)
                                    arPrinterData("printername") = oQueryResult(oQueryResult.GetOrdinal("Name"))
                                    arPrinterData("localIP") = arCachedPrinter("local_ip")
                                    arPrinterData("FQDN") = sSerial & "." & arCachedPrinter("server_hostname")
                                    arAccountPrinters(sSerial) = arPrinterData
                                    End If
                                End While
                            End Using
                        oContext.Response.ContentType = "text/plain"
                        oContext.Response.Write(ZSSOUtilities.oSerializer.Serialize(arAccountPrinters.Values))
                        ZSSOUtilities.WriteLog("ListPrinter : OK : " & ZSSOUtilities.oSerializer.Serialize(arAccountPrinters))
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