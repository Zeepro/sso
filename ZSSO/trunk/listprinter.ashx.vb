Imports System.Web
Imports System.Web.Services
Imports System.Text.RegularExpressions.Regex
Imports System.Data.SqlClient
Imports System.Security.Cryptography
Imports System.Runtime.Caching
Imports System.IO

Public Class listprinter
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim Email As String
        Dim Password As String
        Dim HttpMemory As Caching.Cache = HttpRuntime.Cache

        If ZSSOUtilities.CheckRequests(context.Request.UserHostAddress, "listprinter") > 5 Then
            context.Response.ContentType = "text/plain"
            context.Response.StatusCode = 435
            context.Response.Write("Too many requests")
            ZSSOUtilities.WriteLog("ListPrinter : Too many requests")
            Return
        Else
            If context.Request.HttpMethod = "GET" Then
                context.Response.ContentType = "text/html"
                context.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title></head><body><form  method=""post"" action=""/listprinter.ashx"" accept-charset=""utf-8"">login <input id=""email"" name=""email"" type=""text"" /><br />password <input id=""password"" name=""password"" type=""text"" /><br /><input id=""Submit1"" type=""submit"" value=""Ok"" /></form></body></html>")
            Else
                Email = HttpUtility.UrlDecode(context.Request.Form("email"))
                Password = HttpUtility.UrlDecode(context.Request.Form("password"))

                ZSSOUtilities.WriteLog("ListPrinter : " + ZSSOUtilities.oSerializer.Serialize(context.Request.Form))
                If String.IsNullOrEmpty(Email) Or String.IsNullOrEmpty(Password) Then
                    context.Response.StatusCode = 432
                    context.Response.Write("Missing parameter")
                    ZSSOUtilities.WriteLog("ListPrinter : Missing parameter")
                    Return
                End If

                'check required password pattern
                If Not (ZSSOUtilities.emailExpression.IsMatch(Email)) Then
                    context.Response.ContentType = "text/plain"
                    context.Response.StatusCode = 433
                    context.Response.Write("Incorrect Parameter")
                    ZSSOUtilities.WriteLog("ListPrinter : Incorrect parameter")
                    Return
                End If

                Using oConnexion As New SqlConnection(ZSSOUtilities.getConnectionString())
                    oConnexion.Open()

                    If Not ZSSOUtilities.Login(oConnexion, Email, Password) Then
                        context.Response.ContentType = "text/plain"
                        context.Response.StatusCode = 434
                        context.Response.Write("Login failed")
                        ZSSOUtilities.WriteLog("ListPrinter : Login failed")
                        Return
                    End If

                    Dim QueryString = "SELECT * " & _
                       "FROM Printer " & _
                       "WHERE EmailAccount=@email"

                    Using oSqlCmdSelect As New SqlCommand(QueryString, oConnexion)
                        oSqlCmdSelect.Parameters.AddWithValue("@email", Email)

                        Dim AccountPrinters = New Dictionary(Of String, Dictionary(Of String, String))
                        Using QueryResult As SqlDataReader = oSqlCmdSelect.ExecuteReader()

                            While QueryResult.Read()
                                Dim Serial = QueryResult(QueryResult.GetOrdinal("Serial"))
                                Dim Name = QueryResult(QueryResult.GetOrdinal("Name"))

                                Dim cachedPrinter = TryCast(HttpMemory("printer_" + Serial), Dictionary(Of String, String))
                                If Not IsNothing(cachedPrinter) Then
                                    Dim PrinterData = New Dictionary(Of String, String)
                                    PrinterData("printername") = Name
                                    PrinterData("localIP") = cachedPrinter("local_ip")
                                    PrinterData("FQDN") = Serial + "." + cachedPrinter("server_hostname")
                                    AccountPrinters(Serial) = PrinterData
                                End If
                            End While
                        End Using
                        context.Response.ContentType = "text/plain"
                        context.Response.Write(ZSSOUtilities.oSerializer.Serialize(AccountPrinters.Values))
                        ZSSOUtilities.WriteLog("ListPrinter : OK : " + ZSSOUtilities.oSerializer.Serialize(AccountPrinters))
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