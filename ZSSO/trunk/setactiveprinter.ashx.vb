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

    Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim Serial As String = ""
        Dim Ip As String = ""
        Dim Token As String = ""
        Dim SerialFound As Boolean = False
        Dim HttpMemory As Caching.Cache = HttpRuntime.Cache

        If context.Request.HttpMethod = "GET" Then
            context.Response.ContentType = "text/html"
            context.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title></head><body><form  method=""post"" action=""/setactiveprinter.ashx"" accept-charset=""utf-8"">Serial <input id=""serial"" name=""serial"" type=""text"" /><br />Local IP <input id=""ip"" name=""ip"" type=""text"" /><br />Token <input id=""token"" name=""token"" type=""text"" /><input id=""Submit1"" type=""submit"" value=""Ok"" /></form></body></html>")
        Else
            Serial = HttpUtility.UrlDecode(context.Request.Form("serial"))
            Ip = HttpUtility.UrlDecode(context.Request.Form("ip"))
            Token = HttpUtility.UrlDecode(context.Request.Form("token"))
            ZSSOUtilities.WriteLog("SetActivePrinter : " + ZSSOUtilities.oSerializer.Serialize(context.Request.Form))

            If String.IsNullOrEmpty(Serial) Or String.IsNullOrEmpty(Ip) Or String.IsNullOrEmpty(Token) Then
                context.Response.StatusCode = 432
                context.Response.Write("Missing parameter")
                ZSSOUtilities.WriteLog("SetActivePrinter : Missing parameter")
                Return
            End If

            Dim Ipa As IPAddress = Nothing
            If Not (IPAddress.TryParse(Ip, Ipa)) Then
                context.Response.StatusCode = 433
                context.Response.Write("Incorrect Parameter")
                ZSSOUtilities.WriteLog("SetActivePrinter : Incorrect parameter")
                Return
            End If

            Using oConnexion As New SqlConnection(ZSSOUtilities.getConnectionString())
                oConnexion.Open()

                Dim QueryString = "SELECT TOP 1 Serial " & _
                    "FROM Printer " & _
                    "WHERE Serial=@serial"

                Using oSqlCmd As New SqlCommand(QueryString, oConnexion)

                    oSqlCmd.Parameters.AddWithValue("@serial", Serial)

                    Try
                        Using QueryResult As SqlDataReader = oSqlCmd.ExecuteReader()
                            If QueryResult.HasRows Then
                                SerialFound = True
                            End If
                        End Using
                    Catch ex As Exception
                        'context.Response.Write("Error : " + "Sauvegarde commande " + ex.Message)
                        Return
                    End Try
                End Using

            End Using

            If SerialFound Then
                Dim SerialData = New Dictionary(Of String, String)
                SerialData("local_ip") = Ip
                SerialData("token") = Token
                SerialData("server_hostname") = Dns.GetHostEntry(context.Request.UserHostAddress).HostName + ".zeepro.com"
                HttpMemory.Insert("printer_" + Serial, SerialData, Nothing, DateTime.Now.AddMinutes(20.0), TimeSpan.Zero)
            Else
                context.Response.StatusCode = 436
                context.Response.Write("Unknown printer")
                ZSSOUtilities.WriteLog("SetActivePrinter : Unknown printer")
                Return
            End If
        End If
        ZSSOUtilities.WriteLog("SetActivePrinter : OK")
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class