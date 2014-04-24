Imports System.Web
Imports System.Web.Services
Imports System.Text.RegularExpressions.Regex
Imports System.Data.SqlClient
Imports System.Security.Cryptography
Imports System.Runtime.Caching
Imports System.IO

Public Class addprinter
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim Email As String
        Dim Password As String
        Dim Serial As String
        Dim Name As String
        Dim cacheMemory As ObjectCache = MemoryCache.Default

        If ZSSOUtilities.CheckRequests(context.Request.UserHostAddress, "addprinter") > 5 Then
            context.Response.ContentType = "text/plain"
            context.Response.StatusCode = 435
            context.Response.Write("Too many requests")
            ZSSOUtilities.WriteLog("AddPrinter : Too many requests")
            Return
        Else
            If context.Request.HttpMethod = "GET" Then
                context.Response.ContentType = "text/html"
                context.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title></head><body><form  method=""post"" action=""/addprinter.ashx"" accept-charset=""utf-8"">login <input id=""email"" name=""email"" type=""text"" /><br />password <input id=""password"" name=""password"" type=""text"" /><br />serial <input id=""serial"" name=""serial"" type=""text"" /><br />name <input id=""name"" name=""name"" type=""text"" /><br /><input id=""Submit1"" type=""submit"" value=""Ok"" /></form></body></html>")
            Else
                Email = HttpUtility.UrlDecode(context.Request.Form("email"))
                Password = HttpUtility.UrlDecode(context.Request.Form("password"))
                Serial = HttpUtility.UrlDecode(context.Request.Form("serial"))
                Name = HttpUtility.UrlDecode(context.Request.Form("name"))

                ZSSOUtilities.WriteLog("AddPrinter : " + ZSSOUtilities.oSerializer.Serialize(context.Request.Form))
                If String.IsNullOrEmpty(Email) Or String.IsNullOrEmpty(Password) Or String.IsNullOrEmpty(Serial) Or String.IsNullOrEmpty(Name) Then
                    context.Response.StatusCode = 432
                    context.Response.Write("Missing parameter")
                    ZSSOUtilities.WriteLog("AddPrinter : Missing parameter")
                    Return
                End If

                'check required password pattern
                If Not (ZSSOUtilities.emailExpression.IsMatch(Email)) Then
                    context.Response.ContentType = "text/plain"
                    context.Response.StatusCode = 433
                    context.Response.Write("Incorrect Parameter")
                    ZSSOUtilities.WriteLog("AddPrinter : Incorrect parameter")
                    Return
                End If

                Using oConnexion As New SqlConnection(ZSSOUtilities.getConnectionString())
                    oConnexion.Open()

                    If Not ZSSOUtilities.Login(oConnexion, Email, Password) Then
                        context.Response.ContentType = "text/plain"
                        context.Response.StatusCode = 434
                        context.Response.Write("Login failed")
                        ZSSOUtilities.WriteLog("ChangePassword : Login failed")
                        Return
                    End If

                    Dim QueryString = "UPDATE Printer SET Name = @name, EmailAccount = @email WHERE Serial = @serial " & _
                        "IF @@ROWCOUNT=0 INSERT INTO Printer VALUES (@serial, @name, @email, DEFAULT, NULL)"

                    Using oSqlCmdInsert As New SqlCommand(QueryString, oConnexion)
                        oSqlCmdInsert.Parameters.AddWithValue("@email", Email)
                        oSqlCmdInsert.Parameters.AddWithValue("@name", Name)
                        oSqlCmdInsert.Parameters.AddWithValue("@serial", Serial)

                        Try
                            oSqlCmdInsert.ExecuteNonQuery()
                        Catch ex As Exception
                            'context.Response.Write("Error : " + " commande " + ex.Message)
                            Return
                        End Try

                    End Using

                End Using
            End If
        End If
        ZSSOUtilities.WriteLog("AddPrinter : OK")
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class