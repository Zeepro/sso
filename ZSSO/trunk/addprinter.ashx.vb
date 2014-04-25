Imports System.Web
Imports System.Web.Services
Imports System.Text.RegularExpressions.Regex
Imports System.Data.SqlClient
Imports System.Security.Cryptography
Imports System.Runtime.Caching
Imports System.IO

Public Class addprinter
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal oContext As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim sEmail As String
        Dim sPassword As String
        Dim sSerial As String
        Dim sName As String

        If ZSSOUtilities.CheckRequests(oContext.Request.UserHostAddress, "addprinter") > 5 Then
            oContext.Response.ContentType = "text/plain"
            oContext.Response.StatusCode = 435
            oContext.Response.Write("Too many requests")
            ZSSOUtilities.WriteLog("AddPrinter : Too many requests")
            Return
        Else
            If oContext.Request.HttpMethod = "GET" Then
                oContext.Response.ContentType = "text/html"
                oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title></head><body><form  method=""post"" action=""/addprinter.ashx"" accept-charset=""utf-8"">login <input id=""email"" name=""email"" type=""text"" /><br />password <input id=""password"" name=""password"" type=""text"" /><br />serial <input id=""serial"" name=""serial"" type=""text"" /><br />name <input id=""name"" name=""name"" type=""text"" /><br /><input id=""Submit1"" type=""submit"" value=""Ok"" /></form></body></html>")
            Else
                sEmail = HttpUtility.UrlDecode(oContext.Request.Form("email"))
                sPassword = HttpUtility.UrlDecode(oContext.Request.Form("password"))
                sSerial = HttpUtility.UrlDecode(oContext.Request.Form("serial"))
                sName = HttpUtility.UrlDecode(oContext.Request.Form("name"))

                ZSSOUtilities.WriteLog("AddPrinter : " & ZSSOUtilities.oSerializer.Serialize({sEmail, sSerial, sName}))
                If String.IsNullOrEmpty(sEmail) Or String.IsNullOrEmpty(sPassword) Or String.IsNullOrEmpty(sSerial) Or String.IsNullOrEmpty(sName) Then
                    oContext.Response.StatusCode = 432
                    oContext.Response.Write("Missing parameter")
                    ZSSOUtilities.WriteLog("AddPrinter : Missing parameter")
                    Return
                End If

                If Not (ZSSOUtilities.oEmailRegex.IsMatch(sEmail)) Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 433
                    oContext.Response.Write("Incorrect Parameter")
                    ZSSOUtilities.WriteLog("AddPrinter : Incorrect parameter")
                    Return
                End If

                Using oConnexion As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
                    oConnexion.Open()

                    If Not ZSSOUtilities.Login(oConnexion, sEmail, sPassword) Then
                        oContext.Response.ContentType = "text/plain"
                        oContext.Response.StatusCode = 434
                        oContext.Response.Write("Login failed")
                        ZSSOUtilities.WriteLog("ChangePassword : Login failed")
                        Return
                    End If

                    Dim sQuery = "UPDATE Printer SET Name = @name WHERE Serial = @serial " & _
                        "IF @@ROWCOUNT=0 INSERT INTO Printer (Serial, Name) VALUES (@serial, @name)"

                    Using oSqlCmdUpdate As New SqlCommand(sQuery, oConnexion)
                        oSqlCmdUpdate.Parameters.AddWithValue("@name", sName)
                        oSqlCmdUpdate.Parameters.AddWithValue("@serial", sSerial)

                        Try
                            oSqlCmdUpdate.ExecuteNonQuery()
                        Catch ex As Exception
                            'context.Response.Write("Error : " + " commande " + ex.Message)
                            Return
                        End Try

                    End Using

                    sQuery = "INSERT INTO AccountPrinterAssociation (Serial, Email) VALUES (@serial, @email)"

                    Using oSqlCmdInsert As New SqlCommand(sQuery, oConnexion)
                        oSqlCmdInsert.Parameters.AddWithValue("@email", sEmail)
                        oSqlCmdInsert.Parameters.AddWithValue("@serial", sSerial)

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