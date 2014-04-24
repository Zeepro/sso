﻿Imports System.Web
Imports System.Web.Services
Imports System.Text.RegularExpressions.Regex
Imports System.Data.SqlClient
Imports System.Security.Cryptography
Imports System.Runtime.Caching
Imports System.IO
Imports BCrypt.Net.BCrypt

Public Class createaccount
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim Email As String
        Dim Password As String
        Dim cacheMemory As ObjectCache = MemoryCache.Default

        If ZSSOUtilities.CheckRequests(context.Request.UserHostAddress, "createaccount") > 5 Then
            context.Response.ContentType = "text/plain"
            context.Response.StatusCode = 435
            context.Response.Write("Too many requests")
            ZSSOUtilities.WriteLog("CreateAccount : Too many requests")
            Return
        Else
            If context.Request.HttpMethod = "GET" Then
                context.Response.ContentType = "text/html"
                context.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title></head><body><form  method=""post"" action=""/createaccount.ashx"" accept-charset=""utf-8"">login <input id=""email"" name=""email"" type=""text"" /><br />password <input id=""password"" name=""password"" type=""text"" /><br /><input id=""Submit1"" type=""submit"" value=""Ok"" /></form></body></html>")
            Else
                Email = HttpUtility.UrlDecode(context.Request.Form("email"))
                Password = HttpUtility.UrlDecode(context.Request.Form("password"))

                ZSSOUtilities.WriteLog("CreateAccount : " + ZSSOUtilities.oSerializer.Serialize(context.Request.Form))
                If String.IsNullOrEmpty(Email) Or String.IsNullOrEmpty(Password) Then
                    context.Response.StatusCode = 432
                    context.Response.Write("Missing parameter")
                    ZSSOUtilities.WriteLog("CreateAccount : Missing parameter")
                    Return
                End If

                'check required password pattern
                If Not (ZSSOUtilities.emailExpression.IsMatch(Email)) Then
                    context.Response.ContentType = "text/plain"
                    context.Response.StatusCode = 433
                    context.Response.Write("Incorrect Parameter")
                    ZSSOUtilities.WriteLog("CreateAccount : Incorrect parameter")
                    Return
                End If

                Using oConnexion As New SqlConnection(ZSSOUtilities.getConnectionString())
                    oConnexion.Open()

                    Using oSqlCmd As New SqlCommand( _
                        "INSERT Account (Email, Password) " & _
                        "VALUES (@email, @password)", _
                        oConnexion)

                        Dim Salt = BCrypt.Net.BCrypt.GenerateSalt()
                        Dim PasswordHash As String = BCrypt.Net.BCrypt.HashPassword(Password, Salt)

                        oSqlCmd.Parameters.AddWithValue("email", Email)
                        oSqlCmd.Parameters.AddWithValue("password", PasswordHash)

                        Try
                            oSqlCmd.ExecuteNonQuery()
                        Catch ex As Exception
                            If ex.HResult = -2146232060 Then
                                context.Response.ContentType = "text/plain"
                                context.Response.StatusCode = 437
                                context.Response.Write("Already Exist")
                                ZSSOUtilities.WriteLog("CreateAccount : Already exist")
                            End If
                            'context.Response.Write("Error : " + "Sauvegarde commande " + ex.Message)
                            Return
                        End Try
                    End Using

                End Using

            End If
        End If
        ZSSOUtilities.WriteLog("CreateAccount : OK")
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property
End Class