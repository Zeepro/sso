﻿Imports System.Web
Imports System.Web.Services
Imports System.Text.RegularExpressions.Regex
Imports System.Web.Script.Serialization
Imports System.Data.SqlClient
Imports System.Runtime.Caching
Imports System.IO

Public Class testaccount
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal oContext As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim sEmail As String
        Dim arReturnValue = New Dictionary(Of String, String)

        If ZSSOUtilities.CheckRequests(oContext.Request.UserHostAddress, "testaccount") > 5 Then
            oContext.Response.ContentType = "text/plain"
            oContext.Response.StatusCode = 435
            oContext.Response.Write("Too many requests")
            ZSSOUtilities.WriteLog("TestAccount : Too many requests")
            Return
        Else
            If oContext.Request.HttpMethod = "GET" Then
                oContext.Response.ContentType = "text/html"
                oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title></head><body><form  method=""post"" action=""/testaccount.ashx"" accept-charset=""utf-8"">login <input id=""email"" name=""email"" type=""text"" /><br /><input id=""Submit1"" type=""submit"" value=""Ok"" /></form></body></html>")
            Else
                sEmail = HttpUtility.UrlDecode(oContext.Request.Form("email"))

                ZSSOUtilities.WriteLog("TestAccount : " & ZSSOUtilities.oSerializer.Serialize(oContext.Request.Form))

                If String.IsNullOrEmpty(sEmail) Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 432
                    oContext.Response.Write("Missing parameter")
                    ZSSOUtilities.WriteLog("TestAccount : Missing parameter")
                    Return
                End If

                If Not (ZSSOUtilities.oEmailRegex.IsMatch(sEmail)) Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 433
                    oContext.Response.Write("Incorrect Parameter")
                    ZSSOUtilities.WriteLog("TestAccount : Incorrect parameter")
                    Return
                End If

                If SearchEmail(sEmail) Then
                    arReturnValue("account") = "exist"
                Else
                    arReturnValue("account") = "unknown"
                End If

                oContext.Response.ContentType = "text/plain"
                oContext.Response.Write(ZSSOUtilities.oSerializer.Serialize(arReturnValue))
            End If
        End If
        ZSSOUtilities.WriteLog("TestAccount : OK : " & ZSSOUtilities.oSerializer.Serialize(arReturnValue))
    End Sub

    Public Shared Function SearchEmail(sEmail As String)
        Using oConnexion As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
            oConnexion.Open()

            Dim sQuery = "SELECT TOP 1 Email " & _
                "FROM Account " & _
                "WHERE Email=@email"

            Using oSqlCmdSelect As New SqlCommand(sQuery, oConnexion)

                oSqlCmdSelect.Parameters.AddWithValue("@email", sEmail)

                Try
                    Using oQueryResult As SqlDataReader = oSqlCmdSelect.ExecuteReader()

                        If oQueryResult.HasRows Then
                            Return True
                        End If
                    End Using
                Catch ex As Exception
                End Try
            End Using
        End Using
        Return False
    End Function

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class