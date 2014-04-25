﻿Imports System.Web
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
        Dim bSerialFound As Boolean = False
        Dim oHttpCache As Caching.Cache = HttpRuntime.Cache

        If oContext.Request.HttpMethod = "GET" Then
            oContext.Response.ContentType = "text/html"
            oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title></head><body><form  method=""post"" action=""/setactiveprinter.ashx"" accept-charset=""utf-8"">Serial <input id=""serial"" name=""serial"" type=""text"" /><br />Local IP <input id=""ip"" name=""ip"" type=""text"" /><br />Token <input id=""token"" name=""token"" type=""text"" /><input id=""Submit1"" type=""submit"" value=""Ok"" /></form></body></html>")
        Else
            sSerial = HttpUtility.UrlDecode(oContext.Request.Form("serial"))
            sIp = HttpUtility.UrlDecode(oContext.Request.Form("ip"))
            sToken = HttpUtility.UrlDecode(oContext.Request.Form("token"))
            ZSSOUtilities.WriteLog("SetActivePrinter : " & ZSSOUtilities.oSerializer.Serialize(oContext.Request.Form))

            If String.IsNullOrEmpty(sSerial) Or String.IsNullOrEmpty(sIp) Or String.IsNullOrEmpty(sToken) Then
                oContext.Response.StatusCode = 432
                oContext.Response.Write("Missing parameter")
                ZSSOUtilities.WriteLog("SetActivePrinter : Missing parameter")
                Return
            End If

            Dim oIpa As IPAddress = Nothing
            If Not (IPAddress.TryParse(sIp, oIpa)) Then
                oContext.Response.StatusCode = 433
                oContext.Response.Write("Incorrect Parameter")
                ZSSOUtilities.WriteLog("SetActivePrinter : Incorrect parameter")
                Return
            End If

            Using oConnexion As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
                oConnexion.Open()

                Dim sQuery = "SELECT TOP 1 Serial " & _
                    "FROM Printer " & _
                    "WHERE Serial=@serial"

                Using oSqlCmdSelect As New SqlCommand(sQuery, oConnexion)

                    oSqlCmdSelect.Parameters.AddWithValue("@serial", sSerial)

                    Try
                        Using oQueryResult As SqlDataReader = oSqlCmdSelect.ExecuteReader()
                            If oQueryResult.HasRows Then
                                bSerialFound = True
                            End If
                        End Using
                    Catch ex As Exception
                    End Try
                End Using

            End Using

            If bSerialFound Then
                Dim arSerialData = New Dictionary(Of String, String)
                arSerialData("local_ip") = sIp
                arSerialData("token") = sToken
                arSerialData("server_hostname") = Dns.GetHostEntry(oContext.Request.UserHostAddress).HostName & ".zeepro.com"
                oHttpCache.Insert("printer_" & sSerial, arSerialData, Nothing, DateTime.Now.AddMinutes(20.0), TimeSpan.Zero)
            Else
                oContext.Response.StatusCode = 436
                oContext.Response.Write("Unknown printer")
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