﻿Imports System.Web
Imports System.Web.Services
Imports System.Text.RegularExpressions.Regex
Imports System.Data.SqlClient
Imports System.Security.Cryptography
Imports System.Runtime.Caching
Imports System.IO

Public Class createprinter
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal oContext As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim sEmail As String
        Dim sPassword As String
        Dim sPrinters As String
        Dim arResults As List(Of Dictionary(Of String, String)) = New List(Of Dictionary(Of String, String))

        'If ZSSOUtilities.CheckRequests(oContext.Request.UserHostAddress, "createprinter") > 5 Then
        '    oContext.Response.ContentType = "text/plain"
        '    oContext.Response.StatusCode = 435
        '    oContext.Response.Write("Too many requests")
        '    ZSSOUtilities.WriteLog("CreatePrinter : Too many requests")
        '    Return
        'Else
        If oContext.Request.HttpMethod = "GET" Then
            oContext.Response.ContentType = "text/html"
            oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title></head><body><form  method=""post"" action=""/createprinter.ashx"" accept-charset=""utf-8"">login <input id=""email"" name=""email"" type=""text"" /><br />password <input id=""password"" name=""password"" type=""text"" /><br />Printers <input id=""printers"" name=""printers"" type=""text"" /><br /><input id=""Submit1"" type=""submit"" value=""Ok"" /></form></body></html>")
        Else
            sEmail = HttpUtility.UrlDecode(oContext.Request.Form("email"))
            sPassword = HttpUtility.UrlDecode(oContext.Request.Form("password"))
            sPrinters = HttpUtility.UrlDecode(oContext.Request.Form("printers"))

            ZSSOUtilities.WriteLog("CreatePrinter : " & ZSSOUtilities.oSerializer.Serialize({sEmail, sPrinters}))
            If String.IsNullOrEmpty(sEmail) Or String.IsNullOrEmpty(sPassword) Or String.IsNullOrEmpty(sPrinters) Then
                oContext.Response.StatusCode = 432
                oContext.Response.Write("Missing parameter")
                ZSSOUtilities.WriteLog("CreatePrinter : Missing parameter")
                Return
            End If

            Dim arPrinters As List(Of Dictionary(Of String, String))
            Try
                arPrinters = ZSSOUtilities.oSerializer.Deserialize(Of List(Of Dictionary(Of String, String)))(sPrinters)
            Catch ex As Exception
                oContext.Response.ContentType = "text/plain"
                oContext.Response.StatusCode = 433
                oContext.Response.Write("Incorrect Parameter")
                ZSSOUtilities.WriteLog("CreatePrinter : Incorrect parameter")
                Return
            End Try

            If Not (ZSSOUtilities.oEmailRegex.IsMatch(sEmail)) Then
                oContext.Response.ContentType = "text/plain"
                oContext.Response.StatusCode = 433
                oContext.Response.Write("Incorrect Parameter")
                ZSSOUtilities.WriteLog("CreatePrinter : Incorrect parameter")
                Return
            End If

            Using oConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
                oConnection.Open()

                If String.Compare(sEmail, ZSSOUtilities.oSerialEmail) <> 0 Or Not ZSSOUtilities.Login(oConnection, sEmail, sPassword) Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 434
                    oContext.Response.Write("Login failed")
                    ZSSOUtilities.WriteLog("CreatePrinter : Login failed")
                    Return
                End If

                Dim sQuery As String = "INSERT INTO Printer (Serial, Manufactured, Type, Ean, Upc, Rangecode) VALUES (@serial, @manufactured, @type, @ean, @upc, @rangecode)"
                Dim arPrinterResult As Dictionary(Of String, String)

                For Each oPrinter In arPrinters
                    If oPrinter.ContainsKey("serialnumber") And oPrinter.ContainsKey("manufactured") _
                    And oPrinter.ContainsKey("type") And oPrinter.ContainsKey("EAN") _
                    And oPrinter.ContainsKey("UPC") And oPrinter.ContainsKey("rangecode") Then
                        arPrinterResult = New Dictionary(Of String, String)
                        arPrinterResult("serialnumber") = oPrinter("serialnumber")
                        If (ZSSOUtilities.CheckRangeCode(oPrinter("rangecode")) = False) Then
                            arPrinterResult("result") = "unavailable range"
                        ElseIf (CLng("&H" & oPrinter("serialnumber")) < CLng("&H" & oPrinter("rangecode").Substring(0, 12))) Or (CLng("&H" & oPrinter("serialnumber")) > CLng("&H" & oPrinter("rangecode").Substring(12, 12))) Then
                            arPrinterResult("result") = "incorrect MAC"
                        ElseIf ZSSOUtilities.SearchSerial(oPrinter("serialnumber")) = True Then
                            arPrinterResult("result") = "already exist"
                        ElseIf String.Compare(oPrinter("rangecode"), "testrange") = 0 Then
                            arPrinterResult("result") = "test"
                        Else
                            Using oSqlCmdInsert As New SqlCommand(sQuery, oConnection)
                                oSqlCmdInsert.Parameters.AddWithValue("@serial", oPrinter("serialnumber"))
                                oSqlCmdInsert.Parameters.AddWithValue("@manufactured", DateTime.Parse(oPrinter("manufactured")))
                                oSqlCmdInsert.Parameters.AddWithValue("@type", oPrinter("type"))
                                oSqlCmdInsert.Parameters.AddWithValue("@ean", oPrinter("EAN"))
                                oSqlCmdInsert.Parameters.AddWithValue("@upc", oPrinter("UPC"))
                                oSqlCmdInsert.Parameters.AddWithValue("@rangecode", oPrinter("rangecode"))

                                Try
                                    oSqlCmdInsert.ExecuteNonQuery()
                                    arPrinterResult("result") = "ok"
                                Catch ex As Exception
                                    ZSSOUtilities.WriteLog("CreatePrinter : NOK : " & ex.Message)
                                    arPrinterResult("result") = "incorrect parameter"
                                End Try

                            End Using
                        End If
                        arResults.Add(arPrinterResult)
                    End If
                Next

            End Using
            oContext.Response.ContentType = "text/plain"
            oContext.Response.Write(ZSSOUtilities.oSerializer.Serialize(arResults))
            ZSSOUtilities.WriteLog("CreatePrinter : OK")
        End If
        'End If
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class