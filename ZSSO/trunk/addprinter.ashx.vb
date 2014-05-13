﻿Imports System.Web
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
                oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title></head><body><form  method=""post"" action=""/addprinter.ashx"" accept-charset=""utf-8"">login <input id=""email"" name=""email"" type=""text"" /><br />password <input id=""password"" name=""password"" type=""text"" /><br />serial <input id=""printersn"" name=""printersn"" type=""text"" /><br />name <input id=""printername"" name=""printername"" type=""text"" /><br /><input id=""Submit1"" type=""submit"" value=""Ok"" /></form></body></html>")
            Else
                sEmail = HttpUtility.UrlDecode(oContext.Request.Form("email"))
                sPassword = HttpUtility.UrlDecode(oContext.Request.Form("password"))
                sSerial = HttpUtility.UrlDecode(oContext.Request.Form("printersn"))
                sName = HttpUtility.UrlDecode(oContext.Request.Form("printername"))

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

                Using oConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
                    oConnection.Open()

                    If Not ZSSOUtilities.Login(oConnection, sEmail, sPassword) Then
                        oContext.Response.ContentType = "text/plain"
                        oContext.Response.StatusCode = 434
                        oContext.Response.Write("Login failed")
                        ZSSOUtilities.WriteLog("AddPrinter : Login failed")
                        Return
                    End If

                    If ZSSOUtilities.SearchSerial(sSerial) = False Then
                        oContext.Response.ContentType = "text/plain"
                        oContext.Response.StatusCode = 436
                        oContext.Response.Write("Unknown printer")
                        ZSSOUtilities.WriteLog("AddPrinter : Unknown printer")
                        Return
                    End If

                    Dim oSqlCmd As SqlCommand = oConnection.CreateCommand()
                    Dim oTransaction As SqlTransaction = oConnection.BeginTransaction("addprinter")

                    oSqlCmd.Connection = oConnection
                    oSqlCmd.Transaction = oTransaction

                    Try
                        'Update or Insert into Printer table
                        oSqlCmd.CommandText = "UPDATE Printer SET Name = @name WHERE Serial = @serial "
                        oSqlCmd.Parameters.AddWithValue("@name", sName)
                        oSqlCmd.Parameters.AddWithValue("@serial", sSerial)
                        oSqlCmd.ExecuteNonQuery()

                        'Update old association
                        oSqlCmd.CommandText = "UPDATE AccountPrinterAssociation SET Deleted = GETDATE() WHERE Serial = @serial AND Deleted IS NULL"
                        oSqlCmd.ExecuteNonQuery()

                        'Insert new association
                        oSqlCmd.CommandText = "INSERT INTO AccountPrinterAssociation (Serial, Email) VALUES (@serial, @email)"
                        oSqlCmd.Parameters.AddWithValue("@email", sEmail)
                        oSqlCmd.ExecuteNonQuery()

                        oTransaction.Commit()
                    Catch ex As Exception
                        ZSSOUtilities.WriteLog("AddPrinter : NOK : Rolling back : " & ex.Message)
                        Try
                            oSqlCmd.Transaction.Rollback()
                        Catch
                            ZSSOUtilities.WriteLog("AddPrinter : NOK : Rollback failed : " & ex.Message)
                        End Try
                    End Try
                End Using
                ZSSOUtilities.WriteLog("AddPrinter : OK")
            End If
        End If
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class