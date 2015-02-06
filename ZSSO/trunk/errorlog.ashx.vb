Imports System.Web
Imports System.Web.Services
Imports System.Text.RegularExpressions.Regex
Imports System.Data.SqlClient
Imports System.Security.Cryptography
Imports System.Runtime.Caching
Imports System.IO

Public Class errorlog
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal oContext As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim sSerial As String
        Dim sTime As String
        Dim sLevel As String
        Dim sCode As String
        Dim sMessage As String
        Dim sTemplate As String = ""
        Dim sSubject As String = ""

        If oContext.Request.HttpMethod = "GET" Then
            oContext.Response.ContentType = "text/html"
            oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title>" & _
                                    "<script src=""https://code.jquery.com/jquery-1.10.2.js""></script><script type=""text/javascript"">function load_wait() { $(""#overlay"").addClass(""gray-overlay""); $("".ui-loader"").css(""display"", ""block""); }</script>" & _
                                    "<link rel=""stylesheet"" type=""text/css"" href=""style.css"">" & _
                                    "</head><body><div id=""overlay""></div><div class=""ui-loader ui-corner-all ui-body-a ui-loader-default""><span class=""ui-icon-loading""></span><h1>loading</h1></div><form  method=""post"" action=""/errorlog.ashx"" accept-charset=""utf-8"">Serial <input id=""printersn"" name=""printersn"" type=""text"" /><br />printertime <input id=""printertime"" name=""printertime"" type=""text"" /><br />level <input id=""level"" name=""level"" type=""text"" /><br />code <input id=""code"" name=""code"" type=""text"" /><br />message <input id=""message"" name=""message"" type=""text"" /><br /><input id=""Submit1"" type=""submit"" value=""Ok""  onclick=""javascript: load_wait();"" /></form></body></html>")
        Else
            sSerial = HttpUtility.UrlDecode(oContext.Request.Form("printersn"))
            sTime = HttpUtility.UrlDecode(oContext.Request.Form("printertime"))
            sLevel = HttpUtility.UrlDecode(oContext.Request.Form("level"))
            sCode = HttpUtility.UrlDecode(oContext.Request.Form("code"))
            sMessage = HttpUtility.UrlDecode(oContext.Request.Form("message"))

            If String.IsNullOrEmpty(sSerial) Or String.IsNullOrEmpty(sTime) _
                Or String.IsNullOrEmpty(sLevel) Or String.IsNullOrEmpty(sCode) Then
                oContext.Response.StatusCode = 432
                oContext.Response.Write("Missing parameter")
                ZSSOUtilities.WriteLog("ErrorLog : Missing parameter")
                Return
            End If

            Try
                DateTime.Parse(sTime)
                If Not ZSSOUtilities.SearchSerial(sSerial) Then
                    Throw New Exception
                End If
           Catch ex As Exception
                oContext.Response.ContentType = "text/plain"
                oContext.Response.StatusCode = 433
                oContext.Response.Write("Incorrect Parameter")
                ZSSOUtilities.WriteLog("ErrorLog : Incorrect parameter")
                Return
            End Try

            Using oConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
                oConnection.Open()

                Dim sQuery = "SELECT TOP 1 * FROM ErrorLog ORDER BY Id DESC"

                Dim iLastId = 0
                Using oSqlCmdSelect As New SqlCommand(sQuery, oConnection)
                    Using oQueryResult As SqlDataReader = oSqlCmdSelect.ExecuteReader()

                        If oQueryResult.Read() Then
                            iLastId = Integer.Parse(oQueryResult(oQueryResult.GetOrdinal("Id")))
                        End If
                    End Using
                End Using

                If iLastId >= 100000 Then
                    sQuery = "DELETE FROM ErrorLog WHERE Id <= @lastid"

                    Using oSqlCmdDelete As New SqlCommand(sQuery, oConnection)
                        oSqlCmdDelete.Parameters.AddWithValue("@lastId", iLastId - 99999)
                        Try
                            oSqlCmdDelete.ExecuteNonQuery()
                        Catch ex As Exception
                            oContext.Response.ContentType = "text/plain"
                            oContext.Response.StatusCode = 433
                            oContext.Response.Write("Incorrect Parameter")
                            ZSSOUtilities.WriteLog("ErrorLog : NOK : " & ex.Message)
                            Return
                        End Try
                    End Using
                End If


                sQuery = "INSERT INTO ErrorLog (Serial, PrinterTime, ServerTime, Level, Code, Message) VALUES (@serial, @printertime, @servertime, @level, @code, @message)"

                Using oSqlCmdInsert As New SqlCommand(sQuery, oConnection)
                    oSqlCmdInsert.Parameters.AddWithValue("@serial", sSerial.ToUpper)
                    oSqlCmdInsert.Parameters.AddWithValue("@printertime", DateTime.Parse(sTime))
                    oSqlCmdInsert.Parameters.AddWithValue("@servertime", DateTime.Now)
                    oSqlCmdInsert.Parameters.AddWithValue("@level", sLevel)
                    oSqlCmdInsert.Parameters.AddWithValue("@code", sCode)
                    oSqlCmdInsert.Parameters.AddWithValue("@message", sMessage)

                    oSqlCmdInsert.ExecuteNonQuery()
                End Using

            End Using
        End If
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class