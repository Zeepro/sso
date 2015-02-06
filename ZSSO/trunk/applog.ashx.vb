Imports System.Web
Imports System.Web.Services
Imports System.Data.SqlClient

Public Class applog
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal oContext As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim sApp As String
        Dim sTime As String
        Dim sSerial As String
        Dim sLevel As String
        Dim sMessage As String
        Dim sDetails As String
        Dim sTemplate As String = ""
        Dim sSubject As String = ""

        If oContext.Request.HttpMethod = "GET" Then
            oContext.Response.ContentType = "text/html"
            oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title>" & _
                                    "<script src=""https://code.jquery.com/jquery-1.10.2.js""></script><script type=""text/javascript"">function load_wait() { $(""#overlay"").addClass(""gray-overlay""); $("".ui-loader"").css(""display"", ""block""); }</script>" & _
                                    "<link rel=""stylesheet"" type=""text/css"" href=""style.css"">" & _
                                    "</head><body><div id=""overlay""></div><div class=""ui-loader ui-corner-all ui-body-a ui-loader-default""><span class=""ui-icon-loading""></span><h1>loading</h1></div>" & _
                                    "<form  method=""post"" action=""/applog.ashx"" accept-charset=""utf-8"">" & _
                                    "appname <input id=""appname"" name=""appname"" type=""text"" /><br />" & _
                                    "apptime<input id=""apptime"" name=""apptime"" type=""text"" /><br />" & _
                                    "printersn <input id=""printersn"" name=""printersn"" type=""text"" /><br />" & _
                                    "level <input id=""level"" name=""level"" type=""text"" /><br />" & _
                                    "message <input id=""message"" name=""message"" type=""text"" /><br />" & _
                                    "details <input id=""details"" name=""details"" type=""text"" /><br />" & _
                                    "<input id=""Submit1"" type=""submit"" value=""Ok""  onclick=""javascript: load_wait();"" /></form></body></html>")
        Else
            Try
                sApp = oContext.Request.Form("appname")
                sTime = oContext.Request.Form("apptime")
                sSerial = oContext.Request.Form("printersn")
                sLevel = oContext.Request.Form("level")
                sMessage = oContext.Request.Form("message")
                sDetails = oContext.Request.Form("details")

                If String.IsNullOrEmpty(sApp) Or String.IsNullOrEmpty(sTime) _
                    Or String.IsNullOrEmpty(sLevel) Then
                    oContext.Response.StatusCode = 432
                    oContext.Response.Write("Missing parameter")
                    ZSSOUtilities.WriteLog("ErrorLog : Missing parameter")
                    Return
                End If

                Using oConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
                    oConnection.Open()

                    Dim sQuery = "SELECT TOP 1 * FROM AppLog ORDER BY Id DESC"

                    Dim iLastId = 0
                    Using oSqlCmdSelect As New SqlCommand(sQuery, oConnection)
                        Using oQueryResult As SqlDataReader = oSqlCmdSelect.ExecuteReader()

                            If oQueryResult.Read() Then
                                iLastId = Integer.Parse(oQueryResult(oQueryResult.GetOrdinal("Id")))
                            End If
                        End Using
                    End Using

                    If iLastId >= 100000 Then
                        sQuery = "DELETE FROM AppLog WHERE Id <= @lastid"

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


                    sQuery = "INSERT INTO AppLog (appname, apptime, serial, level, message, details) VALUES (@appname, @apptime, @serial, @level, @message, @details)"

                    Using oSqlCmdInsert As New SqlCommand(sQuery, oConnection)
                        oSqlCmdInsert.Parameters.AddWithValue("@appname", sApp)
                        oSqlCmdInsert.Parameters.AddWithValue("@apptime", DateTime.Parse(sTime))
                        oSqlCmdInsert.Parameters.AddWithValue("@serial", sSerial.ToUpper)
                        oSqlCmdInsert.Parameters.AddWithValue("@level", sLevel)
                        oSqlCmdInsert.Parameters.AddWithValue("@message", sMessage)
                        oSqlCmdInsert.Parameters.AddWithValue("@details", sDetails)

                        oSqlCmdInsert.ExecuteNonQuery()
                    End Using

                End Using
            Catch ex As Exception
                oContext.Response.ContentType = "text/plain"
                oContext.Response.StatusCode = 433
                oContext.Response.Write("Incorrect Parameter")
                ZSSOUtilities.WriteLog("ErrorLog : NOK : " & ex.Message)
            End Try

        End If
    End Sub
    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class