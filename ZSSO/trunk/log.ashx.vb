Imports System.Web
Imports System.Web.Services
Imports System.Text.RegularExpressions.Regex
Imports System.Data.SqlClient
Imports System.Security.Cryptography
Imports System.Runtime.Caching
Imports System.IO

Public Class log
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal oContext As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim sSerial As String
        Dim sVersion As String
        Dim sCategory As String
        Dim sAction As String
        Dim sLabel As String
        Dim sValue As String
        Dim sNonInteraction As String
        Dim sTemplate As String = ""
        Dim sSubject As String = ""

        If oContext.Request.HttpMethod = "GET" Then
            oContext.Response.ContentType = "text/html"
            oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title></head><body><form  method=""post"" action=""/log.ashx"" accept-charset=""utf-8"">Serial <input id=""printersn"" name=""printersn"" type=""text"" /><br />version <input id=""version"" name=""version"" type=""text"" /><br />category <input id=""category"" name=""category"" type=""text"" /><br />action <input id=""action"" name=""action"" type=""text"" /><br />label <input id=""label"" name=""label"" type=""text"" /><br />value <input id=""value"" name=""value"" type=""text"" /><br />non-interaction <input id=""non-interaction"" name=""non-interaction"" type=""text"" /><br /><input id=""Submit1"" type=""submit"" value=""Ok"" /></form></body></html>")
        Else
            sSerial = HttpUtility.UrlDecode(oContext.Request.Form("printersn"))
            sVersion = HttpUtility.UrlDecode(oContext.Request.Form("version"))
            sCategory = HttpUtility.UrlDecode(oContext.Request.Form("category"))
            sAction = HttpUtility.UrlDecode(oContext.Request.Form("action"))
            sLabel = HttpUtility.UrlDecode(oContext.Request.Form("label"))
            sValue = HttpUtility.UrlDecode(oContext.Request.Form("value"))
            sNonInteraction = HttpUtility.UrlDecode(oContext.Request.Form("non-interaction"))

            ZSSOUtilities.WriteLog("Log : " & ZSSOUtilities.oSerializer.Serialize({sSerial, sVersion, sCategory, sAction, sLabel, sValue, sNonInteraction}))
            If String.IsNullOrEmpty(sSerial) Or String.IsNullOrEmpty(sVersion) _
                Or String.IsNullOrEmpty(sCategory) Or String.IsNullOrEmpty(sAction) Then
                oContext.Response.StatusCode = 432
                oContext.Response.Write("Missing parameter")
                ZSSOUtilities.WriteLog("Log : Missing parameter")
                Return
            End If

            Try
                If Not String.IsNullOrEmpty(sValue) Then
                    Dim iValue As Integer = Integer.Parse(sValue)
                End If
                If Not String.IsNullOrEmpty(sNonInteraction) Then
                    Dim bNonInteraction As Boolean = Boolean.Parse(sNonInteraction)
                End If
                If Not ZSSOUtilities.SearchSerial(sSerial) Then
                    Throw New Exception
                End If
            Catch ex As Exception
                oContext.Response.ContentType = "text/plain"
                oContext.Response.StatusCode = 433
                oContext.Response.Write("Incorrect Parameter")
                ZSSOUtilities.WriteLog("Log : Incorrect parameter")
                Return
            End Try

            Using oConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
                oConnection.Open()

                Dim sQuery = "SELECT TOP 1 * FROM StatsLog ORDER BY Id DESC"

                Dim iLastId = 0
                Using oSqlCmdSelect As New SqlCommand(sQuery, oConnection)
                    Using oQueryResult As SqlDataReader = oSqlCmdSelect.ExecuteReader()

                        If oQueryResult.Read() Then
                            iLastId = Integer.Parse(oQueryResult(oQueryResult.GetOrdinal("Id")))
                        End If
                    End Using
                End Using

                If iLastId >= 1000000 Then
                    sQuery = "DELETE FROM StatsLog WHERE Id <= @lastid"

                    Using oSqlCmdDelete As New SqlCommand(sQuery, oConnection)
                        oSqlCmdDelete.Parameters.AddWithValue("@lastId", iLastId - 999999)
                        Try
                            oSqlCmdDelete.ExecuteNonQuery()
                        Catch ex As Exception
                            oContext.Response.ContentType = "text/plain"
                            oContext.Response.StatusCode = 433
                            oContext.Response.Write("Incorrect Parameter")
                            ZSSOUtilities.WriteLog("Log : NOK : " & ex.Message)
                            Return
                        End Try
                    End Using
                End If


                sQuery = "INSERT INTO StatsLog ([Serial], [ServerTime], [Version], [Category], [Action], [Label], [Value], [NonInteraction]) VALUES (@serial, @servertime, @version, @category, @action, @label, @value, @noninteraction)"

                Using oSqlCmdInsert As New SqlCommand(sQuery, oConnection)

                    oSqlCmdInsert.Parameters.AddWithValue("@serial", sSerial)
                    oSqlCmdInsert.Parameters.AddWithValue("@servertime", DateTime.Now)
                    oSqlCmdInsert.Parameters.AddWithValue("@version", sVersion)
                    oSqlCmdInsert.Parameters.AddWithValue("@category", sCategory)
                    oSqlCmdInsert.Parameters.AddWithValue("@action", sAction)
                    oSqlCmdInsert.Parameters.AddWithValue("@label", sLabel)
                    oSqlCmdInsert.Parameters.AddWithValue("@value", sValue)
                    oSqlCmdInsert.Parameters.AddWithValue("@noninteraction", sNonInteraction)

                    Try
                        oSqlCmdInsert.ExecuteNonQuery()
                    Catch ex As Exception
                        oContext.Response.ContentType = "text/plain"
                        oContext.Response.StatusCode = 433
                        oContext.Response.Write("Incorrect Parameter")
                        ZSSOUtilities.WriteLog("Log : NOK : " & ex.Message)
                        Return
                    End Try
                End Using

            End Using

            ZSSOUtilities.WriteLog("Log : OK")
        End If
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class