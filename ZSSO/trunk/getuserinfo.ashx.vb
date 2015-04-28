Imports System.Web
Imports System.Web.Services
Imports System.Data.SqlClient
Imports System.Web.Script.Serialization

Public Class getuserinfo
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal oContext As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim sToken, sAccountEmail As String
        Dim oSerializer As New JavaScriptSerializer

        If ZSSOUtilities.CheckRequests(oContext.Request.UserHostAddress, "getuserinfo") > 5 Then
            oContext.Response.ContentType = "text/plain"
            oContext.Response.StatusCode = 435
            oContext.Response.Write("Too many requests")
            ZSSOUtilities.WriteLog("GetUserInfo : Too many requests")
            Return
        Else
            If oContext.Request.HttpMethod = "GET" Then
                oContext.Response.ContentType = "text/html"
                oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title>" & _
                                        "<script src=""https://code.jquery.com/jquery-1.10.2.js""></script><script type=""text/javascript"">function load_wait() { $(""#overlay"").addClass(""gray-overlay""); $("".ui-loader"").css(""display"", ""block""); }</script>" & _
                                        "<link rel=""stylesheet"" type=""text/css"" href=""style.css"">" & _
                                        "</head><body><div id=""overlay""></div><div class=""ui-loader ui-corner-all ui-body-a ui-loader-default""><span class=""ui-icon-loading""></span><h1>getuserinfo</h1></div>" & _
                                        "<form  method=""post"" action=""/getuserinfo.ashx"" accept-charset=""utf-8"">" & _
                                        "token <input id=""token"" name=""token"" type=""text"" /><br />" & _
                                        "<input id=""Submit1"" type=""submit"" value=""Ok""  onclick=""javascript: load_wait();"" /></form></body></html>")
            Else
                sToken = oContext.Request.Form("token")

                If String.IsNullOrEmpty(sToken) Then
                    oContext.Response.StatusCode = 432
                    oContext.Response.Write("Missing parameter")
                    ZSSOUtilities.WriteLog("GetUserInfo: Missing parameter")
                    Return
                End If

                sAccountEmail = ZSSOUtilities.SearchAccountEmail(sToken)

                If sAccountEmail Is Nothing Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 442
                    oContext.Response.Write("Unauthorized user")
                    ZSSOUtilities.WriteLog("GetUserInfo: Unauthorized user")
                    Return
                End If

                Using oConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
                    oConnection.Open()

                    Dim oSqlCmd As New SqlCommand("SELECT city, country, birthdate, why, what FROM Account WHERE email = @email", oConnection)

                    Try
                        oSqlCmd.Parameters.AddWithValue("@email", sAccountEmail)
                        Using oQueryResult As SqlDataReader = oSqlCmd.ExecuteReader()

                            If oQueryResult.HasRows Then
                                oQueryResult.Read()
                                oContext.Response.ContentType = "text/plain"
                                oContext.Response.Write(oSerializer.Serialize(New Dictionary(Of String, String) From {{"country", IIf(oQueryResult("country") Is System.DBNull.Value, "", oQueryResult("country"))}, _
                                                                                      {"city", IIf(oQueryResult("city") Is System.DBNull.Value, "", oQueryResult("city"))}, _
                                                                                      {"birth_date", IIf(oQueryResult("birthdate") Is System.DBNull.Value, "", oQueryResult("birthdate"))}, _
                                                                                      {"why", IIf(oQueryResult("why") Is System.DBNull.Value, "", oQueryResult("why"))}, _
                                                                                      {"what", IIf(oQueryResult("what") Is System.DBNull.Value, "", oQueryResult("what"))}}))
                                ZSSOUtilities.WriteLog("GetUserInfo: OK")
                            Else
                                ' ?
                                oContext.Response.ContentType = "text/plain"
                                oContext.Response.StatusCode = 442
                                oContext.Response.Write("Unauthorized user")
                                ZSSOUtilities.WriteLog("GetUserInfo: Unauthorized user")
                                Return
                            End If

                        End Using
                    Catch ex As Exception
                        ZSSOUtilities.WriteLog("GetUserInfo: " & ex.Message)
                    End Try
                End Using
            End If
        End If
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class