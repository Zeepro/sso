Imports System.Web
Imports System.Web.Services
Imports System.Data.SqlClient

Public Class setuserinfo
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal oContext As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim sToken, sCountry, sCity, sBirth_date, sWhy, sWhat, sAccountEmail As String
        Dim dBirth_date As Nullable(Of DateTime)

        If ZSSOUtilities.CheckRequests(oContext.Request.UserHostAddress, "setuserinfo") > 5 Then
            oContext.Response.ContentType = "text/plain"
            oContext.Response.StatusCode = 435
            oContext.Response.Write("Too many requests")
            ZSSOUtilities.WriteLog("SetUserInfo : Too many requests")
            Return
        Else
            If oContext.Request.HttpMethod = "GET" Then
                oContext.Response.ContentType = "text/html"
                oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title>" & _
                                        "<script src=""https://code.jquery.com/jquery-1.10.2.js""></script><script type=""text/javascript"">function load_wait() { $(""#overlay"").addClass(""gray-overlay""); $("".ui-loader"").css(""display"", ""block""); }</script>" & _
                                        "<link rel=""stylesheet"" type=""text/css"" href=""style.css"">" & _
                                        "</head><body><div id=""overlay""></div><div class=""ui-loader ui-corner-all ui-body-a ui-loader-default""><span class=""ui-icon-loading""></span><h1>setuserinfo</h1></div>" & _
                                        "<form  method=""post"" action=""/setuserinfo.ashx"" accept-charset=""utf-8"">" & _
                                        "token <input id=""token"" name=""token"" type=""text"" /><br />" & _
                                        "country <input id=""country"" name=""country"" type=""text"" /><br />" & _
                                        "city <input id=""city"" name=""city"" type=""text"" /><br />" & _
                                        "birth date <input id=""birth_date"" name=""birth_date"" type=""text"" /><br />" & _
                                        "why <input id=""why"" name=""why"" type=""text"" /><br />" & _
                                        "what <input id=""what"" name=""what"" type=""text"" /><br />" & _
                                        "<input id=""Submit1"" type=""submit"" value=""Ok""  onclick=""javascript: load_wait();"" /></form></body></html>")
            Else
                sToken = oContext.Request.Form("token")
                sCountry = oContext.Request.Form("country")
                sCity = oContext.Request.Form("city")
                sBirth_date = oContext.Request.Form("birth_date")
                sWhy = oContext.Request.Form("why")
                sWhat = oContext.Request.Form("what")

                If String.IsNullOrEmpty(sToken)Then
                    oContext.Response.StatusCode = 432
                    oContext.Response.Write("Missing parameter")
                    ZSSOUtilities.WriteLog("SetUserInfo: Missing parameter")
                    Return
                End If

                If sCountry Is Nothing Then
                    sCountry = ""
                End If

                If sCity Is Nothing Then
                    sCity = ""
                End If

                If sCity.Length > 250 OrElse sCountry.Length > 250 Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 433
                    oContext.Response.Write("Incorrect Parameter")
                    ZSSOUtilities.WriteLog("SetUserInfo: Incorrect parameter")
                    Return
                End If

                If String.IsNullOrEmpty(sBirth_date) Then
                    dBirth_date = Nothing
                Else
                    Try
                        dBirth_date = DateTime.Parse(sBirth_date)
                    Catch ex As Exception
                        oContext.Response.ContentType = "text/plain"
                        oContext.Response.StatusCode = 433
                        oContext.Response.Write("Incorrect Parameter")
                        ZSSOUtilities.WriteLog("SetUserInfo: Incorrect parameter")
                        Return
                    End Try
                End If

                If sWhy Is Nothing Then
                    sWhy = ""
                End If

                If sWhat Is Nothing Then
                    sWhat = ""
                End If

                sAccountEmail = ZSSOUtilities.SearchAccountEmail(sToken)

                If sAccountEmail Is Nothing Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 442
                    oContext.Response.Write("Unauthorized user")
                    ZSSOUtilities.WriteLog("SetUserInfo: Unauthorized user")
                    Return
                End If

                Using oConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
                    oConnection.Open()

                    Dim oSqlCmd As New SqlCommand("UPDATE Account " & _
                            "SET city = @city, " & _
                            "country = @country, " & _
                            "birthdate = @birthdate, " & _
                            "why = @why, " & _
                            "what = @what " & _
                            "WHERE email = @email", _
                            oConnection)

                    Try
                        oSqlCmd.Parameters.AddWithValue("@city", sCity)
                        oSqlCmd.Parameters.AddWithValue("@country", sCountry)
                        If dBirth_date Is Nothing Then
                            oSqlCmd.Parameters.AddWithValue("@birthdate", System.DBNull.Value)
                        Else
                            oSqlCmd.Parameters.AddWithValue("@birthdate", dBirth_date)
                        End If
                        oSqlCmd.Parameters.AddWithValue("@why", sWhy)
                        oSqlCmd.Parameters.AddWithValue("@what", sWhat)
                        oSqlCmd.Parameters.AddWithValue("@email", sAccountEmail)
                        oSqlCmd.ExecuteNonQuery()
                    Catch ex As Exception
                        ZSSOUtilities.WriteLog("SetUserInfo: " & ex.Message)
                    End Try
                End Using
                ZSSOUtilities.WriteLog("SetUserInfo: OK")
            End If
        End If
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class