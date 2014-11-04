Imports System.Web
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
                oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title>" & _
                                        "<script src=""https://code.jquery.com/jquery-1.10.2.js""></script><script type=""text/javascript"">function load_wait() { console.log(""func""); $(""#overlay"").addClass(""gray-overlay""); $("".ui-loader"").css(""display"", ""block""); console.log(""done"");}</script>" & _
                                        "<link rel=""stylesheet"" type=""text/css"" href=""style.css"">" & _
                                        "</head><body><div id=""overlay""></div><div class=""ui-loader ui-corner-all ui-body-a ui-loader-default""><span class=""ui-icon-loading""></span><h1>loading</h1></div><form  method=""post"" action=""/testaccount.ashx"" accept-charset=""utf-8"">login <input id=""email"" name=""email"" type=""text"" /><br /><input id=""Submit1"" type=""submit"" value=""Ok""  onclick=""javascript: load_wait();"" /></form></body></html>")
            Else
                sEmail = HttpUtility.UrlDecode(oContext.Request.Form("email"))

                ZSSOUtilities.WriteLog("TestAccount : " & ZSSOUtilities.oSerializer.Serialize({sEmail}))

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
                ZSSOUtilities.WriteLog("TestAccount : OK : " & ZSSOUtilities.oSerializer.Serialize(arReturnValue))
            End If
        End If
    End Sub

    Public Shared Function SearchEmail(sEmail As String)
        Using oConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
            oConnection.Open()

            Dim sQuery = "SELECT TOP 1 Email " & _
                "FROM Account " & _
                "WHERE Email=@email AND Confirmed=1"

            Using oSqlCmdSelect As New SqlCommand(sQuery, oConnection)

                oSqlCmdSelect.Parameters.AddWithValue("@email", sEmail.ToLower)

                Using oQueryResult As SqlDataReader = oSqlCmdSelect.ExecuteReader()

                    If oQueryResult.HasRows Then
                        Return True
                    End If
                End Using
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