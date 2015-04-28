Imports System.Web
Imports System.Web.Services
Imports System.Data.SqlClient

Public Class create
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal oContext As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim sToken, sName, sAccountEmail As String

        If ZSSOUtilities.CheckRequests(oContext.Request.UserHostAddress, "userlibcreate") > 5 Then
            oContext.Response.ContentType = "text/plain"
            oContext.Response.StatusCode = 435
            oContext.Response.Write("Too many requests")
            ZSSOUtilities.WriteLog("UserLib/Create: Too many requests")
            Return
        Else
            If oContext.Request.HttpMethod = "GET" Then
                oContext.Response.ContentType = "text/html"
                oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title>" & _
                                        "<script src=""https://code.jquery.com/jquery-1.10.2.js""></script><script type=""text/javascript"">function load_wait() { $(""#overlay"").addClass(""gray-overlay""); $("".ui-loader"").css(""display"", ""block""); }</script>" & _
                                        "<link rel=""stylesheet"" type=""text/css"" href=""/style.css"">" & _
                                        "</head><body><div id=""overlay""></div><div class=""ui-loader ui-corner-all ui-body-a ui-loader-default""><span class=""ui-icon-loading""></span><h1>userlib\create</h1></div>" & _
                                        "<form  method=""post"" action=""/userlib/create.ashx"" accept-charset=""utf-8"">" & _
                                        "token <input id=""token"" name=""token"" type=""text"" /><br />" & _
                                        "model name <input id=""name"" name=""name"" type=""text"" /><br />" & _
                                        "<input id=""Submit1"" type=""submit"" value=""Ok""  onclick=""javascript: load_wait();"" /></form></body></html>")
            Else
                sToken = oContext.Request.Form("token")
                sName = oContext.Request.Form("name")

                If String.IsNullOrEmpty(sToken) OrElse String.IsNullOrEmpty(sName) Then
                    oContext.Response.StatusCode = 432
                    oContext.Response.Write("Missing parameter")
                    ZSSOUtilities.WriteLog("UserLib/Create: Missing parameter")
                    Return
                End If

                If sToken.Length <> 40 Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 433
                    oContext.Response.Write("Incorrect Parameter")
                    ZSSOUtilities.WriteLog("UserLib/Create: Incorrect parameter")
                    Return
                End If

                sAccountEmail = ZSSOUtilities.SearchAccountEmail(sToken)

                If sAccountEmail Is Nothing Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 442
                    oContext.Response.Write("Unauthorized user")
                    ZSSOUtilities.WriteLog("UserLib/Create: Unauthorized user")
                    Return
                End If

                Using oConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
                    oConnection.Open()

                    Dim oSqlCmd As New SqlCommand("INSERT Model (email, name) VALUES (@email, @name)", _
                            oConnection)
                    Try
                        oSqlCmd.Parameters.AddWithValue("@email", sAccountEmail)
                        oSqlCmd.Parameters.AddWithValue("@name", sName)
                        oSqlCmd.ExecuteNonQuery()
                    Catch ex As Exception
                        If DirectCast(ex, System.Data.SqlClient.SqlException).Number = 2601 Then
                            ' Cannot insert duplicate key row in object '%.*ls' with unique index '%.*ls'.
                            ' https://msdn.microsoft.com/en-us/library/cc645728.aspx
                            oContext.Response.ContentType = "text/plain"
                            oContext.Response.StatusCode = 444
                            oContext.Response.Write("Model already exist")
                            ZSSOUtilities.WriteLog("UserLib/Create: Model already exists")
                            Return
                        Else
                            ZSSOUtilities.WriteLog("UserLib/Create: " & ex.Message)
                        End If
                    End Try
                End Using
                ZSSOUtilities.WriteLog("UserLib/Create: OK")
            End If
        End If
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class