Imports System.Web
Imports System.Web.Services
Imports System.Data.SqlClient
Imports System.Web.Script.Serialization

Public Class listuser
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal oContext As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim sToken, sSerial, sAccountEmail As String
        Dim oSerializer As New JavaScriptSerializer

        If ZSSOUtilities.CheckRequests(oContext.Request.UserHostAddress, "listuser") > 5 Then
            oContext.Response.ContentType = "text/plain"
            oContext.Response.StatusCode = 435
            oContext.Response.Write("Too many requests")
            ZSSOUtilities.WriteLog("ListUser : Too many requests")
            Return
        Else
            If oContext.Request.HttpMethod = "GET" Then
                oContext.Response.ContentType = "text/html"
                oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title>" & _
                                        "<script src=""https://code.jquery.com/jquery-1.10.2.js""></script><script type=""text/javascript"">function load_wait() { $(""#overlay"").addClass(""gray-overlay""); $("".ui-loader"").css(""display"", ""block""); }</script>" & _
                                        "<link rel=""stylesheet"" type=""text/css"" href=""style.css"">" & _
                                        "</head><body><div id=""overlay""></div><div class=""ui-loader ui-corner-all ui-body-a ui-loader-default""><span class=""ui-icon-loading""></span><h1>listuser</h1></div>" & _
                                        "<form  method=""post"" action=""/listuser.ashx"" accept-charset=""utf-8"">" & _
                                        "token <input id=""token"" name=""token"" type=""text"" /><br />" & _
                                        "serial <input id=""printersn"" name=""printersn"" type=""text"" /><br />" & _
                                        "<input id=""Submit1"" type=""submit"" value=""Ok""  onclick=""javascript: load_wait();"" /></form></body></html>")
            Else
                sToken = oContext.Request.Form("token")
                sSerial = oContext.Request.Form("printersn")

                If String.IsNullOrEmpty(sToken) OrElse _
                    String.IsNullOrEmpty(sSerial) Then
                    oContext.Response.StatusCode = 432
                    oContext.Response.Write("Missing parameter")
                    ZSSOUtilities.WriteLog("ListUser: Missing parameter")
                    Return
                End If

                If sToken.Length <> 40 OrElse _
                    sSerial.Length <> 12 Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 433
                    oContext.Response.Write("Incorrect Parameter")
                    ZSSOUtilities.WriteLog("ListUser: Incorrect parameter")
                    Return
                End If

                If Not ZSSOUtilities.SearchSerial(sSerial) Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 436
                    oContext.Response.Write("Unknown printer")
                    ZSSOUtilities.WriteLog("ListUser: Unknown printer")
                    Return
                End If

                sAccountEmail = ZSSOUtilities.SearchAccountEmail(sToken, sSerial)

                If sAccountEmail Is Nothing Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 442
                    oContext.Response.Write("Unauthorized user")
                    ZSSOUtilities.WriteLog("ListUser: Unauthorized user")
                    Return
                End If


                Using oConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
                    oConnection.Open()

                    Using oSqlCmd As New SqlCommand("SELECT email, ISNULL(name, '') AS name, ISNULL(accountrestriction, 0) AS accountrestriction, ISNULL(managerestriction, 0) AS managerestriction, ISNULL(viewrestriction, 0) AS viewrestriction " & _
                            "FROM AccountPrinterAssociation " & _
                            "WHERE deleted IS NULL AND serial = @serial", _
                            oConnection)

                        oSqlCmd.Parameters.AddWithValue("@serial", sSerial)

                        Using oQueryResult As SqlDataReader = oSqlCmd.ExecuteReader()

                            Dim aUser As New ArrayList()

                            While oQueryResult.Read
                                aUser.Add(New Dictionary(Of String, String) From {{"email", oQueryResult("email")}, _
                                                                                  {"name", oQueryResult("name")}, _
                                                                                  {"account", IIf(oQueryResult("accountrestriction") = 0, "yes", "no")}, _
                                                                                  {"manage", IIf(oQueryResult("managerestriction") = 0, "yes", "no")}, _
                                                                                  {"view", IIf(oQueryResult("viewrestriction") = 0, "yes", "no")}})
                            End While

                            oContext.Response.ContentType = "text/plain"
                            oContext.Response.Write(oSerializer.Serialize(aUser))
                            ZSSOUtilities.WriteLog("ListUser: OK")
                        End Using
                    End Using
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