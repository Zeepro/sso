Imports System.Web
Imports System.Web.Services
Imports System.Data.SqlClient
Imports System.Web.Script.Serialization

Public Class useraccess
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal oContext As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim sToken, sSerial As String
        Dim oSerializer As New JavaScriptSerializer

        If ZSSOUtilities.CheckRequests(oContext.Request.UserHostAddress, "useraccess") > 5 Then
            oContext.Response.ContentType = "text/plain"
            oContext.Response.StatusCode = 435
            oContext.Response.Write("Too many requests")
            ZSSOUtilities.WriteLog("UserAccess: Too many requests")
            Return
        Else
            If oContext.Request.HttpMethod = "GET" Then
                oContext.Response.ContentType = "text/html"
                oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title>" & _
                                        "<script src=""https://code.jquery.com/jquery-1.10.2.js""></script><script type=""text/javascript"">function load_wait() { $(""#overlay"").addClass(""gray-overlay""); $("".ui-loader"").css(""display"", ""block""); }</script>" & _
                                        "<link rel=""stylesheet"" type=""text/css"" href=""style.css"">" & _
                                        "</head><body><div id=""overlay""></div><div class=""ui-loader ui-corner-all ui-body-a ui-loader-default""><span class=""ui-icon-loading""></span><h1>CreateAccount</h1></div>" & _
                                        "<form  method=""post"" action=""/useraccess.ashx"" accept-charset=""utf-8"">" & _
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
                    ZSSOUtilities.WriteLog("UserAccess: Missing parameter")
                    Return
                End If

                If sToken.Length <> 40 OrElse _
                    sSerial.Length <> 12 Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 433
                    oContext.Response.Write("Incorrect Parameter")
                    ZSSOUtilities.WriteLog("GrantUser: Incorrect parameter")
                    Return
                End If

                If Not ZSSOUtilities.SearchSerial(sSerial) Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 436
                    oContext.Response.Write("Unknown printer")
                    ZSSOUtilities.WriteLog("GrantUser: Unknown printer")
                    Return
                End If

                Using oConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
                    oConnection.Open()

                    Dim sQuery = "DELETE TokenId WHERE date < DATEADD(day, -1, GETDATE());" & _
                        "SELECT TOP 1 AccountPrinterAssociation.email, ISNULL(AccountPrinterAssociation.name, '') AS name, ISNULL(AccountPrinterAssociation.accountrestriction, 0) AS accountrestriction, ISNULL(AccountPrinterAssociation.managerestriction, 0) AS managerestriction, ISNULL(AccountPrinterAssociation.viewrestriction, 0) AS viewrestriction " & _
                        "FROM TokenId " & _
                        "INNER JOIN AccountPrinterAssociation " & _
                        "ON TokenId.email = AccountPrinterAssociation.email " & _
                        "WHERE TokenId.token = @token AND AccountPrinterAssociation.serial = @serial"

                    Using oSqlCmdSelect As New SqlCommand(sQuery, oConnection)
                        oSqlCmdSelect.Parameters.AddWithValue("@token", sToken)
                        oSqlCmdSelect.Parameters.AddWithValue("@serial", sSerial)

                        Using oQueryResult As SqlDataReader = oSqlCmdSelect.ExecuteReader()
                            If oQueryResult.Read() Then
                                oContext.Response.ContentType = "text/plain"

                                oContext.Response.Write(oSerializer.Serialize(New Dictionary(Of String, String) From _
                                                                              {{"email", oQueryResult("email")}, _
                                                                               {"name", oQueryResult("name")}, _
                                                                               {"account", IIf(oQueryResult("accountrestriction") = 0, "yes", "no")}, _
                                                                               {"manage", IIf(oQueryResult("managerestriction") = 0, "yes", "no")}, _
                                                                               {"view", IIf(oQueryResult("viewrestriction") = 0, "yes", "no")}}))
                                ZSSOUtilities.WriteLog("UserAccess: OK")
                            Else
                                oContext.Response.ContentType = "text/plain"
                                oContext.Response.StatusCode = 442
                                oContext.Response.Write("Unauthorized user")
                                ZSSOUtilities.WriteLog("UserAccess: Unauthorized user")
                            End If
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