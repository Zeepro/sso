Imports System.Web
Imports System.Web.Services
Imports System.Data.SqlClient

Public Class revokeuser
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal oContext As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim sToken, sSerial, sEmail As String

        If ZSSOUtilities.CheckRequests(oContext.Request.UserHostAddress, "grantuser") > 5 Then
            oContext.Response.ContentType = "text/plain"
            oContext.Response.StatusCode = 435
            oContext.Response.Write("Too many requests")
            ZSSOUtilities.WriteLog("RevokeUser : Too many requests")
            Return
        Else
            If oContext.Request.HttpMethod = "GET" Then
                oContext.Response.ContentType = "text/html"
                oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title>" & _
                                        "<script src=""https://code.jquery.com/jquery-1.10.2.js""></script><script type=""text/javascript"">function load_wait() { $(""#overlay"").addClass(""gray-overlay""); $("".ui-loader"").css(""display"", ""block""); }</script>" & _
                                        "<link rel=""stylesheet"" type=""text/css"" href=""style.css"">" & _
                                        "</head><body><div id=""overlay""></div><div class=""ui-loader ui-corner-all ui-body-a ui-loader-default""><span class=""ui-icon-loading""></span><h1>CreateAccount</h1></div>" & _
                                        "<form  method=""post"" action=""/revokeuser.ashx"" accept-charset=""utf-8"">" & _
                                        "token <input id=""token"" name=""token"" type=""text"" /><br />" & _
                                        "serial <input id=""printersn"" name=""printersn"" type=""text"" /><br />" & _
                                        "user email <input id=""user_email"" name=""user_email"" type=""text"" /><br />" & _
                                        "<input id=""Submit1"" type=""submit"" value=""Ok""  onclick=""javascript: load_wait();"" /></form></body></html>")
            Else
                sToken = oContext.Request.Form("token")
                sSerial = oContext.Request.Form("printersn")
                sEmail = oContext.Request.Form("user_email")

                If String.IsNullOrEmpty(sToken) OrElse _
                    String.IsNullOrEmpty(sSerial) OrElse _
                    String.IsNullOrEmpty(sEmail) Then
                    oContext.Response.StatusCode = 432
                    oContext.Response.Write("Missing parameter")
                    ZSSOUtilities.WriteLog("RevokeUser: Missing parameter")
                    Return
                End If

                If sToken.Length <> 40 OrElse _
                    sSerial.Length <> 12 OrElse _
                    Not (ZSSOUtilities.oEmailRegex.IsMatch(sEmail)) Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 433
                    oContext.Response.Write("Incorrect Parameter")
                    ZSSOUtilities.WriteLog("RevokeUser: Incorrect parameter")
                    Return
                End If

                If ZSSOUtilities.SearchSerial(sSerial) = False Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 436
                    oContext.Response.Write("Unknown printer")
                    ZSSOUtilities.WriteLog("RevokeUser: Unknown printer")
                    Return
                End If

                If ZSSOUtilities.SearchAccountEmail(sToken, sSerial) Is Nothing Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 442
                    oContext.Response.Write("Unauthorized user")
                    ZSSOUtilities.WriteLog("RevokeUser: Unauthorized user")
                    Return
                End If

                Using oConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
                    oConnection.Open()
                    Using oSqlCmd As New SqlCommand("UPDATE AccountPrinterAssociation SET deleted = GETDATE() WHERE serial = @serial AND email = @email AND deleted IS NULL", oConnection)
                        oSqlCmd.Parameters.AddWithValue("@serial", sSerial)
                        oSqlCmd.Parameters.AddWithValue("@email", sEmail)
                        oSqlCmd.ExecuteNonQuery()
                    End Using
                End Using
                ZSSOUtilities.WriteLog("RevokeUser: OK")
            End If
        End If
    End Sub


    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class