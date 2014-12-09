Imports System.Web
Imports System.Web.Services
Imports System.Data.SqlClient
Imports System.Runtime.Caching

Public Class redirect
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal oContext As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim sSerial As String
        Dim oHttpCache As Caching.Cache = HttpRuntime.Cache
        Dim sRedirection As String

        If ZSSOUtilities.CheckRequests(oContext.Request.UserHostAddress, "redirect") > 5 Then
            oContext.Response.ContentType = "text/plain"
            oContext.Response.StatusCode = 435
            oContext.Response.Write("Too many requests")
            ZSSOUtilities.WriteLog("Redirect: Too many requests")
            Return
        Else
            If oContext.Request.HttpMethod = "GET" Then
                sSerial = HttpUtility.UrlDecode(oContext.Request.QueryString("sn"))

                ZSSOUtilities.WriteLog("Redirect: " & ZSSOUtilities.oSerializer.Serialize({sSerial}))

                If String.IsNullOrEmpty(sSerial) Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 432
                    oContext.Response.Write("Missing parameter")
                    ZSSOUtilities.WriteLog("Redirect: Missing parameter")
                    Return
                End If

                If ZSSOUtilities.SearchSerial(sSerial) = False Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 436
                    oContext.Response.Write("Unknown printer")
                    ZSSOUtilities.WriteLog("Redirect: Unknown printer")
                    Return
                End If

                sRedirection = oHttpCache.Get(sSerial)

                If sRedirection Is Nothing Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 440
                    oContext.Response.Write("No redirection")
                    ZSSOUtilities.WriteLog("Redirect: No redirection")
                    Return
                End If

                oContext.Response.Redirect(sRedirection)

                ZSSOUtilities.WriteLog("Redirect: OK")
            End If
        End If
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class