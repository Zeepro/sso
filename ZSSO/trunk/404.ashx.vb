Imports System.Web

Public Class pagenotfound
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal oContext As HttpContext) Implements IHttpHandler.ProcessRequest
        oContext.Response.ContentType = "text/plain"
        oContext.Response.StatusCode = 200
        '        oContext.Response.Write("404")
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property
End Class