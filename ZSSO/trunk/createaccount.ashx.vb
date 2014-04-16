Imports System.Web
Imports System.Web.Services

Public Class createaccount
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim Login As String
        Dim Password As String

        If context.Request.HttpMethod = "GET" Then
            context.Response.ContentType = "text/html"
            context.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title></head><body><form  method=""post"" action=""/createaccount.ashx"" accept-charset=""utf-8"">login <input id=""login"" name=""login"" type=""text"" /><br />password <input id=""password"" name=""password"" type=""text"" /><br /><input id=""Submit1"" type=""submit"" value=""Ok"" /></form></body></html>")
        Else
            Login = context.Request.Form("login")
            Password = context.Request.Form("password")

            context.Response.ContentType = "text/plain"
            context.Response.Write("Yo!")
        End If
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class