Imports System.Web
Imports System.Web.Services
Imports System.Text.RegularExpressions.Regex
Imports System.Web.Script.Serialization
Imports System.Data.SqlClient
Imports System.Runtime.Caching
Imports System.IO

Public Class sendmail
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal oContext As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim sEmail As String
        Dim sPassword As String
        Dim sAddress As String
        Dim sSubject As String
        Dim sHtmlbody As String
        Dim iCachedCounterByIp As Int32
        Dim oHttpCache As Caching.Cache = HttpRuntime.Cache

        If oContext.Request.HttpMethod = "GET" Then
            oContext.Response.ContentType = "text/html"
            oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title></head><body><form  method=""post"" action=""/sendmail.ashx"" accept-charset=""utf-8"">login <input id=""email"" name=""email"" type=""text"" /><br />password <input id=""password"" name=""password"" type=""text"" /><br />address <input id=""address"" name=""address"" type=""text"" /><br />subject <input id=""subject"" name=""subject"" type=""text"" /><br />htmlbody <input id=""htmlbody"" name=""htmlbody"" type=""text"" /><br /><input id=""Submit1"" type=""submit"" value=""Ok"" /></form></body></html>")
        Else
            Try
                iCachedCounterByIp = CInt(oHttpCache("sendmail_" & oContext.Request.UserHostAddress))
            Catch
                iCachedCounterByIp = 0
            End Try
            If iCachedCounterByIp > 3 Then
                oContext.Response.ContentType = "text/plain"
                oContext.Response.StatusCode = 435
                oContext.Response.Write("Too many requests")
                ZSSOUtilities.WriteLog("SendMail : Too many requests")
                Return
            End If

            iCachedCounterByIp = iCachedCounterByIp + 1
            oHttpCache.Insert("sendmail_" & oContext.Request.UserHostAddress, iCachedCounterByIp, Nothing, DateTime.Now.AddMinutes(3.0), TimeSpan.Zero)

            sEmail = HttpUtility.UrlDecode(oContext.Request.Form("email"))
            sPassword = HttpUtility.UrlDecode(oContext.Request.Form("password"))
            sAddress = HttpUtility.UrlDecode(oContext.Request.Form("address"))
            sSubject = HttpUtility.UrlDecode(oContext.Request.Form("subject"))
            sHtmlbody = HttpUtility.UrlDecode(oContext.Request.Form("htmlbody"))

            ZSSOUtilities.WriteLog("SendMail : " & ZSSOUtilities.oSerializer.Serialize({sEmail, sAddress, sSubject, sHtmlbody}))

            If String.IsNullOrEmpty(sEmail) Or String.IsNullOrEmpty(sPassword) _
                Or String.IsNullOrEmpty(sAddress) Or String.IsNullOrEmpty(sSubject) _
                Or String.IsNullOrEmpty(sHtmlbody) Then
                oContext.Response.ContentType = "text/plain"
                oContext.Response.StatusCode = 432
                oContext.Response.Write("Missing parameter")
                ZSSOUtilities.WriteLog("SendMail : Missing parameter")
                Return
            End If

            If Not (ZSSOUtilities.oEmailRegex.IsMatch(sEmail)) Then
                oContext.Response.ContentType = "text/plain"
                oContext.Response.StatusCode = 433
                oContext.Response.Write("Incorrect Parameter")
                ZSSOUtilities.WriteLog("SendMail : Incorrect parameter")
                Return
            End If
            Using oConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
                oConnection.Open()

                If Not ZSSOUtilities.Login(oConnection, sEmail, sPassword) Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 434
                    oContext.Response.Write("Login failed")
                    ZSSOUtilities.WriteLog("SendMail : Login failed")
                    Return
                End If
            End Using

            Try
                Dim oHtmlEmail As New Mail
                oHtmlEmail.sReceiver = sAddress
                oHtmlEmail.sSubject = sSubject
                oHtmlEmail.sBody = sHtmlbody
                oHtmlEmail.Send()
            Catch ex As Exception
                ZSSOUtilities.WriteLog("SendMail : NOK : " & ex.Message)
                Return
            End Try
            ZSSOUtilities.WriteLog("SendMail : OK")
        End If
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class