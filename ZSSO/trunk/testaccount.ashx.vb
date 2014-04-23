Imports System.Web
Imports System.Web.Services
Imports System.Text.RegularExpressions.Regex
Imports System.Web.Script.Serialization
Imports System.Data.SqlClient
Imports System.Runtime.Caching
Imports System.IO

Public Class testaccount
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim Email As String
        Dim returnValue = New Dictionary(Of String, String)
        Dim cacheMemory As ObjectCache = MemoryCache.Default

        If ZSSOUtilities.CheckRequests(context.Request.UserHostAddress) > 5 Then
            context.Response.ContentType = "text/plain"
            context.Response.StatusCode = 435
            context.Response.Write("Too many requests")
            ZSSOUtilities.WriteLog("TestAccount : Too many requests")
            Return
        Else
            If context.Request.HttpMethod = "GET" Then
                context.Response.ContentType = "text/html"
                context.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title></head><body><form  method=""post"" action=""/testaccount.ashx"" accept-charset=""utf-8"">login <input id=""email"" name=""email"" type=""text"" /><br /><input id=""Submit1"" type=""submit"" value=""Ok"" /></form></body></html>")
            Else
                Email = HttpUtility.UrlDecode(context.Request.Form("email"))

                ZSSOUtilities.WriteLog("TestAccount : " + ZSSOUtilities.oSerializer.Serialize(context.Request.Form))

                If String.IsNullOrEmpty(Email) Then
                    context.Response.ContentType = "text/plain"
                    context.Response.StatusCode = 432
                    context.Response.Write("Missing parameter")
                    ZSSOUtilities.WriteLog("TestAccount : Missing parameter")
                    Return
                End If

                If Not (ZSSOUtilities.emailExpression.IsMatch(Email)) Then
                    context.Response.ContentType = "text/plain"
                    context.Response.StatusCode = 433
                    context.Response.Write("Incorrect Parameter")
                    ZSSOUtilities.WriteLog("TestAccount : Incorrect parameter")
                    Return
                End If

                If SearchEmail(Email) Then
                    returnValue("account") = "exist"
                Else
                    returnValue("account") = "unknown"
                End If

                context.Response.ContentType = "text/plain"
                context.Response.Write(ZSSOUtilities.oSerializer.Serialize(returnValue))
            End If
        End If
        ZSSOUtilities.WriteLog("TestAccount : OK : " + ZSSOUtilities.oSerializer.Serialize(returnValue))
    End Sub

    Public Shared Function SearchEmail(Email As String)
        Using oConnexion As New SqlConnection("Data Source=(LocalDB)\v11.0;AttachDbFilename=C:\Users\ZPFr1\Desktop\zsso\ZSSO\trunk\App_Data\Database1.mdf;Integrated Security=True")
            oConnexion.Open()

            Dim QueryString = "SELECT Email " & _
                "FROM Account " & _
                "WHERE Email=@email"

            Using oSqlCmd As New SqlCommand(QueryString, oConnexion)

                oSqlCmd.Parameters.AddWithValue("@email", Email)

                Try
                    Dim QueryResult As SqlDataReader = oSqlCmd.ExecuteReader()

                    If QueryResult.HasRows Then
                        Return True
                    End If
                Catch ex As Exception
                    'context.Response.Write("Error : " + "Select commande " + ex.Message)
                    Return False
                End Try
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