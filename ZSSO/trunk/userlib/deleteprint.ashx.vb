Imports System.Web
Imports System.Web.Services
Imports System.Data.SqlClient
Imports Microsoft.WindowsAzure.Storage
Imports Microsoft.WindowsAzure.Storage.Auth
Imports Microsoft.WindowsAzure.Storage.Blob

Public Class deleteprint
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal oContext As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim sToken, sId, sDate, sAccountEmail As String
        Dim nId As Integer
        Dim dDate As Nullable(Of DateTime)

        If ZSSOUtilities.CheckRequests(oContext.Request.UserHostAddress, "userlibdeleteprint") > 5 Then
            oContext.Response.ContentType = "text/plain"
            oContext.Response.StatusCode = 435
            oContext.Response.Write("Too many requests")
            ZSSOUtilities.WriteLog("UserLib/DeletePrint: Too many requests")
            Return
        Else
            If oContext.Request.HttpMethod = "GET" Then
                oContext.Response.ContentType = "text/html"
                oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title>" & _
                                        "<script src=""https://code.jquery.com/jquery-1.10.2.js""></script><script type=""text/javascript"">function load_wait() { $(""#overlay"").addClass(""gray-overlay""); $("".ui-loader"").css(""display"", ""block""); }</script>" & _
                                        "<link rel=""stylesheet"" type=""text/css"" href=""/style.css"">" & _
                                        "</head><body><div id=""overlay""></div><div class=""ui-loader ui-corner-all ui-body-a ui-loader-default""><span class=""ui-icon-loading""></span><h1>userlib\deleteprint</h1></div>" & _
                                        "<form  method=""post"" action=""/UserLib/DeletePrint.ashx"" accept-charset=""utf-8"" enctype=""multipart/form-data"">" & _
                                        "token <input id=""token"" name=""token"" type=""text"" /><br />" & _
                                        "model id <input id=""id"" name=""id"" type=""text"" /><br />" & _
                                        "date <input id=""date"" name=""date"" type=""text"" /><br />" & _
                                        "<input id=""Submit1"" type=""submit"" value=""Ok""  onclick=""javascript: load_wait();"" /></form></body></html>")
            Else
                sToken = oContext.Request.Form("token")
                sId = oContext.Request.Form("id")
                sDate = oContext.Request.Form("date")

                If String.IsNullOrEmpty(sToken) OrElse _
                    String.IsNullOrEmpty(sId) OrElse _
                    String.IsNullOrEmpty(sDate) Then
                    oContext.Response.StatusCode = 432
                    oContext.Response.Write("Missing parameter")
                    ZSSOUtilities.WriteLog("UserLib/DeletePrint: Missing parameter")
                    Return
                End If

                If sToken.Length <> 40 Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 433
                    oContext.Response.Write("Incorrect Parameter")
                    ZSSOUtilities.WriteLog("UserLib/DeletePrint: Incorrect parameter")
                    Return
                End If

                Try
                    nId = CInt(sId)
                Catch ex As Exception
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 433
                    oContext.Response.Write("Incorrect Parameter")
                    ZSSOUtilities.WriteLog("UserLib/DeletePrint: Incorrect parameter")
                    Return
                End Try

                Try
                    dDate = DateTime.Parse(sDate)
                Catch ex As Exception
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 433
                    oContext.Response.Write("Incorrect Parameter")
                    ZSSOUtilities.WriteLog("UserLib/DeletePrint: Incorrect parameter")
                    Return
                End Try

                sDate = Format(dDate, "yyyyMMddHHmmss")

                sAccountEmail = ZSSOUtilities.SearchAccountEmail(sToken)

                If sAccountEmail Is Nothing Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 442
                    oContext.Response.Write("Unauthorized user")
                    ZSSOUtilities.WriteLog("UserLib/DeletePrint: Unauthorized user")
                    Return
                End If

                Using oConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
                    oConnection.Open()

                    Dim oSqlCmd As New SqlCommand("SELECT id FROM Model WHERE email = @email AND id = @id", _
                            oConnection)
                    Try
                        oSqlCmd.Parameters.AddWithValue("@email", sAccountEmail)
                        oSqlCmd.Parameters.AddWithValue("@id", nId)
                        If oSqlCmd.ExecuteScalar() <> nId Then
                            oContext.Response.ContentType = "text/plain"
                            oContext.Response.StatusCode = 443
                            oContext.Response.Write("Model unknown")
                            ZSSOUtilities.WriteLog("UserLib/DeletePrint: Model unknown")
                            Return
                        End If

                        oSqlCmd.CommandText = "DELETE ModelPrint WHERE model = @id AND date = @date; SELECT @@ROWCOUNT"
                        oSqlCmd.Parameters.AddWithValue("@date", sDate)
                        If oSqlCmd.ExecuteScalar() <> 1 Then
                            oContext.Response.ContentType = "text/plain"
                            oContext.Response.StatusCode = 445
                            oContext.Response.Write("Print unknown")
                            ZSSOUtilities.WriteLog("UserLib/DeletePrint: Print unknown")
                            Return
                        End If
                    Catch ex As Exception
                        ZSSOUtilities.WriteLog("UserLib/DeletePrint: " & ex.Message)
                        Return
                    End Try
                End Using

                Try
                    Dim oStorageAccount As CloudStorageAccount = CloudStorageAccount.Parse(System.Configuration.ConfigurationManager.AppSettings("StorageConnectionString"))
                    Dim oBlobClient As CloudBlobClient = oStorageAccount.CreateCloudBlobClient()
                    Dim oContainer As CloudBlobContainer = oBlobClient.GetContainerReference("3dprint")

                    For Each oBlob As CloudBlockBlob In oContainer.ListBlobs(sId & "." & sDate)
                        oBlob.DeleteIfExists(DeleteSnapshotsOption.IncludeSnapshots)
                    Next
                Catch ex As Exception
                    ZSSOUtilities.WriteLog("UserLib/DeletePrint: " & ex.Message)
                    Return
                End Try
                ZSSOUtilities.WriteLog("UserLib/DeletePrint: OK")
            End If
        End If
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class