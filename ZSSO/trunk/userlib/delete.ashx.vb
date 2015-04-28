Imports System.Web
Imports System.Web.Services
Imports System.Data.SqlClient
Imports Microsoft.WindowsAzure.Storage
Imports Microsoft.WindowsAzure.Storage.Auth
Imports Microsoft.WindowsAzure.Storage.Blob

Public Class delete
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal oContext As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim sToken, sId, sAccountEmail As String
        Dim nId As Integer

        If ZSSOUtilities.CheckRequests(oContext.Request.UserHostAddress, "userlibdelete") > 5 Then
            oContext.Response.ContentType = "text/plain"
            oContext.Response.StatusCode = 435
            oContext.Response.Write("Too many requests")
            ZSSOUtilities.WriteLog("UserLib/Delete: Too many requests")
            Return
        Else
            If oContext.Request.HttpMethod = "GET" Then
                oContext.Response.ContentType = "text/html"
                oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title>" & _
                                        "<script src=""https://code.jquery.com/jquery-1.10.2.js""></script><script type=""text/javascript"">function load_wait() { $(""#overlay"").addClass(""gray-overlay""); $("".ui-loader"").css(""display"", ""block""); }</script>" & _
                                        "<link rel=""stylesheet"" type=""text/css"" href=""/style.css"">" & _
                                        "</head><body><div id=""overlay""></div><div class=""ui-loader ui-corner-all ui-body-a ui-loader-default""><span class=""ui-icon-loading""></span><h1>userlib\delete</h1></div>" & _
                                        "<form  method=""post"" action=""/UserLib/Delete.ashx"" accept-charset=""utf-8"" enctype=""multipart/form-data"">" & _
                                        "token <input id=""token"" name=""token"" type=""text"" /><br />" & _
                                        "model id <input id=""id"" name=""id"" type=""text"" /><br />" & _
                                        "<input id=""Submit1"" type=""submit"" value=""Ok""  onclick=""javascript: load_wait();"" /></form></body></html>")
            Else
                sToken = oContext.Request.Form("token")
                sId = oContext.Request.Form("id")

                If String.IsNullOrEmpty(sToken) OrElse _
                    String.IsNullOrEmpty(sId) Then
                    oContext.Response.StatusCode = 432
                    oContext.Response.Write("Missing parameter")
                    ZSSOUtilities.WriteLog("UserLib/Delete: Missing parameter")
                    Return
                End If

                If sToken.Length <> 40 Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 433
                    oContext.Response.Write("Incorrect Parameter")
                    ZSSOUtilities.WriteLog("UserLib/Delete: Incorrect parameter")
                    Return
                End If

                Try
                    nId = CInt(sId)
                Catch ex As Exception
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 433
                    oContext.Response.Write("Incorrect Parameter")
                    ZSSOUtilities.WriteLog("UserLib/Delete: Incorrect parameter")
                    Return
                End Try

                sAccountEmail = ZSSOUtilities.SearchAccountEmail(sToken)

                If sAccountEmail Is Nothing Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 442
                    oContext.Response.Write("Unauthorized user")
                    ZSSOUtilities.WriteLog("UserLib/Delete: Unauthorized user")
                    Return
                End If

                Using oConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
                    oConnection.Open()

                    Dim oSqlCmd As New SqlCommand("DELETE Model WHERE email = @email AND id = @id; SELECT @@ROWCOUNT", _
                            oConnection)
                    Try
                        oSqlCmd.Parameters.AddWithValue("@email", sAccountEmail)
                        oSqlCmd.Parameters.AddWithValue("@id", nId)
                        If oSqlCmd.ExecuteScalar() <> 1 Then
                            oContext.Response.ContentType = "text/plain"
                            oContext.Response.StatusCode = 443
                            oContext.Response.Write("Model unknown")
                            ZSSOUtilities.WriteLog("UserLib/Delete: Model unknown")
                            Return
                        End If

                        oSqlCmd.CommandText = "DELETE Model3DFile WHERE model = @id; DELETE ModelPrint WHERE model = @id"
                        oSqlCmd.ExecuteNonQuery()
                    Catch ex As Exception
                        ZSSOUtilities.WriteLog("UserLib/Delete: " & ex.Message)
                        Return
                    End Try
                End Using

                Try
                    Dim oStorageAccount As CloudStorageAccount = CloudStorageAccount.Parse(System.Configuration.ConfigurationManager.AppSettings("StorageConnectionString"))
                    Dim oBlobClient As CloudBlobClient = oStorageAccount.CreateCloudBlobClient()

                    Dim oContainer As CloudBlobContainer = oBlobClient.GetContainerReference("3dmodel")

                    For Each oBlob As CloudBlockBlob In oContainer.ListBlobs(sId & ".")
                        oBlob.DeleteIfExists(DeleteSnapshotsOption.IncludeSnapshots)
                    Next

                    oContainer = oBlobClient.GetContainerReference("3dprint")

                    For Each oBlob As CloudBlockBlob In oContainer.ListBlobs(sId & ".")
                        oBlob.DeleteIfExists(DeleteSnapshotsOption.IncludeSnapshots)
                    Next
                Catch ex As Exception
                    ZSSOUtilities.WriteLog("UserLib/Delete: " & ex.Message)
                    Return
                End Try
                ZSSOUtilities.WriteLog("UserLib/Delete: OK")
            End If
        End If
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class