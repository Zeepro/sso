Imports System.Web
Imports System.Web.Services
Imports System.Data.SqlClient
Imports System.Threading
Imports System.IO

Public Class addprint1
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal oContext As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim sToken, sId, sDate, sAccountEmail As String
        Dim nId As Integer
        Dim dDate As Nullable(Of DateTime)
        Dim cFile As HttpFileCollection

        If ZSSOUtilities.CheckRequests(oContext.Request.UserHostAddress, "userlibaddprint") > 5 Then
            oContext.Response.ContentType = "text/plain"
            oContext.Response.StatusCode = 435
            oContext.Response.Write("Too many requests")
            ZSSOUtilities.WriteLog("UserLib/AddPrint: Too many requests")
            Return
        Else
            If oContext.Request.HttpMethod = "GET" Then
                oContext.Response.ContentType = "text/html"
                oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title>" & _
                                        "<script src=""https://code.jquery.com/jquery-1.10.2.js""></script><script type=""text/javascript"">function load_wait() { $(""#overlay"").addClass(""gray-overlay""); $("".ui-loader"").css(""display"", ""block""); }</script>" & _
                                        "<link rel=""stylesheet"" type=""text/css"" href=""/style.css"">" & _
                                        "</head><body><div id=""overlay""></div><div class=""ui-loader ui-corner-all ui-body-a ui-loader-default""><span class=""ui-icon-loading""></span><h1>userlib\addprint</h1></div>" & _
                                        "<form  method=""post"" action=""/UserLib/AddPrint.ashx"" accept-charset=""utf-8"" enctype=""multipart/form-data"">" & _
                                        "token <input id=""token"" name=""token"" type=""text"" /><br />" & _
                                        "model id <input id=""id"" name=""id"" type=""text"" /><br />" & _
                                        "date <input id=""date"" name=""date"" type=""text"" /><br />" & _
                                        "archive <input id=""archive"" name=""archive"" type=""file"" /><br />" & _
                                        "<input id=""Submit1"" type=""submit"" value=""Ok""  onclick=""javascript: load_wait();"" /></form></body></html>")
            Else
                sToken = oContext.Request.Form("token")
                sId = oContext.Request.Form("id")
                sDate = oContext.Request.Form("date")
                cFile = oContext.Request.Files()

                If String.IsNullOrEmpty(sToken) OrElse _
                    String.IsNullOrEmpty(sId) OrElse _
                    String.IsNullOrEmpty(sDate) Then
                    oContext.Response.StatusCode = 432
                    oContext.Response.Write("Missing parameter")
                    ZSSOUtilities.WriteLog("UserLib/AddPrint: Missing parameter")
                    Return
                End If

                If sToken.Length <> 40 Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 433
                    oContext.Response.Write("Incorrect Parameter")
                    ZSSOUtilities.WriteLog("UserLib/AddPrint: Incorrect parameter")
                    Return
                End If

                Try
                    nId = CInt(sId)
                Catch ex As Exception
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 433
                    oContext.Response.Write("Incorrect Parameter")
                    ZSSOUtilities.WriteLog("UserLib/AddPrint: Incorrect parameter")
                    Return
                End Try

                Try
                    dDate = DateTime.Parse(sDate)
                Catch ex As Exception
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 433
                    oContext.Response.Write("Incorrect Parameter")
                    ZSSOUtilities.WriteLog("UserLib/AddPrint: Incorrect parameter")
                    Return
                End Try

                sDate = Format(dDate, "yyyyMMddHHmmss")

                sAccountEmail = ZSSOUtilities.SearchAccountEmail(sToken)

                If sAccountEmail Is Nothing Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 442
                    oContext.Response.Write("Unauthorized user")
                    ZSSOUtilities.WriteLog("UserLib/AddPrint: Unauthorized user")
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
                            oContext.Response.Write("Unknown model")
                            ZSSOUtilities.WriteLog("UserLib/AddPrint: Unknown model")
                            Return
                        End If

                        oSqlCmd.CommandText = "DELETE ModelPrint WHERE model = @id AND date = @date; INSERT ModelPrint (model, date, gcode, [time-lapse], img) VALUES (@id, @date, 'uploading', 'uploading', 'uploading')"
                        oSqlCmd.Parameters.AddWithValue("@date", sDate)
                        oSqlCmd.ExecuteNonQuery()
                    Catch ex As Exception
                        ZSSOUtilities.WriteLog("UserLib/AddPrint: " & ex.Message)
                        Return
                    End Try
                End Using

                Try
                    If Not Directory.Exists(oContext.Server.MapPath("~\tmp")) Then
                        ' First use (on this VM)
                        Directory.CreateDirectory(oContext.Server.MapPath("~\tmp"))
                    End If
                    ' Chunks cleaning
                    For Each oFile As FileInfo In (New DirectoryInfo(oContext.Server.MapPath("~\tmp"))).GetFiles()
                        If (Now - oFile.CreationTime).Minutes > 2 Then
                            oFile.Delete()
                        End If
                    Next
                    ' Temporary folders cleaning
                    For Each oDirectory As DirectoryInfo In (New DirectoryInfo(oContext.Server.MapPath("~\tmp"))).GetDirectories()
                        If (Now - oDirectory.CreationTime).Minutes > 5 Then
                            oDirectory.Delete(True)
                        End If
                    Next

                    Dim oMatch As Text.RegularExpressions.Match = ZSSOUtilities.rPartArchive.Match(cFile(0).FileName)

                    If oMatch.Success Then
                        ' Part archive
                        Dim nPart As Integer = CInt(oMatch.Groups(2).Value)
                        Dim nNumberPart As Integer = CInt(oMatch.Groups(3).Value)
                        If nPart > 80 Then
                            oContext.Response.ContentType = "text/plain"
                            oContext.Response.StatusCode = 450
                            oContext.Response.Write("Incorrect parameter")
                            ZSSOUtilities.WriteLog("UserLib/AddPrint: Incorrect parameter (" & oContext.Request.UserHostAddress & ")")
                            Return
                        End If
                        If nPart > nNumberPart Then
                            ' Bad number
                            oContext.Response.ContentType = "text/plain"
                            oContext.Response.StatusCode = 433
                            oContext.Response.Write("Incorrect parameter")
                            ZSSOUtilities.WriteLog("UserLib/AddPrint: Incorrect parameter (" & oContext.Request.UserHostAddress & ")")
                            Return
                        Else
                            ' Other chunk 'update'
                            For Each sFile As String In Directory.GetFiles(oContext.Server.MapPath("~\tmp"), oContext.Request.UserHostAddress.Replace(":", "") & "." & sId & "." & sDate & ".*")
                                File.SetCreationTime(sFile, Now)
                            Next
                            ' Chunk save
                            cFile(0).SaveAs(oContext.Server.MapPath("~\tmp\") & oContext.Request.UserHostAddress.Replace(":", "") & "." & sId & "." & sDate & "." & cFile(0).FileName)

                            ' Test if complete
                            For nTmp As Integer = 1 To nNumberPart
                                If Not File.Exists(oContext.Server.MapPath("~\tmp\") & oContext.Request.UserHostAddress.Replace(":", "") & "." & sId & "." & sDate & "." & oMatch.Groups(1).Value & ".zip." & nTmp.ToString() & "." & oMatch.Groups(3).Value) Then
                                    oContext.Response.ContentType = "text/plain"
                                    oContext.Response.StatusCode = 200
                                    oContext.Response.Write("partial")
                                    Return
                                End If
                            Next

                            ' Assemble
                            Using oWrite As FileStream = File.Open(oContext.Server.MapPath("~\tmp\") & oContext.Request.UserHostAddress.Replace(":", "") & "." & sId & "." & sDate & ".zip", FileMode.Create)
                                For nTmp As Integer = 1 To nNumberPart
                                    Using oRead As FileStream = File.OpenRead(oContext.Server.MapPath("~\tmp\") & oContext.Request.UserHostAddress.Replace(":", "") & "." & sId & "." & sDate & "." & oMatch.Groups(1).Value & ".zip." & nTmp.ToString() & "." & oMatch.Groups(3).Value)
                                        oRead.CopyTo(oWrite)
                                    End Using
                                    File.Delete(oContext.Server.MapPath("~\tmp\") & oContext.Request.UserHostAddress.Replace(":", "") & "." & sId & "." & sDate & "." & oMatch.Groups(1).Value & ".zip." & nTmp.ToString() & "." & oMatch.Groups(3).Value)
                                Next
                            End Using
                        End If
                    Else
                        cFile(0).SaveAs(oContext.Server.MapPath("~\tmp\") & oContext.Request.UserHostAddress.Replace(":", "") & "." & sId & "." & sDate & ".zip")
                    End If


                    Using oZip As Ionic.Zip.ZipFile = Ionic.Zip.ZipFile.Read(oContext.Server.MapPath("~\tmp\") & oContext.Request.UserHostAddress.Replace(":", "") & "." & sId & "." & sDate & ".zip")
                        If Not ((oZip.Count = 3 _
                                AndAlso oZip.ContainsEntry("print.gcode") _
                                AndAlso oZip.ContainsEntry("image.jpg") _
                                AndAlso oZip.ContainsEntry("metadata.json")) OrElse _
                            (oZip.Count = 4 _
                                AndAlso oZip.ContainsEntry("print.gcode") _
                                AndAlso oZip.ContainsEntry("time-lapse.mp4") _
                                AndAlso oZip.ContainsEntry("image.jpg") _
                                AndAlso oZip.ContainsEntry("metadata.json"))) Then
                            oContext.Response.ContentType = "text/plain"
                            oContext.Response.StatusCode = 433
                            oContext.Response.Write("Incorrect parameter")
                            ZSSOUtilities.WriteLog("UserLib/AddPrint: Incorrect parameter (" & oContext.Request.UserHostAddress & ")")
                            Return
                        End If

                        If Directory.Exists(oContext.Server.MapPath("~\tmp\") & oContext.Request.UserHostAddress.Replace(":", "") & "." & sId & "." & sDate) Then
                            Directory.Delete(oContext.Server.MapPath("~\tmp\") & oContext.Request.UserHostAddress.Replace(":", "") & "." & sId & "." & sDate, True)
                        End If
                        Directory.CreateDirectory(oContext.Server.MapPath("~\tmp\") & oContext.Request.UserHostAddress.Replace(":", "") & "." & sId & "." & sDate)

                        ' Extract the content of the archive
                        oZip.ExtractAll(oContext.Server.MapPath("~\tmp\") & oContext.Request.UserHostAddress.Replace(":", "") & "." & sId & "." & sDate)
                    End Using
                    File.Delete(oContext.Server.MapPath("~\tmp\") & oContext.Request.UserHostAddress.Replace(":", "") & "." & sId & "." & sDate & ".zip")

                    Using oRead As New StreamReader(oContext.Server.MapPath("~\tmp\") & oContext.Request.UserHostAddress.Replace(":", "") & "." & sId & "." & sDate & "/metadata.json")
                        Dim oTmp As Dictionary(Of String, String) = ZSSOUtilities.oSerializer.Deserialize(Of Dictionary(Of String, String))(oRead.ReadToEnd())
                    End Using

                    ThreadPool.UnsafeQueueUserWorkItem(New WaitCallback(AddressOf ZSSOUtilities.StorePrint), _
                                                       New NameValueCollection() From {{"id", nId}, _
                                                                                       {"date", sDate}, _
                                                                                       {"path", oContext.Server.MapPath("~\tmp\") & oContext.Request.UserHostAddress.Replace(":", "") & "." & sId & "." & sDate}})
                Catch ex As Exception
                    ZSSOUtilities.WriteLog("UserLib/AddPrint: " & ex.Message)
                    Return
                End Try
                ZSSOUtilities.WriteLog("UserLib/AddPrint: OK")
            End If
        End If
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class