Imports System.Web
Imports System.Web.Services
Imports System.Data.SqlClient
Imports System.Globalization
Imports Newtonsoft.Json

Public Class list
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal oContext As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim sToken, sAccountEmail As String
        Dim nId As Integer
        Dim dDate As DateTime
        Dim oList As New List(Of Dictionary(Of String, Object))
        Dim oModel, oPrint As Dictionary(Of String, Object)

        If ZSSOUtilities.CheckRequests(oContext.Request.UserHostAddress, "userliblist") > 5 Then
            oContext.Response.ContentType = "text/plain"
            oContext.Response.StatusCode = 435
            oContext.Response.Write("Too many requests")
            ZSSOUtilities.WriteLog("UserLib/List: Too many requests")
            Return
        Else
            If oContext.Request.HttpMethod = "GET" Then
                oContext.Response.ContentType = "text/html"
                oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title>" & _
                                        "<script src=""https://code.jquery.com/jquery-1.10.2.js""></script><script type=""text/javascript"">function load_wait() { $(""#overlay"").addClass(""gray-overlay""); $("".ui-loader"").css(""display"", ""block""); }</script>" & _
                                        "<link rel=""stylesheet"" type=""text/css"" href=""/style.css"">" & _
                                        "</head><body><div id=""overlay""></div><div class=""ui-loader ui-corner-all ui-body-a ui-loader-default""><span class=""ui-icon-loading""></span><h1>userlib\list</h1></div>" & _
                                        "<form  method=""post"" action=""/UserLib/List.ashx"" accept-charset=""utf-8"">" & _
                                        "token <input id=""token"" name=""token"" type=""text"" /><br />" & _
                                        "<input id=""Submit1"" type=""submit"" value=""Ok""  onclick=""javascript: load_wait();"" /></form></body></html>")
            Else
                sToken = oContext.Request.Form("token")

                If String.IsNullOrEmpty(sToken) Then
                    oContext.Response.StatusCode = 432
                    oContext.Response.Write("Missing parameter")
                    ZSSOUtilities.WriteLog("UserLib/List: Missing parameter")
                    Return
                End If

                If sToken.Length <> 40 Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 433
                    oContext.Response.Write("Incorrect Parameter")
                    ZSSOUtilities.WriteLog("UserLib/List: Incorrect parameter")
                    Return
                End If

                sAccountEmail = ZSSOUtilities.SearchAccountEmail(sToken)

                If sAccountEmail Is Nothing Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 442
                    oContext.Response.Write("Unauthorized user")
                    ZSSOUtilities.WriteLog("UserLib/List: Unauthorized user")
                    Return
                End If

                Try
                    Using oConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
                        oConnection.Open()

                        Using oSqlCmd As New SqlCommand("SELECT Model.id, Model.name, Model3Dfile.description AS modeldescription, Model3Dfile.url1, Model3Dfile.etag1, Model3Dfile.url2, Model3Dfile.etag2, Model3Dfile.img AS modelimg, ModelPrint.date, ModelPrint.description AS printdescription, ModelPrint.gcode, ModelPrint.gcodeetag, ModelPrint.[time-lapse] AS timelapse, ModelPrint.[time-lapseetag] AS timelapseetag, ModelPrint.img AS printimg " & _
                                                        "FROM Model " & _
                                                        "LEFT OUTER JOIN Model3Dfile ON Model.id = Model3Dfile.model " & _
                                                        "LEFT OUTER JOIN ModelPrint ON Model.id = ModelPrint.model " & _
                                                        "WHERE Model.email = @email " & _
                                                        "ORDER BY Model.name, ModelPrint.date", _
                                                        oConnection)

                            oSqlCmd.Parameters.AddWithValue("@email", sAccountEmail)

                            Using oQueryResult As SqlDataReader = oSqlCmd.ExecuteReader()

                                nId = 0
                                dDate = Date.MinValue

                                While oQueryResult.Read()
                                    If nId <> oQueryResult("id") Then
                                        oModel = New Dictionary(Of String, Object)
                                        oModel.Add("id", oQueryResult("id"))
                                        oModel.Add("name", oQueryResult("name"))
                                        If Not IsDBNull(oQueryResult("modeldescription")) Then
                                            oModel.Add("description", oQueryResult("modeldescription"))
                                        End If
                                        If Not IsDBNull(oQueryResult("url1")) Then
                                            oModel.Add("3dfile1", oQueryResult("url1"))
                                        End If
                                        If Not IsDBNull(oQueryResult("etag1")) Then
                                            oModel.Add("3dfile1etag", oQueryResult("etag1"))
                                        End If
                                        If Not IsDBNull(oQueryResult("url2")) Then
                                            oModel.Add("3dfile2", oQueryResult("url2"))
                                        End If
                                        If Not IsDBNull(oQueryResult("etag2")) Then
                                            oModel.Add("3dfile2etag", oQueryResult("etag2"))
                                        End If
                                        If Not IsDBNull(oQueryResult("modelimg")) Then
                                            oModel.Add("img", oQueryResult("modelimg"))
                                        End If
                                        oPrint = Nothing
                                        If Not IsDBNull(oQueryResult("date")) Then
                                            If oPrint Is Nothing Then
                                                oPrint = New Dictionary(Of String, Object)
                                            End If
                                            oPrint.Add("date", DateTime.ParseExact(oQueryResult("date"), "yyyyMMddHHmmss", CultureInfo.InvariantCulture))
                                        End If
                                        If Not IsDBNull(oQueryResult("printdescription")) Then
                                            If oPrint Is Nothing Then
                                                oPrint = New Dictionary(Of String, Object)
                                            End If
                                            oPrint.Add("description", oQueryResult("printdescription"))
                                        End If
                                        If Not IsDBNull(oQueryResult("gcode")) Then
                                            If oPrint Is Nothing Then
                                                oPrint = New Dictionary(Of String, Object)
                                            End If
                                            oPrint.Add("gcode", oQueryResult("gcode"))
                                        End If
                                        If Not IsDBNull(oQueryResult("gcodeetag")) Then
                                            If oPrint Is Nothing Then
                                                oPrint = New Dictionary(Of String, Object)
                                            End If
                                            oPrint.Add("gcodeetag", oQueryResult("gcodeetag"))
                                        End If
                                        If Not IsDBNull(oQueryResult("timelapse")) Then
                                            If oPrint Is Nothing Then
                                                oPrint = New Dictionary(Of String, Object)
                                            End If
                                            oPrint.Add("video", oQueryResult("timelapse"))
                                        End If
                                        If Not IsDBNull(oQueryResult("timelapseetag")) Then
                                            If oPrint Is Nothing Then
                                                oPrint = New Dictionary(Of String, Object)
                                            End If
                                            oPrint.Add("videoetag", oQueryResult("timelapseetag"))
                                        End If
                                        If Not IsDBNull(oQueryResult("printimg")) Then
                                            If oPrint Is Nothing Then
                                                oPrint = New Dictionary(Of String, Object)
                                            End If
                                            oPrint.Add("img", oQueryResult("printimg"))
                                        End If
                                        If Not oPrint Is Nothing Then
                                            oModel.Add("print", New List(Of Dictionary(Of String, Object)) From {oPrint})
                                        End If
                                        oList.Add(oModel)
                                        nId = oQueryResult("id")
                                    Else
                                        oPrint = New Dictionary(Of String, Object)
                                        If Not IsDBNull(oQueryResult("date")) Then
                                            oPrint.Add("date", DateTime.ParseExact(oQueryResult("date"), "yyyyMMddHHmmss", CultureInfo.InvariantCulture))
                                        End If
                                        If Not IsDBNull(oQueryResult("printdescription")) Then
                                            oPrint.Add("description", oQueryResult("printdescription"))
                                        End If
                                        If Not IsDBNull(oQueryResult("gcode")) Then
                                            oPrint.Add("gcode", oQueryResult("gcode"))
                                        End If
                                        If Not IsDBNull(oQueryResult("gcodeetag")) Then
                                            oPrint.Add("gcodeetag", oQueryResult("gcodeetag"))
                                        End If
                                        If Not IsDBNull(oQueryResult("timelapse")) Then
                                            oPrint.Add("video", oQueryResult("timelapse"))
                                        End If
                                        If Not IsDBNull(oQueryResult("timelapseetag")) Then
                                            oPrint.Add("videoetag", oQueryResult("timelapseetag"))
                                        End If
                                        If Not IsDBNull(oQueryResult("printimg")) Then
                                            oPrint.Add("img", oQueryResult("printimg"))
                                        End If
                                        DirectCast(oModel("print"), List(Of Dictionary(Of String, Object))).Add(oPrint)
                                    End If
                                End While
                            End Using
                        End Using
                    End Using
                    oContext.Response.Write(JsonConvert.SerializeObject(oList))
                Catch ex As Exception
                    ZSSOUtilities.WriteLog("UserLib/List: " & ex.Message)
                    Return
                End Try
                ZSSOUtilities.WriteLog("UserLib/List: OK")
            End If
        End If
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class