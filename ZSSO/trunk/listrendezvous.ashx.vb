Imports System.Web
Imports System.Web.Services
Imports System.Text.RegularExpressions.Regex
Imports System.Web.Script.Serialization
Imports System.Data.SqlClient
Imports System.Net
Imports System.Runtime.Caching
Imports System.IO

Public Class listrendezvous
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal oContext As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim sSerial As String = ""
        Dim sToken As String = ""
        Dim oHttpCache As Caching.Cache = HttpRuntime.Cache
        Dim arPrinterLocationData As Dictionary(Of String, String) = ZSSOUtilities.GetLocation(oContext.Request.UserHostAddress)
        'Dim arPrinterLocationData As Dictionary(Of String, String)

        If oContext.Request.HttpMethod = "GET" Then
            oContext.Response.ContentType = "text/html"
            oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title></head><body><form  method=""post"" action=""/listrendezvous.ashx"" accept-charset=""utf-8"">Serial <input id=""serial"" name=""serial"" type=""text"" /><br /><input id=""Submit1"" type=""submit"" value=""Ok"" /></form></body></html>")
        Else
            sSerial = HttpUtility.UrlDecode(oContext.Request.Form("serial"))

            '' Code a supprimer lors de la mise en prod (champs a aussi supprimer dans l'html)
            'Dim sIp As String = HttpUtility.UrlDecode(oContext.Request.Form("ip"))
            'arPrinterLocationData = ZSSOUtilities.GetLocation(sIp)
            '' Fin du code a supprimer

            ZSSOUtilities.WriteLog("ListRDV : " & ZSSOUtilities.oSerializer.Serialize({sSerial}))
            If String.IsNullOrEmpty(sSerial) Then
                oContext.Response.StatusCode = 432
                oContext.Response.Write("Missing parameter")
                ZSSOUtilities.WriteLog("ListRDV : Missing parameter")
                Return
            End If

            Using oConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
                oConnection.Open()

                Dim sQuery = "SELECT TOP 1 * " & _
                    "FROM Printer " & _
                    "WHERE Serial=@serial"

                Using oSqlCmd As New SqlCommand(sQuery, oConnection)

                    oSqlCmd.Parameters.AddWithValue("@serial", sSerial)

                    Try
                        Using oQueryResult As SqlDataReader = oSqlCmd.ExecuteReader()

                            If Not oQueryResult.HasRows Then
                                oContext.Response.StatusCode = 436
                                oContext.Response.Write("Unknown printer")
                                ZSSOUtilities.WriteLog("ListRDV : Unknown printer")
                                Return
                            End If
                        End Using
                    Catch ex As Exception
                        oContext.Response.StatusCode = 436
                        oContext.Response.Write("Unknown printer")
                        ZSSOUtilities.WriteLog("ListRDV : Unknown printer")
                        Return
                    End Try
                End Using
            End Using

            Dim arListRdv = New Dictionary(Of String, Double)
            Dim oCacheEnum As IDictionaryEnumerator = oHttpCache.GetEnumerator()
            While oCacheEnum.MoveNext()
                If oCacheEnum.Current.Key.ToString.StartsWith("rendezvous_server_") Then
                    arListRdv(oCacheEnum.Current.Value("hostname") & ".zeepro.com") = ZSSOUtilities.CalculateDistanceBetweenCoordinates(oCacheEnum.Current.Value("latitude"), oCacheEnum.Current.Value("longitude"), arPrinterLocationData("latitude"), arPrinterLocationData("longitude"))
                End If
            End While

            Dim lnkSorted = From RdvServer In arListRdv
                         Order By RdvServer.Value
            Dim arSortedListRdv = lnkSorted.ToDictionary(Function(p) p.Key, Function(p) p.Value)

            oContext.Response.ContentType = "text/plain"
            oContext.Response.Write(ZSSOUtilities.oSerializer.Serialize(arSortedListRdv.Keys))
            ZSSOUtilities.WriteLog("ListRDV : OK : " + ZSSOUtilities.oSerializer.Serialize(arSortedListRdv.Keys))
        End If

    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property
End Class