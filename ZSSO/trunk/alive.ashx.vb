Imports System.Web
Imports System.Web.Services
Imports System.Text.RegularExpressions.Regex
Imports System.Web.Script.Serialization
Imports System.Data.SqlClient
Imports System.Runtime.Caching
Imports System.IO

Public Class alive
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal oContext As HttpContext) Implements IHttpHandler.ProcessRequest

        Using oConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
            oConnection.Open()

            Dim sQuery = "SELECT TOP 1 * FROM Account"

            Using oSqlCmdSelect As New SqlCommand(sQuery, oConnection)

                Try
                    Using oQueryResult As SqlDataReader = oSqlCmdSelect.ExecuteReader()
                    End Using
                Catch ex As Exception
                    ZSSOUtilities.WriteLog("Alive : NOK : " & ex.Message)
                    Return
                End Try

            End Using
        End Using
        oContext.Response.ContentType = "text/plain"
        oContext.Response.Write("ok")
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class