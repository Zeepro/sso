Imports System.Web
Imports System.Web.Services
Imports System.Text.RegularExpressions.Regex
Imports System.Web.Script.Serialization
Imports System.Data.SqlClient
Imports System.Runtime.Caching
Imports System.IO
Imports System.Security.Cryptography
Imports System.Net.Mail
Imports BCrypt.Net.BCrypt

Public Class sendtipsmail
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal oContext As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim sEmail As String
        Dim bEmailFound As Boolean = False
        Dim oHttpCache As Caching.Cache = HttpRuntime.Cache

        If oContext.Request.HttpMethod = "GET" Then
            sEmail = HttpUtility.UrlDecode(oContext.Request.QueryString("email"))
            ZSSOUtilities.WriteLog("SendTipsMail : " & ZSSOUtilities.oSerializer.Serialize({sEmail}))
            If Not testaccount.SearchEmail(sEmail) Then
                oContext.Response.ContentType = "text/plain"
                oContext.Response.StatusCode = 433
                oContext.Response.Write("Incorrect Parameter")
                ZSSOUtilities.WriteLog("SendTipsMail : Incorrect parameter")
                Return
            End If

            Using oConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
                oConnection.Open()

                'Get account language (english by default)
                Dim sQuery = "SELECT TOP 1 * " & _
                            "FROM Account " & _
                            "WHERE Email=@email AND Deleted IS NULL"

                Dim sLanguage As String = "en"
                Using oSqlCmdSelect As New SqlCommand(sQuery, oConnection)
                    oSqlCmdSelect.Parameters.AddWithValue("@email", sEmail.ToLower)
                    Try
                        Using oQueryResult As SqlDataReader = oSqlCmdSelect.ExecuteReader()
                            If oQueryResult.Read() Then
                                sLanguage = oQueryResult(oQueryResult.GetOrdinal("Language"))
                            End If
                        End Using
                    Catch
                    End Try
                End Using

                'Check email template language
                sQuery = "SELECT TOP 1 * FROM EmailTemplate WHERE Name = @template_name AND Language = @language"

                Using oSqlCmdSelect As New SqlCommand(sQuery, oConnection)
                    oSqlCmdSelect.Parameters.AddWithValue("@template_name", "sendtipsmail")
                    oSqlCmdSelect.Parameters.AddWithValue("@language", sLanguage.ToLower)
                    Try
                        Using oQueryResult As SqlDataReader = oSqlCmdSelect.ExecuteReader()
                            If oQueryResult.HasRows = False Then
                                sLanguage = "en"
                            End If
                        End Using
                    Catch
                    End Try
                End Using

                Dim oSqlCmd As SqlCommand = oConnection.CreateCommand()
                Dim oTransaction As SqlTransaction = oConnection.BeginTransaction("sendtipsmail")

                oSqlCmd.Connection = oConnection
                oSqlCmd.Transaction = oTransaction

                Try
                    'Select email template
                    oSqlCmd.CommandText = "SELECT TOP 1 HtmlTemplate, Subject FROM EmailTemplate WHERE Name = @template_name AND Language = @language"
                    oSqlCmd.Parameters.AddWithValue("@language", sLanguage.ToLower)
                    oSqlCmd.Parameters.AddWithValue("@template_name", "sendtipsmail")

                    oTransaction.Commit()

                    Using oQueryResult As SqlDataReader = oSqlCmd.ExecuteReader()

                        Dim sTemplate As String = ""
                        Dim sSubject As String = ""

                        If oQueryResult.Read() Then
                            sTemplate = oQueryResult(oQueryResult.GetOrdinal("HtmlTemplate"))
                            sSubject = oQueryResult(oQueryResult.GetOrdinal("Subject"))
                        End If

                        If sTemplate.Length > 0 And sSubject.Length > 0 Then
                            Dim sHtmlTemplate = sTemplate
                            Dim oHtmlEmail As New Mail
                            oHtmlEmail.sReceiver = sEmail
                            oHtmlEmail.sSubject = sSubject
                            oHtmlEmail.sBody = sHtmlTemplate
                            oHtmlEmail.Send()
                        End If
                    End Using

                Catch ex As Exception
                    oSqlCmd.Transaction.Rollback()
                    Throw ex
                    Return
                End Try
            End Using
            ZSSOUtilities.WriteLog("SendTipsMail : OK")
        End If
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class