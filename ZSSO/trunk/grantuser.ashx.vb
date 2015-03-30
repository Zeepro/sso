Imports System.Web
Imports System.Web.Services
Imports System.Data.SqlClient

Public Class grantuser
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal oContext As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim sToken, sSerial, sEmail, sName, sAccount, sManage, sView, sLanguage, sAccountEmail As String
        Dim nAccountRestriction, nManageRestriction, nViewRestriction As Integer
 
        If ZSSOUtilities.CheckRequests(oContext.Request.UserHostAddress, "grantuser") > 5 Then
            oContext.Response.ContentType = "text/plain"
            oContext.Response.StatusCode = 435
            oContext.Response.Write("Too many requests")
            ZSSOUtilities.WriteLog("GrantUser : Too many requests")
            Return
        Else
            If oContext.Request.HttpMethod = "GET" Then
                oContext.Response.ContentType = "text/html"
                oContext.Response.Write("<!DOCTYPE html><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title>" & _
                                        "<script src=""https://code.jquery.com/jquery-1.10.2.js""></script><script type=""text/javascript"">function load_wait() { $(""#overlay"").addClass(""gray-overlay""); $("".ui-loader"").css(""display"", ""block""); }</script>" & _
                                        "<link rel=""stylesheet"" type=""text/css"" href=""style.css"">" & _
                                        "</head><body><div id=""overlay""></div><div class=""ui-loader ui-corner-all ui-body-a ui-loader-default""><span class=""ui-icon-loading""></span><h1>CreateAccount</h1></div>" & _
                                        "<form  method=""post"" action=""/grantuser.ashx"" accept-charset=""utf-8"">" & _
                                        "token <input id=""token"" name=""token"" type=""text"" /><br />" & _
                                        "serial <input id=""printersn"" name=""printersn"" type=""text"" /><br />" & _
                                        "user email <input id=""user_email"" name=""user_email"" type=""text"" /><br />" & _
                                        "user name <input id=""user_name"" name=""user_name"" type=""text"" /><br />" & _
                                        "account (yes/no) <input id=""account"" name=""account"" type=""text"" /><br />" & _
                                        "manage (yes/no) <input id=""manage"" name=""manage"" type=""text"" /><br />" & _
                                        "view (yes/no) <input id=""view"" name=""view"" type=""text"" /><br />" & _
                                        "<input id=""Submit1"" type=""submit"" value=""Ok""  onclick=""javascript: load_wait();"" /></form></body></html>")
            Else
                sToken = oContext.Request.Form("token")
                sSerial = oContext.Request.Form("printersn")
                sEmail = oContext.Request.Form("user_email")
                sName = oContext.Request.Form("user_name")
                sAccount = oContext.Request.Form("account")
                sManage = oContext.Request.Form("manage")
                sView = oContext.Request.Form("view")

                If String.IsNullOrEmpty(sToken) OrElse _
                    String.IsNullOrEmpty(sSerial) OrElse _
                    String.IsNullOrEmpty(sEmail) OrElse _
                    String.IsNullOrEmpty(sName) OrElse _
                    String.IsNullOrEmpty(sAccount) OrElse _
                    String.IsNullOrEmpty(sManage) OrElse _
                    String.IsNullOrEmpty(sView) Then
                    oContext.Response.StatusCode = 432
                    oContext.Response.Write("Missing parameter")
                    ZSSOUtilities.WriteLog("GrantUser: Missing parameter")
                    Return
                End If

                sAccount = sAccount.ToLower
                sManage = sManage.ToLower
                sView = sView.ToLower

                If sToken.Length <> 40 OrElse _
                    sSerial.Length <> 12 OrElse _
                    Not (ZSSOUtilities.oEmailRegex.IsMatch(sEmail)) OrElse _
                    (sAccount <> "yes" AndAlso sAccount <> "no") OrElse _
                    (sManage <> "yes" AndAlso sManage <> "no") OrElse _
                    (sView <> "yes" AndAlso sView <> "no") Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 433
                    oContext.Response.Write("Incorrect Parameter")
                    ZSSOUtilities.WriteLog("GrantUser: Incorrect parameter")
                    Return
                End If

                If Not ZSSOUtilities.SearchSerial(sSerial) Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 436
                    oContext.Response.Write("Unknown printer")
                    ZSSOUtilities.WriteLog("GrantUser: Unknown printer")
                    Return
                End If

                sAccountEmail = ZSSOUtilities.SearchAccountToken(sToken, sSerial)

                If sAccountEmail Is Nothing Then
                    oContext.Response.ContentType = "text/plain"
                    oContext.Response.StatusCode = 442
                    oContext.Response.Write("Unauthorized user")
                    ZSSOUtilities.WriteLog("GrantUser: Unauthorized user")
                    Return
                End If

                sName = Left(sName, 50)

                If sAccount = "yes" Then
                    nAccountRestriction = 0
                Else
                    nAccountRestriction = 1
                End If

                If sManage = "yes" Then
                    nManageRestriction = 0
                Else
                    nManageRestriction = 1
                End If

                If sView = "yes" Then
                    nViewRestriction = 0
                Else
                    nViewRestriction = 1
                End If

                Using oConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
                    oConnection.Open()

                    Dim oSqlCmd As SqlCommand = oConnection.CreateCommand()
                    Dim oTransaction As SqlTransaction = oConnection.BeginTransaction("grantuser")

                    oSqlCmd.Connection = oConnection
                    oSqlCmd.Transaction = oTransaction

                    Try
                        If Not testaccount.SearchEmail(sEmail) Then
                            ' Account creation (the email sent to the granted user
                            ' will be in the granting user's language)

                            oSqlCmd.CommandText = "SELECT TOP 1 EmailTemplate.language " & _
                                "FROM EmailTemplate " & _
                                "INNER JOIN Account ON EmailTemplate.language = Account.language " & _
                                "INNER JOIN TokenId ON Account.email = TokenId.email " & _
                                "WHERE Name = @template_name AND TokenId.token = @token"

                            oSqlCmd.Parameters.AddWithValue("@template_name", "creategrantedaccount")
                            oSqlCmd.Parameters.AddWithValue("@token", sToken)

                            sLanguage = oSqlCmd.ExecuteScalar()
                            If sLanguage Is Nothing Then
                                sLanguage = "en"
                            End If

                            Dim oRandom As Random = New Random()
                            Dim iCode = oRandom.Next(1, 9999)
                            Dim sPassword As String = Right("0000" & CStr(iCode), 4)
                            Dim sPasswordHash As String = BCrypt.Net.BCrypt.HashPassword(sPassword, BCrypt.Net.BCrypt.GenerateSalt())

                            'Insert new account
                            oSqlCmd.CommandText = "UPDATE Account SET Password=@password, Code=@code WHERE Email=@email IF @@ROWCOUNT=0 INSERT INTO Account (Email, Password, Code, Language) VALUES (@email, @password, @code, @language)"
                            oSqlCmd.Parameters.Clear()
                            oSqlCmd.Parameters.AddWithValue("@email", sEmail.ToLower)
                            oSqlCmd.Parameters.AddWithValue("@password", sPasswordHash)
                            oSqlCmd.Parameters.AddWithValue("@code", iCode)
                            oSqlCmd.Parameters.AddWithValue("@language", sLanguage)
                            oSqlCmd.ExecuteNonQuery()

                            'Select email template
                            oSqlCmd.CommandText = "SELECT TOP 1 HtmlTemplate, Subject FROM EmailTemplate WHERE Name = @template_name AND Language = @language"
                            oSqlCmd.Parameters.Clear()
                            oSqlCmd.Parameters.AddWithValue("@template_name", "creategrantedaccount")
                            oSqlCmd.Parameters.AddWithValue("@language", sLanguage)

                            'Send email with confirm code
                            Using oQueryResult As SqlDataReader = oSqlCmd.ExecuteReader()

                                Dim sTemplate As String = ""
                                Dim sSubject As String = ""

                                If oQueryResult.Read() Then
                                    sTemplate = oQueryResult(oQueryResult.GetOrdinal("HtmlTemplate"))
                                    sSubject = oQueryResult(oQueryResult.GetOrdinal("Subject"))
                                End If

                                If sTemplate.Length > 0 And sSubject.Length > 0 Then
                                    Dim sHtmlTemplate = sTemplate.Replace("#email#", sAccountEmail).Replace("#password#", sPassword)
                                    Dim oHtmlEmail As New Mail
                                    oHtmlEmail.sReceiver = sEmail
                                    oHtmlEmail.sSubject = sSubject
                                    oHtmlEmail.sBody = sHtmlTemplate
                                    oHtmlEmail.Send()
                                End If
                            End Using
                        End If

                        ' Create guest association
                        oSqlCmd.CommandText = "UPDATE AccountPrinterAssociation " & _
                            "SET name = @name, " & _
                            "accountrestriction = @accountrestriction, " & _
                            "managerestriction = @managerestriction, " & _
                            "viewrestriction = @viewrestriction " & _
                            "WHERE email = @email AND serial = @serial AND deleted IS NULL;" & _
                            "IF @@ROWCOUNT = 0 " & _
                            "INSERT INTO AccountPrinterAssociation " & _
                            "(email, serial, name, accountrestriction, managerestriction, viewrestriction) " & _
                            "VALUES (@email, @serial, @name, @accountrestriction, @managerestriction, @viewrestriction)"
                        oSqlCmd.Parameters.Clear()
                        oSqlCmd.Parameters.AddWithValue("@email", sEmail)
                        oSqlCmd.Parameters.AddWithValue("@serial", sSerial)
                        oSqlCmd.Parameters.AddWithValue("@name", sName)
                        oSqlCmd.Parameters.AddWithValue("@accountrestriction", nAccountRestriction)
                        oSqlCmd.Parameters.AddWithValue("@managerestriction", nManageRestriction)
                        oSqlCmd.Parameters.AddWithValue("@viewrestriction", nViewRestriction)
                        oSqlCmd.ExecuteNonQuery()

                        oTransaction.Commit()
                    Catch ex As Exception
                        oSqlCmd.Transaction.Rollback()
                        Throw ex
                    End Try
                End Using
                ZSSOUtilities.WriteLog("GrantUser: OK")
            End If
        End If
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class