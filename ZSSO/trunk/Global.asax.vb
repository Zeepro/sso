Imports System.Web.SessionState

Public Class Global_asax
    Inherits System.Web.HttpApplication

    Public sEmailError As String = "<html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title></head><body>" & _
                            "Bonjour,<br />" & _
                            "<br/>" & _
                            "Une erreur a &eacute;t&eacute; d&eacute;tect&eacute;e. <br />" & _
                            "<b>M&eacute;thode</b> : {0}<br />" & _
                            "<b>Message</b> : {1}<br />" & _
                            "<br/>" & _
                            "Zeepro" & _
                            "</body></html>"

    Sub Application_Start(ByVal sender As Object, ByVal e As EventArgs)
        ' Se déclenche lorsque l'application est démarrée
    End Sub

    Sub Session_Start(ByVal sender As Object, ByVal e As EventArgs)
        ' Se déclenche lorsque la session est démarrée
    End Sub

    Sub Application_BeginRequest(ByVal sender As Object, ByVal e As EventArgs)
        ' Se déclenche au début de chaque demande
    End Sub

    Sub Application_AuthenticateRequest(ByVal sender As Object, ByVal e As EventArgs)
        ' Se déclenche lors d'une tentative d'authentification de l'utilisation
    End Sub

    Sub Application_Error(ByVal sender As Object, ByVal e As EventArgs)
        Dim oCaughtException = sender.Context.Error
        Dim oHandler = sender.Context.Handler

        ZSSOUtilities.WriteLog(sender.Context.Handler.ToString & " : NOK : " & oCaughtException.Message)
        Try
            Dim oHtmlEmail As New Mail
            oHtmlEmail.sReceiver = "julien.nguyen@zeepro.fr"
            oHtmlEmail.sSubject = "Une erreur est survenue"
            oHtmlEmail.sBody = String.Format(sEmailError, sender.Context.Handler.ToString, oCaughtException.Message)
            oHtmlEmail.Send()
        Catch ex As Exception
            ZSSOUtilities.WriteLog("ApplicationError(Error Handler) : NOK : " & ex.Message)
            Return
        End Try
        ' Se déclenche lorsqu'une erreur se produit
    End Sub

    Sub Session_End(ByVal sender As Object, ByVal e As EventArgs)
        ' Se déclenche lorsque la session se termine
    End Sub

    Sub Application_End(ByVal sender As Object, ByVal e As EventArgs)
        ' Se déclenche lorsque l'application se termine
    End Sub

End Class