﻿Imports System.Net.Mail
Public Class EmailHandler
    Private err_msg As String

    Public Function SendEmail(ByVal strSubject As String, ByVal strMessage As String, ByVal strFile As String, ByVal dsEmail As DataSet) As Boolean
        Dim success As Boolean = True
        Try
            Dim MailMsg As New MailMessage()
            MailMsg.IsBodyHtml = True
            MailMsg.Subject = strSubject.Trim()
            MailMsg.Body = strMessage.Trim() & vbCrLf
            MailMsg.Priority = MailPriority.High
            MailMsg.IsBodyHtml = True


            If Not strFile = "" Then
                If IO.File.Exists(strFile) Then
                    Dim MsgAttach As New Attachment(strFile)
                    MailMsg.Attachments.Add(MsgAttach)
                End If
            End If


            Dim E As Integer
            If dsEmail.Tables(0).Rows.Count > 0 Then
                'MailMsg.From = New MailAddress("notification@atecphil.com")
                'MailMsg.To.Add("dbrazil@atecphil.com")
                'MailMsg.To.Add("kluces@atecphil.com")
                'MailMsg.To.Add("lgestoso@atecphil.com")
                'MailMsg.To.Add("jtorres@atecphil.com")
                'MailMsg.To.Add("apadilla@atecphil.com")
                For E = 0 To dsEmail.Tables(0).Rows.Count - 1
                    If CBool(dsEmail.Tables(0).Rows(E).Item("EMailTo")) Then
                        MailMsg.To.Add(New MailAddress(Trim(dsEmail.Tables(0).Rows(E).Item("Email_Address"))))
                    ElseIf CBool(dsEmail.Tables(0).Rows(E).Item("EMailCC")) Then
                        MailMsg.CC.Add(New MailAddress(Trim(dsEmail.Tables(0).Rows(E).Item("Email_Address"))))
                    ElseIf CBool(dsEmail.Tables(0).Rows(E).Item("EMailBCC")) Then
                        MailMsg.Bcc.Add(New MailAddress(Trim(dsEmail.Tables(0).Rows(E).Item("Email_Address"))))
                    ElseIf CBool(dsEmail.Tables(0).Rows(E).Item("EMailFrom")) Then
                        MailMsg.From = New MailAddress(Trim(dsEmail.Tables(0).Rows(E).Item("Email_Address")))
                    End If
                Next
                E = Nothing
            End If

            '--ATECPHIL--
            Dim sql_handler As New SQLHandler
            Dim ds As New DataSet
            Dim strSQL As String = "SELECT * FROM  ATEC_EmailServer WHERE ID = @ID"
            Dim Username, Password As String
            Dim SmtpMail As New SmtpClient
            sql_handler.CreateParameter(1)
            sql_handler.SetParameterValues(0, "@ID", SqlDbType.Int, 1)

            If sql_handler.OpenConnection Then
                If sql_handler.FillDataSet(strSQL, ds, CommandType.Text) Then

                    SmtpMail.Host = ds.Tables(0).Rows(0).Item("Host").ToString() '"Atec-mail"
                    SmtpMail.Port = ds.Tables(0).Rows(0).Item("Port").ToString()
                    Username = ds.Tables(0).Rows(0).Item("Username").ToString()
                    Password = ds.Tables(0).Rows(0).Item("Password").ToString()
                    SmtpMail.UseDefaultCredentials = True
                    SmtpMail.Credentials = New System.Net.NetworkCredential(Username, Password)
                End If
            End If
            '--ATECPHIL--

            ''--ATECPHIL--
            'Dim SmtpMail As New SmtpClient
            'SmtpMail.Host = "atec-mail" '"202.78.97.8"
            'SmtpMail.Port = 25
            'SmtpMail.UseDefaultCredentials = True
            'SmtpMail.Credentials = New System.Net.NetworkCredential("administrator", "trator#$0809")
            ''--ATECPHIL--

            SmtpMail.Send(MailMsg)

            MailMsg = Nothing
            SmtpMail = Nothing

            Threading.Thread.Sleep(2000)
        Catch exEmail As Exception
            success = False
            err_msg = exEmail.Message
        End Try
        Return success
    End Function

    Public Function SendEmail(ByVal strSubject As String, ByVal strMessage As String, ByVal strFile() As String, ByVal strProcess As String) As Boolean
        Dim success As Boolean = True
        Try
            Dim MailMsg As New MailMessage()
            MailMsg.IsBodyHtml = True
            MailMsg.Subject = strSubject.Trim()
            MailMsg.Body = strMessage.Trim() & vbCrLf
            MailMsg.Priority = MailPriority.High
            MailMsg.IsBodyHtml = True

            If strFile.Length > 0 Then
                For Each file As String In strFile
                    If Not file = "" Then
                        If IO.File.Exists(file) Then
                            Dim MsgAttach As New Attachment(file)
                            MailMsg.Attachments.Add(MsgAttach)
                        End If
                    End If
                Next
            End If

            'Email Recipients
            Dim dsEmail As New DataSet
            Dim strSQL As String = "Select * from tbl_AutoMail_List WHERE ProcessName = @ProcessName"
            Dim sql_handler As New SQLHandler
            sql_handler.CreateParameter(1)
            sql_handler.SetParameterValues(0, "@ProcessName", SqlDbType.NVarChar, strProcess)

            If sql_handler.OpenConnection() Then
                If (sql_handler.FillDataSet(strSQL, dsEmail, CommandType.Text)) Then
                    Dim E As Integer
                    If dsEmail.Tables(0).Rows.Count > 0 Then
                        For E = 0 To dsEmail.Tables(0).Rows.Count - 1
                            If CBool(dsEmail.Tables(0).Rows(E).Item("MailTo")) Then
                                MailMsg.To.Add(New MailAddress(Trim(dsEmail.Tables(0).Rows(E).Item("Email_Address"))))
                            ElseIf CBool(dsEmail.Tables(0).Rows(E).Item("MailCC")) Then
                                MailMsg.CC.Add(New MailAddress(Trim(dsEmail.Tables(0).Rows(E).Item("Email_Address"))))
                            ElseIf CBool(dsEmail.Tables(0).Rows(E).Item("MailFrom")) Then
                                MailMsg.From = New MailAddress(Trim(dsEmail.Tables(0).Rows(E).Item("Email_Address")))
                            End If
                        Next
                        E = Nothing
                    End If
                End If
                sql_handler.CloseConnection()
            End If

            '--ATECPHIL--
            Dim SmtpMail As New SmtpClient
            SmtpMail.Host = "atec-mail" '"202.78.97.8"
            SmtpMail.Port = 25
            SmtpMail.UseDefaultCredentials = True
            SmtpMail.Credentials = New System.Net.NetworkCredential("administrator", "trator#$0809")
            '--ATECPHIL--

            SmtpMail.Send(MailMsg)

            MailMsg = Nothing
            SmtpMail = Nothing

            Threading.Thread.Sleep(5000)
        Catch exEmail As Exception
            success = False
            err_msg = exEmail.Message
        End Try
        Return success
    End Function

    Public Function SendEmail(ByVal strSubject As String, ByVal strMessage As String, ByVal strFile As String, ByVal emailFrom As String, ByVal emailTo As String, ByVal emailCC As String) As Boolean
        Dim success As Boolean = True
        Try
            Dim MailMsg As New MailMessage()
            MailMsg.IsBodyHtml = True
            MailMsg.Subject = strSubject.Trim()
            MailMsg.Body = strMessage.Trim() & vbCrLf
            MailMsg.Priority = MailPriority.High
            MailMsg.IsBodyHtml = True


            If Not strFile = "" Then
                If IO.File.Exists(strFile) Then
                    Dim MsgAttach As New Attachment(strFile)
                    MailMsg.Attachments.Add(MsgAttach)
                End If
            End If

            If emailTo <> "" Then
                MailMsg.To.Add(New MailAddress(emailTo))
            End If
            If emailCC <> "" Then
                MailMsg.CC.Add(New MailAddress(emailCC))
            End If
            If emailFrom <> "" Then
                MailMsg.From = New MailAddress(emailFrom)
            End If


            '--ATECPHIL--
            Dim SmtpMail As New SmtpClient
            SmtpMail.Host = "atec-mail" '"202.78.97.8"
            SmtpMail.Port = 25
            SmtpMail.UseDefaultCredentials = True
            SmtpMail.Credentials = New System.Net.NetworkCredential("administrator", "trator#$0809")
            '--ATECPHIL--

            SmtpMail.Send(MailMsg)

            MailMsg = Nothing
            SmtpMail = Nothing

            Threading.Thread.Sleep(2000)
        Catch exEmail As Exception
            success = False
            err_msg = exEmail.Message
        End Try
        Return success
    End Function

    Public Function GetErroMessage() As String
        Return err_msg
    End Function
End Class
