Imports System.Net.Mail
Imports System.Net
Imports System.Net.Sockets
Imports Microsoft.VisualBasic.VBMath
Module StockManager


	Public LoginTime_ As String = ""
	Public Username_ As String = ""
	Public IsEnabled_ As Boolean = False
	Public CanWorkOnUser_ As Boolean = False
	Public currency_symbol As String = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.CurrencySymbol


	'	Public company_name As String, company_address As String, company_phone As String, company_email As String, company_email_password As String, email_port As String, email_server As String, email_ssl As Boolean, update_email As String

	'Public company_email_ As String = "dennisjamessmith@yahoo.com"

	Public email_port_default As Long = 465
	'	Public port_default As Long = 993
	'Public port_default As Long = 465 ' 587
	Public email_host_default As String = "smtp.gmail.com"
	'Public host_default As String = "plus.smtp.mail.yahoo.com"
	'Public ssl_default As Boolean = True
	Public email_ssl_default As Boolean = False
	Public NewCustomerEmailSubject As String = "New Customer Created"
	Public NewUserEmailSubject As String = "New User Created"
	Public NewSaleEmailSubject As String = "Record Of Sale"
	Public UpdatesSubject As String = "Update from Stock Manager"




	Public email_f As String = "", email_from_f As String = "", email_password_f As String = "", email_to_f As String = "", subject_f As String = "", port_f As Long, enable_ssl_f As Boolean, server_f As String, who_f As String, AsHTML_f As Boolean

	'	Public UserCanRecordStock As Boolean
	'	Public UserCanMakeSale As Boolean
	'	Public UserCanChangeSetting As Boolean
	'	Public UserCanBackupDatabase As Boolean
	'	Public UserCanCreateUser As Boolean
	'	Public UserCanSetDiscount As Boolean
	'	Public UserCanWorkOnClient As Boolean
	'	Public loginName As String

#Region "Stock Manager"

	Public Function Spaces(stringToAdd As String, maximum_length As Integer, Optional EndPoint As Boolean = True) As String
		Dim current_length As String = stringToAdd.Length
		Dim trail_ As String = ""
		Dim return_trail As String = ""
		If current_length < maximum_length Then
			For s__% = 1 To (maximum_length - current_length)
				return_trail &= " "
			Next
			GoTo 3
		ElseIf current_length >= maximum_length Then
			return_trail = ""
		End If

2:
		'		If EndPoint = False Then
		'			If stringToAdd.Length > (maximum_length + 1) Then
		'				stringToAdd = CInt((stringToAdd) - 1) & "+"
		'			End If
		'		End If
3:
		Return stringToAdd & return_trail
	End Function

	Public Sub GetFeedback(email_ As String, email_from As String, email_password As String, email_to As String, subject_ As String, port_ As Long, enable_ssl As Boolean, server_ As String, who As String, Optional AsHTML As Boolean = True)
		email_f = email_
		email_from_f = email_from
		email_password_f = email_password
		email_to_f = email_to
		subject_f = subject_
		port_f = port_
		enable_ssl_f = enable_ssl
		server_f = server_
		who_f = who
		AsHTML_f = AsHTML

		SendFeedback(email_f, email_from_f, email_password_f, email_to_f, subject_f, port_f, enable_ssl_f, server_f, who_f)

		'		Dim r As String = "email_f: " & email_f
		'		r &= vbCrLf & "email_from_f = " & email_from_f
		'		r &= vbCrLf & "email_password_f =  " & email_password_f
		'		r &= vbCrLf & "email_to_f =  " & email_to_f
		'		r &= vbCrLf & "subject_f =  " & subject_f
		'		r &= vbCrLf & "port_f =  " & port_f
		'		r &= vbCrLf & "enable_ssl_f =  " & enable_ssl_f
		'		r &= vbCrLf & "server_f =  " & server_f
		'		r &= vbCrLf & "who_f =  " & who_f
		'		r &= vbCrLf & "AsHTML_f =  " & AsHTML_f

		'		MsgBox(r)
		'		Exit Sub
		'		Dim thread As New System.Threading.Thread(AddressOf SendFeedback)
		'		thread.Start()

	End Sub

	Public Sub SendFeedback(email_ As String, email_from As String, email_password As String, email_to As String, subject_ As String, port_ As Long, enable_ssl As Boolean, server_ As String, who As String, Optional AsHTML As Boolean = True) 'As String
		'	Public Sub SendFeedback()
		'		Try
		On Error Resume Next
		Dim Smtp_Server As New SmtpClient
		Dim e_mail As New MailMessage()
		Smtp_Server.UseDefaultCredentials = True
		Smtp_Server.Credentials = New Net.NetworkCredential(email_from_f, email_password_f)
		Smtp_Server.Port = port_f
		Smtp_Server.EnableSsl = enable_ssl_f
		Smtp_Server.Host = server_f
		'			Smtp_Server.Timeout = 60000 * 60
		e_mail = New MailMessage()
		e_mail.From = New MailAddress(email_from_f)
		e_mail.To.Add(email_to_f)
		e_mail.Subject = subject_f ' "Email Sending"
		e_mail.IsBodyHtml = AsHTML_f
		e_mail.Body = email_f
		Smtp_Server.Send(e_mail)
		'			Return True
		'		Catch error_t As Exception
		'	MsgBox("Email was not sent to the customer due to the following reason: " & vbCrLf & error_t.ToString)
		'		End Try
	End Sub

#End Region

End Module
