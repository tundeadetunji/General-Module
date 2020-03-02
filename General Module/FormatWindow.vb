Imports NModule.Methods
Imports NModule.NFunctions
Imports System.Security.AccessControl
Imports System.Drawing.Drawing2D
Imports System.Collections.ObjectModel
Imports System.Net
Imports System.Net.Sockets
Imports System.Net.Mail
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.IO
Public Class FormatWindow

#Region "References Variables"
#End Region

#Region "OS Functions Variables"
	Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef lParam As Integer) As Integer
	Private Declare Sub ReleaseCapture Lib "User32" ()
	Const WM_NCLBUTTONDOWN As Short = &HA1S
	Const HTCAPTION As Short = 2

#End Region

#Region "Font Variables"
	Public item_f As New Font("Microsoft Sans Serif", 12, GraphicsUnit.Point)
	Public item_f_m As New Font("Bauhaus 93", 12, FontStyle.Bold, GraphicsUnit.Point)
	Public stat_f As New Font("Verdana", 12, GraphicsUnit.Point)

#End Region

#Region "Theme Variables"
	'Yellow
	Public net_border_background_color_yellow As Color = Color.Black
	Public net_background_color_yellow As Color = Color.Wheat ' Color.FromArgb(34, 34, 34)
	Public net_foreground_color_yellow As Color = Color.Brown
	Public net_dialog_foreground_color_yellow As Color = Color.Brown

	'Green
	Public net_border_background_color_green As Color = Color.Black
	Public net_background_color_green As Color = Color.DarkGreen ' Color.FromArgb(128, 128, 255)
	Public net_foreground_color_green As Color = Color.Lime ' Color.FromArgb(255, 120, 120) ' Color.FromArgb(255, 120, 120)
	Public net_dialog_foreground_color_green As Color = Color.Turquoise ' Color.FromArgb(255, 120, 120) ' Color.Lavender

	'Turqoise
	Public net_border_background_color_turqoise As Color = Color.Black
	Public net_background_color_turqoise As Color = Color.Turquoise ' Color.FromArgb(34, 34, 34)
	Public net_foreground_color_turqoise As Color = Color.FromArgb(0, 0, 102)
	Public net_dialog_foreground_color_turqoise As Color = Color.FromArgb(0, 0, 102)

	'Velvet
	Public net_border_background_color_velvet As Color = Color.FromArgb(177, 67, 67) ' Color.Red
	Public net_background_color_velvet As Color = Color.Black
	Public net_foreground_color_velvet As Color = Color.FromArgb(177, 67, 67) ' Color.Lavender
	Public net_dialog_foreground_color_velvet As Color = Color.FromArgb(177, 67, 67) ' Color.Lavender

	'Purple
	Public net_border_background_color_purple As Color = Color.Black
	Public net_background_color_purple As Color = Color.FromArgb(0, 0, 102)
	'	Public net_background_color_purple As Color = Color.Black
	Public net_foreground_color_purple As Color = Color.Magenta ' Color.FromArgb(0, 0, 102) ' Color.Lavender
	Public net_dialog_foreground_color_purple As Color = Color.DarkMagenta ' Color.Red ' Color.FromArgb(0, 0, 102) ' Color.Lavender

	'White
	Public net_border_background_color_white As Color = Color.FromArgb(128, 128, 255) ' Color.Black
	Public net_background_color_white As Color = Color.FromArgb(255, 255, 255)
	Public net_foreground_color_white As Color = Color.Navy ' Color.FromArgb(128, 128, 255)
	Public net_dialog_foreground_color_white As Color = Color.Navy ' Color.FromArgb(128, 128, 255)

	'Brown
	Public net_border_background_color_brown As Color = Color.Black
	Public net_background_color_brown As Color = Color.FromArgb(34, 34, 34)
	Public net_foreground_color_brown As Color = Color.Wheat
	Public net_dialog_foreground_color_brown As Color = Color.White

	Public border_color As Color = Color.Black
	Public background_color As Color = Color.FromArgb(34, 34, 34)
	Public foreground_color As Color = Color.Lavender
#End Region

#Region "Specifics Variables"

	Dim d As Form
	Dim theme_ As String
	Dim dialog__ As Form, LeftBorder_ As PictureBox, RightBorder_ As PictureBox, TopBorder_ As PictureBox, BottomBorder_ As PictureBox, TopLine_ As PictureBox, BottomLine_ As PictureBox, Title_ As Label, AcceptButton_ As Button, MinimizeButton_ As Label, CloseButton_ As Label, HelpButton_ As PictureBox, MenuStrip_ As MenuStrip, SystemCloseButton_ As Button, UseMimize_ As Boolean, UseMenustrip_ As Boolean, UseClose_ As Boolean, NormalWindow_ As Boolean, AppType__ As String, TitleBar_ As PictureBox, FooterBar_ As PictureBox, MaximizeButton_ As Label, Feedback_ As TextBox, UseFeedback_ As Boolean, IsLogin_ As Boolean, ShowTime_ As Boolean, TimeLabel_ As Label
	Dim SetMenuStrip_ As Boolean = False, SetMenuStrip_Ignore As Boolean = True
	Dim SetHelpButton_ As Boolean = False
	Dim out_timer As Timer

	Dim TimeTimer_ As Timer

	Public suffx
	Public prefx
	Public countr
	Public TextHasSpace As Boolean
	Public strSource

	Public minimize_button_text = ChrW(800) '"_"
	Public close_button_text = ChrW(10539) '9645, 9633, 10799
	Public maximize_button_text = ChrW(10064)

	Public time_label As Label

	'	Public database_file As String = My.Application.Info.DirectoryPath & "\d71697869.idf"

	'	Public user_tooltip As String

	Public LoginTime_ As String = ""
	Public Username_ As String = ""
	Public IsEnabled_ As Boolean = False
	Public CanWorkOnUser_ As Boolean = False
	Public Shared currency_symbol As String = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.CurrencySymbol

	Public email_f As String = "", email_from_f As String = "", email_password_f As String = "", email_to_f As String = "", subject_f As String = "", port_f As Long, enable_ssl_f As Boolean, server_f As String, who_f As String, AsHTML_f As Boolean

	Dim _Files As ReadOnlyCollection(Of String)


	Dim user_picture___ As PictureBox, use_user_picture As Boolean

	Private selected_theme_f As String = "Brown"
	Private selected_theme_background_f As String = ""
	Private selected_language_f As String = "English"

#End Region

#Region ""
	Public Property selected_theme As String
		Get
			Return selected_theme_f
		End Get
		Set(value As String)
			selected_theme_f = value
		End Set
	End Property

	Public Property selected_theme_background As String
		Get
			Return selected_theme_background_f
		End Get
		Set(value As String)
			selected_theme_background_f = value
		End Set
	End Property

	Public Property selected_language As String
		Get
			Return selected_language_f
		End Get
		Set(value As String)
			selected_language_f = value
		End Set
	End Property


#End Region

#Region "Logo"
	Public Function Scape_(img_ As Image) As String
		Dim img As Image = img_ ' Image.FromFile(img_file)
		If img.Width > img.Height Then
			Return "l"
		ElseIf img.Width < img.Height Then
			Return "p"
		Else
			Return "s"
		End If
	End Function
	Public Function RightOffset(height_ As Single, dialog As Form) As Single
		Return (78 - height_) / 2
	End Function

	Public Function LogoTop(height_ As Single, dialog As Form) As Single
		Dim t_ = dialog.Height - 78
		Dim top_ = (78 - height_) / 2
		Return t_ + top_
	End Function

	Public Function LogoLeft(width_ As Single, dialog As Form, height_ As Single) As Single
		'		Return dialog.Width - (14 + width_)
		Return dialog.Width - (RightOffset(height_, dialog) + width_)

	End Function

	Public Sub FormatLogo(picture_ As PictureBox, dialog As Form)
		Dim img_w As Single = picture_.Width, img_h As Single = picture_.Height
		If picture_.BackgroundImage IsNot Nothing Then img_w = picture_.BackgroundImage.Width : img_h = picture_.BackgroundImage.Height

		picture_.BackColor = Color.Transparent
		If img_w > img_h Then   'landscape
			'            Dim k As Single = img_h / img_w
			Dim k As Single = img_w / img_h
			Dim target_w As Single = k * 50
			picture_.Width = target_w
			''			picture_.Left = LogoLeft(target_w, dialog)
			picture_.Left = LogoLeft(target_w, dialog, picture_.Height)
			picture_.Top = LogoTop(picture_.Height, dialog)
		ElseIf img_w < img_h Then    'portrait 
			'            Dim k As Single = img_w / img_h
			Dim k As Single = img_h / img_w
			Dim target_h As Single = k * 50
			picture_.Height = target_h
			''			picture_.Left = LogoLeft(picture_.Width, dialog)
			picture_.Left = LogoLeft(picture_.Width, dialog, 50)
			picture_.Top = LogoTop(target_h, dialog)
		Else    'square
			'			picture_.Left = LogoLeft(picture_.Width, dialog)
			picture_.Left = LogoLeft(picture_.Width, dialog, picture_.Height)
			picture_.Top = LogoTop(picture_.Height, dialog)
		End If

	End Sub

	'
	Public Function target_height(img As Image, Optional max_width As Integer = 300) As Integer
		Return (img.Width * img.Height) / max_width
	End Function

	'	Public Function target_width(img As Image, Optional max_width As Integer = 300) As Integer
	'		Return (img.Width * img.Height) / max_width
	'	End Function


	Public Function LogoSize(logo As String) As String

		Dim img As Image = Image.FromFile(logo)
		If img.Width > img.Height Then
			Return "width=200 height=150"
		ElseIf img.Width < img.Height Then
			Return "width=150 height=200"
		Else
			Return "width=200 height=200"
		End If

	End Function

#End Region

#Region "User"
	Public Sub User_(user_picture_ As PictureBox, Optional use_user_picture_ As Boolean = True)
		user_picture___ = user_picture_
		use_user_picture = use_user_picture_
	End Sub

	Public Sub PictureToBackgroundImageOrImage(picture_ As PictureBox, file_ As String, Optional UseImageAndNotBackgroundImage As Boolean = False)
		If file_.Length < 1 Then Exit Sub
		If UseImageAndNotBackgroundImage = True Then
			Try
				picture_.Image = Image.FromFile(file_)
			Catch
			End Try
		Else
			Try
				picture_.BackgroundImage = Image.FromFile(file_)
			Catch
			End Try
		End If
	End Sub
	''' <summary>
	''' Takes image and return it cropped as circle if square or landscape, or elipse if portrait. Just place it inside a picture box/control.
	''' </summary>
	''' <param name="o_"></param>
	''' <param name="n_"></param>
	''' <param name="UseImage"></param>
	''' <param name="N_Picture_Size_Width"></param>
	''' <param name="N_Picture_Size_Height"></param>
	''' <param name="N_Picture_Location_Top"></param>
	''' <param name="N_Picture_Location_Left"></param>
	Public Sub ImageIntoCircularPictureControl(o_ As Image, n_ As PictureBox, Optional UseImage As Boolean = False, Optional N_Picture_Size_Width As Integer = 250, Optional N_Picture_Size_Height As Integer = 250, Optional N_Picture_Location_Top As Integer = 0, Optional N_Picture_Location_Left As Integer = 0)
		If o_ Is Nothing Then Exit Sub
		Select Case UseImage
			Case True
				n_.Top = N_Picture_Location_Top
				n_.Left = N_Picture_Location_Left
				n_.Image = ToCircle(o_, N_Picture_Size_Width, N_Picture_Size_Height)
			Case Else
				n_.Top = N_Picture_Location_Top
				n_.Left = N_Picture_Location_Left
				n_.BackgroundImage = ToCircle(o_, N_Picture_Size_Width, N_Picture_Size_Height)
		End Select
	End Sub

	''' <summary>
	''' Takes image and return it cropped inside picture box as circle if square or landscape, or elipse if portrait. If sys_picture_check = true then it is not cropped.
	''' </summary>
	''' <param name="sys_picture_check"></param>
	''' <param name="o_"></param>
	''' <param name="n_"></param>
	''' <param name="UseImage"></param>
	''' <param name="N_Picture_Size_Width"></param>
	''' <param name="N_Picture_Size_Height"></param>
	''' <param name="N_Picture_Location_Top"></param>
	''' <param name="N_Picture_Location_Left"></param>
	Public Sub UserLoginPicture(sys_picture_check As Boolean, o_ As Image, n_ As PictureBox, Optional UseImage As Boolean = False, Optional N_Picture_Size_Width As Integer = 250, Optional N_Picture_Size_Height As Integer = 250, Optional N_Picture_Location_Left As Integer = 0, Optional N_Picture_Location_Top As Integer = 0)
		If o_ Is Nothing Then Exit Sub
		If sys_picture_check = True Then
			Select Case UseImage
				Case True
					n_.Image = o_
				Case Else
					n_.BackgroundImage = o_
			End Select
		Else
			Select Case UseImage
				Case True
					n_.Top = N_Picture_Location_Top
					n_.Left = N_Picture_Location_Left
					n_.Image = ToCircle(o_, N_Picture_Size_Width, N_Picture_Size_Height)
				Case Else
					n_.Top = N_Picture_Location_Top
					n_.Left = N_Picture_Location_Left
					n_.BackgroundImage = ToCircle(o_, N_Picture_Size_Width, N_Picture_Size_Height)
			End Select
		End If
	End Sub
	Public Sub UserPicture(gender_title_ As Control, picture_ As PictureBox, male_picture As String, female_picture As String, system_picture_check_ As CheckBox, Optional PictureExtension_ As Control = Nothing, Optional IsImageAndNotBackgroundImage As Boolean = False)
		'called from Gender_SelectedIndexChanged
		'		Select Case IsImageAndNotBackgroundImage
		'			Case True
		'				If picture_.Image IsNot Nothing And system_picture_check_.CheckState = CheckState.Unchecked Then
		'					system_picture_check_.CheckState = CheckState.Unchecked
		'					Exit Sub
		'				End If
		'			Case False
		'				If picture_.BackgroundImage IsNot Nothing And system_picture_check_.CheckState = CheckState.Unchecked Then
		'					system_picture_check_.CheckState = CheckState.Unchecked
		'					Exit Sub
		'				End If
		'		End Select

		Select Case gender_title_.Text.ToLower
			Case "male"
				PictureToBackgroundImageOrImage(picture_, male_picture, IsImageAndNotBackgroundImage)
				system_picture_check_.CheckState = CheckState.Checked
				If PictureExtension_ IsNot Nothing Then PictureExtension_.Text = Path.GetExtension(male_picture)
			Case "mr."
				PictureToBackgroundImageOrImage(picture_, male_picture, IsImageAndNotBackgroundImage)
				system_picture_check_.CheckState = CheckState.Checked
				If PictureExtension_ IsNot Nothing Then PictureExtension_.Text = Path.GetExtension(male_picture)
			Case "mr"
				PictureToBackgroundImageOrImage(picture_, male_picture, IsImageAndNotBackgroundImage)
				system_picture_check_.CheckState = CheckState.Checked
				If PictureExtension_ IsNot Nothing Then PictureExtension_.Text = Path.GetExtension(male_picture)
			Case Else
				PictureToBackgroundImageOrImage(picture_, female_picture, IsImageAndNotBackgroundImage)
				system_picture_check_.CheckState = CheckState.Checked
				If PictureExtension_ IsNot Nothing Then PictureExtension_.Text = Path.GetExtension(female_picture)
		End Select

	End Sub

#End Region

#Region "Picture"
	Public Function IsSmaller(image_ As Image, picture_ As PictureBox) As Boolean
		IsSmaller = (image_.Width * image_.Height) < (picture_.Width * picture_.Height)
	End Function

	Public Function IsWide(p_ As PictureBox, Optional isBackgroundImage As Boolean = True)
		If isBackgroundImage = True Then
			IsWide = p_.BackgroundImage.Width > p_.BackgroundImage.Height
		Else
			IsWide = p_.Image.Width > p_.Image.Height
		End If
	End Function

	Public Function IsTall(p_ As PictureBox, Optional isBackgroundImage As Boolean = True)
		If isBackgroundImage = True Then
			IsTall = p_.BackgroundImage.Height > p_.BackgroundImage.Width
		Else
			IsTall = p_.Image.Height > p_.Image.Width
		End If
	End Function

	Public Function IsSquare(p_ As PictureBox, Optional isBackgroundImage As Boolean = True)
		If isBackgroundImage = True Then
			IsSquare = p_.BackgroundImage.Height = p_.BackgroundImage.Width
		Else
			IsSquare = p_.Image.Height = p_.Image.Width
		End If
	End Function

	Public Sub ToEllipse(o_ As PictureBox, n_ As PictureBox)
		If o_.BackgroundImage Is Nothing And o_.Image Is Nothing Then Exit Sub
		'? RESIZE N_ TO SCALE LOOSES PART OF THE OUTPUT
		'MAY RESIZE N_ TO SCALE BEFORE PROCEEDING
		'USUALLY SET N_.IMAGELAYOUT/BACKGROUNDIMAGELAYOUT TO ZOOM 

		'Get the original image.
		Dim originalImage = o_.BackgroundImage

		'Create a new, blank image with the same dimensions.
		Dim croppedImage As New Bitmap(originalImage.Width, originalImage.Height)

		'Prepare to draw on the new image.
		Using g = Graphics.FromImage(croppedImage)
			Dim path As New GraphicsPath

			'Create an ellipse that fills the image in both directions.
			path.AddEllipse(0, 0, croppedImage.Width, croppedImage.Height)

			Dim reg As New Region(path)

			'Draw only within the specified ellipse.
			g.Clip = reg
			g.DrawImage(originalImage, Point.Empty)
		End Using

		'Display the new image.
		n_.BackgroundImage = croppedImage

	End Sub

	Public Function ToCircle(o_ As Image, n_width As Integer, n_height As Integer) As Bitmap
		If o_ Is Nothing Then Exit Function
        'RESIZE N_ TO SCALE LOOSES PART OF THE OUTPUT
        'MAY NEED TO RESIZE N_ TO O_ BEFORE PROCEEDING, MAKING SURE O_ IS SIZED AS ITS IMAGE/BACKGROUNDIMAGE
        'USUALLY SET N_.IMAGELAYOUT/BACKGROUNDIMAGELAYOUT TO ZOOM 


        'Get the original image.
        Dim originalImage = o_

        'Create a new, blank image with the same dimensions.
        Dim croppedImage As New Bitmap(n_width, n_height)

        'Prepare to draw on the new image.
        Using g = Graphics.FromImage(croppedImage)
			Dim path As New GraphicsPath

            'Create an ellipse that fills the image in both directions.
            path.AddEllipse(0, 0, croppedImage.Width, croppedImage.Height)


			Dim reg As New Region(path)

            'Draw only within the specified ellipse.
            g.Clip = reg
			g.DrawImage(originalImage, Point.Empty)
		End Using

        'Display the new image.
        ''        n_.BackgroundImage = croppedImage
        Return croppedImage
	End Function

	Public Sub ToCircles(o_ As PictureBox, n_ As PictureBox)
		If o_.BackgroundImage Is Nothing And o_.Image Is Nothing Then Exit Sub
		'RESIZE N_ TO SCALE LOOSES PART OF THE OUTPUT
		'MAY NEED TO RESIZE N_ TO O_ BEFORE PROCEEDING, MAKING SURE O_ IS SIZED AS ITS IMAGE/BACKGROUNDIMAGE
		'USUALLY SET N_.IMAGELAYOUT/BACKGROUNDIMAGELAYOUT TO ZOOM 


		'Get the original image.
		Dim originalImage = o_.BackgroundImage

		'Create a new, blank image with the same dimensions.
		Dim croppedImage As New Bitmap(n_.Width, n_.Width)

		'Prepare to draw on the new image.
		Using g = Graphics.FromImage(croppedImage)
			Dim path As New GraphicsPath

			'Create an ellipse that fills the image in both directions.
			path.AddEllipse(0, 0, croppedImage.Width, croppedImage.Height)

			Dim reg As New Region(path)

			'Draw only within the specified ellipse.
			g.Clip = reg
			g.DrawImage(originalImage, Point.Empty)
		End Using

		'Display the new image.
		n_.BackgroundImage = croppedImage

	End Sub

	Public Sub ToEllipseImage(o_ As PictureBox, n_ As PictureBox, Optional N_Picture_Size_Width As Integer = 250, Optional N_Picture_Size_Height As Integer = 250, Optional N_Picture_Location_Top As Integer = 0, Optional N_Picture_Location_Left As Integer = 0)
		If o_.BackgroundImage Is Nothing And o_.Image Is Nothing Then Exit Sub
		'? RESIZE N_ TO SCALE LOOSES PART OF THE OUTPUT
		'MAY RESIZE N_ TO SCALE BEFORE PROCEEDING
		'USUALLY SET N_.IMAGELAYOUT/BACKGROUNDIMAGELAYOUT TO ZOOM 

		'Get the original image.
		Dim originalImage = o_.Image

		'Create a new, blank image with the same dimensions.
		Dim croppedImage As New Bitmap(originalImage.Width, originalImage.Height)

		'Prepare to draw on the new image.
		Using g = Graphics.FromImage(croppedImage)
			Dim path As New GraphicsPath

			'Create an ellipse that fills the image in both directions.
			path.AddEllipse(0, 0, croppedImage.Width, croppedImage.Height)

			Dim reg As New Region(path)

			'Draw only within the specified ellipse.
			g.Clip = reg
			g.DrawImage(originalImage, Point.Empty)
		End Using

		'Display the new image.
		With n_
			.Width = N_Picture_Size_Width
			.Height = N_Picture_Size_Height
			.Top = N_Picture_Location_Top
			.Left = N_Picture_Location_Left
			.Image = croppedImage
		End With

	End Sub
	Public Sub ToCircleImage(o_ As PictureBox, n_ As PictureBox, Optional N_Picture_Size_Width As Integer = 250, Optional N_Picture_Size_Height As Integer = 250, Optional N_Picture_Location_Top As Integer = 0, Optional N_Picture_Location_Left As Integer = 0)
		If o_.BackgroundImage Is Nothing And o_.Image Is Nothing Then Exit Sub
		'RESIZE N_ TO SCALE LOOSES PART OF THE OUTPUT
		'MAY NEED TO RESIZE N_ TO O_ BEFORE PROCEEDING, MAKING SURE O_ IS SIZED AS ITS IMAGE/BACKGROUNDIMAGE
		'USUALLY SET N_.IMAGELAYOUT/BACKGROUNDIMAGELAYOUT TO ZOOM 


		'Get the original image.
		Dim originalImage = o_.Image

		'Create a new, blank image with the same dimensions.
		Dim croppedImage As New Bitmap(n_.Width, n_.Width)

		'Prepare to draw on the new image.
		Using g = Graphics.FromImage(croppedImage)
			Dim path As New GraphicsPath

			'Create an ellipse that fills the image in both directions.
			path.AddEllipse(0, 0, croppedImage.Width, croppedImage.Height)

			Dim reg As New Region(path)

			'Draw only within the specified ellipse.
			g.Clip = reg
			g.DrawImage(originalImage, Point.Empty)
		End Using

		'Display the new image.
		With n_
			.Width = N_Picture_Size_Width
			.Height = N_Picture_Size_Height
			.Top = N_Picture_Location_Top
			.Left = N_Picture_Location_Left
			.Image = croppedImage
		End With

	End Sub

#End Region

#Region "Hover"


	'	Public verdana_ As New Font("Verdana", 12, FontStyle.Regular)
	'	Public verdana_bigger As New Font("Verdana", 14, FontStyle.Regular)
	'	Public verdana_underline As New Font("Verdana", 12, FontStyle.Underline)
	'
	'	Public serif_ As New Font("Microsoft Sans Serif", 12, FontStyle.Regular)
	'	Public serif_bigger As New Font("Microsoft Sans Serif", 14, FontStyle.Regular)
	'	Public serif_underline As New Font("Microsoft Sans Serif", 12, FontStyle.Underline)

	Public Function serif_(regular_size As Single) As Font
		Return New Font("Microsoft Sans Serif", regular_size, FontStyle.Regular)
	End Function

	Public Function serif_bigger(hover_size As Single) As Font
		Return New Font("Microsoft Sans Serif", hover_size, FontStyle.Bold)
	End Function

	Public Function serif_underline(regular_size As Single) As Font
		Return New Font("Microsoft Sans Serif", regular_size, FontStyle.Underline)
	End Function

	Public Function verdana_(regular_size As Single) As Font
		Return New Font("Verdana", regular_size, FontStyle.Regular)
	End Function

	Public Function verdana_bigger(hover_size As Single) As Font
		Return New Font("Verdana", hover_size, FontStyle.Bold)
	End Function

	Public Function verdana_underline(regular_size As Single) As Font
		Return New Font("Verdana", regular_size, FontStyle.Underline)
	End Function

	Public Sub MouseHoverOrEnter(l_ As Label, Optional l_caption As Label = Nothing, Optional font_ As String = "verdana_OR_EMPTY", Optional bigger_TRUE_OR_UNDERLINE_FALSE As Boolean = True, Optional regular_size As Integer = 12, Optional hover_size As Integer = 12)
		Select Case font_.ToLower
			Case "verdana"
				Select Case bigger_TRUE_OR_UNDERLINE_FALSE
					Case True
						l_.Font = verdana_bigger(hover_size)
					Case False
						l_.Font = verdana_underline(regular_size)
				End Select
			Case Else
				Select Case bigger_TRUE_OR_UNDERLINE_FALSE
					Case True
						l_.Font = serif_bigger(hover_size)
					Case False
						l_.Font = serif_underline(regular_size)
				End Select
		End Select

		If l_caption IsNot Nothing Then l_caption.Text = l_.Text
	End Sub
	Public Sub MouseLeave(l_ As Label, Optional l_caption As Label = Nothing, Optional font_ As String = "verdana_OR_EMPTY", Optional regular_size As Integer = 12, Optional hover_size As Integer = 12)

		Select Case font_.ToLower
			Case "verdana"
				l_.Font = verdana_(regular_size)
			Case Else
				l_.Font = serif_(regular_size)
		End Select
		If l_caption IsNot Nothing Then l_caption.Text = ""
	End Sub
#End Region

#Region "Panels"
	Private prev_ As Panel, next_ As Panel, left_ As Integer, timer_1 As Timer
	Public Sub NextPanel(previous__ As Panel, next__ As Panel, left__ As Integer, animation_timer__ As Timer, form_ As Form)
		If form_.BackgroundImage IsNot Nothing Then
			previous__.Left = 0 - (previous__.Width + 2)
			next__.Left = left__
			next__.Top = previous__.Top
			Exit Sub
		End If

		'b
		'		previous__.Left = left_
		'		next__.Left = previous__.Left + previous__.Width + left__
		'		next__.Top = previous__.Top

		'		previous__.Left = left_
		next__.Left = previous__.Left + previous__.Width + left__
		next__.Top = previous__.Top

		'b
		left_ = left__
		timer_1 = animation_timer__
		prev_ = previous__
		next_ = next__

		'b
		With timer_1
			.Interval = 2
			AddHandler .Tick, New EventHandler(AddressOf PanelAnimation)
			.Enabled = True
		End With

	End Sub

	Private Sub PanelAnimation()

		'		If prev_.Left > (0 - prev_.Width + left_) And next_.Left > left_ Then
		'			prev_.Left -= left_ * 2
		'			next_.Left = prev_.Left + prev_.Width + left_
		'		ElseIf prev_.Left = (0 - prev_.Width + left_) And next_.Left = left_ Then
		'			timer_1.Enabled = False
		'		End If

		If next_.Left > left_ Then
			next_.Left = prev_.Left + prev_.Width + left_ '+ 2
			prev_.Left -= left_ * 4
		ElseIf next_.Left <= left_ Then
			next_.Left = left_
			timer_1.Enabled = False
		End If

	End Sub

#End Region

#Region "FormatWindow"

	Public Sub Borders(dialog As Form, LeftBorder As PictureBox, RightBorder As PictureBox, TopBorder As PictureBox, BottomBorder As PictureBox, Optional BorderWidth As Integer = 1, Optional BorderColor As Color = Nothing)
		If BorderColor = Nothing Then BorderColor = Color.Black
		With LeftBorder
			.BackColor = BorderColor
			.Width = BorderWidth
			.Height = dialog.Height
			.Left = 0
			.Top = 0
			.BringToFront()
		End With
		With RightBorder
			.BackColor = BorderColor
			.Width = BorderWidth
			.Height = dialog.Height
			.Left = dialog.Width - BorderWidth
			.Top = 0
			.BringToFront()
		End With
		With TopBorder
			.BackColor = BorderColor
			.Width = dialog.Width
			.Height = BorderWidth
			.Left = 0
			.Top = 0
			.BringToFront()
		End With
		With BottomBorder
			.BackColor = BorderColor
			.Width = dialog.Width
			.Height = BorderWidth
			.Left = 0
			.Top = dialog.Height - BorderWidth
			.BringToFront()
		End With
	End Sub

	Public Sub PlaceBorders(diff_ As Integer)
		RightBorder_.Left += diff_
		TopBorder_.Width += diff_
		TopLine_.Width += diff_
		BottomLine_.Width += diff_
		BottomBorder_.Width += diff_
		CloseButton_.Left += diff_
		MinimizeButton_.Left += diff_
		MaximizeButton_.Left += diff_
	End Sub

	Public Sub FormatSplash(dialog As Form, Optional UseTransparency As Boolean = False)
		With dialog
			.BackColor = net_background_color("brown")
			.ForeColor = net_dialog_foreground_color("brown")
			If UseTransparency Then .TransparencyKey = net_background_color("brown")
		End With
	End Sub

	Public Sub FormatDialog(Dialog As Form, Optional theme As String = "brown", Optional UseMenustrip As Boolean = False, Optional MenuStrip As MenuStrip = Nothing, Optional AcceptButton As Button = Nothing, Optional UseClose As Boolean = False, Optional SystemCloseButton As Button = Nothing, Optional EmptyCloseButton As Button = Nothing, Optional Logo As PictureBox = Nothing, Optional AppType As String = "Net")
		'top:48
		If UseMenustrip = True And MenuStrip IsNot Nothing Then
			MenuStrip.Font = item_f
			MenuStrip.Top = 10
			MenuStrip.Left = (Dialog.Width - MenuStrip.Width) / 2
			MenuStrip.Show()
		Else
			If MenuStrip IsNot Nothing Then MenuStrip.Hide()
		End If

		theme_ = theme
		dialog__ = Dialog : AcceptButton_ = AcceptButton : MenuStrip_ = MenuStrip : UseMenustrip_ = UseMenustrip

		d = Dialog

		If AcceptButton IsNot Nothing Then
			If UseClose = True Then
				Dialog.CancelButton = SystemCloseButton
				If SystemCloseButton.Tag = "" Then AddHandler SystemCloseButton.Click, New EventHandler(AddressOf CloseDialog)
			ElseIf UseClose = False And EmptyCloseButton IsNot Nothing Then
				Dialog.CancelButton = EmptyCloseButton
			End If
		End If


		AddHandler Dialog.MouseMove, New MouseEventHandler(AddressOf MouseMove)

		d.BackColor = net_background_color(theme)
		d.ForeColor = net_dialog_foreground_color(theme)

		Dim control_t As TextBox
		Dim control_t_counter As Integer = 0
		For Each control_c As Control In Dialog.Controls
			If TypeOf control_c Is TextBox Then
				control_t = control_c
				If control_t.Multiline = True And control_t.ReadOnly = False Then
					control_t_counter += 1
				End If
				If control_t.Name.StartsWith("stat", StringComparison.CurrentCultureIgnoreCase) Then control_t.TabStop = False : control_t.ReadOnly = True : control_t.Enabled = True : control_t.Font = stat_f
			End If
		Next
		If control_t_counter < 1 Then d.AcceptButton = AcceptButton

		FormatButtons(d, AppType, theme)
		FormatTextBoxes(d, AppType, theme)
		FormatDataGridViews(d, AppType, theme)
		FormatComboBoxes(d, AppType, theme)
		FormatCheckBoxes(d, AppType, theme)
		FormatLabels(d, AppType, theme)
		FormatMenuStrips(d, AppType, theme)
		FormatPictureBoxes(d, AppType, theme)
		FormatNumericUpDowns(d, AppType, theme)
		FormatRadios(d, AppType, theme)
		FormatListBoxes(d, AppType, theme)

		FormatPanels(d, AppType, theme)
		ShieldControls(d)
	End Sub

	Public Sub FormatControls(Dialog As Form, Optional AppType As String = "net", Optional theme As String = "white")
		FormatButtons(d, AppType, theme)
		FormatTextBoxes(d, AppType, theme)
		FormatDataGridViews(d, AppType, theme)
		FormatComboBoxes(d, AppType, theme)
		FormatCheckBoxes(d, AppType, theme)
		FormatCombos(d)
		FormatLists(d)
		FormatLabels(d, AppType, theme)
		FormatMenuStrips(d, AppType, theme)
		FormatPictureBoxes(d, AppType, theme)
		FormatNumericUpDowns(d, AppType, theme)
		FormatRadios(d, AppType, theme)
		FormatListBoxes(d, AppType, theme)

	End Sub

	'Public Sub SetFiles()
	'	theme_file = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\" & App & "\SettingTheme.txt"
	'	theme_background_file = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\" & App & "\SettingThemeBackground.txt"
	'End Sub

	Public Sub FormatMe(Dialog As Form, OutTimer As Timer, LeftBorder As PictureBox, RightBorder As PictureBox, TopBorder As PictureBox, BottomBorder As PictureBox, TopLine As PictureBox, BottomLine As PictureBox, DialogTitle As Label, AcceptButton As Button, MinimizeButton As Label, CloseButton As Label, HelpButton As PictureBox, MenuStrip As MenuStrip, SystemCloseButton As Button, EmptyCloseButton As Button, TitleBar As PictureBox, FooterBar As PictureBox, MaximizeButton As Label, UseMimize As Boolean, UseMenustrip As Boolean, UseClose As Boolean, NormalWindow As Boolean, theme As String, Feedback As TextBox, TimeTimer As Timer, TimeLabel As Label, Optional ShowTime As Boolean = False, Optional IsLogin As Boolean = False, Optional UseFeedback As Boolean = False, Optional AppType As String = "Net", Optional UseTooltip As Boolean = False)
		'OPTIONALLY
		'change any of public variables to another before FormatMe, eg item_f
		'if nothing needs to be called when CloseButton is pressed, then OutTimer should be passed in (in place of Nothing) in FormatMe

		'ADDITIONALLY
		'call SetMenuStrip (after FormatMe, if dialog is UserAccess or expected to have its Menu not at the edge
		'call SetHelpButton (if dialog's tag is set to Main (see below) compulsorily, after FormatMe, else the HelpButton will not show)

		'NORMALLY
		'stat (textbox or label for displaying quick information) should begin with stat, eg StatMain
		'Main (dialog that is expected to be maximized and show the TimeLabel) should have its tag's value set to main
		'button that's expected to be transparent should have its tag set to trans or transparent
		'login dialog's name should be UserAccess
		'logo should not be named logo if dialog's name is UserAccess (logo will be subject to FormatLogo)

		CloseButton.Text = close_button_text
		MinimizeButton.Text = minimize_button_text
		MaximizeButton.Text = maximize_button_text

		Select Case IsLogin
			Case True
				DialogTitle.Visible = False
				MinimizeButton.Visible = False
				'				HelpButton.Left = (Dialog.Width - HelpButton.Width) / 2
				HelpButton.Visible = False
				TopLine.Visible = False
				BottomLine.Visible = False
				MaximizeButton.Visible = False
		End Select

		If UseMimize = True Then MinimizeButton.Visible = True

		time_label = TimeLabel

		With TimeTimer
			If ShowTime Then
				.Interval = 1000
				If .Tag = "" Then AddHandler TimeTimer.Tick, New EventHandler(AddressOf ShowTimeNow)

				.Enabled = True
			End If
		End With

		If DialogTitle.Tag = "" Then DialogTitle.Text = Dialog.Text

		theme_ = theme
		AppType__ = AppType
		dialog__ = Dialog : LeftBorder_ = LeftBorder : RightBorder_ = RightBorder : TopBorder_ = TopBorder : BottomBorder_ = BottomBorder : TopLine_ = TopLine : BottomLine_ = BottomLine : Title_ = DialogTitle : AcceptButton_ = AcceptButton : MinimizeButton_ = MinimizeButton : CloseButton_ = CloseButton : HelpButton_ = HelpButton : MenuStrip_ = MenuStrip : SystemCloseButton_ = SystemCloseButton : UseMimize_ = UseMimize : UseMenustrip_ = UseMenustrip : UseClose_ = UseClose : NormalWindow_ = NormalWindow : AppType__ = AppType : TitleBar_ = TitleBar : FooterBar_ = FooterBar : MaximizeButton_ = MaximizeButton : IsLogin_ = IsLogin : ShowTime_ = ShowTime ': TimeLabel_ = TimeLabel : TimeTimer_ = TimeTimer
		out_timer = OutTimer
		'		AddHandler Dialog.Activated, New EventHandler(AddressOf Dialog_Activated)

		d = Dialog

		If CloseButton.Tag = "" Then AddHandler CloseButton.Click, New EventHandler(AddressOf CloseDialog)
		If UseClose = True Then
			Dialog.CancelButton = SystemCloseButton
			If SystemCloseButton.Tag = "" Then AddHandler SystemCloseButton.Click, New EventHandler(AddressOf CloseDialog)
		ElseIf UseClose = False Then
			Dialog.CancelButton = EmptyCloseButton
		End If

		If MaximizeButton.Tag = "" Then AddHandler MaximizeButton.Click, New EventHandler(AddressOf Restore)
		If TitleBar.Tag = "" Then AddHandler TitleBar.DoubleClick, New EventHandler(AddressOf Restore)

		AddHandler Dialog.MouseMove, New MouseEventHandler(AddressOf MouseMove)
		'		If NormalWindow = True Then
		'			Dialog.WindowState = FormWindowState.Normal
		'		ElseIf NormalWindow = False Then
		'			Dialog.WindowState = FormWindowState.Maximized
		'		End If
		If MinimizeButton.Tag = "" Then AddHandler MinimizeButton.Click, New EventHandler(AddressOf MinimizeDialog)

		'ConnectString(database_file)

		'		Dim s As New ToolTip
		'		s.SetToolTip(CloseButton, f.ToLanguage("Close"))
		'		s.SetToolTip(HelpButton, f.ToLanguage("Need Help? Tap me."))
		'		s.SetToolTip(MinimizeButton, f.ToLanguage("Minimize"))
		'		user_tooltip = f.ToLanguage(" is currently logged in (logged in at ", Username_) & LoginTime_ & ")"
		'		If UseTooltip = True Then s.SetToolTip(Title, Username_ & f.ToLanguage(" is currently logged in (logged in at ", Username_) & LoginTime_ & ")")

		d.BackColor = net_background_color(theme)
		d.ForeColor = net_dialog_foreground_color(theme)
		If AppType.ToLower <> "net" Then
			d.BackColor = background_color
			d.ForeColor = foreground_color
		End If

		d.Font = item_f

		Dim control_t As TextBox
		Dim control_t_counter As Integer = 0
		For Each control_c As Control In Dialog.Controls
			If TypeOf control_c Is TextBox Then
				control_t = control_c
				If control_t.Multiline = True And control_t.ReadOnly = False Then
					control_t_counter += 1
				End If
				If control_t.Name.StartsWith("stat", StringComparison.CurrentCultureIgnoreCase) Then control_t.TabStop = False : control_t.ReadOnly = True : control_t.Enabled = True : control_t.Font = stat_f
			End If
		Next
		If control_t_counter < 1 And AcceptButton IsNot SystemCloseButton And AcceptButton IsNot EmptyCloseButton Then d.AcceptButton = AcceptButton

		FormatButtons(d, AppType, theme)
		FormatTextBoxes(d, AppType, theme)
		FormatDataGridViews(d, AppType, theme)
		FormatComboBoxes(d, AppType, theme)
		FormatCheckBoxes(d, AppType, theme)
		FormatCombos(d)
		FormatLists(d)
		FormatLabels(d, AppType, theme)
		FormatMenuStrips(d, AppType, theme)
		FormatPictureBoxes(d, AppType, theme)
		FormatNumericUpDowns(d, AppType, theme)
		FormatRadios(d, AppType, theme)
		FormatListBoxes(d, AppType, theme)

		FormatPanels(d, AppType, theme)

		SetTitleBar(Dialog, ShowTime, TimeLabel, NormalWindow_, LeftBorder, RightBorder, TopBorder, BottomBorder, TopLine, BottomLine, DialogTitle, MinimizeButton, CloseButton, HelpButton, MenuStrip, TitleBar, FooterBar, MaximizeButton, UseMimize, UseMenustrip, Feedback, UseFeedback, AppType)

		With Feedback
			.Visible = False
			'			.Left = 8
			'			.Top = d.Height - 94 + 14 + 1
			'			.Height = 66
			'			.TextAlign = HorizontalAlignment.Center
			'			.TabStop = False
		End With

		ShieldControls(d)

	End Sub
	Public Sub SetHelpButton(HelpButton_ As Control, Optional ShowHelp As Boolean = False)
		SetHelpButton_ = ShowHelp
		If ShowHelp = True Then
			HelpButton_.Show()
		Else
			HelpButton_.Hide()
		End If
	End Sub

	Public Sub SetMenuStrip(m_ As MenuStrip, d_ As Form, HelpButton As PictureBox, Optional ignore_ As Boolean = True, Optional Always_ As Boolean = True)
		SetMenuStrip_Ignore = ignore_
		SetMenuStrip_ = Always_

		If ignore_ = False Then
			If HelpButton.Visible = True Then
				m_.Left = ((d_.Width - HelpButton.Left) - m_.Width) / 2
				Exit Sub
			Else
				GoTo 2
			End If
		End If

2:
		If ignore_ = True Then m_.Left = ((d_.Width / 2) - m_.Width) / 2
	End Sub

	Public Sub SetTitleBar(Dialog As Form, ShowTime As Boolean, TimeLabel As Label, NormalWindow As Boolean, LeftBorder As PictureBox, RightBorder As PictureBox, TopBorder As PictureBox, BottomBorder As PictureBox, TopLine As PictureBox, BottomLine As PictureBox, Title As Label, MinimizeButton As Label, CloseButton As Label, HelpButton As PictureBox, MenuStrip As MenuStrip, TitleBar As PictureBox, FooterBar As PictureBox, MaximizeButton As Label, UseMimize As Boolean, UseMenustrip As Boolean, Feedback As TextBox, Optional UseFeedback As Boolean = False, Optional AppType As String = "Net")
		Dim ts As Integer, asp As Integer

		For Each l As Control In Dialog.Controls
			If TypeOf l Is Label And l Is TimeLabel Then
				If ShowTime Then
					With l
						.Font = item_f
						'			.Left = Dialog.Width - 23 - CloseButton.Width - 20 - MinimizeButton.Width
						.Left = 14
						If use_user_picture = True Then .Left = (14 * 2) + user_picture___.Width
						Dim top_ '= Dialog.Height - 54  'after bottom line
						If Dialog.Tag IsNot Nothing And Dialog.Tag.ToString.Length > 0 And Dialog.Tag.ToString.ToLower = "main" Then
							top_ = Dialog.Height - 78
						ElseIf Dialog.Tag Is Nothing Or Dialog.Tag.ToString.Length < 1 Or Dialog.Tag.ToString.ToLower <> "main" Then
							top_ = Dialog.Height - 54  'after bottom line
						End If
						Dim available_space = Dialog.Height - top_
						.Top = top_ + ((available_space / 2) - (.Height / 2))
						.Visible = True
						.BringToFront()
					End With
				Else
					TimeLabel.Visible = False
				End If
			End If
		Next

		For Each l As Control In d.Controls
			If TypeOf (l) Is Label Then
				If l Is Title Then
					With l
						.Left = 8
						.Top = 14
						.Cursor = Cursors.Hand
					End With
				End If
				If l Is CloseButton Then
					With l
						'						.Left = Dialog.Width - 23 - CloseButton.Width
						.Left = Dialog.Width - 8 - CloseButton.Width
						.Top = 14
						.Cursor = Cursors.Hand
					End With
				End If
				If l Is MinimizeButton Then
					If UseMimize = True Then
						With l
							.Font = item_f_m
							'							.Left = Dialog.Width - 23 - CloseButton.Width - 20 - MinimizeButton.Width
							.Left = Dialog.Width - 8 - CloseButton.Width - 20 - MinimizeButton.Width
							If NormalWindow = False Then .Left = Dialog.Width - 8 - CloseButton.Width - 20 - MaximizeButton.Width - 20 - MinimizeButton.Width
							''							.Top = CloseButton.Top - 1
							.Top = 13
							.Cursor = Cursors.Hand
							If IsLogin_ = False Then .Visible = True : .Show() : .BringToFront()
						End With
					End If
				End If
				If l Is MaximizeButton Then
					If NormalWindow = False Then
						With l
							''							.Font = item_f_m
							'							.Left = Dialog.Width - 23 - CloseButton.Width - 20 - MinimizeButton.Width
							.Left = Dialog.Width - 8 - CloseButton.Width - 20 - MaximizeButton.Width
							'							If NormalWindow_ = False Then .Left = Dialog.Width - 8 - CloseButton.Width - 20 - MinimizeButton.Width - 20 - MaximizeButton.Width
							.Top = 14
							.Cursor = Cursors.Hand
							.BringToFront()
						End With
					End If
				End If
			End If
		Next

		For Each m As Control In d.Controls
			If TypeOf (m) Is MenuStrip Then
				If UseMenustrip Then
					If m Is MenuStrip Then
						With m
							'					.BackColor = Color.Transparent
							.Font = item_f
							.Left = Title.Left + Title.Width + 7
							If IsLogin_ = True Then
								Dim a_space = Dialog.Width / 2
								.Left = (a_space - .Width) / 2
							End If
							'							.Top = Title.Top - 5 ' 7
							.Top = Title.Top - ((m.Height - Title.Height) / 2)
						End With
					End If
				End If
			End If
		Next

		For Each p As Control In d.Controls
			If TypeOf (p) Is PictureBox Then
				If p Is LeftBorder Then
					With p
						.BackColor = net_border_background_color(selected_theme)
						If AppType.ToLower() <> "net" Then .BackColor = border_color
						.Left = 0
						.Top = 0
						.Width = 1
						If .Tag = "" Then .Height = d.Height
						.BringToFront()
					End With
				End If
				If p Is RightBorder Then
					With p
						.BackColor = net_border_background_color(selected_theme)
						If AppType.ToLower() <> "net" Then .BackColor = border_color
						.Left = d.Width - 1
						.Top = 0
						.Width = 1
						If .Tag = "" Then .Height = d.Height
						.BringToFront()
					End With
				End If
				If p Is TopBorder Then
					With p
						.BackColor = net_border_background_color(selected_theme)
						If AppType.ToLower() <> "net" Then .BackColor = border_color
						.Left = 0
						.Top = 0
						.Width = d.Width
						.Height = 1
						.BringToFront()
					End With
				End If
				If p Is BottomBorder Then
					With p
						.BackColor = net_border_background_color(selected_theme)
						If AppType.ToLower() <> "net" Then .BackColor = border_color
						.Left = 0
						If .Tag = "" Then .Top = d.Height - 1
						.Width = d.Width
						.Height = 1
						.BringToFront()
					End With
				End If
				If p Is TopLine Then
					With p
						If selected_theme.ToLower = "velvet" Then
							.Visible = False
						Else
							.Visible = True
						End If
						.BackColor = net_border_background_color(selected_theme)
						If AppType.ToLower() <> "net" Then .BackColor = border_color
						.Left = 0
						.Top = 2 * (Title.Top) + Title.Height
						.Width = d.Width
						.Height = 1
						.BringToFront()
						'
						If NormalWindow = False Or d.Name.ToLower = "useraccess" Then .Hide()
					End With
				End If
				If p Is BottomLine Then
					With p
						If selected_theme.ToLower = "velvet" Then
							.Visible = False
						Else
							.Visible = True
						End If
						.BackColor = net_border_background_color(selected_theme)
						If AppType.ToLower() <> "net" Then .BackColor = border_color
						.Left = 0
						.Top = d.Height - 94 ' 55
						If Dialog.Tag <> "" Then
							If Dialog.Tag IsNot Nothing And Dialog.Tag.ToLower = "main" Then
								.Top = d.Height - 79
							End If
						End If
						If Dialog.Tag = "" Then .Top = d.Height - 55
						If UseFeedback = True Then .Top = d.Height - Feedback.Height - 28
						.Width = d.Width
						.Height = 1
						.BringToFront()
						'
						If NormalWindow = False Or d.Name.ToLower = "useraccess" Then .Hide()
					End With
				End If
				If p Is HelpButton Then
					With p
						ts = Title.Left + Title.Width
						If UseMenustrip Then ts = MenuStrip.Left + MenuStrip.Width
						'						asp = (CloseButton.Left - ts - HelpButton.Width) / 2
						'						If UseMimize Then asp = (MinimizeButton.Left - ts - HelpButton.Width) / 2
						'						If NormalWindow_ = False Then asp = (MinimizeButton.Left - ts - HelpButton.Width - MaximizeButton.Width) / 2
						'						.Width = Title.Height
						'						.Height = Title.Height
						'						.Left = ts + asp

						If UseMimize Then asp = MinimizeButton.Left
						If NormalWindow = False Then asp = MaximizeButton.Left
						If UseMimize = False And NormalWindow = False Then asp = CloseButton.Left
						Dim aspace As Integer = asp - ts
						.Width = 20
						.Height = 20
						.Left = ts + ((aspace - .Width) / 2)
						.Top = (48 - .Height) / 2 '14
						If SetHelpButton_ = True Then .Show()
						.BringToFront()
					End With
				End If
				If p Is TitleBar Then
					With p
						'						.BackColor = Color.Transparent
						'						.Left = 0
						'						.Top = 0
						'						.Width = Dialog.Width
						'						.Height = 48
						'						.SendToBack()
						'						.Hide()
						.Visible = False
					End With
				End If
				If p Is FooterBar Then
					With p
						'						.BackColor = Color.Transparent
						'						.Hide()
						.Visible = False
					End With
				End If
				If p.Name.ToLower = "logo" Then FormatLogo(p, d)
			End If
		Next

		SetHelpButton(HelpButton, SetHelpButton_)
		If SetMenuStrip_ = True Then SetMenuStrip(MenuStrip, d, HelpButton, SetMenuStrip_Ignore, SetMenuStrip_)

	End Sub

	Public Sub ShieldControls(d As Form)
		''		Exit Sub
		''

		On Error Resume Next
		For Each c_ As Control In d.Controls
			If c_.Left + c_.Width < 0 Or
					c_.Top + c_.Height < 0 Or
					c_.Left > d.Width Or
					c_.Top > d.Height Then

				If TypeOf c_ Is Button Then
					c_.TabStop = False
				ElseIf TypeOf c_ Is TextBox Then
					c_.TabStop = False
					Dim t_ As TextBox = c_
					t_.ReadOnly = True
				ElseIf TypeOf c_ Is ComboBox Then
					c_.TabStop = False
					Dim b_ As ComboBox = c_
					'					AddHandler b_.Click, New EventHandler(AddressOf NoKey)
				ElseIf TypeOf c_ Is CheckBox Then
					c_.TabStop = False
					c_.Enabled = False
				ElseIf TypeOf c_ Is MenuStrip Then
					c_.TabStop = False
					c_.Enabled = False
				ElseIf TypeOf c_ Is NumericUpDown Then
					c_.TabStop = False
					c_.Enabled = False
				ElseIf TypeOf c_ Is ContextMenuStrip Then
					c_.Enabled = False
				ElseIf TypeOf c_ Is DateTimePicker Then
					c_.TabStop = False
					c_.Enabled = False
				ElseIf TypeOf c_ Is RadioButton Then
					c_.TabStop = False
					c_.Enabled = False
				ElseIf TypeOf c_ Is RichTextBox Then
					c_.TabStop = False
					Dim rtb_ As TextBox = c_
					rtb_.ReadOnly = True
				ElseIf TypeOf c_ Is TabControl Then
					c_.TabStop = False
				ElseIf TypeOf c_ Is DataGridView Then
					c_.TabStop = False
				ElseIf TypeOf c_ Is Panel Then
					c_.TabStop = False
				End If
			End If
		Next
	End Sub

	Public Sub ShieldControlsInPanel(d As Panel)
		''		Exit Sub
		''
		On Error Resume Next
		d.TabStop = False : Exit Sub
		For Each c_ As Control In d.Controls
			If c_.Left + c_.Width < 0 Or
					c_.Top + c_.Height < 0 Or
					c_.Left > d.Width Or
					c_.Top > d.Height Then

				If TypeOf c_ Is Button Then
					c_.TabStop = False
				ElseIf TypeOf c_ Is TextBox Then
					c_.TabStop = False
					Dim t_ As TextBox = c_
					t_.ReadOnly = True
				ElseIf TypeOf c_ Is ComboBox Then
					c_.TabStop = False
					Dim b_ As ComboBox = c_
					'					AddHandler b_.Click, New EventHandler(AddressOf NoKey)
				ElseIf TypeOf c_ Is CheckBox Then
					c_.TabStop = False
					c_.Enabled = False
				ElseIf TypeOf c_ Is MenuStrip Then
					c_.TabStop = False
					c_.Enabled = False
				ElseIf TypeOf c_ Is NumericUpDown Then
					c_.TabStop = False
					c_.Enabled = False
				ElseIf TypeOf c_ Is ContextMenuStrip Then
					c_.Enabled = False
				ElseIf TypeOf c_ Is DateTimePicker Then
					c_.TabStop = False
					c_.Enabled = False
				ElseIf TypeOf c_ Is RadioButton Then
					c_.TabStop = False
					c_.Enabled = False
				ElseIf TypeOf c_ Is RichTextBox Then
					c_.TabStop = False
					Dim rtb_ As TextBox = c_
					rtb_.ReadOnly = True
				ElseIf TypeOf c_ Is TabControl Then
					c_.TabStop = False
				ElseIf TypeOf c_ Is DataGridView Then
					c_.TabStop = False
				ElseIf TypeOf c_ Is Panel Then
					c_.TabStop = False
				End If
			End If
		Next
	End Sub

	Public Sub ShieldControlsInTab(d As TabControl)
		For Each c_ As Control In d.Controls
			If c_.Left + c_.Width < 0 Or
					c_.Top + c_.Height < 0 Or
					c_.Left > d.Width Or
					c_.Top > d.Height Then

				If TypeOf c_ Is Button Then
					c_.TabStop = False
				ElseIf TypeOf c_ Is TextBox Then
					c_.TabStop = False
					Dim t_ As TextBox = c_
					t_.ReadOnly = True
				ElseIf TypeOf c_ Is ComboBox Then
					c_.TabStop = False
					Dim b_ As ComboBox = c_
					'					AddHandler b_.Click, New EventHandler(AddressOf NoKey)
				ElseIf TypeOf c_ Is CheckBox Then
					c_.TabStop = False
					c_.Enabled = False
				ElseIf TypeOf c_ Is MenuStrip Then
					c_.TabStop = False
					c_.Enabled = False
				ElseIf TypeOf c_ Is NumericUpDown Then
					c_.TabStop = False
					c_.Enabled = False
				ElseIf TypeOf c_ Is ContextMenuStrip Then
					c_.TabStop = False
				ElseIf TypeOf c_ Is DateTimePicker Then
					c_.TabStop = False
					c_.Enabled = False
				ElseIf TypeOf c_ Is RadioButton Then
					c_.TabStop = False
					c_.Enabled = False
				ElseIf TypeOf c_ Is RichTextBox Then
					c_.TabStop = False
					Dim rtb_ As TextBox = c_
					rtb_.ReadOnly = True
				ElseIf TypeOf c_ Is DataGridView Then
					c_.TabStop = False
				ElseIf TypeOf c_ Is Panel Then
					c_.TabStop = False
				End If
			End If
		Next
	End Sub

	Public Sub Dialog_Activated(sender As Object, e As EventArgs)
		'		FormatMe(dialog__, LeftBorder_, RightBorder_, TopBorder_, BottomBorder_, TopLine_, BottomLine_, Title_, AcceptButton_, MinimizeButton_, CloseButton_, HelpButton_, MenuStrip_, SystemCloseButton_, UseMimize_, UseMenustrip_, UseClose_, NormalWindow_, theme_, AppType__)
	End Sub

	Public Sub FormatButtons(d As Form, AppType As String, theme As String)
		For Each b As Control In d.Controls
			If TypeOf (b) Is Button Then
				FormatButton(b, AppType, theme)
			End If
		Next
	End Sub
	Public Sub FormatTextBoxes(d As Form, AppType As String, theme As String)
		For Each t As Control In d.Controls
			If TypeOf (t) Is TextBox Then
				FormatTextBox(t, AppType, theme)
			End If
		Next
	End Sub
	Public Sub FormatListBoxes(d As Form, AppType As String, theme As String)
		For Each l_ As Control In d.Controls
			If TypeOf (l_) Is ListBox Then
				FormatListBox(l_, AppType, theme)
			End If
		Next
	End Sub
	Public Sub FormatDataGridViews(d As Form, AppType As String, theme As String)
		For Each g As Control In d.Controls
			If TypeOf (g) Is DataGridView Then
				FormatDataGridView(g, AppType, theme)
			End If
		Next
	End Sub
	Public Sub FormatCombos(d As Form)
		For Each c As Control In d.Controls
			If TypeOf (c) Is ComboBox Then
				FormatCombo(c)
			End If
		Next
	End Sub

	Public Sub FormatComboBoxes(d As Form, AppType As String, theme As String)
		For Each c As Control In d.Controls
			If TypeOf (c) Is ComboBox Then
				FormatComboBox(c, AppType, theme)
			End If
		Next
	End Sub

	Public Sub FormatLists(d As Form)
		For Each l_ As Control In d.Controls
			If TypeOf (l_) Is ListBox Then
				FormatList(l_)
			End If
		Next
	End Sub
	Public Sub FormatCheckBoxes(d As Form, AppType As String, theme As String)
		For Each h As Control In d.Controls
			If TypeOf (h) Is CheckBox Then
				FormatCheckBox(h, AppType, theme)
			End If
		Next
	End Sub

	Public Sub FormatRadios(d As Form, AppType As String, theme As String)
		For Each r As Control In d.Controls
			If TypeOf (r) Is RadioButton Then
				FormatRadio(r, AppType, theme)
			End If
		Next
	End Sub

	Public Sub FormatLabels(d As Form, AppType As String, theme As String)
		For Each l As Control In d.Controls
			If TypeOf (l) Is Label Then
				FormatLabel(l, AppType, theme)
			End If
		Next
	End Sub
	Public Sub FormatMenuStrips(d As Form, AppType As String, theme As String)
		For Each m As Control In d.Controls
			If TypeOf (m) Is MenuStrip Then
				FormatMenuStrip(m, AppType, theme)
			End If
		Next
	End Sub
	Public Sub FormatPictureBoxes(d As Form, AppType As String, theme As String)
		For Each p As Control In d.Controls
			If TypeOf (p) Is PictureBox Then
				FormatPictureBox(p, AppType, theme)
			End If
		Next
	End Sub
	Public Sub FormatPanels(d As Form, AppType As String, theme As String)
		For Each n As Control In d.Controls
			If TypeOf (n) Is Panel Then
				FormatPanel(n, AppType, theme)
			End If
		Next
	End Sub
	Public Sub FormatNumericUpDowns(d As Form, AppType As String, theme As String)
		For Each u_ As Control In d.Controls
			If TypeOf (u_) Is NumericUpDown Then
				FormatNumericUpDown(u_, AppType, theme)
			End If
		Next
	End Sub

	''' <summary>
	''' Sends focus to control without text. May be called from _SelectedIndexChanged (to send focus to the next control). Same as LFocusMe.
	''' </summary>
	''' <param name="sender"></param>
	''' <param name="focusTarget_or_sender"></param>
	Public Shared Sub FocusMe(ByVal sender As Control, ByVal focusTarget_or_sender As Control)
		'focus on focusTarget if it's empty, sends cursor to end of field if not
		On Error Resume Next
		If focusTarget_or_sender Is sender Then SendKeys.Send("{End}") : Exit Sub

		Select Case focusTarget_or_sender.Text
			Case ""
				focusTarget_or_sender.Focus()
			Case Else
				focusTarget_or_sender.Focus()
				SendKeys.Send("{End}")
		End Select
	End Sub

	''' <summary>
	''' Sends focus to control without text. May be called from _SelectedIndexChanged (to send focus to the next control).
	''' </summary>
	''' <param name="first_"></param>
	''' <param name="second_"></param>
	Public Shared Sub XFocusMe(ByVal first_ As Control, ByVal second_ As Control)
		Dim c_ As ComboBox
		Try
			If TypeOf first_ Is ComboBox Then
				c_ = first_
				If c_.Text.Trim.Length < 1 Or c_.Items.Count < 1 Then
					Exit Sub
				Else
					LFocusMe(first_, second_)
				End If
			Else
				LFocusMe(first_, second_)
			End If
		Catch ex As Exception
		End Try
		LFocusMe(first_, second_)
	End Sub
	'Public Sub XFocusMe(ByVal first_ As Control, ByVal second_ As Control)
	'	'send focus to control without text
	'	'called from Store_SelectedIndexChanged
	'	On Error Resume Next
	'	Select Case first_.Text
	'		Case ""
	'			first_.Focus()
	'		Case Else
	'			Select Case second_.Text
	'				Case ""
	'					second_.Focus()
	'				Case Else
	'					SendKeys.Send("{End}")
	'			End Select
	'	End Select

	'End Sub

	''' <summary>
	''' Sends focus to control without text. May be called from _SelectedIndexChanged (to send focus to the next control).
	''' </summary>
	''' <param name="first_"></param>
	''' <param name="second_"></param>
	Public Shared Sub LFocusMe(ByVal first_ As Control, ByVal second_ As Control)
		'send focus to control without text
		'called from Store_SelectedIndexChanged
		On Error Resume Next
		Select Case first_.Text
			Case ""
				first_.Focus()
			Case Else
				Select Case second_.Text
					Case ""
						second_.Focus()
					Case Else
						second_.Focus()
						SendKeys.Send("{End}")
				End Select
		End Select

	End Sub

	Public Shared Sub InitialFocus(ByVal first_ As Control, ByVal second_ As Control, ByVal third_ As Control)
		On Error Resume Next
		If first_.Text.Length > 0 And second_.Text.Length > 0 And third_.Text.Length > 0 Then LFocusMe(first_, second_) : Exit Sub

		If first_.Text.Length < 1 And first_.Enabled = True Then
			first_.Focus()
			Exit Sub
		End If

		If second_.Text.Length < 1 And second_.Enabled = True Then
			second_.Focus()
			Exit Sub
		End If

		If third_.Text.Length < 1 And third_.Enabled = True Then
			third_.Focus()
			Exit Sub
		End If
	End Sub

	''' <summary>
	''' Sets AutoComplete to ListItems and Sorted to True.
	''' </summary>
	''' <param name="combobox_"></param>
	Public Sub FormatCombo(combobox_ As ComboBox)
		Try
			'		On Error Resume Next
			Dim c As ComboBox = combobox_
			'		Try
			'			If c.Tag.ToString.Trim.Length < 1 And c.DropDownStyle = ComboBoxStyle.DropDownList Then
			'			c.AutoCompleteMode = AutoCompleteMode.None
			'			Exit Sub
			'			End If
			'		Catch x As Exception
			'			MsgBox(x.ToString)
			'		End Try
			'		Try
			'			If c.Tag.ToString.Trim.Length < 1 Then
			c.AutoCompleteMode = AutoCompleteMode.Suggest
			'			End If
			'		Catch x As Exception
			'			MsgBox(x.ToString)
			'		End Try
			'		Try
			'			If c.Tag.ToString.Trim.Length < 1 Then
			c.AutoCompleteSource = AutoCompleteSource.ListItems
			'			End If
			'		Catch x As Exception
			'			MsgBox(x.ToString)
			'		End Try
			'		On Error Resume Next
			Try
				'			If c.DisplayMember Is Nothing Or c.DataSource Is Nothing Then c.Sorted = True
				If c.Tag.ToString.Trim.Length < 1 Then c.Sorted = True
			Catch x As Exception
				'			MsgBox(x.ToString)
			End Try
		Catch
		End Try
	End Sub

	Public Sub FormatCombo_(combobox_ As ComboBox)
		'		On Error Resume Next
		Dim c As ComboBox = combobox_
		Try
			If c.Tag.ToString.Trim.Length < 1 And c.DropDownStyle = ComboBoxStyle.DropDownList Then
				c.AutoCompleteMode = AutoCompleteMode.None
				'			Exit Sub
			End If
		Catch x As Exception
			'			MsgBox(x.ToString)
		End Try
		Try
			If c.Tag.ToString.Trim.Length < 1 Then c.AutoCompleteMode = AutoCompleteMode.Suggest
		Catch x As Exception
			'			MsgBox(x.ToString)
		End Try
		Try
			If c.Tag.ToString.Trim.Length < 1 Then c.AutoCompleteSource = AutoCompleteSource.ListItems
		Catch x As Exception
			'			MsgBox(x.ToString)
		End Try
		'		On Error Resume Next
		Try
			'			If c.DisplayMember Is Nothing Or c.DataSource Is Nothing Then c.Sorted = True
			If c.Tag.ToString.Trim.Length < 1 Then c.Sorted = True
		Catch x As Exception
			'			MsgBox(x.ToString)
		End Try
	End Sub

	Public Sub FormatList(listbox_ As ListBox)
		'		On Error Resume Next
		Dim l_ As ListBox = listbox_
		'		On Error Resume Next
		Try
			'			If c.DisplayMember Is Nothing Or c.DataSource Is Nothing Then c.Sorted = True
			If l_.Tag.ToString.Trim.Length < 1 Then l_.Sorted = True
		Catch x As Exception
			'			MsgBox(x.ToString)
		End Try
	End Sub

	Public Sub FormatPageSetupDialog(psd As PageSetupDialog, pd As PrintDocument)
		On Error Resume Next
		With psd
			.PageSettings.Margins.Left = 0.01
			.PageSettings.Margins.Top = 0.01
			.PageSettings.Margins.Right = 0.01
			.PageSettings.Margins.Bottom = 0.01
			.Document = pd
		End With

	End Sub


	Public Sub ReportLength(srcControl As Control, tControl As Label, Optional suffixed As Boolean = True)
		On Error Resume Next
		Dim t As TextBox, c As ComboBox

		If TypeOf srcControl Is TextBox Then
			t = srcControl
			If suffixed = True Then
				tControl.Text = t.Text.Length & "/" & t.MaxLength
			ElseIf suffixed = False Then
				tControl.Text = t.Text.Length
			End If
			If tControl.Left < t.Left + t.Width Then
				tControl.Left = t.Left + (t.Width - tControl.Width)
			End If
		ElseIf TypeOf srcControl Is ComboBox Then
			c = srcControl
			If suffixed = True Then
				tControl.Text = c.Text.Length & "/" & c.MaxLength
			ElseIf suffixed = False Then
				tControl.Text = c.Text.Length
			End If
			If tControl.Left < c.Left + c.Width Then
				tControl.Left = c.Left + (c.Width - tControl.Width)
			End If
		End If
	End Sub

#End Region

#Region "Time"
	Public Sub ShowTimeNow()
		Dim f As New Feedback.Feedback()
		f.selected_language = selected_language
		time_label.Text = f.ToLanguage("It is now ") & Now.ToLongTimeString & ",  " & Now.ToShortDateString
	End Sub
#End Region

#Region "FormatControls"

	Public Sub FormatButton(b As Button, AppType As String, theme As String)
		With b
			.FlatStyle = FlatStyle.Flat
			If b.Tag = "" Then .BackColor = net_background_color(theme)
			If b.Tag = "" And AppType.ToLower <> "net" Then .BackColor = background_color
			If b.Tag = "" Then .ForeColor = net_foreground_color(theme)
			If InStr(b.Tag, "trans") Or InStr(b.Tag, "transparent") Then b.FlatAppearance.BorderSize = 0
			If AppType.ToLower <> "net" Then .ForeColor = foreground_color
		End With
	End Sub

	Public Sub FormatTextBox(t As TextBox, AppType As String, theme As String)
		With t
			If t.Tag = "" Then .BackColor = net_background_color(theme)
			If t.Tag = "" And AppType.ToLower <> "net" Then .BackColor = background_color
			If t.Tag = "" Then .ForeColor = net_foreground_color(theme)
			If AppType.ToLower <> "net" Then .ForeColor = foreground_color
			'			.BorderStyle = BorderStyle.None
			Dim t_ As TextBox = t
			With t_
				If .Multiline = True Then
					.ScrollBars = ScrollBars.Both
					.WordWrap = True
				End If
			End With
		End With
	End Sub

	Public Sub FormatListBox(l_ As ListBox, AppType As String, theme As String)
		With l_
			If .Tag = "" Then .BackColor = net_background_color(theme)
			If .Tag = "" And AppType.ToLower <> "net" Then .BackColor = background_color
			If .Tag = "" Then .ForeColor = net_foreground_color(theme)
			If AppType.ToLower <> "net" Then .ForeColor = foreground_color
			'			.BorderStyle = BorderStyle.None
		End With
	End Sub

	Public Sub FormatNumericUpDown(n_ As NumericUpDown, AppType As String, theme As String)
		With n_
			.BackColor = net_background_color(theme)
			If AppType.ToLower <> "net" Then .BackColor = background_color
			.ForeColor = net_foreground_color(theme)
			If AppType.ToLower <> "net" Then .ForeColor = foreground_color
			'			.BorderStyle = BorderStyle.None
		End With
	End Sub

	Public Sub FormatDataGridView(g As DataGridView, AppType As String, theme As String)
		On Error Resume Next
		With g
			.BackgroundColor = net_background_color(theme)
			If AppType.ToLower <> "net" Then .BackColor = background_color
			.DefaultCellStyle.BackColor = net_background_color(theme)
			.DefaultCellStyle.ForeColor = net_foreground_color(theme)
			.ColumnHeadersHeight = 68 '48
			.BorderStyle = BorderStyle.None
			.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
			.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
			.AllowUserToAddRows = False
			.AllowUserToDeleteRows = False

			If .Tag = "" Then .TabStop = False
			.MultiSelect = False
		End With

	End Sub

	Public Sub FormatComboBox(c As ComboBox, AppType As String, theme As String)
		'		FormatCombo(c)
		With c
			.BackColor = net_background_color(theme)
			If AppType.ToLower <> "net" Then .BackColor = background_color
			.ForeColor = net_foreground_color(theme)
			If AppType.ToLower <> "net" Then .ForeColor = foreground_color
		End With
	End Sub

	Public Sub FormatCheckBox(h As CheckBox, AppType As String, theme As String)
		With h
			'			.BackColor = net_background_color(theme)
			'			If AppType.ToLower <> "net" Then .BackColor = background_color
			.BackColor = Color.Transparent
			.ForeColor = net_foreground_color(theme)
			If AppType.ToLower <> "net" Then .ForeColor = foreground_color
		End With
	End Sub

	Public Sub FormatRadio(r As RadioButton, AppType As String, theme As String)
		With r
			'			.BackColor = net_background_color(theme)
			'			If AppType.ToLower <> "net" Then .BackColor = background_color
			.BackColor = Color.Transparent
			.ForeColor = net_foreground_color(theme)
			If AppType.ToLower <> "net" Then .ForeColor = foreground_color
		End With
	End Sub
	Public Sub FormatLabel(l As Label, AppType As String, theme As String)
		With l
			'			If l.Tag = "" Then .BackColor = Color.Transparent
			.BackColor = Color.Transparent
			If l IsNot Title_ And l.Tag = "" Then .ForeColor = net_foreground_color(theme)
			If l Is MaximizeButton_ Or l Is MinimizeButton_ Or l Is CloseButton_ Then .ForeColor = net_foreground_color(theme)
			'			If l Is Title_ Then .BackColor = Color.Transparent
		End With
	End Sub

	Public Sub FormatMenuStrip(m As MenuStrip, AppType As String, theme As String)
		On Error Resume Next
		With m
			'			.BackColor = net_background_color(theme)
			'			If AppType.ToLower <> "net" Then .BackColor = background_color
			.BackColor = Color.Transparent
			.ForeColor = net_foreground_color(theme)
			If AppType.ToLower <> "net" Then .ForeColor = foreground_color
		End With
	End Sub

	Public Sub FormatDateTime(dt_ As DateTimePicker, format_as_year_month_day_short_long_time_custom As String, theme As String, Optional AppType As String = "Net")
		On Error Resume Next
		With dt_
			'			.BackColor = net_background_color(theme)
			'			If AppType.ToLower <> "net" Then .BackColor = background_color
			'			.BackColor = Color.Transparent
			'			.ForeColor = net_foreground_color(theme)
			'			If AppType.ToLower <> "net" Then .ForeColor = foreground_color

			Select Case format_as_year_month_day_short_long_time_custom.ToLower
				Case "year"
					.Format = DateTimePickerFormat.Custom
					.CustomFormat = "yyyy"
				'					.Value = Date.Parse(Now.Year.ToString)
				Case "month"
					.CustomFormat = "mm"
					.Format = DateTimePickerFormat.Custom
					.Value = Date.Parse(Now.Month.ToString)
				Case "day"
					.CustomFormat = "dd"
					.Format = DateTimePickerFormat.Custom
					.Value = Date.Parse(Now.Day.ToString)
				Case "short"
					.Format = DateTimePickerFormat.Short
				Case "long"
					.Format = DateTimePickerFormat.Long
				Case "time"
					.Format = DateTimePickerFormat.Time
				Case Else
					.CustomFormat = format_as_year_month_day_short_long_time_custom
					.Format = DateTimePickerFormat.Custom
			End Select

		End With
	End Sub

	Public Sub FormatPictureBox(p As PictureBox, AppType As String, theme As String)

		With p
			.BackgroundImageLayout = ImageLayout.Zoom
			Try
				If p IsNot LeftBorder_ And p IsNot RightBorder_ And p IsNot TopBorder_ And p IsNot BottomBorder_ And p IsNot TopLine_ And p IsNot BottomLine_ Then p.BackColor = Color.Transparent
			Finally
			End Try
		End With
	End Sub

	Public Sub FormatPanel(p As Panel, AppType As String, theme As String)
		On Error Resume Next

		With p
			'			.BackColor = net_background_color(theme)
			.BackColor = Color.Transparent
			If AppType.ToLower <> "net" Then .BackColor = background_color
			.ForeColor = net_dialog_foreground_color(theme)
			If AppType.ToLower <> "net" Then .ForeColor = foreground_color
			.Font = item_f
			On Error Resume Next
			.TabStop = False
		End With

		For Each pc As Control In p.Controls
			If TypeOf (pc) Is Button Then
				FormatButton(pc, AppType, theme)
			End If
			If TypeOf (pc) Is TextBox Then
				FormatTextBox(pc, AppType, theme)
			End If
			If TypeOf (pc) Is DataGridView Then
				FormatDataGridView(pc, AppType, theme)
			End If
			If TypeOf (pc) Is ComboBox Then
				FormatComboBox(pc, AppType, theme)
			End If
			If TypeOf (pc) Is CheckBox Then
				FormatCheckBox(pc, AppType, theme)
			End If
			If TypeOf (pc) Is Label Then
				FormatLabel(pc, AppType, theme)
			End If
			If TypeOf (pc) Is MenuStrip Then
				FormatMenuStrip(pc, AppType, theme)
			End If
			If TypeOf (pc) Is PictureBox Then
				FormatPictureBox(pc, AppType, theme)
			End If
			If TypeOf (pc) Is NumericUpDown Then
				FormatNumericUpDowns(pc, AppType, theme)
			End If
		Next

		ShieldControlsInPanel(p)
	End Sub

	Public Sub FormatTab(p As TabPage, AppType As String, theme As String)
		On Error Resume Next

		With p
			.BackColor = net_background_color(theme)
			If AppType.ToLower <> "net" Then .BackColor = background_color
			.ForeColor = net_dialog_foreground_color(theme)
			If AppType.ToLower <> "net" Then .ForeColor = foreground_color
			.Font = item_f
		End With

		For Each pc As Control In p.Controls
			If TypeOf (pc) Is Button Then
				FormatButton(pc, AppType, theme)
			End If
			If TypeOf (pc) Is TextBox Then
				FormatTextBox(pc, AppType, theme)
			End If
			If TypeOf (pc) Is DataGridView Then
				FormatDataGridView(pc, AppType, theme)
			End If
			If TypeOf (pc) Is ComboBox Then
				FormatComboBox(pc, AppType, theme)
			End If
			If TypeOf (pc) Is CheckBox Then
				FormatCheckBox(pc, AppType, theme)
			End If
			If TypeOf (pc) Is Label Then
				FormatLabel(pc, AppType, theme)
			End If
			If TypeOf (pc) Is MenuStrip Then
				FormatMenuStrip(pc, AppType, theme)
			End If
			If TypeOf (pc) Is PictureBox Then
				FormatPictureBox(pc, AppType, theme)
			End If
			If TypeOf (pc) Is NumericUpDown Then
				FormatNumericUpDowns(pc, AppType, theme)
			End If
		Next

		'ShieldControlsInTab(p)
	End Sub


#End Region

#Region "GetColors"
	Public Function net_background_color(current_theme As String) As Color
		Select Case current_theme.ToLower
			Case "green"
				net_background_color = net_background_color_green
			Case "turqoise"
				net_background_color = net_background_color_turqoise
			Case "velvet"
				net_background_color = net_background_color_velvet
			Case "purple"
				net_background_color = net_background_color_purple
			Case "white"
				net_background_color = net_background_color_white
			Case "brown"
				net_background_color = net_background_color_brown
			Case "yellow"
				net_background_color = net_background_color_yellow
				'			Case "custom"
				'				Select Case theme
				'					Case "green"
				'						net_background_color = net_background_color_green
				'					Case "turqoise"
				'						net_background_color = net_background_color_turqoise
				'					Case "velvet"
				'						net_background_color = net_background_color_velvet
				'					Case "purple"
				'						net_background_color = net_background_color_purple
				'					Case "white"
				'						net_background_color = net_background_color_white
				'					Case "brown"
				'						net_background_color = net_background_color_brown
				'					Case "yellow"
				'						net_background_color = net_background_color_yellow
				'		End Select
		End Select
	End Function

	Public Function net_dialog_foreground_color(current_theme As String) As Color
		Select Case current_theme.ToLower
			Case "green"
				net_dialog_foreground_color = net_dialog_foreground_color_green
			Case "turqoise"
				net_dialog_foreground_color = net_dialog_foreground_color_turqoise
			Case "velvet"
				net_dialog_foreground_color = net_dialog_foreground_color_velvet
			Case "purple"
				net_dialog_foreground_color = net_dialog_foreground_color_purple
			Case "white"
				net_dialog_foreground_color = net_dialog_foreground_color_white
			Case "brown"
				net_dialog_foreground_color = net_dialog_foreground_color_brown
			Case "yellow"
				net_dialog_foreground_color = net_dialog_foreground_color_yellow
				'			Case "custom"
				'				Select Case theme
				'			Case "green"
				'				net_dialog_foreground_color = net_dialog_foreground_color_green
				'			Case "turqoise"
				'				net_dialog_foreground_color = net_dialog_foreground_color_turqoise
				'			Case "velvet"
				'				net_dialog_foreground_color = net_dialog_foreground_color_velvet
				'			Case "purple"
				'				net_dialog_foreground_color = net_dialog_foreground_color_purple
				'			Case "white"
				'				net_dialog_foreground_color = net_dialog_foreground_color_white
				'			Case "brown"
				'				net_dialog_foreground_color = net_dialog_foreground_color_brown
				'			Case "yellow"
				'				net_dialog_foreground_color = net_dialog_foreground_color_yellow
				'		End Select
		End Select
	End Function

	Public Function net_border_background_color(current_theme As String) As Color
		Select Case current_theme.ToLower
			Case "green"
				net_border_background_color = net_border_background_color_green
			Case "turqoise"
				net_border_background_color = net_border_background_color_turqoise
			Case "velvet"
				net_border_background_color = net_border_background_color_velvet
			Case "purple"
				net_border_background_color = net_border_background_color_purple
			Case "white"
				net_border_background_color = net_border_background_color_white
			Case "brown"
				net_border_background_color = net_border_background_color_brown
			Case "yellow"
				net_border_background_color = net_border_background_color_yellow
				'			Case "custom"
				'				Select Case theme
				'			Case "green"
				'				net_border_background_color = net_border_background_color_green
				'			Case "turqoise"
				'				net_border_background_color = net_border_background_color_turqoise
				'			Case "velvet"
				'				net_border_background_color = net_border_background_color_velvet
				'			Case "purple"
				'				net_border_background_color = net_border_background_color_purple
				'			Case "white"
				'				net_border_background_color = net_border_background_color_white
				'			Case "brown"
				'				net_border_background_color = net_border_background_color_brown
				'			Case "yellow"
				'				net_border_background_color = net_border_background_color_yellow
				'		End Select
		End Select
	End Function

	Public Function net_foreground_color(current_theme As String) As Color
		Select Case current_theme.ToLower
			Case "green"
				net_foreground_color = net_foreground_color_green
			Case "turqoise"
				net_foreground_color = net_foreground_color_turqoise
			Case "velvet"
				net_foreground_color = net_foreground_color_velvet
			Case "purple"
				net_foreground_color = net_foreground_color_purple
			Case "white"
				net_foreground_color = net_foreground_color_white
			Case "brown"
				net_foreground_color = net_foreground_color_brown
			Case "yellow"
				net_foreground_color = net_foreground_color_yellow
				'			Case "custom"
				'				Select Case theme
				'			Case "green"
				'				net_foreground_color = net_foreground_color_green
				'			Case "turqoise"
				'				net_foreground_color = net_foreground_color_turqoise
				'			Case "velvet"
				'				net_foreground_color = net_foreground_color_velvet
				'			Case "purple"
				'				net_foreground_color = net_foreground_color_purple
				'			Case "white"
				'				net_foreground_color = net_foreground_color_white
				'			Case "brown"
				'				net_foreground_color = net_foreground_color_brown
				'			Case "yellow"
				'				net_foreground_color = net_foreground_color_yellow
				'		End Select
		End Select
	End Function

#End Region

#Region "Theme"
	'Public Function theme() As String
	'	'		Dim s_theme As String = My.Computer.FileSystem.ReadAllText(theme_file).Trim
	'	Dim s_theme As String = ReadText(theme_file)
	'	If s_theme.Length < 1 Or s_theme Is Nothing Then s_theme = "brown"
	'	Return s_theme
	'End Function

	Public Shared Sub SetDialogBackground(form_ As Form, background_file As String)
		On Error Resume Next
		If background_file.Length < 1 Or background_file Is Nothing Or My.Computer.FileSystem.FileExists(background_file) = False Then Exit Sub
		Dim img As Image = Image.FromFile(background_file)
		Dim w As Double = img.Width
		Dim h As Double = img.Height
		If w < h Then
			form_.BackgroundImageLayout = ImageLayout.Tile
		ElseIf w > h Then
			form_.BackgroundImageLayout = ImageLayout.Stretch
		Else
			form_.BackgroundImageLayout = ImageLayout.Stretch
		End If
		form_.BackgroundImage = Image.FromFile(background_file)
	End Sub


#End Region

#Region "Dialog"

	Public Sub CloseDialog()
		Try
			If out_timer IsNot Nothing Then
				out_timer.Enabled = True
				Exit Sub
			End If
		Catch
		End Try
		d.Close()
	End Sub

	Public Sub MinimizeDialog()
		'		If NormalWindow_ = False Then Exit Sub
		'		If d.ShowInTaskbar = False Then Exit Sub
		d.WindowState = FormWindowState.Minimized
	End Sub

	Public Sub Restore()
		If d.WindowState = FormWindowState.Maximized Then
			d.WindowState = FormWindowState.Normal
		Else
			d.WindowState = FormWindowState.Maximized
		End If
		SetTitleBar(dialog__, ShowTime_, TimeLabel_, NormalWindow_, LeftBorder_, RightBorder_, TopBorder_, BottomBorder_, TopLine_, BottomLine_, Title_, MinimizeButton_, CloseButton_, HelpButton_, MenuStrip_, TitleBar_, FooterBar_, MaximizeButton_, UseMimize_, UseMenustrip_, Feedback_, UseFeedback_, AppType__)
	End Sub

	Public Sub MouseMove(sender As Object, e As MouseEventArgs)
		Dim Button As Short = e.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = (e.X)
		Dim Y As Single = (e.Y)
		Dim lngReturnValue As Integer
		If Button = 1 Then
			Call ReleaseCapture()
			lngReturnValue = SendMessage(sender.Handle.ToInt32, WM_NCLBUTTONDOWN, HTCAPTION, 0)
		End If

	End Sub

	Public Sub SaveDialogLocation(app_ As String, dialog_ As Form)
		Try
			My.Computer.Registry.SetValue(app_, "Top", dialog_.Top)
			My.Computer.Registry.SetValue(app_, "Left", dialog_.Left)
		Catch
		End Try
	End Sub

	Public Sub OpenAtLocation(app_ As String, dialog_ As Form)
		Try
			dialog_.Top = CInt(My.Computer.Registry.GetValue(app_, "Top", 0))
			dialog_.Left = CInt(My.Computer.Registry.GetValue(app_, "Left", 0))
		Catch
		End Try
	End Sub

#End Region

#Region "App Specific"
	Public Sub ShowHelp(sender As Form)

		On Error GoTo 9
		Process.Start("Main.htm")
		Exit Sub
9:
		MsgBox("Request cannot be completed at this time. Please try again later, or re-install the application.", MsgBoxStyle.Information)

	End Sub

#End Region

#Region "Network"

	Public Function IP()
		Dim server As String = Environment.MachineName
		Dim c As String = ""

		Try
			Dim ASCII As New System.Text.ASCIIEncoding()

            ' Get server related information.
            Dim heserver As IPHostEntry = Dns.Resolve(server)

            ' Loop on the AddressList
            Dim curAdd As IPAddress, counter As Integer = 0, suffix As String = ", "
			For Each curAdd In heserver.AddressList

				' Display the server IP address in the standard format. In 
				' IPv4 the format will be dotted-quad notation, in IPv6 it will be
				' in in colon-hexadecimal notation.
				c &= curAdd.ToString()
				counter += 1
				If counter < heserver.AddressList.Count Then
					c &= suffix
				End If
			Next curAdd
			Return c
		Catch e As Exception

		End Try
	End Function 'IPAddresses

#End Region


#Region "Other Functions"
	Public Shared Function ListFiles(directory_ As String, Optional ext_ As String = "*.txt", Optional search_depth As FileIO.SearchOption = FileIO.SearchOption.SearchTopLevelOnly) As List(Of String)
		Dim Files__ As ReadOnlyCollection(Of String)
		Try
			Files__ = My.Computer.FileSystem.GetFiles(directory_, search_depth, ext_)
			If Files__.Count > 0 Then Return Files__.ToList()
		Catch
		End Try
	End Function

	Public Shared Function GetFiles(listbox_ As ListBox, directory_ As String, file_type As String, Optional ClearBeforeFill As Boolean = False) As Boolean
		If ClearBeforeFill = True Then listbox_.Items.Clear()
		Dim f As New FormatWindow
		Try
			If My.Computer.FileSystem.DirectoryExists(directory_) = True And My.Computer.FileSystem.GetFiles(directory_, FileIO.SearchOption.SearchAllSubDirectories, file_type).Count > 1 Then
				f._Files = My.Computer.FileSystem.GetFiles(directory_, FileIO.SearchOption.SearchAllSubDirectories, file_type)
				For i% = 0 To f._Files.Count - 1
					listbox_.Items.Add(f._Files.Item(i))
				Next
				Return True
			Else
				Return False
			End If
		Catch
			Return False
		End Try
	End Function

	''' <summary>
	''' Gets file name from OpenFileDialog. Use this if language is involved.
	''' </summary>
	''' <param name="audio_video_picture_exec_all_combined"></param>
	''' <param name="control_text"></param>
	''' <returns></returns>
	Public Shared Function GetFile(audio_video_picture_exec_all_combined As String, Optional control_text As String = "") As String
		Dim fw As New FormatWindow
		Dim f As New Feedback.Feedback()
		f.selected_language = fw.selected_language

		Dim f_ As New OpenFileDialog
		Select Case audio_video_picture_exec_all_combined.ToLower
			Case "audio"
			Case "video"
			Case "picture"
				f_.Filter = f.ToLanguage("Images|*.GIF;*.JPG;*.JPEG;*.TIF;*.BMP;*.PNG;*.ICO")
				f_.Title = f.ToLanguage("Select picture")
			Case "exec"
				f_.Filter = f.ToLanguage("App Launcher Files|*.exe")
				f_.Title = f.ToLanguage("Select app")
				f_.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.StartMenu)
			Case "all"
				f_.Filter = f.ToLanguage("Any File|*.*")
				f_.Title = f.ToLanguage("Select File")
			Case "combined"
		End Select
		f_.Multiselect = False
		f_.ShowDialog()
		If f_.FileName.Length > 0 Then
			Return f_.FileName
		ElseIf f_.FileName.Length < 1 Then
			Return control_text
		End If
	End Function

	''' <summary>
	''' Gets file name and extension from OpenFileDialog. Use this if no language is involved.
	''' </summary>
	''' <param name="audio_video_picture_exec_all_combined">What it is expected to find.</param>
	''' <returns>String Array {0:True if no error or False if otherwise, 1:File Path, 2:File Extension</returns>
	Public Shared Function GetFileAndExtension(Optional audio_video_picture_exec_all_combined As String = "all") As Array
		If audio_video_picture_exec_all_combined.Length < 1 Then audio_video_picture_exec_all_combined = "all"
		Dim f_ As New OpenFileDialog
		Dim return_() As String
		Select Case audio_video_picture_exec_all_combined.ToLower
			Case "audio"
			Case "video"
			Case "picture"
				f_.Filter = "Images|*.GIF;*.JPG;*.JPEG;*.TIF;*.BMP;*.PNG;*.ICO"
				f_.Title = "Select picture"
			Case "exec"
				f_.Filter = "App Launcher Files|*.exe"
				f_.Title = "Select app"
			Case "all"
				f_.Filter = "Any File|*.*"
				f_.Title = "Select File"
			Case "combined"
		End Select
		With f_
			f_.Multiselect = False
			f_.ShowDialog()
			If .FileName.Trim <> "" Then
				return_ = {True, .FileName, Path.GetExtension(.FileName)}
			ElseIf .FileName.Trim = "" Then
				return_ = {False, "", ""}
			End If
		End With
		Return return_




		'		Dim value_ As Array = g.GetFileAndExtension("picture")
		'		or, if Shared
		'		Dim value_ As Array = GetFileAndExtension("picture")
		'		'File
		'		If value_(0) = True And value_(1).ToString.Length > 0 Then
		'			Dim file_ As String = value_(1)
		'			'SysPic.CheckState = CheckState.Unchecked	
		'		End If
		'		'File And Extension
		'		If value_(0) = True And value_(1).ToString.Length > 0 And value_(2).ToString.Length > 0 Then
		'			Dim file_ As String = value_(1)
		'			Dim extension_ As String = value_(2)
		'			'SysPic.CheckState = CheckState.Unchecked	
		'		End If



	End Function

	''' <summary>
	''' Gets Folder Path from FolderBrowserDialog. Use if no language is involved.
	''' </summary>
	''' <param name="control_"></param>
	''' <param name="description_"></param>
	''' <param name="ShowNewFolderButton_"></param>
	''' <returns></returns>
	Public Shared Function GetFolder(Optional control_ As Control = Nothing, Optional description_ As String = "Select Folder", Optional ShowNewFolderButton_ As Boolean = True) As String

		Dim t
		If control_ IsNot Nothing And TypeOf control_ Is TextBox Then t = control_
		If control_ IsNot Nothing And TypeOf control_ Is ComboBox Then t = control_
		If control_ IsNot Nothing And TypeOf control_ IsNot TextBox And TypeOf control_ IsNot ComboBox Then Exit Function

		Dim f_ As New FolderBrowserDialog

		With f_
			.Description = description_
			.ShowNewFolderButton = ShowNewFolderButton_
			.ShowDialog()

			If .SelectedPath.Length > 0 Then
				If t IsNot Nothing Then t.Text = .SelectedPath
				Return .SelectedPath
			ElseIf .SelectedPath.Length < 1 Then
				On Error Resume Next
				Return t.text
			End If

		End With
	End Function

	''' <summary>
	''' Gets text from InputBox. Use this if no language is involved.
	''' </summary>
	''' <param name="prompt_">Message to display.</param>
	''' <param name="default_response">String to return if none is given by user.</param>
	''' <param name="title_">Title of the input box.</param>
	''' <returns></returns>
	Public Shared Function GetText(prompt_ As String, Optional default_response As String = "", Optional title_ As String = "") As String
		Return InputBox(prompt_, title_, default_response)
	End Function

	''' <summary>
	''' Gets file's content. Same as ReadText.
	''' </summary>
	''' <param name="file_"></param>
	''' <returns></returns>
	Public Function file_content(file_ As String) As String
		Return ReadText(file_) : Return True : Exit Function
		'Try
		'	Return My.Computer.FileSystem.ReadAllText(file_).Trim
		'Catch ex As Exception
		'End Try
	End Function

#End Region

#Region "Stock Manager"

	Public Function Gross(quantity_bought As String, discount_units As String, unit_price As String, discount As String, use_discount As Boolean)
		quantity_bought = (quantity_bought)
		discount_units = (discount_units)
		unit_price = (unit_price)
		discount = (discount)
		Dim return_value
		If use_discount = False Or quantity_bought < (discount_units + 1) Then
			return_value = quantity_bought * unit_price
			GoTo 5
		End If
		'convert discount to decimal
		Dim discount_decimal = discount / 100
		'how much of price is off
		Dim how_much_off = discount_decimal * unit_price
		'what is the price now
		Dim new_price = unit_price - how_much_off
		'how many items uses the price now, get gross now
		Dim items_to_use_discount = (quantity_bought - discount_units) + 1
		Dim gross_with_discount = items_to_use_discount * new_price
		'how many items uses the old price, get gross then
		Dim items_to_do_without_discount = quantity_bought - items_to_use_discount
		Dim gross_without_discount = items_to_do_without_discount * unit_price
		'return gross now + gross then
		return_value = gross_without_discount + gross_with_discount
5:
		Return return_value
	End Function

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


#End Region

#Region "Default Action"

	Public Sub ActionItems(default_action_drop_down As ComboBox)
		With default_action_drop_down
			With .Items
				.Clear()
				.Add("Do Nothing")
				.Add("Hibernate")
				.Add("Log Off")
				.Add("Open A File")
				.Add("Open An App")
				.Add("Shut down")
				.Add("Sleep")
			End With
			.SelectedIndex = -1
		End With

	End Sub

	Public Sub DefaultAction(action As String, file As String)
		Select Case action
			Case "log_off"
				Try
					Process.Start(Environment.GetFolderPath(Environment.SpecialFolder.System) & "\shutdown.exe", "-l")
				Catch ex As Exception
				End Try
			Case "open_a_file"
				Try
					Process.Start(file)
				Catch ex As Exception
				End Try
			Case "open_an_app"
				Try
					Process.Start(file)
				Catch ex As Exception
				End Try
			Case "shut_down"
				Try
					Process.Start(Environment.GetFolderPath(Environment.SpecialFolder.System) & "\shutdown.exe", "-s -t 0")
				Catch ex As Exception
				End Try
			Case "hibernate"
				Try
					Application.SetSuspendState(PowerState.Hibernate, True, True)
				Catch ex As Exception
				End Try
			Case "sleep"
				Try
					Application.SetSuspendState(PowerState.Suspend, True, True)
				Catch ex As Exception
				End Try
		End Select

	End Sub

	Public Function ActionItemToCode(control As ComboBox) As String
		Dim atc As String = ""
		Select Case control.Text
			Case "Do Nothing"
				atc = "do_nothing"
			Case "Log Off"
				atc = "log_off"
			Case "Open A File"
				atc = "open_a_file"
			Case "Open An App"
				atc = "open_an_app"
			Case "Shut down"
				atc = "shut_down"
			Case "Hibernate"
				atc = "hibernate"
			Case "sleep"
				atc = "sleep"
		End Select
		Return atc
	End Function

	Public Function ActionToDropDown(t_ As String) As String
		Dim atd As String = ""
		Select Case t_.ToLower
			Case "do_nothing"
				atd = "Do Nothing"
			Case "log_off"
				atd = "Log Off"
			Case "open_a_file"
				atd = "Open A File"
			Case "open_an_app"
				atd = "Open An App"
			Case "shut_down"
				atd = "Shut down"
			Case "hibernate"
				atd = "Hibernate"
			Case "sleep"
				atd = "Sleep"
		End Select
		Return atd
	End Function


#End Region



End Class
