<html>
<head>
<title>Check Components</title>
</head>
<body>
<%

	REM	Project:		COM Checker
	REM	Creator:		James Lindën
	REM	Date:		6/25/2001 12:14am

 	'**************************************
***************************

	'	COM Checker version 1.0
	'	© 2001 Ticluse Teknologi, All rights reserved.

	'	This class may be redistributed, as long as copyright and development
	'	information remain intact with class source code. Bug reports and / or
	'	comments may be directed to jlinden@intelidev.com.

 	'**************************************
***************************

	Dim comList(45)
		comList(0) = Array( "AB Mailer","ABMailer.Mailman" )
		comList(1) = Array( "ABC Upload","ABCUpload4.XForm" )
		comList(2) = Array( "ActiveFile","ActiveFile.Post" )
		comList(3) = Array( "ActiveX Data Object","ADODB.Connection" )
		comList(4) = Array( "Adiscon SimpleMail","ADISCON.SimpleMail.1" )
		comList(5) = Array( "ASP HTTP","AspHTTP.Conn" )
		comList(6) = Array( "ASP Image","AspImage.Image" )
		comList(7) = Array( "ASP Mail","SMTPsvg.Mailer" )
		comList(8) = Array( "ASP Simple Upload","ASPSimpleUpload.Upload" )
		comList(9) = Array( "ASP Smart Cache","aspSmartCache.SmartCache" )
		comList(10) = Array( "ASP Smart Mail","aspSmartMail.SmartMail" )
		comList(11) = Array( "ASP Smart Upload","aspSmartUpload.SmartUpload" )
		comList(12) = Array( "ASP Tear","SOFTWING.ASPtear" )
		comList(13) = Array( "ASP Thumbnailer","ASPThumbnailer.Thumbnail" )
		comList(14) = Array( "ASP WhoIs","WhoIs2.WhoIs" )
		comList(15) = Array( "ASPSoft NT Object","ASPSoft.NT" )
		comList(16) = Array( "ASPSoft Upload","ASPSoft.Upload" )
		comList(17) = Array( "CDONTS","CDONTS.NewMail" )
		comList(18) = Array( "Chestysoft Image","csImageFile.Manage" )
		comList(19) = Array( "Chestysoft Upload","csASPUpload.Process" )
		comList(20) = Array( "Dimac JMail","JMail.Message" )
		comList(21) = Array( "Distinct SMTP","DistinctServerSmtp.SmtpCtrl" )
		comList(22) = Array( "Dundas Mailer","Dundas.Mailer" )
		comList(23) = Array( "Dundas Upload","Dundas.Upload.2" )
		comList(24) = Array( "Dynu Encrypt","Dynu.Encrypt" )
		comList(25) = Array( "Dynu HTTP","Dynu.HTTP" )
		comList(26) = Array( "Dynu Mail","Dynu.Email" )
		comList(27) = Array( "Dynu Upload","Dynu.Upload" )
		comList(28) = Array( "Dynu WhoIs","Dynu.Whois" )
		comList(29) = Array( "Easy Mail","EasyMail.SMTP.5" )
		comList(30) = Array( "File System Object","Scripting.FileSystemObject" )
		comList(31) = Array( "Ticluse Teknologi HTTP","InteliSource.Online" )
		comList(32) = Array( "Last Mod","LastMod.FileObj" )
		comList(33) = Array( "Microsoft XML Engine","Microsoft.XMLDOM" )
		comList(34) = Array( "Persits ASP JPEG","Persits.Jpeg" )
		comList(35) = Array( "Persits ASPEmail","Persits.MailSender" )
		comList(36) = Array( "Persits ASPEncrypt","Persits.CryptoManager" )
		comList(37) = Array( "Persits File Upload","Persits.Upload.1" )
		comList(38) = Array( "SMTP Mailer","SmtpMail.SmtpMail.1" )
		comList(39) = Array( "Soft Artisans File Upload","SoftArtisans.FileUp" )
		comList(40) = Array( "Image Size", "ImgSize.Check" )
		comList(41) = Array( "Microsoft XML HTTP", "Microsoft.XMLHTTP" )
		comList(42) = Array( "Grafici Excel", "OWC.Chart" )
		comList(43) = Array( "Excel", "Excel.Application" )
		comList(44) = Array( "ADODB Stream", "ADODB.Stream" ) 
		comList(45) = Array( "CDOSYS", "CDO.Message" )

	'This function was modified from the work of Rob Risner.
	'http://www.planetsourcecode.com/xq/ASP/txtCodeId.6731/lngWId.4/qx/vb/scripts/ShowCode.htm

	Function IsAvailable( comIdentity )
		On Error Resume Next
		IsAvailable = False
		Err = 0
		Set xTestObj = Server.CreateObject( comIdentity )
		If Err = 0 Then IsAvailable = True
		Set xTestObj = Nothing
		Err = 0
	End Function

	Public Function CheckCOM()
		Avail = 0
		strTxt = "<table cellpadding=3 cellspacing=3 border=0 align=center width=300>" & vbNewLine
		For Idx = LBound( comList ) To UBound( comList )
			Provider = Idx
			strTxt = strTxt & vbTab & "<tr><td width=200><font class=norm>" & comList(Idx)(0) & "</font></td>"
			strTxt = strTxt & "<td align=right width=20><font class=norm>[</font></td>"
			If IsAvailable( comList(Idx)(1) ) Then
				strTxt = strTxt & "<td align=center><font face=tahoma size=2 color=blue>Available</font></td>"
				Avail = Avail + 1
			Else
				strTxt = strTxt & "<td align=center><font class=norm>Unavailable</font></td>"
			End If
			strTxt = strTxt & "<td align=left width=20><font class=norm>]</font></td></tr>" & vbNewLine
		Next
		strTxt = strTxt & vbTab & "<tr><td colspan=4 height=30><font class=norm><font color=blue>" & Avail & "</font> of "
		strTxt = strTxt & UBound( comList ) + 1 & " supported components are available.</font></td></tr>" & vbNewLine
		CheckCOM = strTxt & "</table>" & vbNewLine
	End Function

	Response.Write( CheckCOM() )
%>

<p align=center>
	<font class=arial size=2><a href=http://www.intelidev.com/goto.asp?comchecker>COMChecker</a> version 1.1 - by James Lindën<br>© 2001 Ticluse Teknologi, All rights reserved.</font>
</p>
</body>
</html>