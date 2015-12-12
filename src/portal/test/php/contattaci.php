<?PHP
	// recupero la lista di parametri dal FORM
	$nome=$_POST['nome'];
	$cognome=$_POST['cognome'];
	$email=$_POST['email'];
	$telefono=$_POST['telefono'];
	$indirizzo=$_POST['indirizzo'];
	$zipcode=$_POST['zipcode'];
	$city=$_POST['citta'];
	$nazione=$_POST['nazione'];
	$testo=$_POST['testo'];


	// Configuro i dati di invio della mail
	// (destinatario, mittente, oggetto e corpo)
	$mail_to      = $_POST['mail_to'];
	$mail_from    = "info@blackholenet.com";
	$mail_subject = "Invio mail";


$_mail_body_html = '<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">';
$_mail_body_html = $_mail_body_html .'<html>';
$_mail_body_html = $_mail_body_html .'<head>';
$_mail_body_html = $_mail_body_html .'<title>invio mail contatti</title>';
$_mail_body_html = $_mail_body_html .'<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">';
$_mail_body_html = $_mail_body_html .'</head>';
$_mail_body_html = $_mail_body_html .'<body>';
$_mail_body_html = $_mail_body_html .'<div>';		
$_mail_body_html = $_mail_body_html .'nome:&nbsp;'.$nome.'<br><br>';
$_mail_body_html = $_mail_body_html .'cognome:&nbsp;'.$cognome.'<br><br>';
$_mail_body_html = $_mail_body_html .'email:&nbsp;'.$email.'<br><br>';
$_mail_body_html = $_mail_body_html .'telefono:&nbsp;'.$telefono.'<br><br>';
$_mail_body_html = $_mail_body_html .'indirizzo:&nbsp;'.$indirizzo.'<br><br>';
$_mail_body_html = $_mail_body_html .'zipcode:&nbsp;'.$zipcode.'<br><br>';
$_mail_body_html = $_mail_body_html .'città:&nbsp;'.$city.'<br><br>';
$_mail_body_html = $_mail_body_html .'nazione:&nbsp;'.$nazione.'<br><br>';
$_mail_body_html = $_mail_body_html .'messaggio:&nbsp;'.$testo.'<br><br>';
$_mail_body_html = $_mail_body_html .'</div>';
$_mail_body_html = $_mail_body_html .'</body>';
$_mail_body_html = $_mail_body_html .'</html>';

	$mail_body    = $_mail_body_html;

	// Specifico le intestazioni per il formato HTML 
	$mail_in_html  = "MIME-Version: 1.0\r\n";
	$mail_in_html .= "Content-type: text/html; charset=iso-8859-1\r\n";
	$mail_in_html .= "From: <$mail_from>";

	// Invio la mail
	if (mail($mail_to, $mail_subject, $mail_body, $mail_in_html))
	{
	print "Email inviata con successo!";
	}
?>