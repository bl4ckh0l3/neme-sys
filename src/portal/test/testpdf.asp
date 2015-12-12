<!-- #include virtual="/common/include/fpdf.asp" -->
<%
Set pdf=CreateJsObject("FPDF")
pdf.CreatePDF()
pdf.SetPath(Application("baseroot") & "/common/include/fpdf/")
pdf.Open()
pdf.AddPage()

'pdf.Text 10,10,"Questo è un testo che va a capo volta numero "
'pdf.Text 100,10,"altro testo"
'pdf.Text 200,10,"altro testo ancora"

'pdf.MultiCell 100,20,"",0,0,"L",1
'pdf.SetFillColor "148" , "190" , "0"
pdf.Image Application("baseroot") & "/common/img/logo-sanident.jpg",10,1,90,25
pdf.ln(20)
'pdf.SetFillColor "255" , "255" , "255"

m=DatePart("m",Date())
'pdf.write "20","month: "&m
yyyy=DatePart("yyyy",Date())
'pdf.write "20","year: "&yyyy
d=CDate(yyyy&"-"&m-1&"-"&"1")

pdf.SetFont "Arial","B",9
'pdf.write "20","elenco compensi dal "&FormatDateTime(d,2) & " a oggi"
pdf.MultiCell 100, 8, "Medico: pinco pallino", 1, "L"
pdf.ln(2)
pdf.MultiCell 100, 8, "Elenco compensi dal "&FormatDateTime(d,2) & " a oggi", 0, "L"
pdf.ln(5)

pdf.SetFont "Arial","",6
for counter=1 to 100
	if(counter MOD 3=0)then
		pdf.MultiCell 50, 8, "Questo è un testo che va a capo "&counter, 1, "L"
	else
		pdf.Cell 50,8,"Questo è un testo "&counter,1,0,"L"
	end if
next

pdf.Close()
filewrite=server.mappath(Application("baseroot") & "/public/test.pdf")
pdf.Output(filewrite)
Set pdf=nothing
%>