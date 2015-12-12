<!-- #include virtual="/editor/include/IncludeObjectList.inc" -->
<!-- #include virtual="/editor/include/Paginazione.inc" -->
<!-- #include file="include/init.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!-- #include virtual="/editor/include/initCommonMeta.inc" -->
<!-- #include virtual="/editor/include/initCommonJs.inc" -->
<script language="JavaScript">
/**
 * DHTML date validation script. Courtesy of SmartWebby.com (http://www.smartwebby.com/dhtml/)
 */
// Declaring valid date character, minimum year and maximum year
var dtCh= "/";
var minYear=1900;
var maxYear=2100;

function isInteger(s){
	var i;
    for (i = 0; i < s.length; i++){   
        // Check that current character is number.
        var c = s.charAt(i);
        if (((c < "0") || (c > "9"))) return false;
    }
    // All characters are numbers.
    return true;
}

function stripCharsInBag(s, bag){
	var i;
    var returnString = "";
    // Search through string's characters one by one.
    // If character is not in bag, append to returnString.
    for (i = 0; i < s.length; i++){   
        var c = s.charAt(i);
        if (bag.indexOf(c) == -1) returnString += c;
    }
    return returnString;
}

function daysInFebruary (year){
	// February has 29 days in any year evenly divisible by four,
    // EXCEPT for centurial years which are not also divisible by 400.
    return (((year % 4 == 0) && ( (!(year % 100 == 0)) || (year % 400 == 0))) ? 29 : 28 );
}
function DaysArray(n) {
	for (var i = 1; i <= n; i++) {
		this[i] = 31
		if (i==4 || i==6 || i==9 || i==11) {this[i] = 30}
		if (i==2) {this[i] = 29}
   } 
   return this
}

function isDate(dtStr){
	var daysInMonth = DaysArray(12)
	var pos1=dtStr.indexOf(dtCh)
	var pos2=dtStr.indexOf(dtCh,pos1+1)
	var strDay=dtStr.substring(0,pos1)
	var strMonth=dtStr.substring(pos1+1,pos2)
	var strYear=dtStr.substring(pos2+1)
	strYr=strYear
	if (strDay.charAt(0)=="0" && strDay.length>1) strDay=strDay.substring(1)
	if (strMonth.charAt(0)=="0" && strMonth.length>1) strMonth=strMonth.substring(1)
	for (var i = 1; i <= 3; i++) {
		if (strYr.charAt(0)=="0" && strYr.length>1) strYr=strYr.substring(1)
	}
	month=parseInt(strMonth)
	day=parseInt(strDay)
	year=parseInt(strYr)
	if (pos1==-1 || pos2==-1){
		alert("<%=langEditor.getTranslated("backend.logs.lista.js.alert.date_format")%>")
		return false
	}
	if (strMonth.length<1 || month<1 || month>12){
		alert("<%=langEditor.getTranslated("backend.logs.lista.js.alert.valid_month")%>")
		return false
	}
	if (strDay.length<1 || day<1 || day>31 || (month==2 && day>daysInFebruary(year)) || day > daysInMonth[month]){
		alert("<%=langEditor.getTranslated("backend.logs.lista.js.alert.valid_day")%>")
		return false
	}
	if (strYear.length != 4 || year==0 || year<minYear || year>maxYear){
		alert("<%=langEditor.getTranslated("backend.logs.lista.js.alert.valid_year")%> "+minYear+" and "+maxYear)
		return false
	}
	if (dtStr.indexOf(dtCh,pos2+1)!=-1 || isInteger(stripCharsInBag(dtStr, dtCh))==false){
		alert("<%=langEditor.getTranslated("backend.logs.lista.js.alert.valid_date")%>")
		return false
	}
return true
}

function sendSearchLogs(){
	var dtf=document.form_search.dta_from
	var dtt=document.form_search.dta_to
	if(dtf.value != ""){
		if (isDate(dtf.value)==false){
			dtf.focus()
			return
		}
		//document.form_search.dta_from.value = document.form_search.dta_from.value + " 00:00:00"
		//alert(document.form_search.dta_ins_search.value);
		
	}
	if(dtt.value != ""){
		if (isDate(dtt.value)==false){
			dtt.focus()
			return
		}
		//document.form_search.dta_to.value = document.form_search.dta_to.value + " 23:59:59"
		//alert(document.form_search.dta_ins_search.value);
		
	}
    document.form_search.submit();
 }
 
 

function sendDeleteLogs(){
	var dtf=document.form_search.dta_from
	var dtt=document.form_search.dta_to
	if(dtf.value != ""){
		if (isDate(dtf.value)==false){
			dtf.focus()
		return
		}
		//document.form_search.dta_from.value = document.form_search.dta_from.value + " 00:00:00"
		//alert(document.form_search.dta_ins_search.value);

	}
	
	if(dtt.value != ""){
		if (isDate(dtt.value)==false){
			dtt.focus()
		return
		}
		//document.form_search.dta_to.value = document.form_search.dta_to.value + " 23:59:59";
		//alert(document.form_search.dta_ins_search.value);

	}
	
	document.form_search.delete_log.value="1";
	
	if(confirm("<%=langEditor.getTranslated("backend.logs.lista.js.alert.confirm_log_delete")%>")){
		document.form_search.submit();
	}else
		return;
 }

$(function() {
	$('#dta_from').datepicker({
		dateFormat: 'dd/mm/yy',
		changeMonth: true,
		changeYear: true
	});
	$('#ui-datepicker-div').hide();	
});

$(function() {
	$('#dta_to').datepicker({
		dateFormat: 'dd/mm/yy',
		changeMonth: true,
		changeYear: true
	});
	$('#ui-datepicker-div').hide();	
});
</script>
</head>
<body>
<div id="backend-warp">
	<!-- #include virtual="/editor/include/header.inc" -->	
	<div id="container">
		<%cssClass="LL"%>
		<!-- #include virtual="/editor/include/menu.inc" -->
		<div id="backend-content">
	<table class="principal" cellpadding="0" cellspacing="0">
	<form action="<%=Application("baseroot") & "/editor/logs/ListaLogs.asp"%>" method="post" name="form_search">
	<INPUT TYPE="hidden" NAME="delete_log" VALUE="">
	<input type="hidden" value="1" name="page">
	<tr> 
	<th>&nbsp;</th>
	<th><%=langEditor.getTranslated("backend.logs.lista.table.header.type")%></th>
	<th><%=langEditor.getTranslated("backend.logs.lista.table.header.date_from")%></th>
	<th><%=langEditor.getTranslated("backend.logs.lista.table.header.date_to")%></th>
	</tr>
	<tr height="40"> 
	<td align="center">
	<input type="button" class="buttonForm" hspace="4" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.logs.lista.button.cerca.label")%>" onclick="javascript:sendSearchLogs();" />&nbsp;&nbsp;<input type="button" class="buttonForm" hspace="4" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.logs.lista.button.cancella.label")%>" onclick="javascript:sendDeleteLogs();" />
	</td>
	<td>			  
	  <select name="log_type" class="formFieldChangeStato">
	  <option value=""></option>
	  <option value="debug" <%if(paramType = "debug") then response.write("selected") end if%>><%=langEditor.getTranslated("backend.logs.lista.table.select.option.degug")%></option>
	  <option value="info" <%if(paramType = "info") then response.write("selected") end if%>><%=langEditor.getTranslated("backend.logs.lista.table.select.option.info")%></option>
	  <option value="error" <%if(paramType = "error") then response.write("selected") end if%>><%=langEditor.getTranslated("backend.logs.lista.table.select.option.error")%></option>
	  </select>		
	</td>
	<td><input type="text" value="<%=paramDateFrom%>" name="dta_from" id="dta_from" class="formFieldTXT"></td>
	<td><input type="text" value="<%=paramDateTo%>" name="dta_to" id="dta_to" class="formFieldTXT"></td>
	</tr>
          </form>
	</table>
	<table class="principal" border="0" align="top" cellpadding="0" cellspacing="0">
              <tr> 
	      <th>&nbsp;</th>
                <th><%=langEditor.getTranslated("backend.logs.lista.table.header.msg")%></th>
                <th><%=langEditor.getTranslated("backend.logs.lista.table.header.usr")%></th>
                <th><%=langEditor.getTranslated("backend.logs.lista.table.header.type")%></th>
	      <th><%=langEditor.getTranslated("backend.logs.lista.table.header.date_insert")%></th>
              </tr> 
		<%

		if(paramType = "") then paramType = null end if
		if(paramDateFrom = "") then paramDateFrom = null end if
		if(paramDateTo = "") then paramDateTo = null end if
		
		Dim totPages, hasLog
		hasLog = false
		
		on error Resume Next
		Set objListaLog = objLog.getListaLogs(paramType,paramDateFrom,paramDateTo)	

		if(objListaLog.Count > 0) then
			hasLog = true
		end if
			
		if Err.number <> 0 then
		end if	
		
		if(hasLog) then			
				
				Dim intCount
				intCount = 0
				
				Dim newsCounter, iIndex, objTmpLog, objTmpLogKey, FromLog, ToLog, Diff
				iIndex = objListaLog.Count
				FromLog = ((numPage * itemsXpage) - itemsXpage)
				Diff = (iIndex - ((numPage * itemsXpage)-1))
				if(Diff < 1) then
					Diff = 1
				end if
				
				ToLog = iIndex - Diff
				
				totPages = iIndex\itemsXpage
				if(totPages < 1) then
					totPages = 1
				elseif((iIndex MOD itemsXpage <> 0) AND not ((totPages * itemsXpage) >= iIndex)) then
					totPages = totPages +1	
				end if		
				
				Dim styleRow, styleRow2
				styleRow2 = "table-list-on"							
						
				objTmpLog = objListaLog.Items
				objTmpLogKey=objListaLog.Keys		
				for newsCounter = FromLog to ToLog
					styleRow = "table-list-off"
					if(newsCounter MOD 2 = 0) then styleRow = styleRow2 end if%>
				<tr class="<%=styleRow%>">
					<%Set objTmpLog0 = objTmpLog(newsCounter)%>					
					<td>&nbsp;</td>
					<td><%=objTmpLog0.getLogMsg()%></td>
					<td><%response.write(objTmpLog0.getLogUsr())%></td>
					<td><%response.write(objTmpLog0.getLogTipo())%></td>
					<td><%response.write(objTmpLog0.getLogData())%></td>               
				</tr>				
				<%intCount = intCount +1
				next
				Set objListaLog = nothing
				%>
              <tr> 
		<form action="<%=Application("baseroot") & "/editor/logs/ListaLogs.asp"%>" method="post" name="item_x_page">
			<th colspan="5" align="left">
				<input type="text" name="items" class="formFieldTXTNumXPage" value="<%=itemsXpage%>" title="<%=langEditor.getTranslated("backend.commons.lista.table.alt.item_x_page")%>" onblur="javascript:submit();" onkeypress="javascript:return isInteger(event);">
				<%		
				'**************** richiamo paginazione
				call PaginazioneFrontend(totPages, numPage, strGerarchia, "/editor/logs/ListaLogs.asp", "&items="&itemsXpage)
				%>
				</th>
				</form>
              </tr>
	      <%end if
		  Set objLog = Nothing%>
		</table>
		<br><input type="button" class="buttonForm" hspace="2" vspace="4" border="0" align="absmiddle" value="<%=langEditor.getTranslated("backend.logs.lista.button.download_excel.label")%>" onclick="javascript:openWinExcel('<%=Application("baseroot")&"/editor/report/create-log-excel.asp"%>','crea_excel',400,400,100,100);" />		
		<br/><br/>
		</div>
	</div>
	<!-- #include virtual="/editor/include/bottom.inc" -->
</div>
</body>
</html>