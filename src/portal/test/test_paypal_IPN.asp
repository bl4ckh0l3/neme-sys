<!--********************************** ESEMPIO SCARICATO Da PAYPAL DI IPN *********************-->

<%@LANGUAGE="VBScript"%>
<%

' dim some variables
Dim Item_name, Item_number, Payment_status, Payment_amount
Dim Txn_id, Receiver_email, Payer_email
Dim objHttp, str

'define subroutine to handle "all" payments ##
sub allPayments()  ' begin sub ###########################################################

set conn = Server.CreateObject("ADODB.connection")
conn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("paypal.mdb")
set rs = Server.CreateObject("ADODB.Recordset")
rs.open "Payments", conn, 2, 2 
rs.addnew

'add records to the Payments table

rs.Fields("payment_date") = payment_date
rs.Fields("pp_txn_id") = txn_id
rs.Fields("parent_txn_id") = parent_txn_id
rs.Fields("payment_status") = payment_status
rs.Fields("pending_reason") = pending_reason
rs.Fields("reason_code") = reason_code
rs.Fields("txn_type") = txn_type
rs.Fields("payment_type") = payment_type
rs.Fields("mc_gross") = mc_gross
rs.Fields("mc_fee") = mc_fee
rs.Fields("payment_currency") = mc_currency
rs.Fields("settle_amount") = settle_amount
rs.Fields("settle_currency") = settle_currency
rs.Fields("exchange_rate") = exchange_rate
rs.Fields("payer_email") = payer_email
rs.Fields("payment_status") = payer_status
rs.Fields("cust_firstname") = first_name
rs.Fields("cust_lastname") = last_name
rs.Fields("cust_biz_name") = payer_business_name
rs.Fields("gift_address_name") = address_name
rs.Fields("cust_address_street") = address_street
rs.Fields("cust_address_city") = address_city
rs.Fields("cust_address_state") = address_state
rs.Fields("cust_address_zip") = address_zip
rs.Fields("cust_address_country") = address_country
rs.Fields("cust_address_status") = address_status
rs.Fields("notify_version") = notify_version
rs.Fields("for_auction") = for_auction
rs.Fields("auction_buyer_id") = auction_buyer_id
rs.Fields("auction_closing_date") = auction_closing_date

'finish up
rs.Update
rs.Close
Set rs = Nothing
Set conn = Nothing
end sub  'end sub ###########################################################################


'define subroutine to handle subscription payments ##
sub subscriptionPayments()  ' begin sub ###########################################################

set conn = Server.CreateObject("ADODB.connection")
conn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("paypal.mdb")
set rs = Server.CreateObject("ADODB.Recordset")
rs.open "Subscriptions", conn, 2, 2 
rs.addnew

'add records to the Subscriptions table
rs.Fields("subscription_id") = subscr_id
rs.Fields("subscription_date") = subscr_date
rs.Fields("subscr_txn_type") = txn_type
rs.Fields("sub_period1") = period1
rs.Fields("sub_period2") = period2
rs.Fields("sub_period3") = period3
rs.Fields("sub_mcamount1") = amount1
rs.Fields("sub_mcamount2") = amount2
rs.Fields("sub_mcamount3") = amount3
rs.Fields("sub_recurring") = recurring
rs.Fields("sub_reattempt") = reattempt
rs.Fields("sub_retry_at") = retry_at
rs.Fields("sub_recur_times") = recur_times
rs.Fields("sub_username") = username
rs.Fields("sub_password") = password

'finish up
rs.Update
rs.Close
Set rs = Nothing
Set conn = Nothing
end sub  'end sub ###########################################################################

'begin IPN handling
' read post from PayPal system and add 'cmd'
str = Request.Form & "&cmd=_notify-validate"

' post back to PayPal system to validate
set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
' set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP.4.0")
' set objHttp = Server.CreateObject("Microsoft.XMLHTTP")
objHttp.open "POST", "https://www.sandbox.paypal.com/cgi-bin/webscr", false
objHttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
objHttp.Send str

' assign posted variables to local variables
item_name = Request.Form("item_name")
item_number = Request.Form("item_number")
payment_status = Request.Form("payment_status")
txn_id = Request.Form("txn_id")
parent_txn_id = Request.Form("parent_txn_id")
receiver_email = Request.Form("receiver_email")
payer_email = Request.Form("payer_email")
reason_code = Request.Form("reason_code")
business = Request.Form("business")
quantity = Request.Form("quantity")
invoice = Request.Form("invoice")
custom = Request.Form("custom")
tax = Request.Form("tax")
option_name1 = Request.Form("option_name1")
option_selection1 = Request.Form("option_selection1")
option_name2 = Request.Form("option_name2")
option_selection2 = Request.Form("option_selection2")
num_cart_items = Request.Form("num_cart_items")
pending_reason = Request.Form("pending_reason")
payment_date = Request.Form("payment_date")
mc_gross = Request.Form("mc_gross")
mc_fee = Request.Form("mc_fee")
mc_currency = Request.Form("mc_currency")
settle_amount = Request.Form("settle_amount")
settle_currency = Request.Form("settle_currency")
exchange_rate = Request.Form("exchange_rate")
txn_type = Request.Form("txn_type")
first_name = Request.Form("first_name")
last_name = Request.Form("last_name")
payer_business_name = Request.Form("payer_business_name")
address_name = Request.Form("address_name")
address_street = Request.Form("address_street")
address_city = Request.Form("address_city")
address_state = Request.Form("address_state")
address_zip = Request.Form("address_zip")
address_country = Request.Form("address_country")
address_status = Request.Form("address_status")
payer_email = Request.Form("payer_email")
payer_id = Request.Form("payer_id")
payer_status = Request.Form("payer_status")
payment_type = Request.Form("payment_type")
notify_version = Request.Form("notify_version")
verify_sign = Request.Form("verify_sign")

'subscription information
subscr_date = Request.Form("subscr_date")
period1 = Request.Form("period1")
period2 = Request.Form("period2")
period3 = Request.Form("period3")
amount1 = Request.Form("mc_amount1")
amount2 = Request.Form("mc_amount2")
amount3 = Request.Form("mc_amount3")
recurring = Request.Form("recurring")
reattempt = Request.Form("reattempt")
retry_at = Request.Form("retry_at")
recur_times = Request.Form("recur_times")
username = Request.Form("username")
password = Request.Form("password")
subscr_id = Request.Form("subscr_id")

'auction information
for_auction = Request.Form("for_auction")
auction_buyer_id = Request.Form("auction_buyer_id")
auction_closing_date = Request.Form("auction_closing_date")


' Check notification validation
if (objHttp.status <> 200 ) then
' HTTP error handling
elseif (objHttp.responseText = "VERIFIED") then
' check that Payment_status=Completed
' check that Txn_id has not been previously processed
' check that Receiver_email is your Primary PayPal email
' check that Payment_amount/Payment_currency are correct
' process payment


'implement IPN handling logic for DB insertion '#########################################################


'decide what to do based on txn_type - using Select Case

Select Case txn_type
	Case "subscr_signup"
		subscriptionPayments()
	Case "subscr_payment"
		subscriptionPayments()
	Case "subscr_modify"
		subscriptionPayments()
	Case "subscr_failed"
		subscriptionPayments()
	Case "subscr_cancel"
		subscriptionPayments()
	Case "subscr_eot"
		subscriptionPayments()
	Case Else
		allPayments()
End Select


elseif (objHttp.responseText = "INVALID") then
' log for manual investigation
' add code to handle the INVALID scenario

else
' error
end if
set objHttp = nothing
%>