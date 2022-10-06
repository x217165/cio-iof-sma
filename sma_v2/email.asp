<%@ LANGUAGE=VBSCRIPT %>
<%Option Explicit
On Error Resume Next
Response.Buffer = True
%>
<!-- #include file = "smaProcs.inc" -->
<%
private function formatDistribList (distList)
'this function strips ending semi-colons off of 
'the passed in distirbution list

	distList = Trim(distList)
	
	if Right(distList, 1) = ";" then
		distList = Left(distList, len(distList) - 1)
	end if
	
	formatDistribList = distList
	
end function

dim strFrom, strTo, strCC, strBCC, strSubject, strBody
dim arrEmailList, txtEmailAddress, x
dim objRegExp, colMatch

if Request("txtFrom") <> "" then
	'email submitted for delivery
	strFrom = Request("txtFrom")
	strTo = Request("txtTo")
	strCC = Request("txtCC")
	strBCC = Request("txtBCC")
	strSubject = Request("txtsubject")
	strBody = Request("txtBody")
	
	'check TO list
	if strTo <> "" then
		arrEmailList = split(formatDistribList(strTo), ";")
		for x=0 to UBound(arrEmailList)
			txtEmailAddress = Trim(arrEmailList(x))
			'check address syntax
			set objRegExp = new RegExp
			with objRegExp
				.Pattern = "<.+@.+\..+>"
				.Global = true
				.IgnoreCase = true
			end with
			set colMatch = objRegExp.Execute(txtEmailAddress)
			if colMatch.count = 0 then
				DisplayError "BACK", "", 1001, "EMAIL ADDRESS ERROR - [" &txtEmailAddress&"]", "The text you entered does not comply to the standard email address format ""Name <name@company.com>"". Please check it and try again. If the error persists please contact your system administrator."
				Response.end
			end if
		next
	end if
	if err then
		DisplayError "BACK", "", err.Number, "UNEXPECTED ERROR", err.Description
		Response.End
	end if
	'check CC list
	if strCC <> "" then
		arrEmailList = split(formatDistribList(strCC), ";")
		for x=0 to UBound(arrEmailList)
			txtEmailAddress = Trim(arrEmailList(x))
			'check address syntax
			set objRegExp = new RegExp
			with objRegExp
				.Pattern = "<.+@.+\..+>"
				.Global = true
				.IgnoreCase = true
			end with
			set colMatch = objRegExp.Execute(txtEmailAddress)
			if colMatch.count = 0 then
				DisplayError "BACK", "", 1001, "EMAIL ADDRESS ERROR - ["&txtEmailAddress&"]", "The text you entered does not comply to the standard email address format ""Name <name@company.com>"". Please check it and try again. If the error persists please contact your system administrator."
				Response.end
			end if
		next
	end if
	
	'check BCC list
	if strBCC <> "" then
		arrEmailList = split(formatDistribList(strBCC), ";")
		for x=0 to UBound(arrEmailList)
			txtEmailAddress = Trim(arrEmailList(x))
			'check address syntax
			set objRegExp = new RegExp
			with objRegExp
				.Pattern = "<.+@.+\..+>"
				.Global = true
				.IgnoreCase = true
			end with
			set colMatch = objRegExp.Execute(txtEmailAddress)
			if colMatch.count = 0 then
				DisplayError "BACK", "", 1001, "EMAIL ADDRESS ERROR - ["&txtEmailAddress&"]", "The text you entered does not comply to the standard email address format ""Name <name@company.com>"". Please check it and try again. If the error persists please contact your system administrator."
				Response.end
			end if
		next
	end if
	
	'send email
	'SendEmail(strFrom, strTo, strCC, strBCC, strSubject, strBody)
	SendEmail strFrom, strTo, strCC, strBCC, strSubject, strBody
	if err then
		DisplayError "BACK", "", err.Number, "CANNOT SEND EMAIL", err.Description
		Response.End
	else
		Response.Write "<script type=""text/javascript"">alert('Email sent successfully');window.close();</script>"
		Response.End
	end if
else
	'first access - set parameters
	strFrom = Request.Cookies("UserInformation")("email_address")
	strTo = Request.Cookies("txtEmailTo")
	strCC = Request.Cookies("txtEmailCC")
	strBCC = Request.Cookies("txtEmailBCC")
	strSubject = unescape(Request.Cookies("txtEmailSubject"))
	strBody = unescape(Request.Cookies("txtEmailBody"))
	
	'Response.Write "<BR>From: " & strFrom
	'Response.Write "<BR>  To: " & strTo
	'Response.Write "<BR>  CC: " & strCC
	'Response.Write "<BR> BCC: " & strBCC
	'Response.Write "<BR>Subj: " & strSubject
	'Response.Write "<BR>Body: " & strBody
	'Response.End
end if

if err then
	DisplayError "BACK", "", err.Number, "UNEXPECTED ERROR", err.Description
end if
%>
<HTML>
<HEAD>
<link href="stylesheets/styles.css" rel="stylesheet" type="text/css">
<TITLE>SMA - Email Notification Service</TITLE>
</HEAD>
<SCRIPT type="text/javascript" language="javascript" src="AccessLevels.js"></SCRIPT>
<SCRIPT type="text/javascript" language="javascript" src="GeneralJavaFunctions.js"></SCRIPT>
<script type="text/javascript" language="javascript">
function form_onSubmit(){
	if (document.frmEmail.txtFrom.value == '') {alert('You cannot send emails since you don\'t have a personal email address.\nPlease contact your system administrator.'); return(false);}
	return(true);
}

function fct_LookupContact(strField){
	document.frmEmail.hdnDestination.value = strField;
	SetCookie("WinName", 'Popup');
	SetCookie("Case", "E");
	var wndLookup = window.open('SearchFrame.asp?fraSrc=Contact', 'Popup', 'top=100, left=150, height=600, width=800' ) ;
}

function btnCancel_onclick(){
//	if (confirm('Do you want to discard this email?')){
		window.close();
//	}
}

function body_onLoad() {
	DeleteCookie("txtEmailTo");
	DeleteCookie("txtEmailCC");
	DeleteCookie("txtEmailBCC");
	DeleteCookie("txtEmailSubject");
	DeleteCookie("txtEmailBody");
}
</script>
<BODY language="javascript" onLoad="body_onLoad();">
<FORM name="frmEmail" action="email.asp" method="POST" onsubmit="return form_onSubmit();">
<input type="hidden" name="hdnDestination" value="">
<TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
<THEAD>
	<TH colspan=2>Email Notification</TH>
</THEAD>
<TR>
    <TD align="right" valign="top"><br>From:</TD>
    <TD><br><TEXTAREA readonly rows=1 cols=110 name="txtFrom"><%=strFrom%></TEXTAREA></TD>
</TR>
<TR>
	<TD width="10%" align="right" valign="top">To:</TD>
	<TD width="90%"><TEXTAREA rows=3 cols=110 name="txtTO"><%=strTo%></TEXTAREA>&nbsp;<INPUT name=btnToLookup type=button value="..." class="button" onclick="fct_LookupContact('to')">&nbsp;<INPUT name=btnToClear type=button value="X" class="button" style="color:red" onclick="document.frmEmail.txtTO.value=''"></TD>
</TR>
<TR>
    <TD align="right" valign="top">CC:</TD>
    <TD><TEXTAREA rows=1 cols=110 name="txtCC"><%=strCC%></TEXTAREA>&nbsp;<INPUT name=btnCCLookup type=button value="..." class="button" onclick="fct_LookupContact('cc')">&nbsp;<INPUT name=btnCCClear type=button value="X" class="button" style="color:red" onclick="document.frmEmail.txtCC.value=''"></TD>
</TR>
<TR>
    <TD align="right" valign="top">BCC:</TD>
    <TD><TEXTAREA rows=1 cols=110 name="txtBCC"><%=strBCC%></TEXTAREA>&nbsp;<INPUT name=btnBCCLookup type=button value="..." class="button" onclick="fct_LookupContact('bcc')">&nbsp;<INPUT name=btnBCCClear type=button value="X" class="button" style="color:red" onclick="document.frmEmail.txtBCC.value=''"></TD>
</TR>
<TR>
    <TD align="right" valign="top">Subject:</TD>
    <TD><INPUT type="text" name="txtSubject" style="width:100%" value="<%=strSubject%>"></TD>
</TR>
<TR>
	<td colspan=2><hr></td>
</TR>
<TR>
    <TD align="right" valign="top">Message:</TD>
	<TD>
		<TEXTAREA style="width:100%" rows=21 name="txtBody"><%=strBody%></TEXTAREA>
		<div align=center>
		<INPUT type="submit" name="btnSend" value="Send" class="button" title="Click here to send the email.">
		<INPUT type="button" name="btnCancel" onclick="btnCancel_onclick();" value="Cancel" class="button" title="Click here to cancel the email and close this form.">
		</div><br>
	</TD>
</TR>
</TABLE>
</FORM>
</BODY>
</HTML>
