<%@ Language=VBScript %>
<% Option Explicit %>
<!--#include file="smaConstants.inc"-->
<!--#include file="smaProcs.inc"-->
<!--#include file="databaseconnect.asp"-->

<!--
*********************************************************************************************
* Page name:	ContactCleanEntry.asp                                                       *
* Purpose:		To dynamically accept the parameters required to perform Contact cleanup.   *
*				Results are displayed via ContactCleanList.asp                              *
*                                                                                           *
* Created by:	Dan S. Ty	03/30/2002                                                      *
*                                                                                           *
*********************************************************************************************
*		Date		Author			Changes/enhancements made                               *
*       -----		------		------------------------------------------------------      *
*                                                                                           *
*********************************************************************************************
-->
<%

'Check Access rights - check other locations in this page.
dim intAccessLevel
intAccessLevel = CInt(CheckLogon(strConst_ESDCleanup))
If (intAccessLevel And intConst_Access_ReadOnly) <> intConst_Access_ReadOnly  Then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to ESD Cleanup functions. Please contact your system administrator"
End If

dim strSQL

'if the page is called by a lookup function or by Quick Navigation drop-down box
'then following cookies will have values.
dim strContactName, strWinName
strContactName = Request.Cookies("ContactName")
strWinName	= Request.Cookies("WinName")  

%>
<HTML>
<HEAD>
	<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<LINK href="stylesheets/styles.css" rel="stylesheet" type="text/css">
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<TITLE>Service Management Application</TITLE>
	<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
	<script type="text/javascript" SRC="AccessLevels.js"></SCRIPT>
	<script type="text/javascript">
	<!--hide the script from old browsers 
	
	var intAccessLevel = <%=intAccessLevel%>;
	
	// set section title
setPageTitle("SMA - Contact Cleanup");
	

function window_onload() {

	var strContactName,strWinName;
	strWinName = document.frmContactCleanEntry.hdnWinName.value;
	DeleteCookie("WinName");
}

function btnClear_onclick() {
	  
	document.frmContactCleanEntry.txtFRContact.value = "" ;
	document.frmContactCleanEntry.txtTOContact.value = ""  ;
	document.frmContactCleanEntry.selAction.selectedIndex = 0 ;
}
	
function btnContactLookup_onclick(WhichContact) {

	if (document.frmContactCleanEntry.hdnFRContactLName.value == "" &&
	    document.frmContactCleanEntry.hdnFRContactFName.value != "" &&
	    document.frmContactCleanEntry.hdnFRContactMName.value != "" &&
	    document.frmContactCleanEntry.hdnTOContactLName.value != "" &&
	    document.frmContactCleanEntry.hdnTOContactFName.value != "" &&
	    document.frmContactCleanEntry.hdnTOContactMName.value != "" &&
	    WhichContact == 'F') {
		SetCookie("LName", document.frmContactCleanEntry.hdnTOContactLName.value);
		SetCookie("FName", document.frmContactCleanEntry.hdnTOContactFName.value);
	}

	if (document.frmContactCleanEntry.hdnTOContactLName.value == "" &&
	    document.frmContactCleanEntry.hdnTOContactFName.value != "" &&
	    document.frmContactCleanEntry.hdnTOContactMName.value != "" &&
	    document.frmContactCleanEntry.hdnFRContactLName.value != "" &&
	    document.frmContactCleanEntry.hdnFRContactFName.value != "" &&
	    document.frmContactCleanEntry.hdnFRContactMName.value != "" &&
	    WhichContact == 'T') {
		SetCookie("LName", document.frmContactCleanEntry.hdnFRContactLName.value);
		SetCookie("FName", document.frmContactCleanEntry.hdnFRFContactName.value);
	}

	SetCookie("WinName", 'Popup');
	SetCookie("Case", WhichContact);
	window.open('SearchFrame.asp?fraSrc=Contact', 'Popup', 'top=50, left=130, height=600, width=870' ) ;
}	

function validate(theForm) {

	var bolConfirm
	
	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) 
		{	
			alert('Access denied. Please contact your system administrator.'); 
			return (false);
		}
		else
		{
			if (document.frmContactCleanEntry.txtFRContact.value == "" ) 
			{   
				alert('Please select a "From Contact Name"'); 
				document.frmContactCleanEntry.btnFRContactLookup.focus();  
				return(false);
			}
			if (document.frmContactCleanEntry.txtTOContact.value == "" ) 
			{   
				alert('Please select "To Contact Name"'); 
				document.frmContactCleanEntry.btnTOContactLookup.focus();  
				return(false);
			}	
			if (document.frmContactCleanEntry.selAction.value == "" ) 
			{   
				alert('Please select a Cleanup Action'); 
				document.frmContactCleanEntry.selAction.focus();  
				return(false);
			}	
			if (document.frmContactCleanEntry.hdnFRContactID.value == document.frmContactCleanEntry.hdnTOContactID.value ) 
			{   
				alert('"From Contact ID" and "To Contact ID" should be different'); 
				document.frmContactCleanEntry.btnFRContactLookup.focus();  
				return(false);
			}	
			else				
			{
				document.frmContactCleanEntry.submit();
				return(true);
			}
		}
}	
//-->end hide script	
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload();" >
<FORM name = frmContactCleanEntry method=post action="ContactCleanList.asp" target="fraResult" onSubmit="return validate(this);" >

	<!-- hidden variables -->
	<INPUT id=hdnWinName        name=hdnWinName       type=hidden value= "<%=strWinName%>">

	<INPUT id=hdnFRContactID    name=hdnFRContactID    type=hidden value= "">
	<INPUT id=hdnFRContactLName name=hdnFRContactLName type=hidden value= "">
	<INPUT id=hdnFRContactFName name=hdnFRContactFName type=hidden value= "">
	<INPUT id=hdnFRContactMName name=hdnFRContactMName type=hidden value= "">

	<INPUT id=hdnTOContactID    name=hdnTOContactID    type=hidden value= "">
	<INPUT id=hdnTOContactLName name=hdnTOContactLName type=hidden value= "">
	<INPUT id=hdnTOContactFName name=hdnTOContactFName type=hidden value= "">
	<INPUT id=hdnTOContactMName name=hdnTOContactMName type=hidden value= "">

<TABLE border="0" width="100%">    
    <thead><tr><td colspan=4 align=left>Contact Cleanup Parameters</td></tr></thead>
	<tbody>	
l
	<TR><TD align=right width=25%>From Contact Name<font color=red>*</font></TD>
		<TD align=left width=50% colspan=3>
			<input name=txtFRContact type=text disabled size=70 maxlength=70 value="">
			<INPUT align=right type="button"  name=btnFRContactLookup   value="..." onclick="return btnContactLookup_onclick('F')" tabindex=1></TD></TR>

	<TR><TD align=right width=25%>TO Contact Name<font color=red>*</font></TD>
		<TD align=left width=50% colspan=3>
			<input name=txtTOContact type=text disabled size=70 maxlength=70 value="<%if request("hdnTOContactID") <> 0 then Response.Write "(" & routineHTMLString(Request("hdnTOContactName")) & ")" else Response.Write null end if%>" onChange="fct_onChange();">
			<INPUT align=right type="button"  name=btnTOContactLookup   value="..." onclick="return btnContactLookup_onclick('T')" tabindex=2></TD></TR>

	<TR><TD align=right width=15% nowrap>Cleanup Action <font color=red>*</font></TD>
		<TD width=35%>
			<select id=selAction name=selAction tabindex=3 style="width: 110">
				Response.write "<option value="A">Amalgamate</option>  & vbCrLf
				Response.write "<option value="D">De-activate</option> & vbCrLf
				Response.write "<option value="R">Re-activate</option> & vbCrLf
				Response.write "<option value="S">Smart Fix  </option> & vbCrLf</select></TD></TR>
	<TR><TD></TD>
		<TD align=right colspan=2>
			<INPUT id=btnClear name=btnClear type=button style="width: 2cm" value=Clear LANGUAGE=javascript onclick="return btnClear_onclick()">&nbsp;&nbsp;
			<INPUT id=btnGo    name=btnGo    type=submit style="width: 2cm" value=Go > &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD></TR>
	</TABLE>
</FORM>
</BODY>
</HTML>
