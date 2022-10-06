<%@ Language=VBScript %>
<% Option Explicit %>
<!--#include file="smaConstants.inc"-->
<!--#include file="smaProcs.inc"-->
<!--#include file="databaseconnect.asp"-->

<!--
*********************************************************************************************
* Page name:	XLSEntry.asp                                                                *
* Purpose:		To accept the parameters required to generate the Validation Spreadsheets.  *
*				Results are displayed via VXLSCustomer.asp, VXLSContact.asp                 *
*                 VXLSCustService.asp and VXLSServOrder.asp                                 *
*                                                                                           *
* Created by:	Dan S. Ty	03/31/2002                                                      *
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
	dim strCustomerName, strWinName
	strCustomerName = Request.Cookies("CustomerName")
	strWinName	= Request.Cookies("WinName")  

%>
<HTML>
<HEAD>
	<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<LINK href="stylesheets/styles.css" rel="stylesheet" type="text/css">
	<TITLE>Service Management Application</TITLE>
	<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
	<script type="text/javascript" SRC="AccessLevels.js"></SCRIPT>
	<script type="text/javascript">
	<!--hide the script from old browsers 
	
	var intAccessLevel = <%=intAccessLevel%>;
	
	// set section title
setPageTitle("SMA - Validation Spreadsheets");
	

function window_onload() {

	var strCustomerName,strWinName;
	strWinName = document.frmXLSEntry.hdnWinName.value;
	DeleteCookie("WinName");
}

function btnClear_onclick() {
	  
	document.frmXLSEntry.txtCustomer.value = "" ;
	document.frmXLSEntry.selXLS.selectedIndex = 0 ;
}
	
function btnCustomerLookup_onclick(WhichCustomer) {

	SetCookie("WinName", 'Popup');
	SetCookie("ServiceEnd", WhichCustomer);
	window.open('SearchFrame.asp?fraSrc=Cust', 'Popup', 'top=50, left=180, height=600, width=750' );
}	

function btnGo_onclick() {
	if (document.frmXLSEntry.selXLS.value == "CR" ) 
		document.frmXLSEntry.action == "XLSCustList.asp";
//		else
//			if  (document.frmXLSEntry.selXLS.value == "CT" ) 
//			    (document.frmXLSEntry.action.value =  "XLSContList.asp";
//				 return(true);}
//			else
//				if  (document.frmXLSEntry.selXLS.value == "CS" ) 
//					(document.frmXLSEntry.action.value =  "XLSCustServList.asp";
//					 return(true);}
//				else
//					if  (document.frmXLSEntry.selXLS.value == "SO" ) 
//						(document.frmXLSEntry.action.value =  "XLSSOList.asp";
//						 return(true);}
	{document.frmXLSEntry.submit();}
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
			if (document.frmXLSEntry.txtCustomer.value == "" ) 
			{   
				alert('Please select a Customer Name'); 
				document.frmXLSEntry.btnCustomerLookup.focus();  
				return(false);
			}
			if (document.frmXLSEntry.selXLS.value == "" ) 
			{   
				alert('Please select a spreadsheet'); 
				document.frmXLSEntry.selXLS.focus();  
				return(false);
			}	
			else				
			{
				SetCookie("XLS", "Cust");
				document.frmXLSEntry.submit();
				return(true);
			}
		}
}	
//-->end hide script	
//<FORM name = frmXLSEntry method=post action=<%request("selXLS")%> target="fraResult" onSubmit="return validate(this);" >
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload();" >
<FORM name = frmXLSEntry method=post action="XLSCustServList.asp" target="fraResult" onSubmit="return validate(this);" >

	<!-- hidden variables -->
	<INPUT id=hdnWinName      name=hdnWinName      type=hidden value= "<%=strWinName%>">

	<INPUT id=hdnCustomerID   name=hdnCustomerID   type=hidden value= "">
	<INPUT id=hdnCustomerName name=hdnCustomerName type=hidden value= "">

<TABLE border="0" width="100%">    
    <thead><tr><td colspan=4 align=left>Validation Spreadsheet Parameters</td></tr></thead>
	<tbody>	

	<TR><TD align=right width=25%>Customer Name<font color=red>*</font></TD>
		<TD align=left width=50% colspan=3>
			<input name=txtCustomer type=text disabled size=50 maxlength=50 value="">
			<INPUT align=right type="button"  name=btnCustomerLookup   value="..." onclick="return btnCustomerLookup_onclick('X')" tabindex=1></TD></TR>

	<TR><TD align=right width=15% nowrap>Spreadsheet <font color=red>*</font></TD>
		<TD width=35%>
			<select id=selXLS name=selXLS tabindex=2 style="width: 150">
				Response.write "<option value=XLSCustServList.asp >Customer Service </option>" & vbCrLf
				Response.write "<option value=XLSCustList.asp     >Customer</option>"  & vbCrLf
				Response.write "<option value=XLSContList.asp     >Contact</option>" & vbCrLf
				Response.write "<option value=XSLSOList.asp       >Service Order</option>" & vbCrLf</select></TD></TR>

	<TR><TD align=right width=15%nowrap>Active Only</TD>
		<TD align=left width=35%><INPUT id=chkActiveOnly name=chkActiveOnly type=checkbox value=yes checked style="HEIGHT: 24px; WIDTH: 24px" tabindex=3></TD></TR>

	<TR><TD></TD>
		<TD align=right colspan=2>
			<INPUT id=btnClear name=btnClear type=button style="width: 2cm" value=Clear LANGUAGE=javascript onclick="return btnClear_onclick()">&nbsp;&nbsp;
			<INPUT id=btnGo    name=btnGo    type=submit style="width: 2cm" value=Go > &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD></TR>
	</TABLE>
</FORM>
</BODY>
</HTML>
