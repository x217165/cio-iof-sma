<%@ Language=VBScript %>
<% Option Explicit %>
<!--#include file="smaConstants.inc"-->
<!--#include file="smaProcs.inc"-->
<!--#include file="databaseconnect.asp"-->

<!--
*********************************************************************************************
* Page name:	CustCleanEntry.asp                                                          *
* Purpose:		To dynamically accept the parameters required to perform Customer cleanup.  *
*				Results are displayed via CustCleanList.asp                                 *
*                                                                                           *
* Created by:	Dan S. Ty	03/28/2002                                                      *
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
setPageTitle("SMA - Customer Cleanup");
	

function window_onload() {

	var strCustomerName,strWinName;
	strWinName = document.frmCustCleanEntry.hdnWinName.value;
	DeleteCookie("WinName");
}

function btnClear_onclick() {
	  
	document.frmCustCleanEntry.txtFRCustomer.value = "" ;
	document.frmCustCleanEntry.txtTOCustomer.value = ""  ;
	document.frmCustCleanEntry.selAction.selectedIndex = 0 ;
}
	
function btnCustomerLookup_onclick(WhichCustomer) {

	if (document.frmCustCleanEntry.hdnFRCustomerName.value == "" &&
	    document.frmCustCleanEntry.hdnTOCustomerName.value != "" &&
	    WhichCustomer == 'F') {
		SetCookie("CustomerName", document.frmCustCleanEntry.hdnTOCustomerName.value);
	}

	if (document.frmCustCleanEntry.hdnTOCustomerName.value == "" &&
	    document.frmCustCleanEntry.hdnFRCustomerName.value != "" &&
	    WhichCustomer == 'T') {
		SetCookie("CustomerName", document.frmCustCleanEntry.hdnFRCustomerName.value);
	}

	SetCookie("WinName", 'Popup');
	SetCookie("ServiceEnd", WhichCustomer);
	window.open('SearchFrame.asp?fraSrc=Cust', 'Popup', 'top=50, left=180, height=600, width=750' ) ;
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
//			if (document.frmCustCleanEntry.txtFRCustomer.value == "" ) 
//			{   
//				alert('Please select a "From Customer Name"'); 
//				document.frmCustCleanEntry.btnFRCustomerLookup.focus();  
//				return(false);
//			}
			if (document.frmCustCleanEntry.txtTOCustomer.value == "" ) 
			{   
				alert('Please select "To Customer Name"'); 
				document.frmCustCleanEntry.btnTOCustomerLookup.focus();  
				return(false);
			}	
			if (document.frmCustCleanEntry.selAction.value == "" ) 
			{   
				alert('Please select a Cleanup Action'); 
				document.frmCustCleanEntry.selAction.focus();  
				return(false);
			}	
			if (document.frmCustCleanEntry.hdnFRCustomerID.value == document.frmCustCleanEntry.hdnTOCustomerID.value ) 
			{   
				alert('"From Customer ID" and "To Customer ID" should be different'); 
				document.frmCustCleanEntry.btnFRCustomerLookup.focus();  
				return(false);
			}	
			else				
			{
				document.frmCustCleanEntry.submit();
				return(true);
			}
		}
}	
//-->end hide script	
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload();" >
<FORM name = frmCustCleanEntry method=post action="CustCleanList.asp" target="fraResult" onSubmit="return validate(this);" >

	<!-- hidden variables -->
	<INPUT id=hdnWinName        name=hdnWinName        type=hidden value= "<%=strWinName%>">

	<INPUT id=hdnFRCustomerID   name=hdnFRCustomerID   type=hidden value= "">
	<INPUT id=hdnFRCustomerName name=hdnFRCustomerName type=hidden value= "">

	<INPUT id=hdnTOCustomerID   name=hdnTOCustomerID   type=hidden value= "">
	<INPUT id=hdnTOCustomerName name=hdnTOCustomerName type=hidden value= "">

<TABLE border="0" width="100%">    
    <thead><tr><td colspan=4 align=left>Customer Cleanup Parameters</td></tr></thead>
	<tbody>	
l
	<TR><TD align=right width=25%>From Customer Name<font color=red>*</font></TD>
		<TD align=left width=50% colspan=3>
			<input name=txtFRCustomer type=text disabled size=70 maxlength=70 value="">
			<INPUT align=right type="button"  name=btnFRCustomerLookup   value="..." onclick="return btnCustomerLookup_onclick('F')" tabindex=1></TD></TR>

	<TR><TD align=right width=25%>TO Customer Name<font color=red>*</font></TD>
		<TD align=left width=50% colspan=3>
			<input name=txtTOCustomer type=text disabled size=70 maxlength=70 value="<%if request("hdnTOCustomerID") <> 0 then Response.Write "(" & routineHTMLString(Request("hdnTOCustomerName")) & ")" else Response.Write null end if%>" onChange="fct_onChange();">
			<INPUT align=right type="button"  name=btnTOCustomerLookup   value="..." onclick="return btnCustomerLookup_onclick('T')" tabindex=2></TD></TR>

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
