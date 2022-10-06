<%@ Language=VBScript %>
<%Option Explicit	  
  on error resume next %>
<!--#include file="smaConstants.inc"-->
<!--#include file="smaProcs.inc"-->
<!--#include file = "databaseconnect.asp"-->
<!--

********************************************************************************************
* Page name:	OderHistoryCriteria.asp
* Purpose:		To dynamically set the criteria to search for an address.
*				Results are displayed via AddressList.asp
*				
* Created by:	Sara Sangha	Sept. 2nd, 2000
*  
********************************************************************************************
        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
       28-Feb-02	 DTy		Add 'Alias' to 'Customer Service Name' field name.
********************************************************************************************
-->
<HTML>
<HEAD>	
<LINK rel="stylesheet" type="text/css" href="stylesheets/styles.css">
<TITLE>Order History Search</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
<script type="text/javascript">
//********************************************************************************************
//set section title
setPageTitle("SMA - Order History");

function validate(theForm){
// give a warning if no search criteria is entered.
// make sure all fields have correct data type values.
var bolConfirm
if (isWhitespace(theForm.txtCustomerServiceName.value) 
    && isWhitespace(theForm.txtCustomerServiceID.value))
{
   bolConfirm = window.confirm("No Search Criteria have been entered. This search may take a long time..Continue?");
    if (bolConfirm){
      return true;
    }
    else
    {
     return false;
    }
  }
  
  if (isNaN(Number(document.frmOrderHistorySearch.txtCustomerServiceID.value ))) {
		alert("Customer Service ID must be a number."); 
		document.frmOrderHistorySearch.txtCustomerServiceID.focus();
		document.frmOrderHistorySearch.txtCustomerServiceID.select();
		return(false) ;
		}
  else 
	{return(true); }
  
   return true;
}

function btnClear_onclick() {
	
	document.frmOrderHistorySearch.txtCustomerServiceName.value = "" ;
	document.frmOrderHistorySearch.txtCustomerServiceID.value = "" ;

}

function window_onload() {
	
	var strCustomerServiceName = document.frmOrderHistorySearch.txtCustomerServiceName.value ;
	var intCustomerServiceID = document.frmOrderHistorySearch.txtCustomerServiceID.value ;
	
	DeleteCookie("CustomerServiceID");
	DeleteCookie("CustomerServiceName");
	
	if (( strCustomerServiceName != "" )|| ( intCustomerServiceID != "" )) {
		document.frmOrderHistorySearch.submit() ; 
	}
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">
<%

dim objRsRegion ,objRsSupportGroup, objRsStatus, strSQL
dim strCustomerServiceName, strCustomerServiceID, strProjectCode
	
	strCustomerServiceName = Request.Cookies("CustomerServiceName")
	strCustomerServiceID = Request.Cookies("CustomerServiceID")
		
%>

<FORM name=frmOrderHistorySearch method=post action="OrderHistoryList.asp" target="fraResult" onSubmit="return validate(this)">
<TABLE>
	<THEAD>
		<TR>
		<TD COLSPAN=2>Order History Search</TD></TR></THEAD>
    <TR>
        <TD align=right nowrap width=15%>Customer Service Name/Alias</TD>
        <TD><INPUT id=txtCustomerServiceName name=txtCustomerServiceName style="HEIGHT: 22px; WIDTH: 400px" value="<%=strCustomerServiceName%>"></TR>				
        
    <TR>
		<TD align=right nowrap width=15%>Customer Service ID</TD>
        <TD><INPUT id=txtCustomerServiceID name=txtCustomerServiceID style="HEIGHT: 22px; WIDTH: 133px" value="<%=strCustomerServiceID%>" ></TD></TR>
    
	<TR>
		<TD colSpan=2 align=right > 
			<INPUT id=btnClear name=btnClear type=button value=Clear style="width: 2cm" LANGUAGE=javascript onclick="return btnClear_onclick()">&nbsp;&nbsp;
			<INPUT id=btnSearch name=btnSearch type=submit value=Search style="width: 2cm"> &nbsp;&nbsp;
			
        </TD></TR>
</TABLE>
</FORM>
</BODY>
</HTML>