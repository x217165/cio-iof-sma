<%@ Language=VBScript %>
<% Option Explicit %>
<!--#include file="smaConstants.inc"-->
<!--#include file="smaProcs.inc"-->
<!--#include file="databaseconnect.asp"-->

<!--
*********************************************************************************************
* Page name:	CPCriteria.asp                                                            *
* Purpose:		To dynamically set the criteria to search for a customer profile            *
*				via the CP web service.                   									*
*				Results are displayed on CPList.asp                              *
*                                                                                           *
* Created by:	Anthony Cheung	09/05/2013                                                  *
*                                                                                           *
*********************************************************************************************
-->
<%
	'*************SECURITY********************************************************************
	dim intAccessLevel
	intAccessLevel = CInt(CheckLogon(strConst_Customer))
	'Response.Write ("intAccessLevel:" & intAccessLevel & "<BR>")
	'********************************************************************************************

'	dim rsRegion, rsStatus
'	dim strSQL
	'get Region List
'	strSQL = "select noc_region_lcode, noc_region_desc" & _
'			 " from crp.lcode_noc_region" & _
'			 " where record_status_ind = 'A'" & _
'			 " order by noc_region_desc"
'	set rsRegion = Server.CreateObject("ADODB.Recordset")
'	rsRegion.CursorLocation = adUseClient
'	rsRegion.Open strSQL, objConn
'	if err then
'	end if
	'release the active connection, keep the recordset open
'	set rsRegion.ActiveConnection = nothing
		
	'get status list
'	strSQL = "select customer_status_lcode, customer_status_desc " &_
'			"from crp.lcode_customer_status " &_
'			"where record_status_ind = 'A' " &_
'			"order by customer_status_desc "
'	set rsStatus = Server.CreateObject("ADODB.Recordset")
'	rsStatus.CursorLocation = adUseClient
'	rsStatus.Open strSQL, objConn
'	if err then
'	end if
	'release the active connection, keep the recordset open
'	set rsStatus.ActiveConnection = nothing	
	
'	set objConn = nothing
	
	'if the page is called by a lookup function or by Quick Navigation drop-down box
	'then following cookies will have values.
	dim strWinName,strServiceEnd,strCustomerProfileName, strCustomerProfileID
	strCustomerProfileName = Request.Cookies("CustomerProfileName")
	strCustomerProfileID = Request.Cookies("strCustomerProfileID")
	strWinName	= Request.Cookies("WinName")  
	strServiceEnd = Request.Cookies("ServiceEnd")  
%>
<HTML>
<HEAD>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<LINK href="stylesheets/styles.css" rel="stylesheet" type="text/css">
	<TITLE>Service Management Application</TITLE>
	<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
	<script type="text/javascript" SRC="AccessLevels.js"></SCRIPT>
	<script type="text/javascript">
	<!--hide the script from old browsers 
	
	var intAccessLevel = <%=intAccessLevel%>;
	
	// set section title
setPageTitle("SMA - Customer");
	
	
	function window_onload() {
	
	/************************************************************************************************
	*  Function:	window_onload																	*
	*																								*
	*  Purpose:		To submit the form automatically when txtCustomerName has a value.				*
	*				This will happen when this page is called by a lookup function or by the Quick	*
	*    		    Navigation drop-down box, which had already saved some values in the cookies	*
	* 				and this form's VBScript has read those values and put in the form controls.	*
	*																								*			
	*  Created By:	Sara Sangha Aug 25, 2000														*
	*																								*
	*  Updated By:																					*
	*************************************************************************************************/
		
		var strCustomerProfileName,strWinName;
	 	
	 	strCustomerProfileName = document.frmCPCriteria.txtCustomerProfileName.value ;
	 	strWinName = document.frmCPCriteria.hdnWinName.value;
	 	
	 	DeleteCookie("CustomerProfileName");
	 	DeleteCookie("strCustomerProfileID");	
 		DeleteCookie("WinName");
 	//	DeleteCookie("ServiceEnd"); - never delete "ServiceEnd" cookie as this is needed to determine the parent window the pop-up returns
	 	
	 	if ((strCustomerProfileName!=  "") ) {
 			document.frmCPCriteria.submit() ;  
 		}	
	}

function btnClear_onclick() {
	  
	document.frmCPCriteria.txtCustomerProfileName.value = "" ;   
	document.frmCPCriteria.txtCustomerProfileID.value = ""  ;
}

function validate(theForm) {
 var bolConfirm
	if(isWhitespace(theForm.txtCustomerProfileName.value) 
    && isWhitespace(theForm.txtCustomerProfileID.value))
  {
   bolConfirm = window.confirm("No Search Criteria have been entered. This search may generate no records.Continue?");
    if (bolConfirm){
      return true;
    }
    else
    {
     return false;
    }
  }
   return true;
    
}	
//-->end hide script	
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload();" >
<FORM name = frmCPCriteria method=post action="CPList.asp" target="fraResult" onSubmit="return validate(this);" >

	<!-- hidden variables -->
	<INPUT id=hdnWinName name=hdnWinName type=hidden value="<%=strWinName%>">
	<INPUT id=hdnServiceEnd name=hdnServiceEnd type=hidden value="<%=strServiceEnd%>">

<TABLE border="0" width="100%">    
    <thead><tr><td colspan=4 align=left>Customer Profile Search</td></tr></thead>
	<tbody>	
		<TR>
			<TD align=right nowrap width=20%>Customer Profile Name</TD>
			<TD align=left width=30%><INPUT name=txtCustomerProfileName size=40 maxlength=50 tabindex=1 value="<%=routineHTMLString(strCustomerProfileName)%>"></TD>
		</TR>
		<TR>
			<TD align=right nowrap width=20%>Customer Profile ID (CPID)</TD>
			<TD align=left width=30%><INPUT name=txtCustomerProfileID size=40 maxlength=50 tabindex=1 value="<%=routineHTMLString(strCustomerProfileID)%>"></TD>
		</TR>
		<TR><TD align=right colspan=4>
				<%if strWinName <> "Popup" then%>
					&nbsp;&nbsp;
				<%end if%>
				<INPUT id=btnClear name=btnClear type=button style="width: 2cm" value=Clear LANGUAGE=javascript onclick="return btnClear_onclick()">&nbsp;&nbsp;
				<INPUT id=btnSearch name=btnSearch type=submit style="width: 2cm" value=Search > &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td></TR>
	</TABLE>
</FORM>
</BODY>
</HTML>

