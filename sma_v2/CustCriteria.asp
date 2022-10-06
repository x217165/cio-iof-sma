<%@ Language=VBScript %>
<% Option Explicit %>
<!--#include file="smaConstants.inc"-->
<!--#include file="smaProcs.inc"-->
<!--#include file="databaseconnect.asp"-->

<!--
*********************************************************************************************
* Page name:	CustCriteria.asp                                                              *
* Purpose:		To dynamically set the criteria to search for a customer.                     *
*				Results are displayed via CustList.asp                                              *
*                                                                                           *
* Created by:	Nancy Mooney	08/01/2000                                                      *
*                                                                                           *
*       29-Jul-15   PSmith  Set Cookies in validation so the back key works                 *
*       05-Oct-15   PSmith  Only sumbit() for pop-up windows
*       03-Feb-16   PSmith  Don't pre-populate search criteria
*********************************************************************************************
-->
<%
	'*************SECURITY********************************************************************
	dim intAccessLevel
	intAccessLevel = CInt(CheckLogon(strConst_Customer))
	'Response.Write ("intAccessLevel:" & intAccessLevel & "<BR>")
	'********************************************************************************************
 
	dim rsRegion, rsStatus
	dim strSQL
	'get Region List
	strSQL = "select noc_region_lcode, noc_region_desc" & _
			 " from crp.lcode_noc_region" & _
			 " where record_status_ind = 'A'" & _
			 " order by noc_region_desc"
	set rsRegion = Server.CreateObject("ADODB.Recordset")
	rsRegion.CursorLocation = adUseClient
	rsRegion.Open strSQL, objConn
	if err then
	end if
	'release the active connection, keep the recordset open
	set rsRegion.ActiveConnection = nothing
		
	'get status list
	strSQL = "select customer_status_lcode, customer_status_desc " &_
			"from crp.lcode_customer_status " &_
			"where record_status_ind = 'A' " &_
			"order by customer_status_desc "
	set rsStatus = Server.CreateObject("ADODB.Recordset")
	rsStatus.CursorLocation = adUseClient
	rsStatus.Open strSQL, objConn
	if err then
	end if
	'release the active connection, keep the recordset open
	set rsStatus.ActiveConnection = nothing	
	
	set objConn = nothing
	
	'if the page is called by a lookup function or by Quick Navigation drop-down box
	'then following cookies will have values.
	dim strCustomerName, strWinName,strServiceEnd
	dim strCustShort
	strCustomerName = Request.Cookies("CustomerName")
	strCustShort = Request("txtCustShort")
	strWinName	= Request.Cookies("WinName")  
	strServiceEnd = Request.Cookies("ServiceEnd")  
%>
<HTML>
<HEAD>
	<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
    	<meta http-equiv="Pragma" content="no-cache">
    	<meta http-equiv="Expires" content="0">
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
		
		

		var strCustomerName,strWinName,strCustomerNameAux ;
	 	var strCustShort,strCustShortAux;	
	 	strCustomerName = document.frmCustCriteria.txtCustomerName.value ;
	        strCustShort = document.frmCustCriteria.txtCustShort.value;
	    
	 	strWinName = document.frmCustCriteria.hdnWinName.value;
		
		strCustomerNameAux=document.getElementById("txtCustomerName").value;
	        strCustShortAux=document.getElementById("txtCustShort").value;
		console.log(strCustomerName );
	        
		
	 	DeleteCookie("CustomerName");
	 	DeleteCookie("txtCustShort");
 		DeleteCookie("WinName");
 		DeleteCookie("ServiceEnd");

	 	if (strWinName == "Popup" && ((strCustomerName !=  "") || (strCustShort != ""))) {
		  SetCookie("CustomerName", document.frmCustCriteria.txtCustomerName.value);
		  SetCookie("txtCustShort", document.frmCustCriteria.txtCustShort.value);
		  SetCookie("ServiceEnd", document.frmCustCriteria.hdnServiceEnd.value);
		  thinking(parent.fraResult);
 			document.frmCustCriteria.submit() ;  
 		}
 		

		 
	 	
		

	}

function btnNew_onclick(){
		if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {alert('Access denied. Please contact your system administrator.'); return;}
		parent.document.location.href ="CustDetail.asp?CustomerID=NEW";
	}

function btnClear_onclick() {
	  
	document.frmCustCriteria.txtCustomerName.value = "" ;   
	document.frmCustCriteria.txtCustShort.value="";
	document.frmCustCriteria.txtSMRFName.value = ""  ;
	document.frmCustCriteria.txtSMRLName.value = "" ;
	document.frmCustCriteria.selRegion.selectedIndex = 0 ;
	//maintain the status default as Current
	document.frmCustCriteria.selStatus.selectedIndex=1 ;
	document.frmCustCriteria.chkActiveOnly.checked=true;      
}



function validate(theForm) {
 var strCustomerNameAux=document.getElementById("txtCustomerName").value;
 var strCustShortAux=document.getElementById("txtCustShort").value;
 
 console.log("validate function: "+strCustomerNameAux);
	        
 var bolConfirm;
 	console.log(document.getElementById("txtCustomerName").value);
	if(isWhitespace(theForm.txtCustomerName.value) 
	&& isWhitespace(theForm.txtCustShort.value)
    && isWhitespace(theForm.txtSMRLName.value) 
    && isWhitespace(theForm.txtSMRFName.value) 
    && theForm.selRegion.selectedIndex == 0 
    && theForm.selStatus.selectedIndex == 0 )
  {
   bolConfirm = window.confirm("No Search Criteria have been entered. This search may take a long time..Continue?");
    if (!bolConfirm){
     return false;
    }
  }
   thinking(parent.fraResult);
   
   return true;
    
}	
//-->end hide script	
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload();" >
<FORM name = frmCustCriteria method=post action="CustList.asp" target="fraResult" onSubmit="return validate(this);" >

	<!-- hidden variables -->
	<INPUT id=hdnWinName name=hdnWinName type=hidden value="<%=strWinName%>">
	<INPUT id=hdnServiceEnd name=hdnServiceEnd type=hidden value="<%=strServiceEnd%>">

<TABLE border="0" width="100%">    
    <thead><tr><td colspan=4 align=left>Customer Search</td></tr></thead>
	<tbody>	
		<TR>
			<TD align=right nowrap width=20%>Customer Name</TD>
			<TD align=left width=30%><INPUT id=txtCustomerName name=txtCustomerName size=40 maxlength=50 tabindex=1 value="<%=routineHTMLString(strCustomerName)%>"></TD>
			<TD align=right width=15% nowrap>Region</TD>
			<td width=35%><select id=selRegion name=selRegion tabindex=4 style="width: 160">
				<option value="All"> </option>
				<%while not rsRegion.EOF
				Response.write "<option value='" & rsRegion(0) & "'>" & routineHtmlString(rsRegion(1)) & "</option>" & vbCrLf
				rsRegion.movenext
				wend
				rsRegion.Close
				%>
			</select></td></TR>
		<TR>
			<TD align=right nowrap width=20%>Customer Short Name</TD>
			<TD align=left width=30%><INPUT id=txtCustShort name=txtCustShort size=40 maxlength=50 tabindex=1 value="<%=routineHTMLString(strCustShort)%>"></TD>
			<TD align=right width=15%>Status</TD>
			<TD align=left width=35%>
				<SELECT id=selStatus name=selStatus tabindex=5 style="width: 160" >
				 <option value="All"></option>
					<%while not rsStatus.EOF
						Response.write "<option value= '" &rsStatus(0)& "'>" & routineHtmlString(rsStatus(1)) & "</option>" & vbCrLf
						rsStatus.MoveNext
					wend
					rsStatus.Close
					%>
				</SELECT></TD><TR>

			
			
			
		<TR>
			<TD align=right width=15%>&nbsp;Service Management Rep:</TD>
			<td width=35%>&nbsp;</td>   
    		<TD align=right width=15%nowrap>Active Only</TD>
			<TD align=left width=35%><INPUT id=chkActiveOnly name=chkActiveOnly type=checkbox value=yes checked style="HEIGHT: 24px; WIDTH: 24px" tabindex=6></TD>
        <TR>
		<TR>
			<TD align=right width=20%>Last Name</TD>
			<TD align=left width=30%><INPUT name=txtSMRLName size="20" maxlength="20" tabindex=2 ></TD> 
			<td colSpan=2 align=right width=50%></TD>
		</TR> 
		<TR>
			<td align=right width=20%>First Name</td>
			<td align=left width=30%><input id=txtSMRFName name=txtSMRFName size="20" maxlength="20" tabindex=3></td>
			<TD align=right colspan=4>
				<%if strWinName <> "Popup" then%>
					<INPUT id=btnNew name=btnNew type=button style="width: 2cm" value=New onclick="return btnNew_onclick()">&nbsp;&nbsp;
				<%end if%>
				<INPUT id=btnClear name=btnClear type=button style="width: 2cm" value=Clear LANGUAGE=javascript onclick="return btnClear_onclick()">&nbsp;&nbsp;
				<INPUT id=btnSearch name=btnSearch type=submit style="width: 2cm" value=Search > &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>		
		</TR>
	</TABLE>
</FORM>
</BODY>
</HTML>

