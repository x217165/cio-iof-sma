<%@ Language=VBScript %>
<% Option Explicit %>
<!--#include file="smaConstants.inc"-->
<!--#include file="smaProcs.inc"-->
<!--#include file="databaseconnect.asp"-->
<!--
*********************************************************************************************
* Page name:	ContactRoleCriteria.asp                                                     *
* Purpose:		To dynamically set the criteria to search for a contact role.               *
*				Results are displayed via ContactRoleList.asp                               *
*																							*
* Navigation:	Quick Navigation drop-down box:												*
*					CustDetail.asp															*
*					ContactDetail.asp														*
*				Lookup: None																*
*																							*				                                                                                         *
* Created by:	Nancy Mooney	08/30/2000                                                  *
*																							*
*       29-Jul-15   PSmith  Set Cookies in validation so the back key works
*       05-Oct-15   PSmith  Only sumbit() for pop-up windows
*       03-Feb-16   PSmith  Don't pre-populate search criteria
*********************************************************************************************
-->
<%
    '***
    dim intAccessLevel
    intAccessLevel = CInt(CheckLogon(strConst_ContactRole))
    '***

	dim rsRole,rsRegion
	dim strSQL
	'get Role List
	strSQL = "select distinct customer_contact_type_lcode, customer_contact_type_desc" & _
			 " from crp.lcode_customer_contact_type" & _
			 " where record_status_ind = 'A'" & _
			 " order by Upper(customer_contact_type_lcode)"
	set rsRole = Server.CreateObject("ADODB.Recordset")
	rsRole.CursorLocation = adUseClient
	rsRole.Open strSQL, objConn
	if err then
		DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
	end if
	if rsRole.EOF then
		DisplayError "BACK", "", 999, "CANNOT CREATE OBJECT TYPE LIST", "EOF condition occurred in rsRole recorset."
	end if
	'release the active connection, keep the recordset open
	set rsRole.ActiveConnection = nothing
	
	'noc region
	strSQL = "select noc_region_lcode, noc_region_desc" & _
			 " from crp.lcode_noc_region" & _
			 " where record_status_ind = 'A'" & _
			 " order by noc_region_desc"
	set rsRegion = Server.CreateObject("ADODB.Recordset")
	rsRegion.CursorLocation = adUseClient
	rsRegion.Open strSQL, objConn
	if err then
		DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
	end if
	if rsRegion.EOF then
		DisplayError "BACK", "", 999, "CANNOT CREATE OBJECT TYPE LIST", "EOF condition occurred in rsRegion recordset."
	end if
	'release the active connection, keep the recordset open
	set rsRegion.ActiveConnection = nothing
	
	set objConn = nothing
	
	'retrieve cookie variables
	dim strCustomerName, strWorkFor, strLName, strFName, strWinName
	
	strCustomerName = Request.Cookies("CustomerName")
	strWorkFor = Request.Cookies("WorkFor")
	strLName = Request.Cookies("LName")
	strFName = Request.Cookies("FName")
	strWinName	= Request.Cookies("WinName") 
%>

<HTML>
<HEAD>
	<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<LINK href="stylesheets/styles.css" rel="stylesheet" type="text/css">
	<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
	<script type="text/javascript" SRC="AccessLevels.js"></script>
	<script type="text/javascript">
	<!--hide the script from old browsers 
	
	var intAccessLevel=<%=intAccessLevel%>;
	
	// set section title
setPageTitle("SMA - Contact Roles");
	
	function window_onload() {
	
	/************************************************************************************************
	*  Function:	window_onload																	*
	*																								*
	*  Purpose:		To submit the form automatically when values have been received from a cookie   *				
	*				and have been stored in hidden form controls.									*	
	*																								*			
	*  Created By:	Nancy Mooney 08/30/2000															*
	*																								*
	*************************************************************************************************/
	
		var strCustomerName, strWorkFor, strLName, strFName, strWinName
		
		strCustomerName = document.frmContactRoleCriteria.txtCustomerName.value;
	 	strLName = document.frmContactRoleCriteria.txtLName.value;
	 	strFName = document.frmContactRoleCriteria.txtFName.value;
	 	strWorkFor = document.frmContactRoleCriteria.hdnWorkFor.value;
	 	strWinName = document.frmContactRoleCriteria.hdnWinName.value;
	 	
	 	DeleteCookie("CustomerName");
 		DeleteCookie("LName");
 		DeleteCookie("FName");
 		DeleteCookie("WorkFor");
 		DeleteCookie("WinName");
 			
	 	if (strWinName == "Popup" && ((strCustomerName != "")||(strLName != "")||(strFName != "")||(strWorkFor !=  "" ))){
			SetCookie("CustomerName",document.frmContactRoleCriteria.txtCustomerName.value);
			SetCookie("LName",document.frmContactRoleCriteria.txtLName.value);
			SetCookie("FName",document.frmContactRoleCriteria.txtFName.value);
			SetCookie("WorkFor",document.frmContactRoleCriteria.hdnWorkFor.value);
		  thinking(parent.fraResult);
 			document.frmContactRoleCriteria.submit() ;  
 		}	
	}
	
	function fct_onChangeRole() {
	
		var strWhole;
		var strRoleDesc, intStart, intIndex;
		
		intIndex = document.frmContactRoleCriteria.selContactRole.selectedIndex;
		strWhole = document.frmContactRoleCriteria.selContactRole.options[intIndex].value;
 		intStart = strWhole.indexOf('<%=strDelimiter%>');
		document.frmContactRoleCriteria.txtRoleDesc.value = strWhole.substr(intStart+1);
	}
	
	function fct_Clear(){
		//clear the hidden variables
		//document.frmContactRoleCriteria.hdnLName.value
		//document.frmContactRoleCriteria.hdnFName.value
		//clear input areas
		document.frmContactRoleCriteria.txtCustomerName.value="";
		document.frmContactRoleCriteria.selContactRole.value="";
		document.frmContactRoleCriteria.txtRoleDesc.value="";
		document.frmContactRoleCriteria.txtLName.value= "";
		document.frmContactRoleCriteria.txtFName.value= "";
		document.frmContactRoleCriteria.chkActiveOnly.checked=true;
		document.frmContactRoleCriteria.radSort.value="Role";
	}
	
	
	function fct_addNew(){
	
	  if ((intAccessLevel & intConst_Access_Create)!= intConst_Access_Create) 
		{
			alert('Access denied. Please contact your system administrator.'); 
			return;
		}

		{
		parent.document.location.href ="ContactRoleDetail.asp?hdnCustomerContactID=0" ;
		}
	
	}
	
	
	function validate(theForm){
	//**********************************************************************************************	
	// Function:	validate()																	   *
	// Purpose:		To alert user that criteria should be entered to avoid a full database search  *
	// Created By:	Nancy Mooney		09/25/2000												   *
	// Updated By:	Shawn for use on Contact Role Criteria  10/03/2000																			   *
	//**********************************************************************************************

	var bolConfirm ;
				
	if (isWhitespace(theForm.txtCustomerName.value) && 
		    (theForm.selContactRole.selectedIndex == 0 ) && 
		    isWhitespace(theForm.txtLName.value) &&
		    isWhitespace(theForm.txtFName.value) &&
		    (theForm.selRegion.selectedIndex == 0 )) 	
		 {
		   bolConfirm = window.confirm("No search criteria have been entered. This search may take a long time...Continue?")
		    if (!bolConfirm){
			 // abort search
		     return false;			
		    }
		  }

	  // Start thinking
		  thinking(parent.fraResult);
		  // search critiera have been entered so continue search
		  return true ;				
	}
		

	//-->end hide script
	</SCRIPT>
	
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload();" >
<FORM name = frmContactRoleCriteria method=post action="ContactRoleList.asp" target="fraResult" onsubmit="return validate(this);">
	
	<!--hidden variables -->
	<INPUT name=hdnWinName type=hidden value="<%=strWinName%>">
	<INPUT name=hdnWorkFor type=hidden value="<%=strWorkFor%>">
	
<TABLE border="0" width="100%">    
	<thead><tr><td align=left colspan=4>Contact Role Search</td></tr></thead>
	<tbody>	
		<TR>
			<TD align=right width=15% nowrap >Customer Name</TD>
			<TD width=40% align=left><INPUT name=txtCustomerName size=50 maxlength=50 tabindex=1 value="<%=strCustomerName%>"></TD>
			<TD align=right nowrap width=15%>Region</TD>
			<td><select id=selRegion name=selRegion tabindex=5 >
					<option value="All"> </option>
					<%while not rsRegion.EOF
					Response.write "<option value='" & rsRegion(0) & "'>" & routineHtmlString(rsRegion(1)) & "</option>" & vbCrLf
					rsRegion.movenext
					wend
					rsRegion.Close
					%>
				</select>
			</td>
		</TR>
		<TR>
			<TD align=right width=15%>Role</TD>   
			<TD align=left width=40% >
				<SELECT name="selContactRole" tabindex=2 onChange="fct_onChangeRole();">
					<option value="All"></option>
					<%while not rsRole.EOF
						Response.write "<option value=""" & rsRole(0)& strDelimiter & rsRole(1) & """>" & Ucase(routineHtmlString(rsRole(0))) & "</option>" & vbCrLf
						rsRole.MoveNext
					wend
					rsRole.Close
					%>
				</SELECT>
				<input type=text name=txtRoleDesc disabled size=35 maxlength=50 >
			</TD>
			<TD align=right width=15% nowrap>Active Only</TD>
			<TD align=left ><INPUT name=chkActiveOnly type=checkbox tabindex=6 checked ></TD>
		</tr>
		<TR>
			<td align=right width=15%>Contact Name:</td>
			<td width=40%>&nbsp;</td>
			<td align=right width=15%>&nbsp;</td>
		</tr>
		<tr>
			<td align=right width=15%>Last</td>
			<td align=left width=40%><INPUT align=left name=txtLName size="20" maxlength="20" tabindex=3 value="<%=strLName%>"></td>
			<td align=right width=15%>Order by:</td>
			<td align=left>
				<INPUT name=radSort type=radio tabindex=7 value=Role checked>Role
				<INPUT name=radSort type=radio tabindex=7 value=Contact>Contact
			</td>
		<tr>
			<td align=right width=15%>First</td>
			<td align=left width=40%><input align=left name=txtFName size="20" maxlength="20" tabindex=4 value="<%=strFName%>"></td>
		</tr>
		<TR>
			<td colSpan=4 align=right width=100%> 
			 <% if strWinName <> "Popup" then %>
				<input name=btnNew type=button value=New style="width: 2cm" tabindex=10 style="HEIGHT: 24px; WIDTH: 62px" onClick="fct_addNew();">&nbsp;&nbsp;
			 <% end if %>
				<INPUT name=btnClear type=button value=Clear style="width: 2cm" tabindex=8 style="HEIGHT: 24px; WIDTH: 62px" onClick="fct_Clear();">&nbsp;&nbsp;
				<INPUT name=btnSearch type=submit value=Search style="width: 2cm" tabindex=9 style="HEIGHT: 24px; WIDTH: 62px" > &nbsp;&nbsp;
				
				
			</td>
		</TR>
	</TABLE>
</FORM>
</BODY>
</HTML>

