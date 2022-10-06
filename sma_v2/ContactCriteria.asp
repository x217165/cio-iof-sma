<%@ Language=VBScript %>
<%
Option Explicit 
'on error resume next
%>
<!--
********************************************************************************************
* Page name:	ContactCriteria.asp
* Purpose:		To dynamically set the criteria to search for a contact.
*				Results are displayed via ContactList.asp
*
* Created by:	Nancy Mooney	08/15/2000
*  
********************************************************************************************
        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
       20-Jul-01	  DTy		Select 'All' as default for 'Range' radio button.
       19-Feb-02	  DTy		Increase email address size from 50 t0 60.
								Increase Work For search field size from 25 to 50.
       29-Jul-15   PSmith  Set Cookies in validation so the back key works
       05-Oct-15   PSmith  Only sumbit() for pop-up windows
       03-Feb-16   PSmith  Don't pre-populate search criteria
********************************************************************************************
-->

<!--#include file = "smaConstants.inc" -->
<!--#include file = "databaseconnect.asp"-->
<!--#include file = "smaProcs.inc" -->

<% 
'*** SECURITY ********************************************************************
dim intAccessLevel
intAccessLevel = CInt(CheckLogon(strConst_Contact))

'navigation: read cookies
dim strWorkFor, strLName, strFName, strWinName, strTelusOnly, strCase, strEmail

	strLName = Trim(Request.Cookies("LName"))
	strFName = Trim(Request.Cookies("FName"))
	strWinName = Trim(Request.Cookies("WinName"))
	strWorkFor = Request.Cookies("WorkFor")
	strTelusOnly = Request.Cookies("TelusOnly")
	strCase = Request.Cookies("Case")
	strEmail = Request.Cookies("Email")

	
'create list
dim strSQL, rsRegion
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
%>

<HTML>
<HEAD>
	<LINK rel="stylesheet" type="text/css" href="stylesheets/styles.css">
	<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
	<script type="text/javascript" SRC="AccessLevels.js"></script> 

<SCRIPT LANGUAGE=javascript>
<!--
	intAccessLevel = <%=intAccessLevel%>

	//set page heading
setPageTitle("SMA - Contacts");
	
	function window_onload() {
	/**********************************************************************************************
	*  Function:	window_onload																  *
	*  Purpose:		Submit the form automatically when called via lookup or Quick Navigation box; *
	*				EXCEPT do not submit if lookup calling field is blank.						  *
	*  Created By:	Nancy Mooney 08/31/2000														  *
	***********************************************************************************************/
	
		var strLName, strFName, strWorkFor, strWinName, strEmail;
		
		//get search criteria
		strLName = document.frmContactCriteria.txtLName.value;
		strFName = document.frmContactCriteria.txtFName.value;
	 	strWorkFor = document.frmContactCriteria.txtWorksForName.value; 
	 	strWinName = document.frmContactCriteria.hdnWinName.value;
	 	strEmail = document.frmContactCriteria.txtEmail.value;
	 	
	 	DeleteCookie("LName");
	 	DeleteCookie("FName");
 		DeleteCookie("WorkFor");
 		DeleteCookie("WinName");
 		DeleteCookie("TelusOnly");
 		DeleteCookie("Case");
 		DeleteCookie("Email"); 		
 				
		if ( strWinName == "Popup" && ((strLName !=  "") || (strFName !=  "")||(strWorkFor !=  "")))
	 	{
			SetCookie("LName",document.frmContactCriteria.txtLName.value);
			SetCookie("FName",document.frmContactCriteria.txtFName.value);
			SetCookie("WorkFor",document.frmContactCriteria.txtWorksForName.value);
		  thinking(parent.fraResult);
 			document.frmContactCriteria.submit();  
		}	
	}	
	
	function validate(theForm){
	//**********************************************************************************************	
	// Function:	validate()																	   *
	// Purpose:		To alert user that criteria should be entered to avoid a full database search  *
	// Created By:	Nancy Mooney		09/25/2000												   *
	// Updated By:																				   *
	//**********************************************************************************************

	var bolConfirm ;
					
	if (isWhitespace(theForm.txtWorksForName.value) && 
		    isWhitespace(theForm.txtLName.value) && 
		    isWhitespace(theForm.txtFName.value) &&
		    isWhitespace(theForm.txtWPhoneArea.value) &&
		    isWhitespace(theForm.txtWPhoneMid.value) &&
		    isWhitespace(theForm.txtWPhoneEnd.value) &&
		    isWhitespace(theForm.txtEmail.value) &&
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
				
	function btnNew_onclick(){
		if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {alert('Access denied. Please contact your system administrator.'); return;}
		parent.document.location.href ="ContactDetail.asp?ContactID=0";
	}
		
	function fct_Clear(){
		with (document.frmContactCriteria) { 
		
			//clear input fields
			txtWorksForName.value="" ;
			txtLName.value="" ;
			txtFName.value="" ;
			txtEmail.value="" ;
			txtWPhoneArea.value="" ;
			txtWPhoneMid.value="" ;
			txtWPhoneEnd.value="" ;
			//set lists, checkboxes & radio buttons
			selRegion.selectedIndex=0 ;
			radContactType.value='1';
			chkActiveOnly.checked=true; }
	}
			
	//End of script hiding-->
</SCRIPT>
</HEAD>
<BODY onload="window_onload()" >
<FORM name = frmContactCriteria method=post action="ContactList.asp" target="fraResult" onsubmit="return validate(this);">
	
	<INPUT name=hdnWinName type=hidden value="<%=strWinName%>">
	<INPUT name=hdnTelusOnly type=hidden value="<%=strTelusOnly%>">
	<INPUT name=hdnCase type=hidden value="<%=strCase%>">
	
<TABLE BORDER=0 width=100%>
	<thead><TR><TD colspan=4 align=left>Contact Search</TR></TD></thead>
    <TR>
        <TD align=right nowrap width=15%>Works For</TD>   
        <TD align=left width=20%><INPUT name=txtWorksForName size=50 tabindex=1 value="<%=strWorkFor%>"></TD> 
        <TD align=right nowrap width=15%>Region</TD>
		<td><select id=selRegion name=selRegion tabindex=8 >
				<option value="All"> </option>
				<%while not rsRegion.EOF
				Response.write "<option value='" & rsRegion(0) & "'>" & routineHtmlString(rsRegion(1)) & "</option>" & vbCrLf
				rsRegion.movenext
				wend
				rsRegion.Close
				%>
			</select>
		</td>
    <TR>
        <TD align=right nowrap width=15%>Last Name</TD>
        <TD align=left width=20%><INPUT name=txtLName size=20 tabindex=2 value="<%=strLName%>"></TD>
        <td align=right width=15%>Range:</td>
        <TD align=left nowrap width=50%>
			<INPUT type=radio name=radContactType tabindex=9 value=1 <% if (strTelusOnly = "yes") then Response.write  " checked  " end if %>>TELUS Only 
			<INPUT type=radio name=radContactType tabindex=10 value=2 >External Only
			<INPUT type=radio name=radContactType tabindex=11 checked value=3>All
		</TD>
    <TR>
    	<TD align=right nowrap width=15%>First Name</TD>   
        <TD align=left width=20%><INPUT name=txtFName size=20 tabindex=3 value="<%=strFName%>"></TD>
        <TD align=right nowrap width=15%>Active Only</TD>
		<TD align=left><INPUT name=chkActiveOnly type=checkbox value=yes 
			checked style="HEIGHT: 22px; WIDTH: 24px" tabindex=12></TD>
    </TR>
    <TR>
		<td align=right nowrap width=15%>Email</td>
		<td align=left width=20%><INPUT name=txtEmail size=80 maxlength=80 tabindex=4 value="<%=strEmail%>"></TD>
		<TD width=140 ></TD>
    </TR>
    <TR>
		<TD align=right nowrap width=15%>Work Phone</TD> 
		<TD align=left width=20%>(&nbsp;<INPUT name=txtWPhoneArea size=3 maxlength=3 tabindex=5 >
			)&nbsp<INPUT name=txtWPhoneMid size=3 maxlength=3 tabindex=6 >
			-&nbsp<INPUT name=txtWPhoneEnd size=4 maxlength=4 tabindex=7 >
		</TD>
		<TD colSpan=2 align=right nowrap width=100%> 
			<%if strWinName <> "Popup" then %>
				<INPUT name=btnNew type=button value=New style="width: 2cm" tabindex=13 LANGUAGE=javascript onclick="return btnNew_onclick()">&nbsp;&nbsp;
			<%end if%>
			<INPUT name=btnClear type=button value=Clear style="width: 2cm" tabindex=14 onClick = fct_Clear(); >&nbsp;&nbsp;
			<INPUT name=btnSearch type=submit style="width: 2cm" tabindex=15 value=Search> &nbsp;&nbsp;
        </TD>
	</TR> 
</TABLE>
</FORM>
</BODY>
</HTML>











