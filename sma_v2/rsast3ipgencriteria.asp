<%@ Language=VBScript %>
<% Option Explicit %>
<!--#include file="smaConstants.inc"-->
<!--#include file="smaProcs.inc"-->
<!--#include file="databaseconnect.asp"-->
<!--
*********************************************************************************************
* Page name:	RSAST3IPGenCriteria.asp                                                     *
* Purpose:		To dynamically set the criteria to search or                                *
*				generated Gateway IP and its IP addresses                                   *
*																							*
* Navigation:	                            												*
* Lookup:    																                *
*																							*				                                                                                         *
* Created by:	Dan Ty	01/28/2002                                                          *
********************************************************************************************
*
        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
*********************************************************************************************
-->
<%
    '***
    dim intAccessLevel
    intAccessLevel = CInt(CheckLogon(strConst_RSAS))
    if intAccessLevel and intConst_Access_ReadOnly <> intConst_Access_ReadOnly then
	   DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to POS PLUS IP Admin. Please contact your system administrator"
    end if
    '***

	dim rsIPAddress, rsCode, rsLocation
	dim strSQL

	'Get Gateway IP Address
	strSQL = "select distinct gateway_ip_address, ip_address, code, location " & _
			 " from crp.rsas_ip_address" & _
			 " where record_status_ind = 'A'" & _
			 " order by gateway_ip_address, ip_address"

	set rsIPAddress = Server.CreateObject("ADODB.Recordset")
	rsIPAddress.CursorLocation = adUseClient
	rsIPAddress.Open strSQL, objConn

	if err then
		DisplayError "BACK", "", err.Number, "CANNOT RETRIEVE GATEWAY IP ADDRESS", err.Description
	end if
	if rsIPAddress.EOF then
		DisplayError "BACK", "", 999, "NO GATEWAY IP ADDRESS FOUND", "EOF condition occurred in rsRole recorset."
	end if

	'release the active connection, keep the recordset open
	set rsIPAddress = nothing
	
	'Get Gateway Code
	strSQL = "select distinct code" & _
			 " from crp.rsas_ip_address" & _
			 " where record_status_ind = 'A'" & _
			 " order by code"
	set rsCode = Server.CreateObject("ADODB.Recordset")
	rsCode.CursorLocation = adUseClient
	rsCode.Open strSQL, objConn
	if err then
		DisplayError "BACK", "", err.Number, "CANNOT RETRIEVE GATEWAY CODE", err.Description
	end if
	if rsCode.EOF then
		DisplayError "BACK", "", 999, "NO GATEWAY CODE FOUND", "EOF condition occurred in rsRegion recordset."
	end if
	
	'Get Gateway Location
	strSQL = "select distinct location" & _
			 " from crp.rsas_ip_address" & _
			 " where record_status_ind = 'A'" & _
			 " order by location"
	set rsLocation = Server.CreateObject("ADODB.Recordset")
	rsLocation.CursorLocation = adUseClient
	rsLocation.Open strSQL, objConn
	if err then
		DisplayError "BACK", "", err.Number, "CANNOT RETRIEVE GATEWAY LOCATION", err.Description
	end if
	if rsCode.EOF then
		DisplayError "BACK", "", 999, "NO GATEWAY LOCATION FOUND", "EOF condition occurred in rsRegion recordset."
	end if
	'release the active connection, keep the recordset open
	set rsCode.ActiveConnection = nothing
	set objConn = nothing

	'retrieve cookie variables
	strWinName	= Request.Cookies("WinName") 

	dim strGatewayIP, strIPAddress, strCode, strLocation, strAvailable, strWinName
	
	strGatewayIP = Request("txtGatewayIP")
	strIPAddress = Request("txtIPAddress")
	strCode      = Request("selCode")
	strLocation  = Request("selLocation")
	strAvailable = Request("selAvailable")

	strWinName	 = Request.Cookies("WinName") 

%>

<HTML>
<HEAD>
	<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<LINK href="stylesheets/styles.css" rel="stylesheet" type="text/css">
	<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
	<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
	<script type="text/javascript" SRC="AccessLevels.js"></script>
	<script type="text/javascript">
	<!--hide the script from old browsers 
	
	var intAccessLevel=<%=intAccessLevel%>;
	
	// set section title
setPageTitle("SMA - POS PLUS Admin");
	
	/************************************************************************************************
	*  Function:	window_onload																	*
	*																								*
	*  Purpose:		To submit the form automatically when values have been received from a cookie   *				
	*				and have been stored in hidden form controls.									*	
	*************************************************************************************************/
	function window_onload()
	{
	  var strGatewayIP, strIPAddress, selCode, selLocation, selAvailable, strWinName
		
	  strGatewayIP = document.frmRSAST3IPGenCriteria.txtGatewayIP.value;
	  strIPAddress = document.frmRSAST3IPGenCriteria.txtIPAddress.value;
	  selCode      = document.frmRSAST3IPGenCriteria.selCode.value;
	  selLocation  = document.frmRSAST3IPGenCriteria.selLocation.value;
	  selAvailable = document.frmRSAST3IPGenCriteria.selAvailable.value;
	  strWinName   = document.frmRSAST3IPGenCriteria.hdnWinName.value;

	  DeleteCookie("WinName");
 			
	  if ((strWinName !=  "" )||(strGatewayIP != "")||(strIPAddress != "")||(selCode != "")||(selLocation !=  "" )||(selAvailable !=  "" ))
	  {
		  document.frmRSAST3IPGenCriteria.submit() ;  
	  }
	}	
	function fct_Clear()
	{
	  //clear input areas
	  document.frmRSAST3IPGenCriteria.txtGatewayIP.value="";
	  document.frmRSAST3IPGenCriteria.txtIPAddress.value="";
	  document.frmRSAST3IPGenCriteria.selCode.indexvalue=0;
	  document.frmRSAST3IPGenCriteria.selLocation.indexvalue=0;
	  document.frmRSAST3IPGenCriteria.selAvailable.indexvalue=0;
	  document.frmRSAST3IPGenCriteria.chkActiveOnly.checked=true;
    }	

	function fct_IPGen()
	{
	  if ((intAccessLevel & intConst_Access_Create) = intConst_Access_Create) 
		{
			alert('Access denied. Please contact your system administrator.'); 
			return;
		}
	  else
		{
		//Call generate program here.
		}
	}
	//**********************************************************************************************	
	// Function:	validate()																	   *
	// Purpose:		To alert user that criteria should be entered to avoid a full database search  *
	//**********************************************************************************************
    function validate(theForm)
	{
	  var bolConfirm ;
	
	  if (isWhitespace(theForm.txtGatewayIP.value) && 
	      isWhitespace(theForm.txtIPAddress.value) &&
	      theForm.selCode.selectedIndex == 0       && 
	      theForm.selLocation.selectedIndex == 0   && 
	      theForm.selAvailable.selectedIndex == 0)
	      {
		  bolConfirm = window.confirm("No search criteria have been entered. This search may take a long time...Continue?")
		  if (bolConfirm)
		     {
		     return true;			 
		     }
		  else
		     {
		     return false;			
	         }
	      }
	  // search critiera have been entered so continue search
	  return true ;
   }
	//-->end hide script

	</SCRIPT>
	
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload();" >
<FORM name = frmRSAST3IPGenCriteria method=post action="RSAST3IPGenList.asp" target="fraResult" onsubmit="return validate(this);">
	<!--hidden variables -->
	<INPUT name=hdnWinName type=hidden value="<%=strWinName%>">
	
<TABLE border="0" width="100%">    
	<thead><tr><td align=left colspan=4>Gateway IP Search</td></tr></thead>
	<tbody>	
		<TR>
			<TD align=right width=15% nowrap >Gateway IP</TD>
			<TD width=40% align=left><INPUT name=txtGatewayIP size=20 maxlength=20 tabindex=1 value="<%=strGatewayIP%>"></TD>
		</TD>

		<TR>
			<TD align=right width=15% nowrap >IP Address</TD>
			<TD width=40% align=left><INPUT name=txtIPAddress size=20 maxlength=20 tabindex=2 value="<%=strIPAddress%>"></TD>
		</TD>

		<TR>
			<TD align=right nowrap width=15%>Code</TD>
			<TD><select id=selCode name=selCode tabindex=3>
				<option value=""> </option>
				<%while not rsCode.EOF
				    Response.write "<option value='" & rsCode(0) & "'>" & routineHtmlString(rsCode(0)) & "</option>" & vbCrLf
				    rsCode.movenext
				  wend
				  rsCode.Close
				%>
				</select>
			</TD>
		</TD>
		<TR>
			<TD align=right nowrap width=15%>Location</TD>
			<TD><select id=selLocation name=selLocation tabindex=4>
					<option value=""> </option>
					<%while not rsLocation.EOF
					Response.write "<option value='" & rsLocation(0) & "'>" & routineHtmlString(rsLocation(0)) & "</option>" & vbCrLf
					rsLocation.movenext
					wend
					rsLocation.Close
					%>
				</select>
			</TD>
		</TR>

		<TR>
			<TD align=right nowrap width=15%>Available?</TD>
			<TD><select id=selAvailable name=selAvailable tabindex=5 >
				<option value='' selected> </option>
				  Response.write "<option value='Y'> Yes</option>" & vbCrLf
				  Response.write "<option value='N'> No</option>" & vbCrLf
				</select>
			</TD>
		</TR>

		<TR>
			<TD align=right width=15% nowrap>Active Only</TD>
			<TD align=left ><INPUT name=chkActiveOnly type=checkbox tabindex=6 checked ></TD>
		</TR>
		
		<TR>
			<td colSpan=4 align=right width=100%> 
			 <% if strWinName <> "Popup" then %>
				<input name=btnNew type=button value="IP Gen" style="width: 2cm" tabindex=10 style="HEIGHT: 24px; WIDTH: 62px" onClick="fct_IPGen();">&nbsp;&nbsp;
			 <% end if %>
				<INPUT name=btnClear type=button value=Clear style="width: 2cm" tabindex=8 style="HEIGHT: 24px; WIDTH: 62px" onClick="fct_Clear();">&nbsp;&nbsp;
				<INPUT name=btnSearch type=submit value=Search style="width: 2cm" tabindex=9 style="HEIGHT: 24px; WIDTH: 62px" > &nbsp;&nbsp;
			</td>
		</TR>
	</TABLE>
</FORM>
</BODY>
</HTML>

