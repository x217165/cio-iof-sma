<%@ Language=VBScript %>
<% Option Explicit %>
<% on error resume next %>
<!--#include file="smaConstants.inc"-->
<!--#include file="databaseconnect.asp"-->
<!--#include file="smaProcs.inc"-->
<!--
********************************************************************************************
* Page name:	CustServAliasDetail.asp
* Purpose:		To display Customer Service Name alias and allow user to make changes.
*
* Created by:	Dan S. Ty	02/27/2002
********************************************************************************************
        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
***************************************************************************************************
-->
<%
    stop
'check user's rights
dim intAccessLevel

intAccessLevel = CInt(CheckLogon(strConst_CustomerService))
if intAccessLevel and intConst_Access_ReadOnly <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to Customer Service. Please contact your system administrator"
end if

dim sql, strWinMessage, rsAlias

dim strAction
strAction = Request("action")			'get the action code from caller
if strAction = "" then
	Response.write "No action requested"
	Response.End						'no action requested
end if

dim strMasterID
strMasterID = Request("masterID")		'get master id
dim strAliasID
strAliasID = Request("aliasID")			'get alias id
dim strRealUserID
strRealUserID = Session("username")
if err then
	'unexpected error
	DisplayError "BACK", "", 0, "UNEXPECTED ERROR", "Close alias window to return to Customer Serive form."
end if
strLastUpdate = Request("hdnLastUpdate")

'save changes?
if strAction = "save" then
	dim strNameAlias, strLastUpdate
	strMasterID = Request("masterID")
	strNameAlias = Request("txtNameAlias")
	'call stored proc to save the record

	if (strAliasID <> "") then
		if (intAccessLevel and intConst_Access_Update) <> intConst_Access_Update then
			DisplayError "REFRESH", strWinLocation, 0, "UPDATE DENIED", "You don't have access to update customer service. Please contact your system administrator."
		end if

		'create command object for update stored proc

		dim cmdUpdateObj
		set cmdUpdateObj = server.CreateObject("ADODB.Command")
		set cmdUpdateObj.ActiveConnection = objConn
		cmdUpdateObj.CommandType = adCmdStoredProc
		cmdUpdateObj.CommandText = "sma_sp_userid.spk_sma_cserv_inter.sp_cs_alias_update"
		'create params
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_userid",         adVarChar,     adParamInput, 20, strRealUserID)
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_alias_id",       adNumeric,     adParamInput,   , CLng(strAliasID))
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_alias_desc",     adVarChar,     adParamInput, 80, strNameAlias)
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_cs_id",          adNumeric,     adParamInput,   , CLng(strMasterID))
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput,   , CDate(strLastUpdate))

		'call the update stored proc
		on error resume next
		cmdUpdateObj.Execute
		if objConn.Errors.Count <> 0 then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT UPDATE OBJECT", objConn.Errors(0).Description
			objConn.Errors.Clear
		end if
		strWinMessage = "Record saved successfully. You can now see the changes you made."
	elseif (strAliasID = "") and (intAccessLevel and intConst_Access_Create = intConst_Access_Create) then
		'create command object for insert stored proc
		dim cmdInsertObj
		set cmdInsertObj = server.CreateObject("ADODB.Command")
		set cmdInsertObj.ActiveConnection = objConn
		cmdInsertObj.CommandType = adCmdStoredProc
		cmdInsertObj.CommandText = "sma_sp_userid.spk_sma_cserv_inter.sp_cs_alias_insert"

		'create params
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_userid",     adVarChar, adParamInput, 20, strRealUserID)
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_alias_id",   adNumeric, adParamOutput,  , null)
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_alias_desc", adVarChar, adParamInput, 80, strNameAlias)
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_cs_id",      adNumeric, adParamInput,  , CLng(strMasterID))

		'call the update stored proc
		objConn.Errors.Clear
		on error resume next
		cmdInsertObj.Execute
		if objConn.Errors.Count <> 0 then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT INSERT OBJECT", objConn.Errors(0).Description
			objConn.Errors.Clear
		end if
		strAliasID = cmdInsertObj.Parameters("p_alias_id").Value	'set return parameter
		if strAliasID = "" then
			DisplayError "BACK", "", 2100, "CANNOT DISPLAY NEW ALIAS.", "Most probably the new alias has been saved successfully even if there was an error retrieving the new id. Close the alias window to return to the customer service screen."
			objConn.Errors.Clear
		end if
		strWinMessage = "Record saved successfully. You can now see the changes you made."
	else
		DisplayError "BACK", "", err.Number, "ACCESS DENIED. PLEASE CONTACT YOUR SYSTEM ADMINISTRATOR.", err.Description
	end if
	if err then
		DisplayError "BACK", "", err.Number, "CANNOT UPDATE ALIAS - TRY AGAIN", err.Description
	end if
end if

'delete alias?
if strAction = "delete" then
	'call stor proc to delete current alias
	if intAccessLevel and intConst_Access_Delete = intConst_Access_Delete then
		'create command object for update stored proc
		dim cmdDeleteObj
		set cmdDeleteObj = server.CreateObject("ADODB.Command")
		set cmdDeleteObj.ActiveConnection = objConn
		cmdDeleteObj.CommandType = adCmdStoredProc
    stop
		cmdDeleteObj.CommandText = "sma_sp_userid.spk_sma_cserv_inter.sp_cs_alias_delete"
		'create params
		cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_alias_id", adNumeric, adParamInput,, CLng(strAliasID))
		cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , CDate(strLastUpdate))
                cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("P_USER_ID", adVarChar , adParamInput, 30, strRealUserID)
                

		'call the update stored proc
		if objConn.Errors.Count <> 0 then
	'		DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT - PARAMETER ERROR", objConn.Errors(0).Description
			objConn.Errors.Clear
		end if
		cmdDeleteObj.Execute
		if objConn.Errors.Count <> 0 then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT", objConn.Errors(0).Description
			objConn.Errors.Clear
		end if
		strWinMessage = "Record deleted successfully. You can now create a new alias."
	else
		DisplayError "BACK", "", err.Number, "ACCESS DENIED. PLEASE CONTACT YOUR SYSTEM ADMINISTRATOR.", err.Description
	end if
	if err then
		DisplayError "BACK", "", err.Number, "CANNOT DELETE ALIAS", err.Description
	end if
	'called from main form?
	if Request("back") = "true" then
		Response.Redirect "CustServAlias.asp?cs_id="&strMasterID
	end if
	'ready to enter a new alias?
	strAliasID=""
	strAction="new"
end if

'display the alias info
if strAction <> "new" then
	sql =	"SELECT "&_
				"CUSTOMER_SERVICE_DESC_ALIAS_ID, "&_
				"CUSTOMER_SERVICE_ID, "&_
				"CUSTOMER_SERVICE_DESC_ALIAS, "&_
				"CREATE_DATE_TIME, "&_
				"sma_sp_userid.spk_sma_library.sf_get_full_username(CREATE_REAL_USERID) as create_real_userid, "&_
				"UPDATE_DATE_TIME," &_
				"sma_sp_userid.spk_sma_library.sf_get_full_username(UPDATE_REAL_USERID) as update_real_userid, "&_
				"RECORD_STATUS_IND "&_
			"FROM CRP.CUSTOMER_SERVICE_DESC_ALIAS "&_
			"WHERE CUSTOMER_SERVICE_DESC_ALIAS_ID = " & strAliasID

	set rsAlias=server.CreateObject("ADODB.Recordset")
	rsAlias.CursorLocation = adUseClient
	rsAlias.Open sql, objConn
	if err then
		DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
	end if
	if rsAlias.EOF then
		DisplayError "BACK", "", err.Number, "CANNOT FIND ALIAS", err.Description
	end if
	set rsAlias.ActiveConnection = nothing
end if

%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK REL="stylesheet" TYPE="text/css" HREF="stylesheets/styles.css">
<TITLE>Alias Detail</TITLE>

<script type="text/javascript" SRC="AccessLevels.js"></script>
<SCRIPT type="text/javascript">
var bolSaveRequired = false;
intAccessLevel=<%=intAccessLevel%>;
var intConst_MessageDisplay=<%=intConst_MessageDisplay%>;

function fct_onChange(){
	bolSaveRequired = true;
}

function btnNew_click(){
	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {alert('Access denied. Please contact your system administrator.'); return;}
	var strMasterID = "<%=strMasterID%>";
	document.location.href ="CustServAliasDetail.asp?action=new&masterID="+strMasterID;
}

function fct_onDelete(){
	if (document.frmAlias.AliasID.value != '') {
	   var strMasterID = "<%=strMasterID%>";
	   if (((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) || (document.frmAlias.txtRecordStatusInd.value == "D")) {alert('Access denied. Please contact your system administrator.'); return;}
	   if (confirm('Do you really want to delete this object?')){
	       document.location = "CustServAliasDetail.asp?action=delete&back=false&masterID="+strMasterID+"&AliasID="+document.frmAlias.AliasID.value+"&hdnLastUpdate="+escape(document.frmAlias.hdnLastUpdate.value);
	   }
	}else{fct_displayStatus('There is no need to delete an empty alias.');}
}

function btnClose_onclick(){
	window.close();
}

function frmAlias_onsubmit() {
	if	((((intAccessLevel & intConst_Access_Create) == intConst_Access_Create) && (document.frmAlias.AliasID.value == "")) || (((intAccessLevel & intConst_Access_Update) == intConst_Access_Update) && (document.frmAlias.AliasID.value != ""))) {
		if (document.frmAlias.txtNameAlias.value != "")
			{document.frmAlias.action.value = "save";
			bolSaveRequired = false;
			document.frmAlias.submit();
			return(true);}
		else
			{alert("You cannot save an empty alias. Please click DELETE button if you want to delete the alias.");
			return(false);}

	}
	else {alert('Access denied. Please contact your system administrator.'); return(false);}
}

function fct_clearStatus() {
	window.status = "";
}

function fct_displayStatus(strMessage){
	window.status = strMessage;
	setTimeout('fct_clearStatus()',intConst_MessageDisplay);
}

function body_onLoad(strWinStatus){
	var strWinStatus='<%=strWinMessage%>';
	fct_displayStatus(strWinStatus);
}

function body_onBeforeUnload(){
	if (bolSaveRequired) {
		event.returnValue = "There is unsaved data on the screen. To save changes, click CANCEL below then click SAVE on the main form.";
	}
}

function btnReset_onclick(){
	bolSaveRequired = false ;


}

function body_onUnload(){
	opener.document.frmCustServDetail.btn_iframe1Refresh.click();
}
</SCRIPT>
</HEAD>

<BODY onLoad="body_onLoad();" onBeforeUnload="body_onBeforeUnload();" onUnload="body_onUnload();">
<FORM name=frmAlias LANGUAGE=javascript onsubmit="return frmAlias_onsubmit()">
<INPUT type=hidden name=action        value="">
<INPUT type=hidden name=hdnLastUpdate value="<%if strAliasID <> "" then Response.Write rsAlias.Fields("UPDATE_DATE_TIME").value%>">
<INPUT type=hidden name=AliasID       value="<%if strAliasID <> "" then Response.Write rsAlias("CUSTOMER_SERVICE_DESC_ALIAS_ID")%>">
<INPUT type=hidden name=MasterID      value="<%=strMasterID%>">

<TABLE border=0 width=100%>
<THEAD>
	<TR ><TD colspan=2>Customer Service Name Alias Detail</td></tr>
</THEAD>

<TBODY>
<TR>
	<TD ALIGN=RIGHT width=20% NOWRAP>Alias Name</TD>
	<TD width=80%><INPUT size=40 maxlength=40 name=txtNameAlias value="<%if strAliasID <> "" then Response.write rsAlias("CUSTOMER_SERVICE_DESC_ALIAS")%>" onchange ="fct_onChange();">
</TR>
</TBODY>

<TFOOT>
	<TR><TD align=right colspan=5>
	  	<INPUT type=button name=btnClose  value=Close  style= "width: 2cm" onclick="return btnClose_onclick();"  >&nbsp;&nbsp;
	  	<INPUT type=button name=btnDelete value=Delete style= "width: 2cm" onclick="return fct_onDelete();"      >&nbsp;&nbsp;
	  	<INPUT type=button name=btnReset  value=Reset  style= "width: 2cm" onclick="return btnReset_onclick();"  >&nbsp;&nbsp;
	  	<INPUT type=button name=btnNew    value=New    style= "width: 2cm" onclick="return btnNew_click();"      >&nbsp;&nbsp;
	  	<INPUT type=button name=btnSave   value=Save   style= "width: 2cm" onclick="return frmAlias_onsubmit();" >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	</TD></TR>
</TFOOT>
</TABLE>

<FIELDSET>
<LEGEND align=right><B>Audit Information</B></LEGEND>
<Div SIZE=8pt ALIGN=RIGHT>
	Record Status Indicator
	<INPUT align = left name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 18px"disabled value="<%if strAliasID <> "" then Response.write rsAlias("RECORD_STATUS_IND")%>" >&nbsp;&nbsp;&nbsp;
	Create Date&nbsp;<INPUT align=center name=txtCreateDateTime type=text style="HEIGHT: 20x; WIDTH: 150px" disabled value="<%if strAliasID <> "" then Response.write rsAlias("CREATE_DATE_TIME")%>" >&nbsp;
	Created By&nbsp; <INPUT align=right  name=txtCreateRealUser type=text style="HEIGHT: 20px; WIDTH: 200px" disabled value="<%if strAliasID <> "" then Response.write rsAlias("CREATE_REAL_USERID")%>" ><BR>
	Update Date&nbsp;<INPUT align=center name=txtUpdateDateTime type=text style="HEIGHT: 20px; WIDTH: 150px" disabled value="<%if strAliasID <> "" then Response.write rsAlias("UPDATE_DATE_TIME")%>" >
	Updated By&nbsp; <INPUT align=right  name=txtUpdateRealUser type=text style="HEIGHT: 20px; WIDTH: 200px" disabled value="<%if strAliasID <> "" then Response.write rsAlias("UPDATE_REAL_USERID")%>" >
</DIV>
</FIELDSET>
<%
'Clean up our ADO objects
if strAliasID <> "" then
	rsAlias.close
	set rsAlias = Nothing
	objConn.close
	set objConn = Nothing
end if
%>

</FORM>
</BODY>
</HTML>
