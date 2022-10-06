<%@ Language=VBScript %>
<% Option Explicit %>
<% Response.Buffer = True %>
meen<!--% on error resume next %-->
<!--#include file="smaConstants.inc"-->
<!--#include file="databaseconnect.asp"-->
<!--#include file="smaProcs.inc"-->

<!--
*************************************************************************************
* File Name:	RSAST3DevDetail.asp
*
* Purpose:		To display the detailed information about a Device entry.
*
* In Param:
*
* Out Param:
*
* Created By:
**************************************************************************************

        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
       10-23-01	     DTy		Change field names and variables.

								Local DNA to Local X25 DNA
								LOCAL_DNA to LOCAL_X25_DNA
								strLocalDNA to strLocalxx25DNA
								txtLocalDNA to txtLocalxx25DNA

								HOST_DNA_ID to HOST_X25_DNA_ID
								strHostDNAID to strHostX25DNAID
								selHostDNAID to selHostX25DNAID

								Change CInt to CLng use on HOST_X25_DNA_ID
								Pass on HOST_X25_DNA_ID on SELECT not HOST_X25_DNA
**************************************************************************************
-->

<%

'check user's rights
dim intAccessLevel
intAccessLevel = CInt(CheckLogon(strConst_RSAS))
if intAccessLevel and intConst_Access_ReadOnly <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to POS PLUS Tier 3 device. Please contact your system administrator"
end if

dim sql, strWinMessage, rsDevice, strSQL

dim strAction
strAction = Request("action")			'get the action code from caller

if strAction = "" then
	Response.write "No action requested"
	Response.End						'no action requested
end if

dim strMasterID
strMasterID = Request("masterID")		'get master id

dim strDeviceID
strDeviceID = Request("hdnDeviceID")			'get device id

if strDeviceID ="" then
	strDeviceID = 0
end if

dim strRealUserID
strRealUserID = Session("username")
if err then
	'unexpected error
	DisplayError "BACK", "", 0, "UNEXPECTED ERROR", "Close alias window to return to customer form."
end if
strLastUpdate = Request("hdnLastUpdate")

'Response.Write ("strDeviceID = " & strDeviceID & ", " & "strLastUpdate = " & strLastUpdate )
'Response.Write ("strMasterID = " & strMasterID & ", " & "strAction = " & strAction )
'Response.end

'save changes?
if strAction = "save" then
	dim strLocalX25DNA, strPollCode, strHostX25DNAID, strLastUpdate
	strMasterID = Request("masterID")
	strLocalX25DNA = Request("txtLocalX25DNA")
	strPollCode = Request("selPollCode")
	strHostX25DNAID = Request("selHostX25DNAID")
	'call stored proc to save the record

	if (strDeviceID <> 0)  then

		'create command object for update stored proc

		dim cmdUpdateObj
		set cmdUpdateObj = server.CreateObject("ADODB.Command")
		set cmdUpdateObj.ActiveConnection = objConn

		cmdUpdateObj.CommandType = adCmdStoredProc
		cmdUpdateObj.CommandText = "sma_sp_userid.spk_sma_rsas_inter.sp_device_update"

		'create required params
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_user_id", adVarChar , adParamInput, 30, strRealUserID)
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_device_id", adNumeric , adParamInput,, CLng(strDeviceID))
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , CDate(strLastUpdate))

		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_tail_circuit_id", adNumeric , adParamInput,, CLng(strMasterID))
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_local_x25_dna", adVarChar , adParamInput, 10, strLocalX25DNA)
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_poll_code", adVarChar, adParamInput, 4, strPollCode)
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_host_x25_dna_id", adNumeric , adParamInput,, CLng(strHostX25DNAID))

		'call the update stored proc
		if err then
			DisplayError "BACK", "", err.Number, "CANNOT UPDATE OBJECT - PARAMETER ERROR", err.Description
			objConn.Errors.Clear
		end if

		'Response.Write "<b> count = " & cmdUpdateObj.Parameters.count & "<br>"
		'dim nx
		'for nx = 0 to cmdUpdateObj.Parameters.Count-1
		'	Response.Write cmdUpdateObj.Parameters.Item(nx).Name & " = " & cmdInsertObj.Parameters.Item(nx).Value & " <br>"
		'next
		'Response.end

		cmdUpdateObj.Execute
		if objConn.Errors.Count <> 0 then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT UPDATE OBJECT", objConn.Errors(0).Description
			objConn.Errors.Clear
		end if
		strWinMessage = "Record saved successfully."

	elseif (strDeviceID = 0) and (intAccessLevel and intConst_Access_Create = intConst_Access_Create) then

		'create command object for insert stored proc
		dim cmdInsertObj
		set cmdInsertObj = server.CreateObject("ADODB.Command")
		set cmdInsertObj.ActiveConnection = objConn
		cmdInsertObj.CommandType = adCmdStoredProc
		cmdInsertObj.CommandText = "sma_sp_userid.spk_sma_rsas_inter.sp_device_insert"

		'create params
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_user_id", adVarChar , adParamInput, 30, strRealUserID)
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_device_id", adNumeric , adParamOutput)

		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_tail_circuit_id", adNumeric , adParamInput,, CLng(strMasterID))
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_local_x25_dna", adVarChar , adParamInput, 10, strLocalX25DNA)
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_poll_code", adVarChar, adParamInput, 4, strPollCode)
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_host_x25_dna_id", adNumeric , adParamInput,, CLng(strHostX25DNAID))

		'Response.Write "<b> count = " & cmdInsertObj.Parameters.count & "<br>"
		'dim nx
		'for nx = 0 to cmdInsertObj.Parameters.Count-1
		'	Response.Write cmdInsertObj.Parameters.Item(nx).Name & " = " & cmdInsertObj.Parameters.Item(nx).Value & " <br>"
		'next
		'Response.end

		if objConn.Errors.Count <> 0 then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT INSERT OBJECT - PARAMETER ERROR", objConn.Errors(0).Description
			objConn.Errors.Clear
		end if

		'call the insert stored proc
		cmdInsertObj.Execute
		if objConn.Errors.Count <> 0 then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT INSERT OBJECT", objConn.Errors(0).Description
			objConn.Errors.Clear
		end if
		strDeviceID = cmdInsertObj.Parameters("p_device_id").Value	'set return parameter
		if strDeviceID = 0 then
			DisplayError "BACK", "", 2100, "CANNOT DISPLAY NEW DEVICE.", "Most probably the new device has been saved successfully even if there was an error retrieving the new id. Close the device window to return to the tail circuit screen."
			objConn.Errors.Clear
		end if
		strWinMessage = "Record saved successfully."
	else
		DisplayError "BACK", "", err.Number, "ACCESS DENIED. PLEASE CONTACT YOUR SYSTEM ADMINISTRATOR.", err.Description
	end if
	if err then
		DisplayError "BACK", "", err.Number, "CANNOT UPDATE DEVICE - TRY AGAIN", err.Description
	end if
end if

'delete device?
if strAction = "delete" then

	'call stor proc to delete current alias
	if intAccessLevel and intConst_Access_Delete = intConst_Access_Delete then

		'create command object for delete stored proc

		dim cmdDeleteObj
		set cmdDeleteObj = server.CreateObject("ADODB.Command")
		set cmdDeleteObj.ActiveConnection = objConn
		cmdDeleteObj.CommandType = adCmdStoredProc
		cmdDeleteObj.CommandText = "sma_sp_userid.spk_sma_rsas_inter.sp_device_delete"

		'create params
		cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_device_id", adNumeric , adParamInput,, CInt(strDeviceID))
		cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_last_update", adDBTimeStamp, adParamInput,, CDate(strLastUpdate))

		'call the delete stored proc

		if objConn.Errors.Count <> 0 then
	'		DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT - PARAMETER ERROR", objConn.Errors(0).Description
			objConn.Errors.Clear
		end if
		cmdDeleteObj.Execute
		if objConn.Errors.Count <> 0 then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT", objConn.Errors(0).Description
			objConn.Errors.Clear
		end if
		strWinMessage = "Record deleted successfully."
	else
		DisplayError "BACK", "", err.Number, "ACCESS DENIED. PLEASE CONTACT YOUR SYSTEM ADMINISTRATOR.", err.Description
	end if
	if err then
		DisplayError "BACK", "", err.Number, "CANNOT DELETE DEVICE", err.Description
	end if

	'called from main form? ADAM
	if Request("back") = "true" then
		Response.Redirect "RSAST3DevList.asp?TailCircuitID="&strMasterID
	end if
	'ready to enter a new device?

	strDeviceID=0
	strAction="new"
end if

'display the device info
if strAction <> "new" then
	sql =	"SELECT " &_
				"D.DEVICE_ID, " &_
				"D.TAIL_CIRCUIT_ID, " &_
				"D.LOCAL_X25_DNA, " &_
				"D.POLL_CODE, " &_
				"H.HOST_X25_DNA_ID, " &_
				"H.HOST_DNA, " &_
				"H.HOST_DNA_MNEMONIC, " &_
				"TO_CHAR(D.CREATE_DATE_TIME,'MON-DD-YY HH24:MI:SS') AS CREATE_DATE_TIME, " &_
				"sma_sp_userid.spk_sma_library.sf_get_full_username(D.CREATE_REAL_USERID) as create_real_userid, " &_
				"D.UPDATE_DATE_TIME, " &_
				"TO_CHAR(D.UPDATE_DATE_TIME,'MON-DD-YY HH24:MI:SS') AS LAST_UPDATE_DATE_TIME, " &_
				"sma_sp_userid.spk_sma_library.sf_get_full_username(D.UPDATE_REAL_USERID) as update_real_userid, "&_
				"D.RECORD_STATUS_IND " &_
			"FROM CRP.RSAS_DEVICE D, CRP.RSAS_HOST_DNA H " &_
			"WHERE D.DEVICE_ID = " & strDeviceID & " and D.HOST_X25_DNA_ID = H.HOST_X25_DNA_ID (+)"

	'Response.Write (sql)
	'Response.end

	set rsDevice=server.CreateObject("ADODB.Recordset")
	rsDevice.CursorLocation = adUseClient
	rsDevice.Open sql, objConn
	if err then
		DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
	end if
	if rsDevice.EOF then
		DisplayError "BACK", "", err.Number, "CANNOT FIND DEVICE", err.Description
	end if
	set rsDevice.ActiveConnection = nothing



	'else insert sql for NEW action here........





end if


dim objRsPollCode
   'get a list of available Poll Codes
		strSQL = "SELECT C.CODE_ID, C.CODE_DESC, C.CODE_ORDER " &_
			"FROM CRP.RSAS_CODE C " &_
			"WHERE C.CODE_TYPE_CODE='PC' "&_
			"ORDER BY C.CODE_ORDER"

	'Response.Write (strSQL)
	'Response.End

	set objRsPollCode = objConn.Execute(strSQL)


dim objRsHostDNA
   'get a list of available Host DNA
		strSQL = "SELECT H.HOST_X25_DNA_ID, H.HOST_DNA, H.HOST_DNA_MNEMONIC " &_
			"FROM CRP.RSAS_HOST_DNA H " &_
			"WHERE H.RECORD_STATUS_IND='A' "&_
			"ORDER BY H.HOST_DNA"

	'Response.Write (strSQL)
	'Response.End

	set objRsHostDNA = objConn.Execute(strSQL)


dim myval, myval1

'myval = strDeviceID
'myval1 = (rsDevice("POLL_CODE")

'Response.Write "the value of strDeviceID is " & Vartype(myval)
'Response.Write "the value of pollcode is " & Vartype(myval1)

'Response.end

%>

<HTML>

<HEAD>

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

<LINK REL="stylesheet" TYPE="text/css" HREF="stylesheets/styles.css">

<TITLE>POS PLUS Device Detail</TITLE>

<script type="text/javascript" SRC="AccessLevels.js"></script>
<SCRIPT type="text/javascript" language="javascript" src="GeneralJavaFunctions.js"></SCRIPT>
<SCRIPT type="text/javascript">

<!--

var bolSaveRequired = false;
intAccessLevel=<%=intAccessLevel%>;

var intConst_MessageDisplay=<%=intConst_MessageDisplay%>;

function fct_onChange(){
	bolSaveRequired = true;
}

function btnNew_click(){
	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {alert('Access denied. Please contact your system administrator.'); return;}
	var strMasterID = "<%=strMasterID%>";
	document.location.href ="RSAST3DevDetail.asp?action=new&hdnDeviceID=0&masterID="+strMasterID;
}

function fct_onDelete(){
	if (document.frmRSASDeviceDetail.hdnDeviceID.value != '') {
	var strMasterID = "<%=strMasterID%>";
	if (((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) || (document.frmRSASDeviceDetail.txtRecordStatusInd.value == "D")) {alert('Access denied. Please contact your system administrator.'); return;}
	if (confirm('Do you really want to delete this object?')){
		document.location = "RSAST3DevDetail.asp?action=delete&masterID="+strMasterID+"&hdnDeviceID="+document.frmRSASDeviceDetail.hdnDeviceID.value+"&hdnLastUpdate="+escape(document.frmRSASDeviceDetail.hdnLastUpdate.value);
	}
	}else{fct_displayStatus('There is no need to delete an empty device.');}
}

function btnClose_onclick(){
	window.close();
}

function frmRSASDeviceDetail_onsubmit() {


	if	((((intAccessLevel & intConst_Access_Create) == intConst_Access_Create) && (document.frmRSASDeviceDetail.hdnDeviceID.value == "")) || (((intAccessLevel & intConst_Access_Update) == intConst_Access_Update) && (document.frmRSASDeviceDetail.hdnDeviceID.value != ""))) {
		if ((document.frmRSASDeviceDetail.txtLocalX25DNA.value != "") &&
			(document.frmRSASDeviceDetail.selPollCode.value != "") &&
			(document.frmRSASDeviceDetail.selHostX25DNAID.value != ""))
			{document.frmRSASDeviceDetail.action.value = "save";
			bolSaveRequired = false;
			document.frmRSASDeviceDetail.submit();
			return(true);}
		else
			{alert("You cannot save an empty device ... all fields are required. Please click DELETE button if you want to delete the device.");
			return(false);}
	} else {alert('Access denied. Please contact your system administrator.'); return(false);}
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

function body_onUnload(){
	//ADAM ... fix once can open separate window
	opener.document.frmRSASDetail.btn_iFrameRefresh.click();
}

//-->
</SCRIPT>
</HEAD>

<BODY onLoad="body_onLoad();" onBeforeUnload="body_onBeforeUnload();" onUnload="body_onUnload();">
<FORM name=frmRSASDeviceDetail LANGUAGE=javascript >
<INPUT type="hidden" name=action value="">
<INPUT type=hidden name=hdnLastUpdate value="<%if strDeviceID <> 0 then Response.Write rsDevice.Fields("UPDATE_DATE_TIME").value%>">
<INPUT type=hidden name=hdnDeviceID value="<%if strDeviceID <> 0 then Response.Write rsDevice("DEVICE_ID")%>">
<INPUT type=hidden name=MasterID value="<%=strMasterID%>">

<TABLE border=0 width=100%>
<THEAD>
	<TR ><TD colspan=2>POS PLUS Device Detail</td></tr>
</THEAD>

<TBODY>

<TR>
	<TD ALIGN=RIGHT width=20% NOWRAP>Local X25 DNA</TD>
	<TD width=80%><INPUT size=10 maxlength=10 name=txtLocalX25DNA value="<%if strDeviceID <> 0 then Response.write rsDevice("LOCAL_X25_DNA")%>" onchange ="fct_onChange();">
</TR>

<TR>
	<TD ALIGN="right"  NOWRAP>Poll Code</TD>
	<TD align=left>
	<SELECT id=selPollCode name=selPollCode tabindex=2 style="HEIGHT: 22px; WIDTH: 425px" onchange ="fct_onChange();">
		"<OPTION></OPTION>"
				<%
				Do while Not objRsPollCode.EOF

				Response.write "<OPTION "

					if (strDeviceID <> 0) then
					  if  ((objRsPollCode("CODE_DESC")) = (rsDevice("POLL_CODE"))) then
						Response.Write " SELECTED "
					  end if
					end if
				Response.Write 	" VALUE=" &objRsPollCode(1)& ">" &objRsPollCode(1) & "</OPTION>" &vbCrLf
				objRsPollCode.MoveNext
				Loop	%>
	</TD>
</TR>

<TR>
	<TD ALIGN="right"  NOWRAP>Host X25 DNA</TD>
	<TD align=left>
	<SELECT id=selHostX25DNAID name=selHostX25DNAID tabindex=3 style="HEIGHT: 22px; WIDTH: 425px" onchange ="fct_onChange();">
		"<OPTION></OPTION>"
				<%Do while Not objRsHostDNA.EOF
				Response.write "<OPTION "
					if strDeviceID <> 0 then
					  if clng(objRsHostDNA("HOST_X25_DNA_ID")) = clng(rsDevice("HOST_X25_DNA_ID")) then
						Response.Write " SELECTED "
					  end if
					end if
				Response.Write " VALUE=" &objRsHostDNA("HOST_X25_DNA_ID") & "> " &objRsHostDNA("HOST_DNA") & " ===> " &objRsHostDNA("HOST_DNA_MNEMONIC") & "</OPTION>" &vbCrLf
				objRsHostDNA.MoveNext
				Loop	%>
	</TD>
</TR>

</TBODY>

<TFOOT>
	<TR><TD align=right colspan=5>
	  	<INPUT type=button name=btnClose tabindex=4 style= "width: 2cm" value="Close" onclick="return btnClose_onclick()">&nbsp;
	  	<INPUT type=button name=btnDelete tabindex=5 style= "width: 2cm" value="Delete" onclick="return fct_onDelete();">&nbsp;
	  	<INPUT type=reset  name=btnReset tabindex=6 style= "width: 2cm" value="Reset" style="HEIGHT: 24px; WIDTH: 51px">&nbsp;
	  	<INPUT type=button name=btnAddNew tabindex=7 style= "width: 2cm" value="New" onclick="return btnNew_click()">&nbsp;
	  	<INPUT type=button name=btnSave tabindex=8 style= "width: 2cm" value="Save" onclick="return frmRSASDeviceDetail_onsubmit()">&nbsp;&nbsp;
	</TD></TR>
</TFOOT>
</TABLE>

<FIELDSET>
<LEGEND align=right><B>Audit Information</B></LEGEND>
<Div SIZE=8pt ALIGN=RIGHT>
	Record Status Indicator
	<INPUT align = left name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 18px"disabled value="<%if strDeviceID <> 0 then Response.write rsDevice("RECORD_STATUS_IND")%>" >&nbsp;&nbsp;&nbsp;
	Create Date&nbsp;<INPUT align=center name=txtCreateDateTime type=text style="HEIGHT: 20px; WIDTH: 150px" disabled value="<%if strDeviceID <> 0 then Response.write rsDevice("CREATE_DATE_TIME")%>" >&nbsp;
	Created By&nbsp;<INPUT align=right name=txtCreateRealUser type=text style="HEIGHT: 20px; WIDTH: 200px" disabled value="<%if strDeviceID <> 0 then Response.write rsDevice("CREATE_REAL_USERID")%>" ><BR>
	Update Date&nbsp;<INPUT align=center name=txtUpdateDateTime type=text style="HEIGHT: 20px; WIDTH: 150px" disabled value="<%if strDeviceID <> 0 then Response.write rsDevice("LAST_UPDATE_DATE_TIME")%>" >
	Updated By&nbsp;<INPUT align=right name=txtUpdateRealUser type=text style="HEIGHT: 20px; WIDTH: 200px" disabled value="<%if strDeviceID <> 0 then Response.write rsDevice("UPDATE_REAL_USERID")%>" >
</DIV>
</FIELDSET>

<%
'Clean up our ADO objects
if strDeviceID <> 0 then
	rsDevice.close
	set rsDevice = Nothing
	objConn.close
	set objConn = Nothing
end if
%>

</FORM>

</BODY>

</HTML>
