<%@  language="VBScript" %>
<% Option Explicit %>
<% on error resume next %>
<!--#include file="sma_env.inc"-->
<!--#include file="smaConstants.inc"-->
<!--#include file="databaseconnect.asp"-->
<!--#include file="smaProcs.inc"-->
<%

'check user's rights
dim intAccessLevel
intAccessLevel = CInt(CheckLogon(strConst_CustomerNameAlias))
if intAccessLevel and intConst_Access_ReadOnly <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to customer alias name. Please contact your system administrator"
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
	DisplayError "BACK", "", 0, "UNEXPECTED ERROR", "Close alias window to return to customer form."
end if
strLastUpdate = Request("hdnLastUpdate")

'Response.Write ("strAliasID = " & strAliasID & ", " & "strLastUpdate = " & strLastUpdate )
'Response.end

'save changes?
if strAction = "save" then
	dim strNameAlias, strLastUpdate
	strMasterID = Request("masterID")
	strNameAlias = Request("txtNameAlias")
	'call stored proc to save the record

	if (strAliasID <> "") and (intAccessLevel and intConst_Access_Update = intConst_Access_Update) then
		'create command object for update stored proc
		dim cmdUpdateObj
		set cmdUpdateObj = server.CreateObject("ADODB.Command")
		set cmdUpdateObj.ActiveConnection = objConn
		cmdUpdateObj.CommandType = adCmdStoredProc
		cmdUpdateObj.CommandText = "sma_sp_userid.spk_sma_cust_inter.sp_cust_alias_update"
		'create params
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_user_id", adVarChar , adParamInput, 30, strRealUserID)
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_cna_id", adNumeric , adParamInput,, CLng(strAliasID))
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_customer_id", adNumeric , adParamInput,, CLng(strMasterID))
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_cna_upper", adVarChar, adParamInput, 30, UCase(strNameAlias))
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , CDate(strLastUpdate))
		'call the update stored proc
		if err then
			'DisplayError "BACK", "", err.Number, "CANNOT UPDATE OBJECT - PARAMETER ERROR", err.Description
			objConn.Errors.Clear
		end if
		cmdUpdateObj.Execute
		if objConn.Errors.Count <> 0 then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT UPDATE OBJECT", objConn.Errors(0).Description
			objConn.Errors.Clear
		end if
		strWinMessage = "Record saved successfully."
	elseif (strAliasID = "") and (intAccessLevel and intConst_Access_Create = intConst_Access_Create) then
		'create command object for insert stored proc
		dim cmdInsertObj
		set cmdInsertObj = server.CreateObject("ADODB.Command")
		set cmdInsertObj.ActiveConnection = objConn
		cmdInsertObj.CommandType = adCmdStoredProc
		cmdInsertObj.CommandText = "sma_sp_userid.spk_sma_cust_inter.sp_cust_alias_insert"
		'create params
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_user_id", adVarChar , adParamInput, 30, strRealUserID)
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_cna_id", adNumeric , adParamOutput)
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_customer_id", adNumeric , adParamInput,, CLng(strMasterID))
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_cna_upper", adVarChar, adParamInput, 30, Ucase(strNameAlias))

		'Response.Write "<b> count = " & cmdInsertObj.Parameters.count & "<br>"
		'dim nx
		'for nx = 0 to cmdInsertObj.Parameters.Count-1
		'	Response.Write cmdInsertObj.Parameters.Item(nx).Name & " = " & cmdInsertObj.Parameters.Item(nx).Value & " <br>"
		'next
		'Response.end

		if objConn.Errors.Count <> 0 then
			'DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT UPDATE OBJECT - PARAMETER ERROR", objConn.Errors(0).Description
			objConn.Errors.Clear
		end if
		'call the update stored proc
		cmdInsertObj.Execute
		if objConn.Errors.Count <> 0 then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT UPDATE OBJECT", objConn.Errors(0).Description
			objConn.Errors.Clear
		end if
		strAliasID = cmdInsertObj.Parameters("p_cna_id").Value	'set return parameter
		if strAliasID = "" then
			DisplayError "BACK", "", 2100, "CANNOT DISPLAY NEW ALIAS.", "Most probably the new alias has been saved successfully even if there was an error retrieving the new id. Close the alias window to return to the customer screen."
			objConn.Errors.Clear
		end if
		strWinMessage = "Record saved successfully."
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

    if objConn.Errors.Count <> 0 then
	'		DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT - PARAMETER ERROR", objConn.Errors(0).Description
			objConn.Errors.Clear
		end if

		'create command object for update stored proc
		dim cmdDeleteObj
		set cmdDeleteObj = server.CreateObject("ADODB.Command")
		set cmdDeleteObj.ActiveConnection = objConn
		cmdDeleteObj.CommandType = adCmdStoredProc
		cmdDeleteObj.CommandText = "sma_sp_userid.spk_sma_cust_inter.sp_cust_alias_delete"
		'create params
        'Response.Write(Request("aliasID"))
        'DisplayError Request("aliasID")

		cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_cna_id", adNumeric , adParamInput,, Request("aliasID"))
       ' cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_cna_id", adNumeric , adParamInput,, 17553)
		cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_last_update", adDBTimeStamp, adParamInput,, CDate(strLastUpdate)) 
        'cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_last_update", adDBTimeStamp, adParamInput,, CDate("12/19/2016 10:16:55 AM"))
        cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("P_USER_ID", adVarChar , adParamInput, 30, strRealUserID)
		'call the update stored proc
		
		cmdDeleteObj.Execute
		if objConn.Errors.Count <> 0 then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT", objConn.Errors(0).Description
             'DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT", CInt(Request("aliasID"))
			objConn.Errors.Clear
		end if
		strWinMessage = "Record deleted successfully."
	else
		DisplayError "BACK", "", err.Number, "ACCESS DENIED. PLEASE CONTACT YOUR SYSTEM ADMINISTRATOR.", err.Description
	end if
	if err then
	DisplayError "BACK", "", err.Number, "CANNOT DELETE ALIAS", err.Description
    'DisplayError "BACK", "", err.Number, "CANNOT DELETE ALIAS", CInt(Request("aliasID"))
	end if
	'called from main form?
	if Request("back") = "true" then
		Response.Redirect "CustAlias.asp?CustomerID="&strMasterID
	end if
	'ready to enter a new alias?
	strAliasID=""
	strAction="new"
end if

'display the alias info
if strAction <> "new" then
	sql =	"SELECT " &_
				"CUSTOMER_NAME_ALIAS_ID, " &_
				"CUSTOMER_ID, " &_
				"CUSTOMER_NAME_ALIAS_UPPER, " &_
				"TO_CHAR(CREATE_DATE_TIME,'MON-DD-YY HH24:MI:SS')AS CREATE_DATE_TIME, " &_
				"sma_sp_userid.spk_sma_library.sf_get_full_username(CREATE_REAL_USERID) as create_real_userid, " &_
				"UPDATE_DATE_TIME, " &_
				"TO_CHAR(UPDATE_DATE_TIME,'MON-DD-YY HH24:MI:SS')AS LAST_UPDATE_DATE_TIME, " &_
				"sma_sp_userid.spk_sma_library.sf_get_full_username(UPDATE_REAL_USERID) as update_real_userid, "&_
				"RECORD_STATUS_IND " &_
			"FROM CRP.CUSTOMER_NAME_ALIAS " &_
			"WHERE CUSTOMER_NAME_ALIAS_ID = " & strAliasID

	'Response.Write (sql)
	'Response.end

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
<html>
<head>
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
    <link rel="stylesheet" type="text/css" href="stylesheets/styles.css">
    <title>Customer Name Alias Detail</title>

    <script type="text/javascript" src="AccessLevels.js"></script>
    <script type="text/javascript">
        var bolSaveRequired = false;
        intAccessLevel=<%=intAccessLevel%>;
        var intConst_MessageDisplay=<%=intConst_MessageDisplay%>;

        function fct_onChange(){
            bolSaveRequired = true;
        }

        function btnNew_click(){
            if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {alert('Access denied. Please contact your system administrator.'); return;}
            var strMasterID = "<%=strMasterID%>";
            document.location.href ="CustAliasDetail.asp?action=new&masterID="+strMasterID;
        }

        function fct_onDelete(){
            if (document.frmCustAliasDetail.AliasID.value != '') {
                var strMasterID = "<%=strMasterID%>";
                if (((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) || (document.frmCustAliasDetail.txtRecordStatusInd.value == "D")) {alert('Access denied. Please contact your system administrator.'); return;}
                if (confirm('Do you really want to delete this object?')){
                    document.location = "CustAliasDetail.asp?action=delete&masterID="+strMasterID+"&AliasID="+document.frmCustAliasDetail.AliasID.value+"&hdnLastUpdate="+escape(document.frmCustAliasDetail.hdnLastUpdate.value);
                }
            }else{fct_displayStatus('There is no need to delete an empty alias.');}
        }

        function btnClose_onclick(){
            window.close();
        }

        function frmCustAliasDetail_onsubmit() {
            if	((((intAccessLevel & intConst_Access_Create) == intConst_Access_Create) && (document.frmCustAliasDetail.AliasID.value == "")) || (((intAccessLevel & intConst_Access_Update) == intConst_Access_Update) && (document.frmCustAliasDetail.AliasID.value != ""))) {
                if (document.frmCustAliasDetail.txtNameAlias.value != "")
                {document.frmCustAliasDetail.action.value = "save";
                    bolSaveRequired = false;
                    document.frmCustAliasDetail.submit();
                    return(true);}
                else
                {alert("You cannot save an empty alias. Please click DELETE button if you want to delete the alias.");
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
            opener.document.frmCustDetail.btn_iFrameRefresh.click();
        }
    </script>
</head>

<body onload="body_onLoad();" onbeforeunload="body_onBeforeUnload();" onunload="body_onUnload();">
    <form name="frmCustAliasDetail" language="javascript">
        <input type="hidden" name="action" value="">
        <input type="hidden" name="hdnLastUpdate" value="<%if strAliasID <> "" then Response.Write rsAlias.Fields("UPDATE_DATE_TIME").value%>">
        <input type="hidden" name="AliasID" value="<%if strAliasID <> "" then Response.Write rsAlias("CUSTOMER_NAME_ALIAS_ID")%>">
        <input type="hidden" name="MasterID" value="<%=strMasterID%>">

        <table border="0" width="100%">
            <thead>
                <tr>
                    <td colspan="2">Customer Name Alias Detail</td>
                </tr>
            </thead>

            <tbody>
                <tr>
                    <td align="RIGHT" width="20%" nowrap>Alias Name</td>
                    <td width="80%">
                        <input size="40" maxlength="30" name="txtNameAlias" value="<%if strAliasID <> "" then Response.write rsAlias("CUSTOMER_NAME_ALIAS_UPPER")%>" onchange="fct_onChange();">
                </tr>
            </tbody>

            <tfoot>
                <tr>
                    <td align="right" colspan="5">
                        <input type="button" name="btnClose" style="width: 2cm" value="Close" onclick="return btnClose_onclick()">&nbsp;
	  	<input type="button" name="btnDelete" style="width: 2cm" value="Delete" onclick="return fct_onDelete();">&nbsp;
	  	<input type="reset" name="btnReset" style="width: 2cm" value="Reset" style="height: 24px; width: 51px">&nbsp;
	  	<input type="button" name="btnAddNew" style="width: 2cm" value="New" onclick="return btnNew_click()">&nbsp;
	  	<input type="button" name="btnSave" style="width: 2cm" value="Save" onclick="return frmCustAliasDetail_onsubmit()">&nbsp;&nbsp;
                    </td>
                </tr>
            </tfoot>
        </table>

        <fieldset>
            <legend align="right"><b>Audit Information</b></legend>
            <div size="8pt" align="RIGHT">
                Record Status Indicator
	<input align="left" name="txtRecordStatusInd" type="text" style="height: 20px; width: 18px" disabled value="<%if strAliasID <> "" then Response.write rsAlias("RECORD_STATUS_IND")%>">&nbsp;&nbsp;&nbsp;
	Create Date&nbsp;<input align="center" name="txtCreateDateTime" type="text" style="height: 20px; width: 150px" disabled value="<%if strAliasID <> "" then Response.write rsAlias("CREATE_DATE_TIME")%>">&nbsp;
	Created By&nbsp;<input align="right" name="txtCreateRealUser" type="text" style="height: 20px; width: 200px" disabled value="<%if strAliasID <> "" then Response.write rsAlias("CREATE_REAL_USERID")%>"><br>
                Update Date&nbsp;<input align="center" name="txtUpdateDateTime" type="text" style="height: 20px; width: 150px" disabled value="<%if strAliasID <> "" then Response.write rsAlias("LAST_UPDATE_DATE_TIME")%>">
                Updated By&nbsp;<input align="right" name="txtUpdateRealUser" type="text" style="height: 20px; width: 200px" disabled value="<%if strAliasID <> "" then Response.write rsAlias("UPDATE_REAL_USERID")%>">
            </div>
        </fieldset>
        <%
'Clean up our ADO objects
if strAliasID <> "" then
	rsAlias.close
	set rsAlias = Nothing
	objConn.close
	set objConn = Nothing
end if
        %>
    </form>
</body>
</html>
