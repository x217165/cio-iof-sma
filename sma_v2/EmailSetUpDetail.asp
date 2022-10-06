<%@ Language=VBScript %>
<% option explicit %>
<% Response.Buffer = true %>
<%on error resume next%>

<!--
********************************************************************************************
* Page name:	EmailSetupDetail.asp

* Purpose:		To display the detailed information about an email setup.
*				Entry is chosen via EmailSetupList.asp
*
* Created by:	Shawn Myers	10/04/2000
*
********************************************************************************************
-->
<!--#include file="SmaConstants.inc"-->
<!--#include file="smaProcs.inc"-->
<!--#include file="databaseconnect.asp"-->

<%



'********************************
'check the present user's rights*
'********************************

dim intAccessLevel
intAccessLevel = CInt(CheckLogon(strConst_AssetCatalogue))


if intAccessLevel and intConst_Access_ReadOnly <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to managed object. Please contact your system administrator"
end if



'*****variables*****

dim strID

dim strWinLocation
dim strWinMessage
dim strRealUserID

'*****variables*****



'get the service status change id from the list page string
strID = Request("hdnServiceStatusChangeID")

'Response.Write "the passed id from the list screen is equal to " & strID & "<BR>"


'get the hidden window location
strWinLocation = "EmailSetUpDetail.asp?ServiceStatusChangeID="& Request("strID")

'set the variable for the UserInfo cookie
strRealUserID = Session("username")


'************************
'do save, insert, delete*
'************************


dim aFromStatus		'used to get the ID from the "From Status" drop down list
dim aToStatus		'used to get the ID from the "To Status" drop down list

select case Request("hdnFrmAction")

	case "SAVE"

		'get value of flags from the seven checkboxes

		dim strCustCareStaffFlag, strPortfolioStaffFlag, strDesignStaffFlag, strImplManStaffFlag
		dim strImplStaffFlag, strInstallationStaffFlag, strOperationsStaffFlag

		if Lcase(Request.Form("chkCustCareStaff")) = "on" then
			strCustCareStaffFlag = "Y"
		else
			strCustCareStaffFlag = "N"
		end if

		if Lcase(Request.Form("chkPortfolioStaff")) = "on" then
			strPortfolioStaffFlag = "Y"
		else
			strPortfolioStaffFlag = "N"
		end if

		if Lcase(Request.Form("chkDesignStaff")) = "on" then
			strDesignStaffFlag = "Y"
		else
			strDesignStaffFlag = "N"
		end if

		if Lcase(Request.Form("chkImplementManager")) = "on" then
			strImplManStaffFlag = "Y"
		else
			strImplManStaffFlag = "N"
		end if

		if Lcase(Request.Form("chkImplementStaff")) = "on" then
			strImplStaffFlag = "Y"
		else
			strImplStaffFlag = "N"
		end if

		if Lcase(Request.Form("chkInstallationStaff")) = "on" then
			strInstallationStaffFlag = "Y"
		else
			strInstallationStaffFlag = "N"
		end if

		if Lcase(Request.Form("chkOperationsStaff")) = "on" then
			strOperationsStaffFlag = "Y"
		else
			strOperationsStaffFlag = "N"
		end if


	'check to see if entry exists already in database by checking for the existence of id


	  if Request("hdnServiceStatusChangeID")  <> "" then  ' it is an existing record so save the changes


		if intAccessLevel and intConst_Access_Update <> intConst_Access_Update then

			DisplayError "REFRESH", strWinLocation, 0, "UPDATE DENIED", "You don't have access to update managed objects. Please contact your system administrator"

		end if

		dim cmdUpdateObj

		set cmdUpdateObj = server.CreateObject("ADODB.Command")
		set cmdUpdateObj.ActiveConnection = objConn
		cmdUpdateObj.CommandType = adCmdStoredProc

		aFromStatus = split(Request("selFromServStatCode"),"¿")
		aToStatus = split(Request("selToServStatCode"),"¿")

		'get the email_setup_detail stored update procedure <schema.package.procedure>
		cmdUpdateObj.CommandText = "sma_sp_userid.spk_sma_cserv_inter.sp_serv_stat_change_update"

		'create the associated parameters
		'user id associated with time stamp


		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_userid",adVarChar,adParamInput, 20,strRealUserID)
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_service_status_change_id", adNumeric, adParamInput, 9, Clng(Request("hdnServiceStatusChangeID")))
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_from_service_status_code",adVarChar, adParamInput, 6, aFromStatus(0))
        cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_to_service_status_code", adVarChar, adParamInput, 6, aToStatus(0))
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_notify_cust_care_staff_flag",adChar, adParamInput, 1, strCustCareStaffFlag)
        cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_notify_portfolio_staff_flag", adChar, adParamInput, 1, strPortfolioStaffFlag)
        cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_notify_design_staff_flag", adChar, adParamInput, 1, strDesignStaffFlag)
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_notify_impl_manager_flag", adChar, adParamInput, 1, strImplManStaffFlag)
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_notify_impl_staff_flag", adChar, adParamInput, 1, strImplStaffFlag)
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_notify_inst_staff_flag", adChar, adParamInput, 1, strInstallationStaffFlag)
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_notify_oper_staff_flag", adChar, adParamInput, 1, strOperationsStaffFlag)
        cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_last_update_dt", adDBTimeStamp,adParamInput, , CDate(Request("hdnUpdateDateTime")))
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_addition_distribution_list", adVarChar, adParamInput, 2000, Request("txtDistList"))


		'execute the update object


		cmdUpdateObj.Execute

			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT UPDATE OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			end if

			strWinMessage = "Record saved successfully. You can now see the changes you made."





	  else 'create a new record


	    if intAccessLevel and intConst_Access_Create<> intConst_Access_Create then

			DisplayError "REFRESH", strWinLocation, 0, "INSERT DENIED", "You don't have access to create managed objects. Please contact your system administrator"

		end if

		dim cmdInsertObj
		set cmdInsertObj = server.CreateObject("ADODB.Command")
		set cmdInsertObj.ActiveConnection = objConn
		cmdInsertObj.CommandType = adCmdStoredProc

		aFromStatus = split(Request("selFromServStatCode"),"¿")
		aToStatus = split(Request("selToServStatCode"),"¿")


		'get the asset_catalog_detail insert procedure <schema.package.procedure>
		cmdInsertObj.CommandText = "sma_sp_userid.spk_sma_cserv_inter.sp_serv_stat_change_insert"



		'create the insert parameters

		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_userid", adVarChar, adParamInput, 20,strRealUserID)
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_service_status_change_id", adNumeric, adParamOutput,, null)
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_from_service_status_code",adVarChar, adParamInput, 6, aFromStatus(0))
        cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_to_service_status_code", adVarChar, adParamInput, 6, aToStatus (0))
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_notify_cust_care_staff_flag",adChar, adParamInput, 1, strCustCareStaffFlag)
        cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_notify_portfolio_staff_flag", adChar, adParamInput, 1, strPortfolioStaffFlag)
        cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_notify_design_staff_flag", adChar, adParamInput, 1, strDesignStaffFlag)
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_notify_impl_manager_flag", adChar, adParamInput, 1, strImplManStaffFlag)
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_notify_impl_staff_flag", adChar, adParamInput, 1, strImplStaffFlag)
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_notify_inst_staff_flag", adChar, adParamInput, 1, strInstallationStaffFlag)
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_notify_oper_staff_flag", adChar, adParamInput, 1, strOperationsStaffFlag)
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_addition_distribution_list", adVarChar, adParamInput, 2000, Request("txtDistList"))


		' execute the insert object

		cmdInsertObj.Execute

			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT CREATE NEW OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			else
				strID = cmdInsertObj.Parameters("p_service_status_change_id").Value
			end if
			strWinMessage = "Record created successfully. You can now see the new record."

	  end if


	case "DELETE"


	        if intAccessLevel and intConst_Access_Delete<> intConst_Access_Delete then

				DisplayError "REFRESH", strWinLocation, 0, "DELETE DENIED", "You don't have access to delete managed objects. Please contact your system administrator"

			end if

			dim cmdDeleteObj
			set cmdDeleteObj = server.CreateObject("ADODB.Command")
			set cmdDeleteObj.ActiveConnection = objConn
			cmdDeleteObj.CommandType = adCmdStoredProc

			'get the email setup detail delete procedure <schema.package.procedure>
			cmdDeleteObj.CommandText = "sma_sp_userid.spk_sma_cserv_inter.sp_serv_stat_change_delete"

			'create the delete parameters
			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_service_status_change_id", adNumeric, adParamInput, , CLng(strID))
			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, ,CDate(Request("hdnUpdateDateTime")))


			'execute the delete object
			cmdDeleteObj.Execute

			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			end if

			strID = 0
			strWinMessage = "Record deleted successfully."

end select


'*************************
'end save, insert, delete*
'*************************









'ok, now go get the detailed Email Setup information


'declare the connection and sql variables
Dim strSQL, strSelectClause, strFromClause, strWhereClause
Dim rsEmailDetail

dim objCmd
dim rsStatusCode

'connect to the database
'<<CONNECT>>

'use the sqlstring to extract the necessary information from the database

	if strID <> 0 then

		strSelectClause = "SELECT " &_
					"t1.service_status_change_id, " & _
					"t1.from_service_status_code, " & _
					"t1.to_service_status_code, " & _
					"t1.notify_cust_care_staff_flag, " & _
					"t1.notify_portfolio_staff_flag, " & _
					"t1.notify_design_staff_flag, " & _
					"t1.notify_implement_manager_flag, " & _
					"t1.notify_implement_staff_flag, " & _
					"t1.notify_installation_staff_flag, " & _
					"t1.notify_operations_staff_flag, " & _
					"t1.addition_distribution_list, " & _
					"to_char(t1.create_date_time,'MON-DD-YYYY HH24:MI:SS') as create_date, " & _
					"sma_sp_userid.spk_sma_library.sf_get_full_username(t1.create_real_userid) as create_real_userid, " & _
					"to_char(t1.update_date_time,'MON-DD-YYYY HH24:MI:SS') as update_date, " & _
					"sma_sp_userid.spk_sma_library.sf_get_full_username(t1.update_real_userid) as update_real_userid, " & _
					"t1.update_date_time as last_update_date_time, " & _
					"t1.record_status_ind "


		strFromClause =	" FROM crp.service_status_change  t1 "


		strWhereClause = " WHERE " & _
					"t1.service_status_change_id = " & strID


		strSQL =  strSelectClause & strFromClause & strWhereClause

		'show SQL for debugging if necessary by using>>
		'Response.Write "<BR>" & strsql	 & "<br>"

		'set and open the email_detail recordset and database connection

		set rsEmailDetail = Server.CreateObject("ADODB.Recordset")

		rsEmailDetail.CursorLocation = adUseClient
		rsEmailDetail.Open strSQL, objConn

		if err then
			DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
		end if
		if rsEmailDetail.EOF then
			DisplayError "BACK", "", 999, "CANNOT FIND REQUESTED OBJECT", "EOF condition occurred in rsEmailDetail recordset."
		end if


	end if



'Load the From and To Service Status Code Listboxes with their own sql

strsql = " SELECT service_status_code" &_
		 " ,      service_status_name" &_
		 " FROM   crp.service_status" &_
		 " ORDER  BY service_status_code"

'Create the Command object

set objCmd = Server.CreateObject("ADODB.command")
    objCmd.ActiveConnection = objconn
    objCmd.CommandText = strsql
    objCmd.CommandType = adCmdText

'Create the StatusCode Recordset object

set rsStatusCode = Server.CreateObject("ADODB.Recordset")

		rsStatusCode.CursorLocation = adUseClient
		rsStatusCode.Open strsql, objConn

		if err then
			DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
		end if
		if rsStatusCode.EOF then
			DisplayError "BACK", "", 999, "CANNOT FIND REQUESTED OBJECT", "EOF condition occurred in rsStatusCode recordset."
		end if


%>



<HTML>
<HEAD>
	<LINK rel="stylesheet" type="text/css" href="stylesheets/styles.css">
	<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
	<script type="text/javascript" SRC="AccessLevels.js"></script>

	<SCRIPT LANGUAGE=JavaScript>
	<!--

	var strWinMessage = '<%=strWinMessage%>';
    var intAccessLevel = '<%=intAccessLevel%>';
    var bolNeedToSave = false ;

setPageTitle("SMA - Email Setup Detail");

	function fct_NewEmailSetupEntry(){

	//alert ('in the new function');

	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create)

			{

			alert('Access denied. Please contact your system administrator.');
			return (false);

			}


			self.document.location.href = "EmailSetupDetail.asp?hdnServiceStatusChangeID=0";



		}


	//OK

	function fct_onSave(){

	//alert('in the save function' + <%=intAccessLevel%>);

	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create)
			{
				alert('Access denied. Please contact your system administrator.');
				return (false);
			}
			else
			{
				if (document.frmEmailSetupDetail.txtFromStatusCode.value == "" )
					{
						alert('Please select a From Service Status Code');
						//document.txtFromStatus.btnMakeLookup.focus();
						return(false);
					}
				if (document.frmEmailSetupDetail.txtToStatusCode.value == "" )
					{
						alert('Please select a To Service Status Code');
						//document.frmAssetCatDetail.btnModelLookup.focus();
						return(false);
					}
					else
					{
					document.frmEmailSetupDetail.hdnFrmAction.value = "SAVE";
					bolNeedToSave = false;
					return(true);
					}
			}
    }
	//OK

	function fct_onDelete() {
	if ((intAccessLevel & intConst_Access_Delete)!= intConst_Access_Delete)
				{
					alert('Access denied. Please contact your system administrator.');
					return (false);
				}
		var strID = document.frmEmailSetupDetail.hdnServiceStatusChangeID.value;
		var strUpdateDate = document.frmEmailSetupDetail.hdnUpdateDateTime.value;

	    //alert ('the value of the statusid is '+ document.frmEmailSetupDetail.hdnServiceStatusChangeID.value);
	    //alert ('the value of the update date time is '+ document.frmEmailSetupDetail.hdnUpdateDateTime.value);
				{
					if (confirm('Do you really want to delete this object?'))
						{
						document.location = "EmailSetUpDetail.asp?hdnFrmAction=DELETE&hdnServiceStatusChangeID="+strID+"&hdnUpdateDateTime="+strUpdateDate ;
						}
				}
		}

	function fct_onReset() {
		if(confirm('All changes will be lost. Do you really want to reset the page?')){
			bolNeedToSave = false ;
			document.location = 'EmailSetUpDetail.asp?hdnServiceStatusChangeID='+ "<%=strID%>" ;
		}
	}



	function body_onbeforeunload(){

		document.frmEmailSetupDetail.btnSave.focus();

		if ((bolNeedToSave == true) && ((<%=intAccessLevel%> & <%=intConst_Access_Update%>) == <%=intConst_Access_Update%>))
				{
					event.returnValue = "There is unsaved data on the screen. To save changes, click CANCEL below then click SAVE on the main form.";
				}
		}

	function fct_clearStatus() {
		window.status = "";
	}

	function fct_DisplayStatus(strWindowStatus){
		window.status=strWindowStatus;
		setTimeout('fct_clearStatus()', '<%=intConst_MessageDisplay%>');
    }

	function fct_onChangeFromStatus() {

		var strWhole;
		var strStatusDesc, intStart, intIndex;

		intIndex = document.frmEmailSetupDetail.selFromServStatCode.selectedIndex;
		strWhole = document.frmEmailSetupDetail.selFromServStatCode.options[intIndex].value;
		intStart = strWhole.indexOf('<%=strDelimiter%>');
		document.frmEmailSetupDetail.txtFromStatusCode.value = strWhole.substr(intStart+1);
		fct_onChange();
	}


	function fct_onChangeToStatus() {

		var strWhole;
		var strStatusDesc, intStart, intIndex;

		intIndex = document.frmEmailSetupDetail.selToServStatCode.selectedIndex;
		strWhole = document.frmEmailSetupDetail.selToServStatCode.options[intIndex].value;
		intStart = strWhole.indexOf('<%=strDelimiter%>');
		document.frmEmailSetupDetail.txtToStatusCode.value = strWhole.substr(intStart+1);
		fct_onChange();
	}

    function fct_onChange(){

		bolNeedToSave = true;
	}


	function btnReferences_onclick() {

		var strOwner = 'CRP' ;
		var strTableName = 'SERVICE_STATUS_CHANGE' ;
		var strRecordID = document.frmEmailSetupDetail.hdnServiceStatusChangeID.value ;
		var URL ;

		if ( strID = 0)
				{
					alert("No references. This is a new record.");
				}
		else
				{
				URL ='Dependency.asp?Owner=' + strOwner + '&TableName=' + strTableName + '&RecordID='+ strRecordID   ;
				window.open(URL, 'Popup', 'top=100, left=100, WIDTH=500, HEIGHT=300'  ) ;
				}
	}

//-->
</SCRIPT>

</HEAD>

<BODY onLoad="fct_DisplayStatus(strWinMessage);" onbeforeunload="body_onbeforeunload();">
<FORM name=frmEmailSetupDetail action="EmailSetUpDetail.asp"  method="POST" onsubmit="return fct_onSave()">

	<!--
	<INPUT name=hdnServiceStatusChangeID type=hidden value= <%if strID <> 0 then Response.Write rsEmailDetail("SERVICE_STATUS_CHANGE_ID") else Response.Write """""" end if%>>
	-->
	<INPUT name=hdnServiceStatusChangeID type=hidden value="<%if strID <> 0 then  Response.Write rsEmailDetail("SERVICE_STATUS_CHANGE_ID") else Response.Write """""" end if%>">
	<INPUT name=hdnUpdateDateTime type=hidden value="<%if strID <> 0 then  Response.Write rsEmailDetail("LAST_UPDATE_DATE_TIME") else Response.Write """""" end if%>">
	<INPUT id=hdnFrmAction name=hdnFrmAction type=hidden value= "">

	<!-- user interface -->

	<TABLE border=0 width=100%>

	<thead>
		<tr>
			<td colspan=3 align=left>Email Setup Detail</td>
		</tr>
	</thead>

	<TR>

		<TD ALIGN=RIGHT width=20% NOWRAP>From Service Status Code<font color=red>*</font></TD>
		<TD colspan="2">
			<SELECT id=selFromServStatCode name=selFromServStatCode style="HEIGHT: 20px; WIDTH: 120px" onChange="return fct_onChangeFromStatus();">
			<%
			Dim fromDesc	'used to set the intial value of txtFromStatusCode
			Do while Not rsStatusCode.EOF
				Response.write "<OPTION "
				if fromDesc = "" then
					fromDesc = rsStatusCode("SERVICE_STATUS_NAME")
				end if
				if strID <> 0 then
					if rsStatusCode("SERVICE_STATUS_CODE") = rsEmailDetail("FROM_SERVICE_STATUS_CODE") then
						Response.Write " selected "
						fromDesc = rsStatusCode("SERVICE_STATUS_NAME")
					end if
				end if
				Response.Write " VALUE =""" & routineHTMLString(rsStatusCode("SERVICE_STATUS_CODE")& strDelimiter & rsStatusCode("SERVICE_STATUS_NAME")) & """>" & routineHTMLString(rsStatusCode("SERVICE_STATUS_CODE")) & "</OPTION>"
				rsStatusCode.MoveNext
			Loop
			%>
			</SELECT>
			<INPUT id=txtFromStatusCode name=txtFromStatusCode value="<%=fromDesc%>" disabled style="WIDTH: 240px">
		</TD>

	</TR>


	<TR>

		<TD ALIGN=RIGHT width=20% NOWRAP>To Service Status Code<font color=red>*</font></TD>
		<TD colspan="2">
			<SELECT id=selToServStatCode name=selToServStatCode style="HEIGHT: 20px; WIDTH: 120px" onChange="return fct_onChangeToStatus();">
			<%
			Dim toDesc	'used to set the initial value of txtToStatusCode
			rsStatusCode.movefirst
			Do while Not rsStatusCode.EOF
				Response.write "<OPTION "
				if toDesc = "" then
					toDesc = rsStatusCode("SERVICE_STATUS_NAME")
				end if
				if strID <> 0 then
					if rsStatusCode("SERVICE_STATUS_CODE") = rsEmailDetail("TO_SERVICE_STATUS_CODE") then
						Response.Write " selected "
						toDesc = rsStatusCode("SERVICE_STATUS_NAME")
					end if
				end if
				Response.Write " VALUE =""" & routineHTMLString(rsStatusCode("SERVICE_STATUS_CODE")& strDelimiter & rsStatusCode("SERVICE_STATUS_NAME")) & """>" & routineHTMLString(rsStatusCode("SERVICE_STATUS_CODE")) & "</OPTION>"
				rsStatusCode.MoveNext
			Loop
			%>
			</SELECT>
			<INPUT id=txtToStatusCode name=txtToStatusCode value="<%=toDesc%>" disabled style="WIDTH: 240px">
		</TD>

	</TR>

	<TR>
		<TD width=15% align=right nowrap>Notify Customer Care Staff</TD>
		<TD align=left width=27px><INPUT id=chkCustCareStaff name=chkCustCareStaff type=checkbox <%if strID <> 0 then IF rsEmailDetail("Notify_Cust_Care_Staff_Flag")="Y" THEN  Response.Write "CHECKED"  end if END IF%> onChange="fct_onChange();"></TD>

		<TD align=left nowrap>Notify Implement Staff
		<INPUT id=chkImplementStaff name=chkImplementStaff type=checkbox <%if strID <> 0 then IF rsEmailDetail("Notify_Implement_Staff_Flag")="Y" THEN  Response.Write "CHECKED"  end if END IF%> onChange="fct_onChange();"></TD>

	</TR>

	<TR>
		<TD width=15% align=right nowrap>Notify Portfolio Staff</TD>
		<TD width=27px align=left><INPUT id=chkPortfolioStaff name=chkPortfolioStaff type=checkbox <%if strID <> 0 then IF rsEmailDetail("Notify_Portfolio_Staff_Flag")="Y" THEN  Response.Write "CHECKED"  end if END IF%> onChange="fct_onChange();"></TD>

		<TD align=left wrap>Notify Installation Staff
		<INPUT id=chkInstallationStaff name=chkInstallationStaff type=checkbox <%if strID <> 0 then IF rsEmailDetail("Notify_Installation_Staff_Flag")="Y" THEN  Response.Write "CHECKED"  end if END IF%> onChange="fct_onChange();"></TD>

	</TR>

	<TR>
		<TD width=15% align=right nowrap>Notify Design Staff</TD>
		<TD width=27px align=left><INPUT id=chkDesignStaff name=chkDesignStaff type=checkbox <%if strID <> 0 then IF rsEmailDetail("Notify_Design_Staff_Flag")="Y" THEN  Response.Write "CHECKED"  end if END IF%> onChange="fct_onChange();"></TD>

		<TD align=left wrap>Notify Operations Staff
		<INPUT id=chkOperationsStaff name=chkOperationsStaff type=checkbox <%if strID <> 0 then IF rsEmailDetail("Notify_Operations_Staff_Flag")="Y" THEN  Response.Write "CHECKED"  end if END IF%> onChange="fct_onChange();"></TD>

	</TR>

	<tr>
		<TD width=15% align=right nowrap>Notify Implement Manager</TD>
		<TD width=27px align=left><INPUT id=chkImplementManager name=chkImplementManager type=checkbox <%if strID <> 0 then IF rsEmailDetail("Notify_Implement_Manager_Flag")="Y" THEN  Response.Write  "CHECKED"  end if END IF%> onChange="fct_onChange();"></TD>

	</tr>

	<TR>
		<td  align=right valign=top>Addition Distribution List</td>
		<td  colspan=2 rowspan=3 valign="top"><TEXTAREA cols=25 name=txtDistList rows=6 style="width: 800px" onChange="fct_onChange();"><%if strID <> 0 then Response.Write routineHTMLString(rsEmailDetail("ADDITION_DISTRIBUTION_LIST")) else Response.Write null end if%></TEXTAREA></td>

	</TR>


	<tfoot>
	<tr>
		<td width="100%" colspan="4" align="right">
			<INPUT name=btnReferences type=button value=References style= "width: 2 cm" LANGUAGE=javascript onclick="return btnReferences_onclick()">
			<INPUT name="btnDelete" type="button" value="Delete" style= "width: 2 cm" onClick="return fct_onDelete();">
			<INPUT name="btnReset" type="button" value="Reset" style= "width: 2 cm" onClick= "return fct_onReset();">
			<INPUT name="btnNew" type="button" value="New" style= "width: 2 cm" onClick="return fct_NewEmailSetupEntry();">
			<INPUT id="btnSave" name="btnSave" type="submit" value="Save" style= "width: 2 cm">
		</td>
	</tr>
	</tfoot>

</TABLE>
	<FIELDSET>
	<LEGEND align=right><B>Audit Information</B></LEGEND>
	<Div SIZE=8pt ALIGN=RIGHT>
		Record Status Indicator
		<INPUT align=left name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 18px"disabled value=<%if strID <> 0 then Response.Write """"&rsEmailDetail("record_status_ind")&"""" else Response.Write """""" end if%> >&nbsp;&nbsp;&nbsp;
		Create Date
		<INPUT align=center name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 140px"disabled value=<%if strID <> 0 then Response.Write """"&rsEmailDetail("create_date")&"""" else Response.Write """""" end if%>>&nbsp;
		Created By
		<INPUT align=right name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 100px"disabled value=<%if strID <> 0 then Response.Write """"&rsEmailDetail("create_real_userid")&"""" else Response.Write """""" end if%> ><BR>
		Update Date
		<INPUT align=center name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 140px"disabled value=<%if strID <> 0 then Response.Write """"&rsEmailDetail("update_date")&"""" else Response.Write """""" end if%> >
		Updated By
		<INPUT align=right name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 100px"disabled value=<%if strID <> 0 then Response.Write """"&rsEmailDetail("update_real_userid")&"""" else Response.Write """""" end if%> >
	</DIV>
	</FIELDSET>
</FORM>

<%

	'objConn.close
	'set objConn = nothing

	if strID <> 0 then
		rsEmailDetail.close
		set rsEmailDetail = nothing
		objConn.close
		set objConn = nothing
	end if


%>


</BODY>
</HTML>

