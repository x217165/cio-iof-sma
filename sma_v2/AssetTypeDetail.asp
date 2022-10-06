<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<!--#include file = "smaConstants.inc"-->
<!--#include file = "smaProcs.inc"-->
<!--#include file = "databaseconnect.asp" -->
<!--
*************************************************************************************************
* Page name:	AssetTypeDetail.asp																*
* Purpose:		To display the Asset Type														*
*																								*
* Created by:	Nancy Mooney Oct.17,2000														*
*																								*
*************************************************************************************************
-->
<%
Dim intAccessLevel
Dim strAssetTypeID, txtAssetTypeDesc, datUpdateDateTime, strWinMessage, strWinLocation
Dim lRow, arrAssetClassTypeList, arrAssetClassList, arrAssetSubclassList
Dim	objRS, objRSSelect, objCommand, strSQL, strErrMessage

	'Check user's rights
	intAccessLevel = CInt(CheckLogon(strConst_AssetTypeClassification))
	If (intAccessLevel And intConst_Access_ReadOnly) <> intConst_Access_ReadOnly  Then
		DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to Asset Type. Please contact your system administrator"
	End If

	strWinMessage = ""
	strAssetTypeID = Request.QueryString("AssetTypeID")
	datUpdateDateTime = Request.Form("UpdateDateTime")

	Select Case UCase(Request("hdnFrmAction"))
		Case "SAVE"
			Set objCommand = Server.CreateObject("ADODB.Command")
			Set objCommand.ActiveConnection = objConn
			objCommand.CommandType = adCmdStoredProc

			If IsNumeric(Request("hdnAssetTypeID")) Then	'Save existing Asset Type
				If (intAccessLevel And intConst_Access_Update) <> intConst_Access_Update Then
					DisplayError "BACK", strWinLocation, 0, "UPDATE DENIED", "You don't have access to update asset types. Please contact your system administrator"
				End If
				'all parameters are required fields
				objCommand.CommandText = "sma_sp_userid.spk_sma_asset_inter.sp_asset_type_update"
				objCommand.Parameters.Append objCommand.CreateParameter("p_userid", adVarChar, adParamInput, 20, Session("username"))
				objCommand.Parameters.Append objCommand.CreateParameter("p_asset_type_id", adNumeric, adParamInput, , CLng(Request("hdnAssetTypeID")))
				objCommand.Parameters.Append objCommand.CreateParameter("p_asset_type_desc", adVarChar, adParamInput, 50, Trim(Request("txtAssetTypeDesc")))
				objCommand.Parameters.Append objCommand.CreateParameter("p_asset_sub_class_id", adNumeric, adParamInput, , CLng(Request("hdnAssetSubClassID")))
				objCommand.Parameters.Append objCommand.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , CDate(Request("hdnUpdateDateTime")))

				strErrMessage = "CANNOT UPDATE OBJECT"

			Else										'Create a new Asset Type
				If (intAccessLevel And intConst_Access_Create) <> intConst_Access_Create Then
					DisplayError "BACK", strWinLocation, 0, "INSERT DENIED", "You don't have access to create asset types. Please contact your system administrator"
				End If
				'all parameters are required fields
				objCommand.CommandText = "sma_sp_userid.spk_sma_asset_inter.sp_asset_type_insert"
				objCommand.Parameters.Append objCommand.CreateParameter("p_userid", adVarChar, adParamInput, 20, Session("username"))
				objCommand.Parameters.Append objCommand.CreateParameter("p_asset_type_id", adNumeric, adParamOutput, , Null)
				objCommand.Parameters.Append objCommand.CreateParameter("p_asset_type_desc", adVarChar, adParamInput, 50, Trim(Request("txtAssetTypeDesc")))
				objCommand.Parameters.Append objCommand.CreateParameter("p_asset_sub_class_id", adNumeric, adParamInput, , CLng(Request("hdnAssetSubClassID")))

				strErrMessage = "CANNOT CREATE OBJECT"

				'parameter check - debugging
				'Response.Write "<b> count = " & objCommand.Parameters.count & "<br>"
				'dim nx
				'for nx = 0 to objCommand.Parameters.Count-1
				'	Response.Write objCommand.Parameters.Item(nx).Name & " = " & objCommand.Parameters.Item(nx).Value & " <br>"
				'next
				'Response.end

			End If

			On Error Resume Next
			objCommand.Execute
			If objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, strErrMessage, objConn.Errors(0).Description
				objConn.Errors.Clear
			End If
			strAssetTypeID = CStr(objCommand.Parameters("p_asset_type_id").Value)
			strWinMessage = "Record saved successfully."

		Case "DELETE"
			If (intAccessLevel And intConst_Access_Delete) <> intConst_Access_Delete Then
				DisplayError "BACK", strWinLocation, 0, "DELETE DENIED", "You don't have access to delete asset types. Please contact your system administrator"
			End If

			Set objCommand = Server.CreateObject("ADODB.Command")
			Set objCommand.ActiveConnection = objConn
			objCommand.CommandType = adCmdStoredProc
			objCommand.CommandText = "sma_sp_userid.spk_sma_asset_inter.sp_asset_type_delete"
			objCommand.Parameters.Append objCommand.CreateParameter("p_asset_type_id", adNumeric, adParamInput, , CLng(strAssetTypeID))
			objCommand.Parameters.Append objCommand.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, ,CDate(Request("hdnUpdateDateTime")))

  			On Error Resume Next
			objCommand.Execute
			If objConn.Errors.Count <> 0 Then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			End If
			strAssetTypeID = "DEL"
			strWinMessage = "Record deleted successfully."
	End Select

	If IsNumeric(strAssetTypeID) Then
		strSQL = "SELECT AT.ASSET_TYPE_ID, " &_
			"AT.ASSET_TYPE_DESC, " &_
			"ASB.ASSET_SUB_CLASS_ID, " &_
			"ASB.ASSET_SUB_CLASS_DESC, " &_
			"AC.ASSET_CLASS_ID, " &_
			"AC.ASSET_CLASS_DESC, " &_
			"ACT.ASSET_CLASS_TYPE_ID, " &_
			"ACT.ASSET_CLASS_TYPE_DESC, " &_
			"TO_CHAR(AT.CREATE_DATE_TIME, 'MON-DD-YYYY HH24:MI:SS') AS CREATE_DATE_TIME, " &_
			"sma_sp_userid.spk_sma_library.sf_get_full_username(AT.CREATE_REAL_USERID) AS CREATE_REAL_USERID, " &_
			"TO_CHAR(AT.UPDATE_DATE_TIME, 'MON-DD-YYYY HH24:MI:SS') AS UPDATE_DATE_TIME, " &_
			"sma_sp_userid.spk_sma_library.sf_get_full_username(AT.UPDATE_REAL_USERID) AS UPDATE_REAL_USERID, " &_
			"AT.RECORD_STATUS_IND, " &_
			"AT.UPDATE_DATE_TIME AS LAST_UPDATE_DATE_TIME " &_
			"FROM " &_
			"CRP.ASSET_TYPE AT, " &_
			"CRP.ASSET_SUB_CLASS ASB, " &_
			"CRP.ASSET_CLASS AC, " &_
			"CRP.ASSET_CLASS_TYPE ACT " &_
			"WHERE AT.ASSET_SUB_CLASS_ID = ASB.ASSET_SUB_CLASS_ID " &_
			"AND ASB.ASSET_CLASS_ID = AC.ASSET_CLASS_ID " &_
			"AND AC.ASSET_CLASS_TYPE_ID = ACT.ASSET_CLASS_TYPE_ID " & _
			"AND AT.ASSET_TYPE_ID =	" & strAssetTypeID

		'Response.Write strSQL
		'Response.End

		'Create Recordset object
		Set objRS = Server.CreateObject("ADODB.Recordset")
		On Error Resume Next
		objRS.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
		If objConn.Errors.Count <> 0 Then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, "ERROR LOADING DATA (Asset Type)", objConn.Errors(0).Description
			objConn.Errors.Clear
		End If
	End If

%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link REL="stylesheet" TYPE="text/css" HREF="stylesheets/styles.css">
<SCRIPT type="text/javascript" language="javascript" src="AccessLevels.js"></SCRIPT>
<SCRIPT type="text/javascript" language="javascript" src="GeneralJavaFunctions.js"></SCRIPT>
<SCRIPT type="text/javascript" language="javascript">
<!-- //Hide Client-Side SCRIPT
var strWinMessage = "<%=strWinMessage%>";
var intAccessLevel = <%=intAccessLevel%>;
var bolSaveRequired = false;

var arrAssetSubClassList = new Array();

setPageTitle("SMA - Asset Type");

/*function fct_selNavigate(){

var strPageName = document.frmAssetTypeDetail.selNavigate.item(document.frmAssetTypeDetail.selNavigate.selectedIndex).value ;

	switch (strPageName) {
		case "AssetClass":
			document.frmAssetTypeDetail.selNavigate.selectedIndex = 0;
			var strAssetClassID = document.frmAssetTypeDetail.hdnAssetClassID.value;
			self.location.href = "AssetClassDetail.asp?AssetClassID=" + strAssetClassID;
			break ;

		case "AssetSubClass":
			document.frmAssetTypeDetail.selNavigate.selectedIndex = 0;
			var strAssetSubClassID = document.frmAssetTypeDetail.hdnAssetSubClassID.value;
			self.location.href = "AssetSubClass.asp?AssetSubClassID=" + strAssetSubClassID;
			break ;

		case "DEFAULT":
			// do nothing ;
	}
}*/

function btnDelete_onClick() {

	var strAssetTypeID = document.frmAssetTypeDetail.hdnAssetTypeID.value;
	var strUpdateDateTime = document.frmAssetTypeDetail.hdnUpdateDateTime.value;

	if ((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) {
		alert('You do not have permission to DELETE an Asset Type.  Please contact your System Administrator.');
		return false;
	}

	if (strAssetTypeID == "") {
		alert('This Asset Type does not exist in the database.');
		return false;
	}

	if (confirm('Do you really want to delete this asset type?')){
		self.document.location.href = "AssetTypeDetail.asp?AssetTypeID=" + strAssetTypeID + "&hdnFrmAction=DELETE" + "&hdnUpdateDateTime=" + strUpdateDateTime;
	}
}

function btnNew_onClick() {
	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
		alert('You do not have permission to CREATE an Asset Type.  Please contact your System Administrator.');
		return false;
	}
	document.location = "AssetTypeDetail.asp?AssetTypeID=NEW";
}

function fct_onChange() {
	bolSaveRequired = true;
}

function form_onSubmit() {
	if ((intAccessLevel & intConst_Access_Update) != intConst_Access_Update) {
		alert('You do not have permission to UPDATE an Asset Type.  Please contact your System Administrator.');
		return false;
	}
	//validate required fields
	if (document.frmAssetTypeDetail.hdnAssetSubClassID.value == "") {
		alert('MISSING REQUIRED FIELD. Please select an Asset Subclass from the list box.');
		document.frmAssetTypeDetail.btnAssetSubClassLookup.focus();
		return false;
	}
	if (document.frmAssetTypeDetail.txtAssetTypeDesc.value == "") {
		alert('MISSING REQUIRED FIELD. Please enter an Asset Type.');
		document.frmAssetTypeDetail.txtAssetTypeDesc.focus();
		return false;
	}

	document.frmAssetTypeDetail.hdnFrmAction.value = "SAVE";
	bolSaveRequired = false;
	document.frmAssetTypeDetail.submit();
	return true;
}

function btnReferences_onClick() {
var strOwner = 'CRP';
var strTableName = 'ASSET_TYPE';
var strRecordID = document.frmAssetTypeDetail.hdnAssetTypeID.value ;
var strURL;

	if (strRecordID == "") {
		alert("No references. This is a new record.");
		return false;
	}

	strURL = "Dependency.asp?Owner=" + strOwner + "&TableName=" + strTableName + "&RecordID=" + strRecordID;
	window.open(strURL, 'Popup', 'top=100, left=100, width=500, height=300');
}

function window_onBeforeUnload() {
	//Ensure that fct_onChange() fires for any changed data.
	document.frmAssetTypeDetail.btnSave.focus();

	if (bolSaveRequired) {
		event.returnValue = "There is unsaved data on the screen. To save changes, click CANCEL below then click SAVE on the main FORM.";
	}
}

function ClearStatus() {
	window.status = "";
}

function DisplayStatus(strWinStatus) {
	window.status = strWinStatus;
	setTimeout('ClearStatus()', <%=intConst_MessageDisplay%>);
}

function btnReset_onClick() {
	if(confirm('All changes will be lost. Do you really want to reset the page?')){
		bolSaveRequired = false;
		document.location.href = "AssetTypeDetail.asp?AssetTypeID=<%=strAssetTypeID%>";
	}
}

function btnAssetSubClassLookup_onClick()
{
		if (document.frmAssetTypeDetail.txtAssetSubClassDesc.value != "" ) {
			SetCookie("AssetSubClassDesc",document.frmAssetTypeDetail.txtAssetSubClassDesc.value) ;
		}
		SetCookie("WinName", 'Popup');
		bolNeedToSave = true;
		window.open('SearchFrame.asp?fraSrc=AssetSubclass', 'Popup', 'top=50, left=100, WIDTH=800, HEIGHT=600' ) ;
}
// end hide Client-Side SCRIPT -->
</SCRIPT>
</HEAD>
<BODY language="javascript" onLoad="DisplayStatus(strWinMessage);" onBeforeUnload="window_onBeforeUnload();" >
<FORM name="frmAssetTypeDetail" action="AssetTypeDetail.asp" method="post" >
	<INPUT type="hidden" id="hdnAssetClassTypeID" name="hdnAssetClassTypeID" value="<%If IsNumeric(strAssetTypeID) Then Response.Write objRS.Fields("ASSET_CLASS_TYPE_ID").Value%>">
	<INPUT type="hidden" id="hdnAssetClassID" name="hdnAssetClassID" value="<%If IsNumeric(strAssetTypeID) Then Response.Write objRS.Fields("ASSET_CLASS_ID").Value%>">
	<INPUT type="hidden" id="hdnAssetSubClassID" name="hdnAssetSubClassID" value="<%If IsNumeric(strAssetTypeID) Then Response.Write objRS.Fields("ASSET_SUB_CLASS_ID").Value%>">
	<INPUT type="hidden" id="hdnAssetTypeID" name="hdnAssetTypeID" value="<%If IsNumeric(strAssetTypeID) Then Response.Write objRS.Fields("ASSET_TYPE_ID").Value%>">
	<INPUT type="hidden" id="hdnUpdateDateTime" name="hdnUpdateDateTime" value="<%If IsNumeric(strAssetTypeID) Then Response.Write objRS.Fields("LAST_UPDATE_DATE_TIME").Value%>">
	<INPUT type="hidden" id="hdnFrmAction" name="hdnFrmAction" value="">

<TABLE border="0" cols="4" width="100%">
<THEAD>
<TR>
	<TD align="left" colspan="4">Asset Type Detail</TD>
	<!--
	<TD align="right">
		<SELECT valign="top" id="selNavigate" name="selNavigate" onChange="fct_selNavigate();">
			<OPTION value="DEFAULT" selected>Quickly Goto ...</OPTION>
			<OPTION value="AssetClass">Asset Class</OPTION>
			<OPTION value="AssetSubClass">Asset Subclass</OPTION>
		</SELECT>
	</TD>
	-->
</TR>
</THEAD>
<TBODY>
<TR>
	<TD align="right" nowrap>Asset Class Type<FONT color="red">*</FONT></TD>
	<TD align="left" nowrap><INPUT disabled name="txtAssetClassTypeDesc" value="<%If IsNumeric(strAssetTypeID) Then Response.Write objRS.Fields("ASSET_CLASS_TYPE_DESC").Value%>" ></TD>
</TR>
<TR>
	<TD align="right" nowrap>Asset Class<FONT color="red">*</FONT></TD>
	<TD align="left" nowrap><INPUT disabled name="txtAssetClassDesc" style="width: 365px" value="<%If IsNumeric(strAssetTypeID) Then Response.Write objRS.Fields("ASSET_CLASS_DESC").Value%>" ></TD>
</TR>
<TR>
	<TD align="right" nowrap>Asset Subclass<FONT color="red">*</FONT></TD>
	<td align=left >
		<INPUT disabled name="txtAssetSubClassDesc" style="width: 365px" onChange="fct_AssetSubClass_OnChange();" value="<%If IsNumeric(strAssetTypeID) Then Response.Write objRS.Fields("ASSET_SUB_CLASS_DESC").Value%>" >
		<INPUT name=btnAssetSubClassLookup type=button value="..." language=javascript onClick="btnAssetSubClassLookup_onClick();">
	</td>
</TR>
<TR>
	<TD align="right" nowrap>Asset Type<FONT color="red">*</FONT></TD>
	<TD align="left" nowrap><INPUT name=txtAssetTypeDesc type="text" maxlength=50 size=50 onChange="return fct_onChange();" value="<%If IsNumeric(strAssetTypeID) Then Response.Write objRS.Fields("ASSET_TYPE_DESC").Value%>"></TD>
</TR>
</TBODY>
<TFOOT>
<TR>
	<TD colspan="4" align="right">
	<INPUT id="btnReferences" name="btnReferences" type="button" value="References" style="width: 2.2cm" onClick="return btnReferences_onClick();">&nbsp;
	<INPUT id="btnDelete"     name="btnDelete"     type="button" value="Delete"     style="width: 2cm"   onClick="return btnDelete_onClick();">&nbsp;
	<INPUT id="btnReset"      name="btnReset"      type="button" value="Reset"      style="width: 2cm"   onClick="return btnReset_onClick();">&nbsp;
	<INPUT id="btnNew"        name="btnNew"        type="button" value="New"        style="width: 2cm"   onClick="return btnNew_onClick();">&nbsp;
	<INPUT id="btnSave"       name="btnSave"       type="button" value="Save"       style="width: 2cm"   onClick="return form_onSubmit();">&nbsp;</TD>
</TR>
</TFOOT>
</TABLE>
<FIELDSET width="100%">
	<LEGEND align="right"><b>Audit Information</b></LEGEND>
	<DIV size="8pt" align="right">
	Record Status Indicator&nbsp;<INPUT align="left"   name="txtRecordStatusInd" type="text" style="width: 18px"  disabled value="<%If IsNumeric(strAssetTypeID) Then Response.Write objRS.Fields("RECORD_STATUS_IND").Value%>">&nbsp;&nbsp;&nbsp;
	Create Date&nbsp;            <INPUT align="center" name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If IsNumeric(strAssetTypeID) Then Response.Write objRS.Fields("CREATE_DATE_TIME").Value%>">&nbsp;&nbsp;
	Created By&nbsp;             <INPUT align="right"  name="txtRecordStatusInd" type="text" style="width: 200px" disabled value="<%If IsNumeric(strAssetTypeID) Then Response.Write objRS.Fields("CREATE_REAL_USERID").Value%>">&nbsp;&nbsp;<br>
	Update Date&nbsp;            <INPUT align="center" name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If IsNumeric(strAssetTypeID) Then Response.Write objRS.Fields("UPDATE_DATE_TIME").Value%>">&nbsp;
	Updated By&nbsp;             <INPUT align="right"  name="txtRecordStatusInd" type="text" style="width: 200px" disabled value="<%If IsNumeric(strAssetTypeID) Then Response.Write objRS.Fields("UPDATE_REAL_USERID").Value%>">&nbsp;&nbsp;
	</DIV>
</FIELDSET>
</FORM>
<%
	'Clean up our ADO objects
	Set objRS = Nothing
	objConn.Close
	Set ObjConn = Nothing
%>
</BODY>
</HTML>
