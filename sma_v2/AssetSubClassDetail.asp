<%@ Language=VBScript     %>
<% Option Explicit        %>
<% on error resume next   %>
<% Response.Buffer = true %>
<!--#include file="SmaConstants.inc"-->
<!--#include file="databaseconnect.asp"-->
<!--#include file="smaProcs.inc"-->
<%
Const ASP_NAME = "AssetSubClassDetail.asp"
Const NO_ID = "null"
Const DESC = "Asset Subclass"     'used in javascript to set the title
Const TABLE = "ASSET_SUB_CLASS"   'used in Javascript for references button and must be uppercase
Const UPDATE_PROC = "sma_sp_userid.spk_sma_asset_inter.sp_asset_sub_class_update"
Const INSERT_PROC = "sma_sp_userid.spk_sma_asset_inter.sp_asset_sub_class_insert"
Const DELETE_PROC = "sma_sp_userid.spk_sma_asset_inter.sp_asset_sub_class_delete"


'check user's rights
Dim intAccessLevel
Dim strNew
Dim strRealUserID
Dim strWinMessage

intAccessLevel = CInt(CheckLogon(strConst_AssetCatalogue))
strRealUserID = Session("username")

if (intAccessLevel and intConst_Access_ReadOnly) <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to Asset. Please contact your system administrator"
end if

Dim StrAssetSubclassID, StrSql, objRS

StrAssetSubclassID  = Request("hdnAssetSubclassID")
strNew = Request("NewRecord")

if  strNew = "NEW" or StrAssetSubclassID = "" THEN
	StrAssetSubclassID = NO_ID
END IF

select case Request("hdnFrmAction")
	case "SAVE"
		if StrAssetSubclassID <> NO_ID then
			'update existing record
			if (intAccessLevel and intConst_Access_Update) <> intConst_Access_Update then
				DisplayError "REFRESH", strWinLocation, 0, "UPDATE DENIED", "You don't have access to update makes. Please contact your system administrator."
			end if

			dim cmdUpdateObj
			set cmdUpdateObj = server.CreateObject("ADODB.Command")
			set cmdUpdateObj.ActiveConnection = objConn
			cmdUpdateObj.CommandType = adCmdStoredProc
			cmdUpdateObj.CommandText = UPDATE_PROC

			'create params
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_userid", adVarChar, adParamInput, 20, strRealUserID)
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_asset_sub_class_id", adNumeric, adParamInput, , Clng(Request("hdnAssetSubclassID")))
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_asset_sub_class_desc", adVarChar, adParamInput, 50 , Request("txtDesc"))
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_asset_sub_class_id", adNumeric, adParamInput, , Clng(Request("hdnAssetClassID")))
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("last_update", adDBTimeStamp, adParamInput, , CDate(Request("hdnUpdateDateTime")))

			'call the update stored proc
			cmdUpdateObj.Execute
			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT UPDATE OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			end if

			strWinMessage = "Record saved successfully."

		else
			'create new record
			if (intAccessLevel and intConst_Access_Create) <> intConst_Access_Create then
				DisplayError "REFRESH", strWinLocation, 0, "INSERT DENIED", "You don't have access to create makes. Please contact your system administrator."
			end if

			dim cmdInsertObj
			set cmdInsertObj = server.CreateObject("ADODB.Command")
			set cmdInsertObj.ActiveConnection = objConn
			cmdInsertObj.CommandType = adCmdStoredProc
			cmdInsertObj.CommandText = INSERT_PROC

			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_userid", adVarChar, adParamInput, 20, strRealUserID)
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_asset_sub_class_id", adNumeric, adParamOutput, , null)
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_asset_sub_class_desc", adVarChar, adParamInput, 50 , Request("txtDesc"))
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_asset_class_id", adNumeric, adParamInput, , Request("hdnAssetClassID"))

			'call the proc
			cmdInsertObj.Execute
			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT CREATE NEW OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			else
				StrAssetSubclassID = cmdInsertObj.Parameters("p_asset_sub_class_id").Value
			end if

			strWinMessage = "Record created successfully."

		end if

	case "DELETE"
		if (intAccessLevel and intConst_Access_Delete) <> intConst_Access_Delete then
			DisplayError "REFRESH", strWinLocation, 0, "DELETE DENIED", "You don't have access to delete makes. Please contact your system administrator"
		end if

		dim cmdDeleteObj
		set cmdDeleteObj = server.CreateObject("ADODB.Command")
		set cmdDeleteObj.ActiveConnection = objConn
		cmdDeleteObj.CommandType = adCmdStoredProc
		cmdDeleteObj.CommandText = DELETE_PROC

		cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_id", adNumeric, adParamInput, , clng(StrAssetSubclassID))					'number(9)
		cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_last_update", adDBTimeStamp, adParamInput, ,Cdate(Request("hdnUpdateDateTime")))		'Date

		cmdDeleteObj.Execute
		if objConn.Errors.Count <> 0 then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT", objConn.Errors(0).Description
			objConn.Errors.Clear
		end if

		StrAssetSubclassID = NO_ID

		StrWinMessage = "Record deleted successfully."

end select

if StrAssetSubclassID <> NO_ID then

	StrSql = " SELECT sc.asset_sub_class_id" &_
			 " ,      sc.asset_sub_class_desc" &_
			 " ,      c.asset_class_id" &_
			 " ,      c.asset_class_desc" &_
			 " ,      t.asset_class_type_id" &_
			 " ,      t.asset_class_type_desc" &_
			 " ,      sc.create_date_time" &_
			 " ,      sc.create_db_userid" &_
			 " ,      sma_sp_userid.spk_sma_library.sf_get_full_username(sc.create_real_userid) create_real_userid" &_
			 " ,      sc.update_date_time last_update_date_time" &_
			 " ,      sc.update_db_userid" &_
			 " ,      sma_sp_userid.spk_sma_library.sf_get_full_username(sc.update_real_userid) update_real_userid" &_
			 " ,      sc.record_status_ind " &_
			 " FROM   crp.asset_sub_class  sc" &_
			 " ,      crp.asset_class      c" &_
			 " ,      crp.asset_class_type t" &_
			 " WHERE  t.asset_class_type_id = c.asset_class_type_id" &_
			 " AND    c.asset_class_id = sc.asset_class_id" &_
			 " AND    sc.asset_sub_class_id = " & StrAssetSubclassID


	'Create Recordset object
	set objRs = objConn.Execute(StrSql)

	if err then
		DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
	end if
end if

%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK REL="stylesheet" TYPE="text/css" HREF="stylesheets/styles.css">
<TITLE></TITLE>
<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
<script type="text/javascript" SRC="AccessLevels.js"></script>
<SCRIPT LANGUAGE=javascript>
<!--
//set the heading
setPageTitle("SMA - Make");

var intAccessLevel = <%=intAccessLevel%>;
var intID = <%=StrAssetSubclassID%>;
var boolNeedToSave = false;


function fct_onChange()
{
	boolNeedToSave = true;
}

function  btnReset_onClick()
{
	if(confirm('All changes will be lost. Are you sure you want to reset the page?')){
		boolNeedToSave = false;
		document.location.href = "<%=ASP_NAME%>?hdnAssetSubClassID=" + intID;
	}
}

function btnNew_onClick()
{

	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create)
	{
		alert('Access denied.  Please contact your system administrator.');
		return false;
	}


	self.document.location.href = "<%=ASP_NAME%>?NewRecord=NEW";

}

function btnReferences_onclick()
{

	if ('<%=StrAssetSubclassID%>' != '<%=NO_ID%>')
	{
		var strOwner = 'CRP' ;			// owner name must be in Uppercase
		var strTableName = '<%=TABLE%>' ;		// replace ADDRESS with your own table name and table name must be in Uppercase
		var strRecordID = document.frmDetail.hdnAssetSubclassID.value ;   // insert your record id
		var URL ;

		URL ='Dependency.asp?Owner=' + strOwner + '&TableName=' + strTableName + '&RecordID='+ strRecordID   ;
		window.open(URL, 'Popup', 'top=100, left=100, WIDTH=500, HEIGHT=300'  ) ;
	}
	else
	{
		alert('This is a new record, therefore, there are no references.');
	}
}

function btnDelete_onClick()
{
	if ((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete)
	{
		alert('Access denied. Please contact your system administrator.');
		return false;
	}

	var lngID = document.frmDetail.hdnAssetSubclassID.value ;
	var strUpdateDateTime = document.frmDetail.hdnUpdateDateTime.value ;

	if (lngID != "<%=NO_ID%>")
	{
		if (confirm("Do you really want to delete this object?"))
		{
			boolNeedToSave = false;
			document.location = "<%=ASP_NAME%>?hdnFrmAction=DELETE&hdnAssetSubclassID=" + lngID + "&hdnUpdateDateTime=" + strUpdateDateTime ;
		}
	}
}
function btnClassLookup_onClick()
{

	if (document.frmDetail.txtAssetClassDesc.value != "")
	{
		SetCookie("AssetClassDesc", document.frmDetail.txtAssetClassDesc.value);
	}

	SetCookie("WinName", 'Popup');
	window.open('SearchFrame.asp?fraSrc=AssetClass', 'Popup', 'top=50, left=100, height=600, width=800' ) ;
	fct_onChange();
}

function form_onSubmit()
{
	//no need to validate if the user cannot save the record
	if ( ((<%=intAccessLevel%> & <%=intconst_Access_Create%>) == <%=intconst_Access_Create%>) || ( (<%=intAccessLevel%> & <%=intconst_Access_Update%>) == <%=intconst_Access_Update%>) )
	{
		if (document.frmDetail.txtDesc.value == "" )
		{
			alert("Please type a make description.");
			document.frmDetail.txtDesc.focus();
			return(false);
		}

	}
	else
	{
		alert('Access denied.  Please contact your system administrator.');
		return (false);
	}
	document.frmDetail.hdnFrmAction.value = "SAVE"
	boolNeedToSave = false;
	document.forms[0].submit();
	return(true);
}

function fct_onBeforeUnload()
{

	//must set focus to save button because if user has changed only one field and has not
	//left it the on_change event will not have fired and the flag that determines whether
	//you need to save or not will be false
	document.frmDetail.btnSave.focus();

	if ((intAccessLevel & intConst_Access_Update) == intConst_Access_Update || (intAccessLevel & intConst_Access_Create) == intConst_Access_Create)
	{
		if (boolNeedToSave == true)
		{
			event.returnValue = "There is unsaved data on the screen. To save changes, click CANCEL below then click SAVE on the main form.";
		}
	}

}

function fct_displayStatus(msg)
{
		window.status=msg;
		setTimeout('fct_clearStatus()', <%=intConst_MessageDisplay%>);
}

function fct_clearStatus()
{
        window.status='';
}

function window_onload()
{
	fct_displayStatus('<%=routineJavaScriptString(strWinMessage)%>');
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload();" onbeforeunload="return fct_onBeforeUnload();">
<FORM NAME=frmDetail METHOD=POST ACTION="<%=ASP_NAME%>">
<!-- hidden fields -->
	<INPUT id=hdnUpdateDateTime  name=hdnUpdateDateTime  type=hidden value="<%if StrAssetSubclassID <> NO_ID then Response.Write objRs("last_update_date_time") else Response.Write null end if%>">
	<INPUT id=hdnAssetSubclassID name=hdnAssetSubclassID type=hidden value="<%if StrAssetSubclassID <> NO_ID then Response.Write StrAssetSubclassID else Response.Write null end if%>">
	<INPUT id=hdnAssetClassID    name=hdnAssetClassID    type=hidden value="<%if StrAssetSubclassID <> NO_ID then Response.Write objRs("ASSET_CLASS_ID") else Response.Write null end if%>">
	<INPUT id=hdnFrmAction       name=hdnFrmAction       type=hidden value="">
<!-- end hidden fields -->
<table width="100%" border=0>
	<thead>
		<TR><TD align=left colspan=2><%=DESC%> Detail</TD>
	</thead>
	<tbody>
		<TR>
			<TD ALIGN=RIGHT NOWRAP>Asset Class Type</TD>
			<TD>
				<INPUT name=txtAssetClassTypeDesc DISABLED style="HEIGHT: 23px; WIDTH: 300px" value="<%if StrAssetSubclassID <> NO_ID then  Response.Write routineHtmlString(objRS("ASSET_CLASS_TYPE_DESC")) else Response.Write null end if%>" onchange ="fct_onChange();">
			</TD>

		</TR>
		<TR>
			<TD ALIGN=RIGHT NOWRAP>Asset Class<font color="red">*</font></TD>
			<TD>
				<INPUT name=txtAssetClassDesc DISABLED style="HEIGHT: 23px; WIDTH: 300px" value="<%if StrAssetSubclassID <> NO_ID then  Response.Write routineHtmlString(objRS("ASSET_CLASS_DESC")) else Response.Write null end if%>" onchange ="fct_onChange();">
				<INPUT id=btnContactLookup  name=btnContactLookup  style="HEIGHT: 21px; WIDTH: 19px"  type=button value="..." onclick="return  btnClassLookup_onClick()">
			</TD>

		</TR>
		<TR>
			<TD ALIGN=RIGHT NOWRAP>Asset Subclass<font color="red">*</font></TD>
			<TD>
				<INPUT name=txtDesc style="HEIGHT: 23px; WIDTH: 300px" value="<%if StrAssetSubclassID <> NO_ID then  Response.Write routineHtmlString(objRS("ASSET_SUB_CLASS_DESC")) else Response.Write null end if%>" onchange ="fct_onChange();">
			</TD>

		</TR>
	</tbody>
	<tfoot>
		<TR>
			<TD align=right colspan=3>
				<input name=btnReferences type=button value=References style= "width: 2.2cm" LANGUAGE=javascript onclick="return btnReferences_onclick()">&nbsp;&nbsp;
				<INPUT name=btnDelete type=button value=Delete style="width: 2cm" onclick="return btnDelete_onClick();">&nbsp;&nbsp;
				<INPUT name=btnReset type=button value=Reset style="width: 2cm" onclick="return btnReset_onClick();">&nbsp;&nbsp;
				<INPUT name=btnNew type=button value="New" style="width: 2cm" onclick="return btnNew_onClick();">&nbsp;&nbsp;
				<INPUT name=btnSave type=button value=Save style="width: 2cm" onclick="form_onSubmit();">&nbsp;&nbsp;
			</TD>
		</TR>
	</tfoot>
</table>
<FIELDSET>
	<LEGEND align=RIGHT><B>Audit Information</B></LEGEND>
	<div size=8pt align=right>
		Record Status Indicator:
		<INPUT align=left name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 18px"disabled value=<%if StrAssetSubclassID <> NO_ID then  Response.Write """"&objRS("RECORD_STATUS_IND")&"""" else Response.Write """""" end if%> >
		Create Date:
		<INPUT align =center name=txtcrdate type=text style="HEIGHT: 20px; WIDTH: 100px"disabled value=<%if StrAssetSubclassID <> NO_ID then  Response.Write """"&objRS("CREATE_DATE_TIME")&"""" else Response.Write """""" end if%> >
		&nbsp;Created By:
		<INPUT align =right name=txtcrby type=text style="HEIGHT: 20px; WIDTH: 100px"disabled value=<%if StrAssetSubclassID <> NO_ID then  Response.Write """"&objRS("CREATE_REAL_USERID")&"""" else Response.Write """""" end if%> ><BR>
		Update Date:
		<INPUT align=center name=txtupdate type=text style="HEIGHT: 20px; WIDTH: 101px"disabled value=<%if StrAssetSubclassID <> NO_ID then  Response.Write """"&objRS("LAST_UPDATE_DATE_TIME")&"""" else Response.Write """""" end if%> >
		Updated By:
		<INPUT align=right name=txtupby type=text style="HEIGHT: 20px; WIDTH: 100px"disabled value=<%if StrAssetSubclassID <> NO_ID then  Response.Write """"&objRS("UPDATE_REAL_USERID")&"""" else Response.Write """""" end if%> >
	</DIV>
</FIELDSET>

</FORM>
<%

 'Clean up our ADO objects
if StrAssetSubclassID <> NO_ID then
    objRS.close
    set objRS = Nothing
end if

%>

</BODY>
</HTML>
