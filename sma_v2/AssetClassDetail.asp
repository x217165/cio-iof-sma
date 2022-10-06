<%@ Language=VBScript %>
<% option explicit %>
<% Response.Buffer = true %>
<%on error resume next%>

<!--
********************************************************************************************
* Page name:	AssetClassDetail.asp

* Purpose:		To display the detailed information about an Asset Class.
*				Entry is chosen via AssetClassList.asp
*
* Created by:	Shawn Myers	10/17/2000
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
intAccessLevel = CInt(CheckLogon(strConst_AssetTypeClassification))


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
strID = Request("hdnAssetClassID")

'Response.Write "the value of the passed id is " & strID
'response.end

'get the hidden window location
strWinLocation = "AssetClassDetail.asp?AssetClassID="& Request("strID")

'set the variable for the UserInfo cookie
strRealUserID = Session("username")




'************************
'do save, insert, delete*
'************************


dim aClassType		'used to get the ID from the "AssetClassType" drop down list

select case Request("hdnFrmAction")

	case "SAVE"

	'check to see if entry exists already in database by checking for the existence of id


	  if Request("hdnAssetClassID")  <> "" then  ' it is an existing record so save the changes


		if intAccessLevel and intConst_Access_Update <> intConst_Access_Update then

			DisplayError "REFRESH", strWinLocation, 0, "UPDATE DENIED", "You don't have access to update managed objects. Please contact your system administrator"

		end if

		dim cmdUpdateObj

		set cmdUpdateObj = server.CreateObject("ADODB.Command")
		set cmdUpdateObj.ActiveConnection = objConn
		cmdUpdateObj.CommandType = adCmdStoredProc

		'get value from dropdown
		aClassType = split(Request("selAssetClassType"),"¿")


		'get the asset class detail stored update procedure <schema.package.procedure>
		cmdUpdateObj.CommandText = "sma_sp_userid.spk_sma_asset_inter.sp_asset_class_update"

		'create the associated parameters
		'user id associated with time stamp


		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_userid",adVarChar,adParamInput, 20,strRealUserID)
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_asset_class_id", adNumeric, adParamInput, 9, Clng(Request("hdnAssetClassID")))
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_asset_class_desc", adVarChar, adParamInput, 50, Request("txtAssetClassDescription"))
        cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_asset_class_type_id", adVarChar, adParamInput, 9, aClassType(0))
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_last_update_dt", adDBTimeStamp,adParamInput, , CDate(Request("hdnUpdateDateTime")))


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


		'get value from dropdown
		aClassType = split(Request("selAssetClassType"),"¿")



		'get the asset_class_detail insert procedure <schema.package.procedure>
		cmdInsertObj.CommandText = "sma_sp_userid.spk_sma_asset_inter.sp_asset_class_insert"



		'create the insert parameters

		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_userid", adVarChar, adParamInput, 20,strRealUserID)
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_asset_class_id", adNumeric, adParamOutput,, null)
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_asset_class_desc",adVarChar, adParamInput, 50, Request("txtAssetClassDescription"))
        cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_asset_class_type_id", adVarChar, adParamInput, 9, aClassType(0))

		' execute the insert object

		cmdInsertObj.Execute

			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT CREATE NEW OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			else
				strID = cmdInsertObj.Parameters("p_asset_class_id").Value
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

			'get the asset class detail delete procedure <schema.package.procedure>
			cmdDeleteObj.CommandText = "sma_sp_userid.spk_sma_asset_inter.sp_asset_class_delete"

			'create the delete parameters
			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_asset_class_id", adNumeric, adParamInput, , CLng(strID))
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












'ok, now go get the detailed  Asset Class information


'declare the connection and sql variables
Dim strSQL, strSelectClause, strFromClause, strWhereClause
Dim rsAssetClassDetail

dim objCmd

'connect to the database
'<<CONNECT>>

'use the sqlstring to extract the necessary information from the database

	if strID <> 0 then

		strSelectClause = "SELECT " &_
					"t1.asset_class_id, " & _
					"t1.asset_class_desc, " & _
					"t2.asset_class_type_id, " & _
					"t2.asset_class_type_desc, " &_
					"to_char(t1.create_date_time,'MON-DD-YYYY HH24:MI:SS') as create_date, " & _
					"sma_sp_userid.spk_sma_library.sf_get_full_username(t1.create_real_userid) as create_real_userid, " & _
					"to_char(t1.update_date_time,'MON-DD-YYYY HH24:MI:SS') as update_date, " & _
					"sma_sp_userid.spk_sma_library.sf_get_full_username(t1.update_real_userid) as update_real_userid, " & _
					"t1.update_date_time as last_update_date_time, " & _
					"t1.record_status_ind "


		strFromClause =	" FROM crp.asset_class  t1, " &_
						" crp.asset_class_type t2"


		strWhereClause = " WHERE " & _
					"t1.asset_class_id = " & strID &_
					" AND t1.asset_class_type_id = t2.asset_class_type_id"


		strSQL =  strSelectClause & strFromClause & strWhereClause

		'show SQL for debugging if necessary by using>>
		'Response.Write "<BR>" & strsql	 & "<br>"
		'Response.end

		'set and open the asset class detail recordset and database connection

		set rsAssetClassDetail = Server.CreateObject("ADODB.Recordset")

		rsAssetClassDetail.CursorLocation = adUseClient
		rsAssetClassDetail.Open strSQL, objConn

		if err then
			DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET A", err.Description
		end if
		if rsAssetClassDetail.EOF then
			DisplayError "BACK", "", 999, "CANNOT FIND REQUESTED OBJECT", "EOF condition occurred in rsAssetClassDetail recordset."
		end if


	end if



'Load the Asset Class Type Dropdown with its own sql

dim rsACType, strTypeSQL

strTypeSQL = " SELECT asset_class_type_id" &_
		     " ,      asset_class_type_desc" &_
		     " FROM   crp.asset_class_type" &_
			 " ORDER  BY asset_class_type_desc"

'Response.Write strTypeSQL
'Response.end

'Create the Command object

set objCmd = Server.CreateObject("ADODB.command")
    objCmd.ActiveConnection = objconn
    objCmd.CommandText = strTypeSQL
    objCmd.CommandType = adCmdText

'Create the StatusCode Recordset object

set rsACType = Server.CreateObject("ADODB.Recordset")

		rsACType.CursorLocation = adUseClient
		rsACType.Open strTypeSQL, objconn

		if err then
			DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET B", err.Description
		end if
		if rsACType.EOF then
			DisplayError "BACK", "", 999, "CANNOT FIND REQUESTED OBJECT", "EOF condition occurred in rsACType recordset."
		end if





%>



<HTML>
<HEAD>
	<LINK rel="stylesheet" type="text/css" href="stylesheets/styles.css">
	<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
	<script type="text/javascript" SRC="AccessLevels.js"></script>

	<SCRIPT LANGUAGE=JavaScript>
	<!--

	var strWinMessage = '<%=strWinMessage%>';
    var intAccessLevel = '<%=intAccessLevel%>';
    var bolNeedToSave = false ;

setPageTitle("SMA - Asset Class Detail");

	function fct_NewAssetClassEntry(){

	//alert ('in the new function');

	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create)

			{

			alert('Access denied. Please contact your system administrator.');
			return (false);

			}


			self.document.location.href = "AssetClassDetail.asp?hdnAssetClassID=0";



		}


	//OK

	function fct_OnSave(){

	//alert('in the save function' + <%=intAccessLevel%>);

	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create)

			{
				alert('Access denied. Please contact your system administrator.');
				return (false);
			}

	else

			{

				if (document.frmAssetClassDetail.txtAssetClassDescription.value == "" )
					{
						alert('Please enter a description.');
						return(false);
					}




				else

					{
						document.frmAssetClassDetail.hdnFrmAction.value = "SAVE";
						bolNeedToSave = false;
						document.frmAssetClassDetail.submit();
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

		var strID = document.frmAssetClassDetail.hdnAssetClassID.value;
		var strUpdateDate = document.frmAssetClassDetail.hdnUpdateDateTime.value;

	    //alert ('the value of the assetclassid is '+ document.frmAssetClassDetail.hdnAssetClassID.value);
	    //alert ('the value of the update date time is '+ document.frmAssetClassDetail.hdnUpdateDateTime.value);

				{

					if (confirm('Do you really want to delete this object?'))

						{
						document.location = "AssetClassDetail.asp?hdnFrmAction=DELETE&hdnAssetClassID="+strID+"&hdnUpdateDateTime="+strUpdateDate ;
						}

				}


		}

	//OK

	function fct_onReset() {
		if(confirm('All changes will be lost. Do you really want to reset the page?')){
			bolNeedToSave = false;
			document.location = 'AssetClassDetail.asp?hdnAssetClassID=' + "<%=strID%>" ;
		}
	}

	//OK


	function fct_onBeforeUnload()

	{


		document.frmAssetClassDetail.btnSave.focus();

		if ((intAccessLevel & intConst_Access_Update) == intConst_Access_Update || (intAccessLevel & intConst_Access_Create) == intConst_Access_Create)
		{
			if (bolNeedToSave == true)
			{
				event.returnValue = "There is unsaved data on the screen. To save changes, click CANCEL below then click SAVE on the main form.";
			}
		}

	}


	//OK

	function fct_clearStatus() {
		window.status = "";
	}


	//OK

	function fct_DisplayStatus(strWindowStatus){

	window.status=strWindowStatus;
	setTimeout('fct_clearStatus()', '<%=intConst_MessageDisplay%>');

    }




	function btnReferences_onclick() {

	var strOwner = 'CRP';
	var strTableName = 'ASSET_CLASS';
	var strRecordID = document.frmAssetClassDetail.hdnAssetClassID.value;
	var URL ;


	if ( strID = 0)

			{
				alert("No references. This is a new record.");
			}

	else

			{
			URL ='Dependency.asp?Owner=' + strOwner + '&TableName=' + strTableName + '&RecordID='+ strRecordID;
			window.open(URL, 'Popup', 'top=100, left=100, WIDTH=500, HEIGHT=300');
			}

	}

	function fct_onChangeAssetType() {

	var strWhole;
	var strAssetTypeDesc, intStart, intIndex;

	intIndex = document.frmAssetClassDetail.selAssetClassType.selectedIndex;
	strWhole = document.frmAssetClassDetail.selAssetClassType.options[intIndex].value;
	intStart = strWhole.indexOf('<%=strDelimiter%>');

	fct_onChange();
	}



	//OK

    function fct_onChange(){

		bolNeedToSave = true;
	}







	//-->
	</SCRIPT>

</HEAD>





<BODY onLoad="fct_DisplayStatus(strWinMessage);" onbeforeunload="fct_onBeforeUnload();">
<FORM name=frmAssetClassDetail action="AssetClassDetail.asp"  method="POST">



	<INPUT name=hdnAssetClassID type=hidden value="<%if strID <> 0 then  Response.Write rsAssetClassDetail("ASSET_CLASS_ID") else Response.Write """""" end if%>">
	<INPUT name=hdnUpdateDateTime type=hidden value="<%if strID <> 0 then  Response.Write rsAssetClassDetail("LAST_UPDATE_DATE_TIME") else Response.Write """""" end if%>">
	<INPUT id=hdnFrmAction name=hdnFrmAction type=hidden value= "">


	<!-- user interface -->

	<TABLE border=0 width=100%>

	<thead>
		<tr>
			<td colspan=4 align=left>Asset Class Detail</td>
		</tr>
	</thead>

	<TR>
	<TD ALIGN=RIGHT width=20% NOWRAP>Asset Class Type<font color=red>*</font></TD>
		<TD colspan="2">
			<SELECT id=selAssetClassType name=selAssetClassType style="HEIGHT: 20px; WIDTH: 120px" onChange="return fct_onChangeAssetType();">
			<%

			Do while Not rsACType.EOF
				Response.write "<OPTION "
				if strID <> 0 then
					if rsACType("ASSET_CLASS_TYPE_DESC") = rsAssetClassDetail("ASSET_CLASS_TYPE_DESC") then
						Response.Write " selected "
					end if
				end if
				Response.Write " VALUE =""" & routineHTMLString(rsACType("ASSET_CLASS_TYPE_ID")& strDelimiter & rsACType("ASSET_CLASS_TYPE_DESC")) & """>" & routineHTMLString(rsACType("ASSET_CLASS_TYPE_DESC")) & "</OPTION>"
				rsACType.MoveNext
			Loop
			%>
			</SELECT>
		</TD>

    </TR>


	<TR>
		<TD align=right width=25%>Asset Class Description<font color=red>*</font></TD>
		<TD align=left width=50% colspan=2>
			<input name=txtAssetClassDescription type=text size=50 maxlength=50 value="<%if strID <> 0 then Response.Write routineHTMLString(rsAssetClassDetail("ASSET_CLASS_DESC")) else Response.Write null end if%>" onChange ="fct_onChange();">
		</td>
		<td width=25% >&nbsp;</td>
	</TR>



	<tfoot>
	<tr>
		<td width="100%" colspan="4" align="right">
			<INPUT name=btnReferences type=button value=References style= "width: 2.2cm" LANGUAGE=javascript onclick="return btnReferences_onclick();">
			<INPUT name="btnDelete" type="button" value="Delete" style= "width: 2cm" LANGUAGE=javascript onClick="return fct_onDelete();">
			<INPUT name="btnReset" type="button" value="Reset" style= "width: 2cm" LANGUAGE=javascript onClick="return fct_onReset();">
			<INPUT name="btnNew" type="button" value="New" style= "width: 2cm" LANGUAGE=javascript onClick="return fct_NewAssetClassEntry();">
			<INPUT name="btnSave" type="button" value="Save" style= "width: 2cm" LANGUAGE=javascript onClick="return fct_OnSave();">
		</td>
	</tr>
	</tfoot>

</TABLE>
	<FIELDSET>
	<LEGEND align=right><B>Audit Information</B></LEGEND>
	<Div SIZE=8pt ALIGN=RIGHT>
		Record Status Indicator
		<INPUT align=left name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 18px"disabled value=<%if strID <> 0 then Response.Write """"&rsAssetClassDetail("record_status_ind")&"""" else Response.Write """""" end if%> >&nbsp;&nbsp;&nbsp;
		Create Date
		<INPUT align=center name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 140px"disabled value=<%if strID <> 0 then Response.Write """"&rsAssetClassDetail("create_date")&"""" else Response.Write """""" end if%>>&nbsp;
		Created By
		<INPUT align=right name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 100px"disabled value=<%if strID <> 0 then Response.Write """"&rsAssetClassDetail("create_real_userid")&"""" else Response.Write """""" end if%> ><BR>
		Update Date
		<INPUT align=center name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 140px"disabled value=<%if strID <> 0 then Response.Write """"&rsAssetClassDetail("update_date")&"""" else Response.Write """""" end if%> >
		Updated By
		<INPUT align=right name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 100px"disabled value=<%if strID <> 0 then Response.Write """"&rsAssetClassDetail("update_real_userid")&"""" else Response.Write """""" end if%> >
	</DIV>
	</FIELDSET>
</FORM>

<%



	if strID <> 0 then
		rsAssetClassDetail.close
		set rsAssetClassDetail = nothing
		objConn.close
		set objConn = nothing
	end if


%>


</BODY>
</HTML>


