<%@ Language=VBScript %>
<% option explicit %>
<% Response.Buffer = true %>
<% on error resume next %>


<!--
********************************************************************************************
* Page name:	AssetCatalogueDetail.asp
* Purpose:		To display the detailed information about an asset catalogue entry.
*				Entry is chosen via AssetCatList.asp
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
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to managed objects. Please contact your system administrator."
end if


'****************************
'declare necessary variables*
'****************************
dim lngAssetCatalogueID

dim strWinLocation
dim strWinMessage
dim strRealUserID



'get the hidden asset catalogue id from string
lngAssetCatalogueID = Request("hdntxtAssetCatalogueID")

'get the hidden window location
strWinLocation = "AssetCatDetail.asp?AssetCatalogueID="& Request.Form("hdntxtAssetCatalogueID")

'set the variable for the UserInfo cookie
strRealUserID = Session("username")

'************************
'do save, insert, delete*
'************************

select case Request("hdnFrmAction")

	case "SAVE"

'check to see if catalogue entry exists already in database by
'checking for the existence of the hidden asset catalogue id

	  if Request.Form("hdntxtAssetCatalogueID")  <> "" then  ' it is an existing record so save the changes

		if intAccessLevel and intConst_Access_Update <> intConst_Access_Update then

			DisplayError "REFRESH", strWinLocation, 0, "UPDATE DENIED", "You don't have access to update managed objects. Please contact your system administrator"

		end if

		dim cmdUpdateObj,arole

		set cmdUpdateObj = server.CreateObject("ADODB.Command")
		set cmdUpdateObj.ActiveConnection = objConn
		cmdUpdateObj.CommandType = adCmdStoredProc

		'get the asset_catalogue_detail stored update procedure <schema.package.procedure>
		cmdUpdateObj.CommandText = "sma_sp_userid.spk_sma_asset_inter.sp_asset_catalogue_update"


		'create the associated parameters
		'user id associated with time stamp


		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_userid", adVarChar, adParamInput, 20,strRealUserID)
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_asset_catalog_id", adNumeric, adParamInput, 22, Clng(Request("hdntxtAssetCatalogueID")))
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_make_id",adNumeric, adParamInput, 22, Clng(Request("hdnMakeID")))
        cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_model_id", adNumeric, adParamInput, 22, Clng(Request("hdnModelID")))
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_part_number_id", adNumeric, adParamInput, 22, Clng(Request("hdnPartNumID")))
        cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_last_update_dt", adDBTimeStamp,adParamInput, , CDate(Request("hdnUpdateDateTime")))

		'these fields can be empty

		if Request("txtSAPMatItemNumber") <> "" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_sap_material_item_number", adVarChar, adParamInput, 50, (Request("txtSAPMatItemNumber")))
		else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_sap_material_item_number", adVarChar, adParamInput, 50, null)
		end if


		if Request("txtComments") <> "" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_comments", adVarChar, adParamInput, 255, (Request("txtComments")))
		else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_comments", adVarChar, adParamInput, 255, null)
		end if




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


		'get the asset_catalog_detail insert procedure <schema.package.procedure>
		cmdInsertObj.CommandText = "sma_sp_userid.spk_sma_asset_inter.sp_asset_catalogue_insert"



		'create the insert parameters

		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_userid", adVarChar, adParamInput, 20, strRealUserID)
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_asset_catalogue_id", adNumeric, adParamOutput, , null)
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_make_id",adNumeric, adParamInput, , Clng(Request("hdnMakeID")))
        cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_model_id", adNumeric, adParamInput, , Clng(Request("hdnModelID")))
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_part_number_id", adNumeric, adParamInput, , Clng(Request("hdnPartNumID")))

		if Request("txtSAPMatItemNumber") <> "" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_sap_material_item_number", adVarChar, adParamInput, 50, (Request("txtSAPMatItemNumber")))
		else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_sap_material_item_number", adVarChar, adParamInput, 50, null)
		end if

		if Request("txtComments") <> "" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_comments", adVarChar, adParamInput, 255, (Request("txtComments")))
		else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_comments", adVarChar, adParamInput, 255, null)
		end if


		' execute the insert object

		cmdInsertObj.Execute

			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT CREATE NEW OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			else
				lngAssetCatalogueID = cmdInsertObj.Parameters("p_asset_catalogue_id").Value
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

			'get the asset_catalog_detail delete procedure <schema.package.procedure>
			cmdDeleteObj.CommandText = "sma_sp_userid.spk_sma_asset_inter.sp_asset_catalogue_delete"

			'create the delete parameters
			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_asset_catalogue_id", adNumeric, adParamInput, , CLng(lngAssetCatalogueID))
			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, ,CDate(Request("hdnUpdateDateTime")))



			'execute the delete object

			cmdDeleteObj.Execute

			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			end if

			lngAssetCatalogueID = 0
			strWinMessage = "Record deleted successfully."

end select


'*************************
'end save, insert, delete*
'*************************


'ok, now go get the detailed Asset Catalogue information


'declare the connection and sql variables
Dim strSQL
Dim strSelectClause
Dim strFromClause
Dim strWhereClause
Dim rsAssetCatalogue
Dim rsPartNumDefault


'declare the detail variables which will be used to populate the
'displayed and hidden fields
dim strMakeID
dim strMakeDesc
dim strModelID
dim strModelDesc
dim strPartNumberID
dim strPartNumDesc
dim strSAPMatItemNumber
dim strComments


'connect to the database using databaseconnect inc/smaconstants connection string
'<<CONNECT>>

'use the sqlstring to extract the necessary information from the database

	if lngAssetCatalogueID <> 0 then

		strSelectClause = "select " &_
					"t1.asset_catalogue_id, " & _
					"t1.make_id, " & _
					"t1.model_id, " & _
					"t1.part_number_id, " & _
					"t1.sap_material_item_number, " & _
					"t1.comments, " & _
					"t2.make_id, " & _
					"t2.make_desc, " & _
					"t3.model_id, " & _
					"t3.model_desc, " & _
					"t4.part_number_id, " & _
					"t4.part_number_desc, " & _
					"to_char(t1.create_date_time,'MON-DD-YYYY HH24:MI:SS') as create_date, " & _
					"sma_sp_userid.spk_sma_library.sf_get_full_username(t1.create_real_userid) as create_real_userid, " & _
					"to_char(t1.update_date_time,'MON-DD-YYYY HH24:MI:SS') as update_date, " & _
					"sma_sp_userid.spk_sma_library.sf_get_full_username(t1.update_real_userid) as update_real_userid, " & _
					"t1.update_date_time as last_update_date_time, " & _
					"t1.record_status_ind "

		strFromClause =	" from crp.asset_catalogue  t1, " &_
					"crp.make  t2, " & _
					"crp.model  t3, " & _
					"crp.part_number  t4 "

		 strWhereClause = " where " & _
					"t1.asset_catalogue_id = " & lngAssetCatalogueID & " and " & _
					"t1.make_id = t2.make_id and " & _
					"t1.model_id = t3.model_id and " & _
					"t1.part_number_id = t4.part_number_id "

		strSQL =  strSelectClause & strFromClause & strWhereClause

		'show SQL for debugging if necessary by using>>
		'Response.Write "<BR>" & strSQL	 & "<br>"

		'set and open the asset catalogue recordset and database connection

		set rsAssetCatalogue = Server.CreateObject("ADODB.Recordset")

		rsAssetCatalogue.CursorLocation = adUseClient
		rsAssetCatalogue.Open strSQL, objConn

		if err then
			DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
		end if
		if rsAssetCatalogue.EOF then
			DisplayError "BACK", "", 999, "CANNOT FIND REQUESTED OBJECT", "EOF condition occurred in rsAssetCatalogue recordset."
		end if


		'fill the detail variables with values from the recordset


		strMakeID = rsAssetCatalogue("make_id")
		strMakeDesc = rsAssetCatalogue("make_desc")
		strModelID = rsAssetCatalogue("model_id")
		strModelDesc = rsAssetCatalogue("model_desc")
		strPartNumberID = rsAssetCatalogue("part_number_id")
		strPartNumDesc = rsAssetCatalogue("part_number_desc")
		strSAPMatItemNumber = rsAssetCatalogue("sap_material_item_number")
		strComments = rsAssetCatalogue("comments")

	else

		'this defaults the part number to <none> and its id to -1 on NEW

		strsql= "SELECT part_number_id, part_number_desc FROM crp.part_number WHERE part_number_desc = '<none>'"

		set rsPartNumDefault = Server.CreateObject("ADODB.Recordset")

		rsPartNumDefault.CursorLocation = adUseClient
		rsPartNumDefault.Open strSQL, objConn

		if err then
			DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
		end if
		if rsPartNumDefault.EOF then
			DisplayError "BACK", "", 999, "CANNOT FIND REQUESTED OBJECT", "EOF condition occurred in rsPartNumDefault recordset."
		end if


	end if



%>



<HTML>
<HEAD>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<LINK rel="stylesheet" type="text/css" href="stylesheets/styles.css">
	<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
	<script type="text/javascript" SRC="AccessLevels.js"></script>
	<SCRIPT LANGUAGE=JavaScript>
	<!--

	var strWinMessage = '<%=strWinMessage%>';
    var intAccessLevel = '<%=intAccessLevel%>';
    var bolNeedToSave = false ;

	setPageTitle("SMA - Asset Catalogue Detail");

	//***************************************************************************************************
	// Function:	fct_selNavigate															            *
	//																									*
	// Purpose:		To display the page selected by the user from Quick Navigation drop-down box. The	*
	//				function saves Customer Name and Contact Name in cookes, which may be retrieved     *
	//				by the called page(s).                                    							*
	//																									*
	// 																									*																				*
	//**************************************************************************************************

	function fct_selNavigate(){

	 var strPageName;
	 var lngMakeID;
	 var lngModelID;
	 var lngPartNumberID;
	 var strMake = document.frmAssetCatDetail.txtMakeDesc.value ;
	 var strModel = document.frmAssetCatDetail.txtModelDesc.value ;

	strPageName = document.frmAssetCatDetail.selNavigate.item(document.frmAssetCatDetail.selNavigate.selectedIndex).value ;

		switch(strPageName)

			{
			case 'Make':
				document.frmAssetCatDetail.selNavigate.selectedIndex=0;
				lngMakeID = document.frmAssetCatDetail.hdnMakeID.value ;
				self.location.href = "MakeDetail.asp?hdnID=" + lngMakeID;
				break;
			case 'Model':
				document.frmAssetCatDetail.selNavigate.selectedIndex=0;
				lngModelID = document.frmAssetCatDetail.hdnModelID.value ;
				self.location.href = "ModelDetail.asp?hdnID=" + lngModelID;
				break;
			case 'PartNumber':
				document.frmAssetCatDetail.selNavigate.selectedIndex=0;
				lngPartNumberID = document.frmAssetCatDetail.hdnPartNumID.value ;
				self.location.href = "PartNumDetail.asp?hdnID=" + lngPartNumberID;
				break;
			case 'Asset':
			//This pulls up a list with your make and/or model parameters on the "Asset" screen
		        document.frmAssetCatDetail.selNavigate.selectedIndex=0;
		        if (strMake != ""){SetCookie("Make", strMake)};
				if (strModel != ""){SetCookie("Model", strModel)};
		        self.location.href = "SearchFrame.asp?fraSrc=" + strPageName ;
		        break;
            case 'DEFAULT':
				//do nothing
			}
		}




	//OK

	function fct_NewAssetCatEntry(){

	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create)

			{

			alert('Access denied. Please contact your system administrator.');
			return (false);

			}


			self.document.location.href = "AssetCatDet.asp?hdntxtAssetCatalogueID=0";



		}






	//OK

	function fct_OnSave(){

	var strComments = document.frmAssetCatDetail.txtComments.value;

	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create)

			{
				alert('Access denied. Please contact your system administrator.');
				return (false);
			}

			else
			{
				if (document.frmAssetCatDetail.txtMakeDesc.value == "" )
					{
						alert('Please select a Make');
						document.frmAssetCatDetail.btnMakeLookup.focus();
						return(false);

					}


				if (document.frmAssetCatDetail.txtModelDesc.value == "" )
					{
						alert('Please select a Model');
						document.frmAssetCatDetail.btnModelLookup.focus();
						return(false);

					}

				if (document.frmAssetCatDetail.txtPartNumDesc.value == "" )
					{
						alert('Please select a Part Number');
						document.frmAssetCatDetail.btnPartNumLookup.focus();
						return(false);

					}

				if (strComments.length > 255)
					{
						alert('Comments can be at most 255 characters.\n\nYou entered ' + strComments.length + ' character(s).');
						document.frmAssetCatDetail.txtComments.focus();
						return false;
					}


					else

					{
					document.frmAssetCatDetail.hdnFrmAction.value = "SAVE";
					bolNeedToSave = false;
					document.frmAssetCatDetail.submit();
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

		var lngAssetCatalogueID = document.frmAssetCatDetail.hdntxtAssetCatalogueID.value;
		var strUpdateDate = document.frmAssetCatDetail.hdnUpdateDateTime.value;

	    //alert ('the value of the catid is '+ document.frmAssetCatDetail.hdntxtAssetCatalogueID.value);
	    //alert ('the value of the update date time is '+ document.frmAssetCatDetail.hdnUpdateDateTime.value);

				{

					if (confirm('Do you really want to delete this object?'))

						{
							document.location = "AssetCatDet.asp?hdnFrmAction=DELETE&hdntxtAssetCatalogueID="+lngAssetCatalogueID+"&hdnUpdateDateTime="+strUpdateDate ;
							//alert(document.location);
						}
				}


		}




	//OK

	/*

	function fct_onReset()
	{
		var bolConfirm ;

		bolConfirm = window.confirm("Are you sure you want to Reset the form?");

	    if (bolConfirm){
			return true;
		}
		else {
			return false;
		}
	}

	**/

	function fct_onReset() {
		if(confirm('All changes will be lost. Do you really want to reset the page?')){
			bolNeedToSave = false ;
			document.location = 'AssetCatDet.asp?hdntxtAssetCatalogueID='+ '<%=lngAssetCatalogueID%>' ;
		}
	}

	//OK

    function fct_onChange(){

		bolNeedToSave = true;
	}


	//OK

	function fct_onBeforeUnload()

	{


		document.frmAssetCatDetail.btnSave.focus();

		if ((intAccessLevel & intConst_Access_Update) == intConst_Access_Update || (intAccessLevel & intConst_Access_Create) == intConst_Access_Create)
		{
			if (bolNeedToSave == true)
			{
				event.returnValue = "There is unsaved data on the screen. To save changes, click CANCEL below then click SAVE on the main form.";
			}
		}

	}




	/************************************
	   *BEGIN LOOKUP BUTTON FUNCTIONS*
	*************************************/

	//OK

	function btnMakeLookup_onclick(){

		if (document.frmAssetCatDetail.txtMakeDesc.value != "")
		{
			 SetCookie("MakeDesc", document.frmAssetCatDetail.txtMakeDesc.value);
		}

		SetCookie("WinName", 'Popup');
		bolNeedToSave = true ;
		//opens MakeCriteria.asp for search frame, MakeCriteriaList for list
		window.open('SearchFrame.asp?fraSrc=Make', 'Popup', 'top=50, left=100, height=600, width=800' ) ;
		document.frmAssetCatDetail.btnSave.disabled = false;

	}



	//OK

	function btnModelLookup_onclick(){
	    //alert(document.frmAssetCatDetail.txtModelDesc.value);

		if (document.frmAssetCatDetail.txtModelDesc.value != "")
		{
			SetCookie("ModelDesc", document.frmAssetCatDetail.txtModelDesc.value);
		}
		SetCookie("WinName", 'Popup');
		bolNeedToSave = true ;
		window.open('SearchFrame.asp?fraSrc=Model', 'Popup', 'top=50, left=100, height=600, width=800' ) ;
		document.frmAssetCatDetail.btnSave.disabled = false;
	}



	//OK

	function btnPartNumLookup_onclick(){
	    //alert(document.frmAssetCatDetail.txtPartNumDesc.value);

		if (document.frmAssetCatDetail.txtPartNumDesc.value != "")
		{
			SetCookie("PartNumDesc", document.frmAssetCatDetail.txtPartNumDesc.value);
		}

		SetCookie("WinName", 'Popup');
		bolNeedToSave = true ;
		window.open('SearchFrame.asp?fraSrc=PartNum', 'Popup', 'top=50, left=100, height=600, width=800' ) ;
		document.frmAssetCatDetail.btnSave.disabled = false;
	}



	/************************************
	   *END LOOKUP BUTTON FUNCTIONS*
	************************************/


	//OK

	function fct_clearStatus() {
		window.status = "";
	}

	//OK
	function fct_DisplayStatus(strWindowStatus){

	window.status=strWindowStatus;
	setTimeout('fct_clearStatus()', '<%=intConst_MessageDisplay%>');

    }

    //OK
	function btnReferences_onclick() {

		var strOwner = 'CRP' ;
		var strTableName = 'ASSET_CATALOGUE' ;
		var strRecordID = document.frmAssetCatDetail.hdntxtAssetCatalogueID.value ;
		var URL ;


		if ( lngAssetCatalogueID = 0)

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

<BODY onLoad="fct_DisplayStatus(strWinMessage);" onbeforeunload="fct_onBeforeUnload();">
<FORM name=frmAssetCatDetail action="AssetCatDet.asp"  method="POST">

	<!-- hidden variables where requested values are stored-->

	<INPUT name=hdntxtAssetCatalogueID type=hidden value= <%if lngAssetCatalogueID <> 0 then Response.Write rsAssetCatalogue("asset_catalogue_id") else Response.Write """""" end if%>>
	<INPUT name=hdnMakeID  type=hidden value=<%if lngAssetCatalogueID <> 0 then Response.Write rsAssetCatalogue("make_id") else Response.Write """""" end if%>>
	<INPUT name=hdnModelID type=hidden value=<%if lngAssetCatalogueID <> 0 then Response.Write rsAssetCatalogue("model_id") else Response.Write """""" end if%>>
	<INPUT name=hdnPartNumID type=hidden value=<%if lngAssetCatalogueID <> 0 then Response.Write rsAssetCatalogue("part_number_id") else Response.Write rsPartNumDefault("PART_NUMBER_ID") end if%>>
	<INPUT name=hdnUpdateDateTime type=hidden value="<%if lngAssetCatalogueID <> 0 then  Response.Write rsAssetCatalogue("last_update_date_time") else Response.Write """""" end if%>">
    <INPUT id=hdnFrmAction name=hdnFrmAction type=hidden value= "">




	<!-- user interface -->

	<TABLE border=0 width=100%>

	<thead>
		<tr>
			<td colspan=3 align=left>Asset Catalogue Detail</td>
			<td align=right><SELECT <%if lngAssetCatalogueID = 0 then Response.Write "disabled" end if%> align=right valign=top name=selNavigate LANGUAGE=javascript onchange="fct_selNavigate();">
				<OPTION value="DEFAULT">Quickly Goto ...</OPTION>
				<OPTION value="Make" >Make</OPTION>
				<OPTION value="Model" >Model</OPTION>
				<OPTION value="PartNumber" >Part Number</OPTION>
				<OPTION value="Asset" >Asset List</OPTION>
				</SELECT>
			</td>
	</thead>

	<TR>
		<TD align=right width=25%>Make<font color=red>*</font></TD>
		<TD align=left width=50% colspan=2>
			<input name=txtMakeDesc type=text disabled size=50 maxlength=50 value="<%if lngAssetCatalogueID <> 0 then Response.Write (routineHtmlString(strMakeDesc)) else Response.Write null end if%>" onChange ="fct_onChange();">
			<INPUT align=right type="button"  name=btnMakeLookup  value="..." onclick="return btnMakeLookup_onclick()">
		</td>
		<td width=25% >&nbsp;</td>
	</TR>


	<TR>
		<TD align=right width=25%>Model<font color=red>*</font></TD>
		<TD align=left width=50% colspan=2>
			<input name=txtModelDesc type=text  disabled size=50 maxlength=50 value="<%if lngAssetCatalogueID <> 0 then Response.Write (routineHtmlString(strModelDesc)) else Response.Write null end if%>" onChange ="fct_onChange();">
			<INPUT align=right type="button"  name=btnModelLookup   value="..." onclick="return btnModelLookup_onclick()">
		</td>
		<td width=25% >&nbsp;</td>
	</TR>

	<TR>
		<TD align=right width=25%>Part Number<font color=red>*</font></TD>
		<TD align=left width=50% colspan=2>
			<INPUT name=txtPartNumDesc type=text disabled size=50 maxlength=50 value="<%if lngAssetCatalogueID <> 0 then Response.Write (routineHtmlString(strPartNumDesc)) else Response.Write rsPartNumDefault("PART_NUMBER_DESC") end if%>" onChange ="fct_onChange();">
			<INPUT align=right type="button"  name=btnPartNumberLookup   value="..." onclick="return btnPartNumLookup_onclick()">
		</td>
		<td width=25% >&nbsp;</td>
	</TR>

	<TR>
		<TD align=right width=25%>SAP Material Item Number</TD>
		<TD align=left width=50% colspan=2>
			<input name=txtSAPMatItemNumber type=text  size=50 maxlength=50 value="<%if lngAssetCatalogueID <> 0 then Response.Write (routineHtmlString(strSAPMatItemNumber)) else Response.Write null end if%>" onChange ="fct_onChange();">
		</td>
		<td width=25% >&nbsp;</td>
	</TR>

	<tr>
		<td  align=right valign=top>Comments&nbsp;</td>
		<td  rowspan=3 valign="top">
		<TEXTAREA cols=25 name=txtComments rows=6 style="width: 360" type=text
		onChange="fct_onChange();"><%if lngAssetCatalogueID <> 0 then Response.Write (routineHtmlString(strComments)) else Response.Write null end if%></TEXTAREA></td>
		<td  align="left" >

	</tr>

	<td width=25%>&nbsp;</td>
	</tr><tr>
		<td width=25%>&nbsp;</td>
		<td width=25%>&nbsp;</td>
		<td width=25%>&nbsp;</td>
		<td width=25%>&nbsp;</td>


	<tfoot>
	<tr>
		<td width="100%" colspan="4" align="right">
			<INPUT name="btnReferences" type="button" value="References" style= "width: 2.2cm" LANGUAGE=javascript onclick="return btnReferences_onclick()">&nbsp;&nbsp;
			<INPUT name="btnDelete" type="button" value="Delete" style= "width: 2cm" onClick="return fct_onDelete();">
			<INPUT name="btnReset" type="button" value="Reset" style= "width: 2cm" onClick= "return fct_onReset();">
			<INPUT name="btnNew" type="button" value="New" style= "width: 2cm" onClick="return fct_NewAssetCatEntry();">
			<INPUT id="btnSave" name="btnSave" type="button" value="Save" style= "width: 2cm" onClick="return fct_OnSave();">
		</td>
	</tr>
	</tfoot>

</TABLE>
	<FIELDSET>
	<LEGEND align=right><B>Audit Information</B></LEGEND>
	<Div SIZE=8pt ALIGN=RIGHT>
		Record Status Indicator
		<INPUT align=left name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 18px"disabled value=<%if lngAssetCatalogueID <> 0 then Response.Write """"&rsAssetCatalogue("record_status_ind")&"""" else Response.Write """""" end if%> >&nbsp;&nbsp;&nbsp;
		Create Date
		<INPUT align=center name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 140px"disabled value=<%if lngAssetCatalogueID <> 0 then Response.Write """"&rsAssetCatalogue("create_date")&"""" else Response.Write """""" end if%>>&nbsp;
		Created By
		<INPUT align=right name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 100px"disabled value=<%if lngAssetCatalogueID <> 0 then Response.Write """"&rsAssetCatalogue("create_real_userid")&"""" else Response.Write """""" end if%> ><BR>
		Update Date
		<INPUT align=center name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 140px"disabled value=<%if lngAssetCatalogueID <> 0 then Response.Write """"&rsAssetCatalogue("update_date")&"""" else Response.Write """""" end if%> >
		Updated By
		<INPUT align=right name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 100px"disabled value=<%if lngAssetCatalogueID <> 0 then Response.Write """"&rsAssetCatalogue("update_real_userid")&"""" else Response.Write """""" end if%> >
	</DIV>
	</FIELDSET>
</FORM>

<%



	if lngAssetCatalogueID <> 0 then

		rsAssetCatalogue.close
		set rsAssetCatalogue = nothing

	else

		rsPartNumDefault.Close
		set rsPartNumDefault=nothing

	end if

	objConn.close
	set objConn = nothing

%>


</BODY>
</HTML>

