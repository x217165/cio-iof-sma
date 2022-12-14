<%@ Language=VBScript %>
<% option explicit 
Response.Buffer = true
on error resume next

'/////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'//	This page displays the details of a network element. To jump to a specific network element pass the 
'//	network_element_id as "txtNE_ID" in a QueryString, as a parameter in a POST method or as a cookie.
'//
'/////////////////////////////////////////////////////////////////////////////////////////////////////////
%>
<!--#include file="smaConstants.inc"-->
<!--#include file="databaseconnect.asp"-->
<!--#include file="smaProcs.inc"-->
<!--
*************************************************************************************
* File Name:	manobjdet.asp
*
* Purpose:	
*
* In Param:		
*				
* Out Param:
*
* Created By:	
**************************************************************************************
		 Date		Author			Changes/enhancements made

       25-Jan-02   Adam Haydey  Added Customer Service City, Customer Service Address,
									TAC Assset Code and Non-Correlated Only search fields.
				                TAC Asset Code was added to the search results.
       11-Mar-02   DTy	        Add Port Name and IP - primarily used by HRDC.
       20-Oct-03   DTy          Revise Port Information layout.
								Expand screen, add 'Clone' button.
	13-Nov-03   DTy		Make 'IP Address' mandatory.
	16-Aug-04   ACheung	Add LYNX repair priority
**************************************************************************************
-->

<%
'check user's rights

dim intAccessLevel
intAccessLevel = CInt(CheckLogon(strConst_ManagedObjects))
if intAccessLevel and intConst_Access_ReadOnly <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to managed object. Please contact your system administrator"
end if

dim sql, ne_id, strNE_ID, strWinMessage, strServLocAddress, bolClone
dim strReadOnly
dim rsNE

bolClone = false

'get requested network element id
strNE_ID = Request("ne_id")
if strNE_ID = "" then
	strNE_ID = Request.Cookies("txtNE_ID")
end if

dim strRealUserID
strRealUserID = Request.Cookies("UserInformation")("username")
if err then
	'unexpected error
	DisplayError "BACK", "", 0, "UNEXPECTED ERROR", "Close alias window to return to managed objects form."
end if

select case Request("txtFrmAction")
	case "SAVE" 
		if Request("txtObjID") <> "" then
 			if (intAccessLevel and intConst_Access_Update <> intConst_Access_Update) then
				DisplayError "BACK", "", 0, "UPDATE DENIED", "You don't have access to update managed objects. Please contact your system administrator"
			end if
			'save the network element id and do a save
			strNE_ID = Request("txtObjID")
			'create command object for update stored proc
			dim cmdUpdateObj
			set cmdUpdateObj = server.CreateObject("ADODB.Command")
			set cmdUpdateObj.ActiveConnection = objConn
			cmdUpdateObj.CommandType = adCmdStoredProc
			cmdUpdateObj.CommandText = "sma_sp_userid.spk_sma_mo_inter.sp_mo_update"
			'create params 
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_userid", adVarChar, adParamInput, 20, strRealUserID)											'varchar2(30)	means: Real User ID
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_network_element_id",adNumeric , adParamInput,, CLng(Request("txtObjID"))) 					'number(9)		means: Managed Object Id
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_network_element_name", adVarChar,adParamInput, 30, Request("txtObjName"))					'varchar2(30)	means: Managed Object Name
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_network_element_type_code", adVarChar,adParamInput, 6, Request("selObjType"))				'varchar2(6)	means: Managed Object Type
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_customer_id", adNumeric, adParamInput,, CLng(Request("hdnCustomerID")))						'number(9)		means: Customer (id)
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_service_location", adNumeric, adParamInput,, CLng(Request("hdnServLocID")))					'number(9)		means: Service Location (id)
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_asset_catalogue_id", adNumeric, adParamInput,, CLng(Request("hdnAssetCatalogueID")))			'number(9)		means: Id of Make/Model/Port reference
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_network_element_desc", adVarChar,adParamInput, 80, Request("txtObjDesc"))					'varchar2(80)	means: Managed Object Description
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_repair_priority", adVarChar, adParamInput, 30, Request("selRepairPriority"))	        'LYNX repair priority
			if Request("hdnAssetID") <> "" then
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_asset_id", adNumeric, adParamInput,, clng(Request("hdnAssetID")))
			else 
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_asset_id", adNumeric, adParamInput,, null)
			end if
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_managed_ip_address", adVarChar,adParamInput, 30, Request("txtIPAddress"))					'varchar2(30)	means: Managed IP Address
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_trusted_host_mac_address", adVarChar, adParamInput, 30, Request("txtMACAddress"))			
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_comments", adVarChar,adParamInput, 2000, Request("txtComments"))								
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_out_of_band_dialup", adVarChar, adParamInput, 30, Request("txtOBDialUp"))					
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_support_group", adVarChar,adParamInput, 15, Request("selSupportGroup"))						
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_serial_no", adVarChar, adParamInput, 30, Request("txtSerialNumber"))							
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_barcode", adVarChar, adParamInput, 30, Request("txtBarcode"))								
			if Request("selSupportContactRole") <> "" then
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_remedy_contact", adVarChar, adParamInput,15 , Request("selSupportContactRole"))			
			else
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_remedy_contact", adVarChar, adParamInput,15, null)			
			end if
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , CDate(Request("hdnNEUpdateDateTime")))		'date			means: update_date_time from Network_Element record
			
				
			'call the insert stored proc 
  			'dim objparm
  			'for each objparm in cmdUpdateObj.Parameters
  			 ' Response.Write "<b>" & objparm.name & "</b>"
  			 ' Response.Write " has size:  " & objparm.Size & " "
  			 ' Response.Write " and value:  " & objparm.value & " "
  			 'Response.Write " and datatype:  " & objparm.Type & "<br> "
  		  'next
  		   
  		   'Response.Write "<b> count = " & cmdUpdateObj.Parameters.count & "<br>"
  		'	dim nx
  		'	 for nx=0 to cmdUpdateObj.Parameters.count-1
  		'	   Response.Write " parm value= " & cmdUpdateObj.Parameters.Item(nx).Value  & "<br>"
  		'	  next 
	
			
			
			'call the update stored proc
			'on error resume next
			cmdUpdateObj.Execute
			if err then
				if instr(1, objConn.Errors(0).Description, "ORA-20040" ) then
					dim strWinLocation
					strWinLocation = "manobjdet.asp?ne_id="&Request("txtObjID")
					DisplayError "REFRESH", strWinLocation, objConn.Errors(0).NativeError, "OBJECT UPDATED", objConn.Errors(0).Description
				else
					DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT UPDATE OBJECT", objConn.Errors(0).Description
				end if
				objConn.Errors.Clear
			end if
			strWinMessage = "Record saved successfully. You can now see the changes you made."
		else
			'create a new record
			if (intAccessLevel and intConst_Access_Create <> intConst_Access_Create) then
				DisplayError "BACK", "", 0, "INSERT DENIED", "You don't have access to create managed objects. Please contact your system administrator"
			end if
			dim cmdInsertObj
			set cmdInsertObj = server.CreateObject("ADODB.Command")
			set cmdInsertObj.ActiveConnection = objConn
			cmdInsertObj.CommandType = adCmdStoredProc
			cmdInsertObj.CommandText = "sma_sp_userid.spk_sma_mo_inter.sp_mo_insert"
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_userid", adVarChar, adParamInput, 20, strRealUserID)									'varchar2(20)	means: Real User ID
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_network_element_id", adNumeric, adParamOutput) 										'number(9)		means: Managed Object Id
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_network_element_name", adVarChar, adParamInput, 30, Request("txtObjName"))			'varchar2(30)	means: Managed Object Name
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_network_element_type_code", adVarChar, adParamInput, 6, Request("selObjType"))		'varchar2(6)	means: Managed Object Type
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_customer_id", adNumeric, adParamInput,, CLng(Request("hdnCustomerID")))				'number(9)		means: Customer (id)
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_service_location", adNumeric, adParamInput,, CLng(Request("hdnServLocID")))			'number(9)		means: Service Location (id)
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_asset_catalogue_id", adNumeric, adParamInput,, CLng(Request("hdnAssetCatalogueID")))	'number(9)		means: Id of Make/Model/Port reference
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_warning_message", adVarChar, adParamOutput, 200, null) 								'varchar        only returned when we have a warning message
			
			'optional fields
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_network_element_desc", adVarChar, adParamInput, 80, Request("txtObjDesc"))			'varchar2(80)	means: Managed Object Description
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_repair_priority", adVarChar, adParamInput, 30, Request("selRepairPriority"))	        'LYNX repair priority
			if Request("hdnAssetID") <> "" then
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_asset_id", adNumeric, adParamInput,, CLng(Request("hdnAssetID")))
			else 
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_asset_id", adNumeric, adParamInput,, null)
			end if
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_managed_ip_address", adVarChar, adParamInput, 30, Request("txtIPAddress"))			'varchar2(30)	means: Managed IP Address
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_trusted_host_mac_address", adVarChar, adParamInput, 30, Request("txtMACAddress"))	'varchar2(30)	means: Trusted Host Mac Address
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_comments", adVarChar, adParamInput, 2000, Request("txtComments"))					'varchar2(2000)	means: Comments
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_out_of_band_dialup", adVarChar, adParamInput, 30, Request("txtOBDialUp"))			'varchar2(30)	means: Out of Band Dialup
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_support_group", adVarChar, adParamInput, 15, Request("selSupportGroup"))				'varchar2(15)	means: Support Group
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_serial_no", adVarChar, adParamInput, 30, Request("txtSerialNumber"))					'varchar2(30)	means: Serial Number
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_barcode", adVarChar, adParamInput, 30, Request("txtBarcode"))						'varchar2(30)	emans: Barcode
			if Request("selSupportContactRole") <> "" then
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_remedy_contact", adVarChar, adParamInput, 15, Request("selSupportContactRole"))		
			else
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_remedy_contact", adVarChar, adParamInput, 15, null)			
			end if
			
			'on error resume next
			cmdInsertObj.Execute
			'on error goto 0

			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT CREATE NEW OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			else
				Dim strWarning
				strNE_ID = cmdInsertObj.Parameters("p_network_element_id").Value
				strWarning = cmdInsertObj.Parameters("p_warning_message").Value
				'Response.Write "warning: " & strWarning
				if strWarning <> "" then
					strWinLocation = "manobjdet.asp?ne_id=" & strNE_ID
					DisplayError "REFRESH", strWinLocation, "-20040", "OBJECT INSERTED", strWarning
				end if
			end if

			strWinMessage = "Record created successfully. You can now see the new record."
		end if
	case "DELETE"
		'delete record
		if (intAccessLevel and intConst_Access_Delete <> intConst_Access_Delete) then
			DisplayError "BACK", "", 0, "DELETE DENIED", "You don't have access to delete managed objects. Please contact your system administrator"
		end if
			dim cmdDeleteObj
			set cmdDeleteObj = server.CreateObject("ADODB.Command")
			set cmdDeleteObj.ActiveConnection = objConn
			cmdDeleteObj.CommandType = adCmdStoredProc
			cmdDeleteObj.CommandText = "sma_sp_userid.spk_sma_mo_inter.sp_mo_delete"
			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_network_element_id", adNumeric, adParamInput, ,CLng(Request("txtObjID")))					'number(9)		means: Managed Object Id
			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , CDate(Request("hdnNEUpdateDateTime")))		'date			means: update_date_time from Network_Element record
			cmdDeleteObj.Execute
			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			end if
			strNE_ID = ""
			strWinMessage = "Record deleted successfully."
	case "CLONE"
		bolClone = true
end select

if strNE_ID <> "" then
	'build query
	sql = "SELECT " &_
			"NE.NETWORK_ELEMENT_ID, "&_
			"NE.ASSET_ID, "&_
			"NE.ASSET_CATALOGUE_ID, "&_
			"NE.CUSTOMER_ID, "&_
			"NE.SERVICE_LOCATION_ID, "&_
			"NE.NETWORK_ELEMENT_NAME, " &_
			"NE.NETWORK_ELEMENT_DESC, " &_
			"NE.MANAGED_IP_ADDRESS, "&_
			"NE.TRUSTED_HOST_MAC_ADDRESS, "&_
			"NE.OUT_OF_BAND_DIALUP, "&_
			"NE.SERIAL_NUMBER, "&_
			"NE.BARCODE, "&_
			"AST.TAC_NAME ASSET_TAC_NAME, "&_
			"AMAK.MAKE_DESC ASSET_MAKE_DESC, "&_
			"AMOD.MODEL_DESC ASSET_MODEL_DESC, "&_
			"APN.PART_NUMBER_DESC ASSET_PART_NO_DESC, "&_
			"CUS.CUSTOMER_NAME, "&_
			"CUS.CUSTOMER_SHORT_NAME, "&_
			"SL.SERVICE_LOCATION_NAME, "&_
			"NE.NETWORK_ELEMENT_TYPE_CODE, "&_
			"CUS.NOC_REGION_LCODE, " &_
			"NE.REMEDY_SUPPORT_GROUP_ID, " &_
			"NE.REMEDY_CONTACT_ROLE_ID, "&_
			"NE.COMMENTS, "&_
			"NE.LYNX_DEF_SEV_LCODE, "&_
			"SLA.BUILDING_NAME, "&_
			"SLA.STREET, "&_
			"SLA.MUNICIPALITY_NAME, "&_
			"SLA.PROVINCE_STATE_LCODE, "&_
			"SLA.POSTAL_CODE_ZIP, "&_
			"TO_CHAR(NE.CREATE_DATE_TIME,'MON-DD-YYYY HH24:MI:SS') CREATE_DATE_TIME, "&_
			"sma_sp_userid.spk_sma_library.sf_get_full_username(NE.CREATE_REAL_USERID) as create_real_userid, "&_
			"TO_CHAR(NE.UPDATE_DATE_TIME,'MON-DD-YYYY HH24:MI:SS') UPDATE_DATE_TIME, "&_
			"sma_sp_userid.spk_sma_library.sf_get_full_username(NE.UPDATE_REAL_USERID) as update_real_userid, "&_
			"NE.RECORD_STATUS_IND "&_
		"FROM " &_
			"CRP.NETWORK_ELEMENT				NE, "&_
			"CRP.ASSET							AST, "&_
			"CRP.ASSET_CATALOGUE				ACAT, "&_
			"CRP.MAKE							AMAK, "&_
			"CRP.MODEL							AMOD, "&_
			"CRP.PART_NUMBER					APN, "&_
			"CRP.CUSTOMER						CUS, "&_
			"CRP.SERVICE_LOCATION				SL, "&_
			"CRP.V_ADDRESS_CONSOLIDATED_STREET	SLA "&_
		"WHERE " &_
			"NE.CUSTOMER_ID = CUS.CUSTOMER_ID "&_
			"AND NE.ASSET_ID = AST.ASSET_ID (+) "&_
			"AND NE.ASSET_CATALOGUE_ID = ACAT.ASSET_CATALOGUE_ID "&_
			"AND ACAT.MAKE_ID = AMAK.MAKE_ID "&_
			"AND ACAT.MODEL_ID = AMOD.MODEL_ID "&_
			"AND ACAT.PART_NUMBER_ID = APN.PART_NUMBER_ID "&_
			"AND NE.SERVICE_LOCATION_ID = SL.SERVICE_LOCATION_ID "&_
			"AND SL.ADDRESS_ID = SLA.ADDRESS_ID "&_
			"AND NE.NETWORK_ELEMENT_ID = " & strNE_ID
	
	'get the network element recordset
	if err then
		DisplayError "BACK", "", err.Number, "UNEXPECTED ERROR", err.Description
	end if
	set rsNE=server.CreateObject("ADODB.Recordset")
	rsNE.CursorLocation = adUseClient
	rsNE.Open sql, objConn
	if err then
		DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
	end if

	if rsNE.EOF then
		DisplayError "BACK", "", 999, "CANNOT FIND REQUESTED OBJECT", "EOF condition occured in rsNE recordset."
	else
	'preformat Service Location Address
		if len(rsNE("building_name") ) > 0 then
			strServLocAddress = rsNE("building_name") & vbNewLine & rsNE("street") & vbNewLine&_
			rsNE("municipality_name") & " " & rsNE("province_state_lcode")
		else
			strServLocAddress = rsNE("street") & rsNE("municipality_name") & " " & rsNE("province_state_lcode")
		end if 
	end if

	set rsNE.ActiveConnection = nothing
end if


'get the network element type recordset
dim rsNET
sql = "SELECT NETWORK_ELEMENT_TYPE_CODE FROM CRP.NETWORK_ELEMENT_TYPE WHERE RECORD_STATUS_IND='A' ORDER BY NETWORK_ELEMENT_TYPE_CODE"
set rsNET=server.CreateObject("ADODB.Recordset")
rsNET.CursorLocation = adUseClient
rsNET.Open sql, objConn
if err then
	DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
end if

if rsNET.EOF then
	DisplayError "BACK", "", 999, "CANNOT CREATE OBJECT TYPE LIST", "EOF condition occured in rsNET recordset."
end if

set rsNET.ActiveConnection = nothing


'get the support group recordset
dim rsSG
sql = "SELECT REMEDY_SUPPORT_GROUP_ID, GROUP_NAME FROM CRP.V_REMEDY_SUPPORT_GROUP ORDER BY GROUP_NAME"
set rsSG=server.CreateObject("ADODB.Recordset")
rsSG.CursorLocation = adUseClient
rsSG.Open sql, objConn
if err then
	DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
end if

if rsSG.EOF then
	DisplayError "BACK", "", 999, "CANNOT CREATE SUPPORT GROUP LIST", "EOF condition occured in rsSG recordset."
end if

set rsSG.ActiveConnection = nothing


'get the support contact role recordset
dim rsSCR
sql = "SELECT REMEDY_CONTACT_ROLE_ID, CONTACT_ROLE_NAME  FROM CRP.V_REMEDY_CONTACT_ROLE  ORDER BY CONTACT_ROLE_NAME"
set rsSCR=server.CreateObject("ADODB.Recordset")
rsSCR.CursorLocation = adUseClient
rsSCR.Open sql, objConn
if err then
	DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
end if

if rsSCR.EOF then
	DisplayError "BACK", "", 999, "CANNOT CREATE CONTACT ROLE LIST", "EOF condition occured in rsSCR recordset."
end if

set rsSCR.ActiveConnection = nothing

'get the LYNX default repair priority
dim rsLYNXrp

sql = "select LYNX_DEF_SEV_DESC, LYNX_DEF_SEV_LCODE from CRP.LCODE_LYNX_DEF_SEV where RECORD_STATUS_IND='A' ORDER BY LYNX_DEF_SEV_LCODE"
set rsLYNXrp=server.CreateObject("ADODB.Recordset")
rsLYNXrp.CursorLocation = adUseClient
rsLYNXrp.Open sql, objConn
if err then
	DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
end if

if rsLYNXrp.EOF then
	DisplayError "BACK", "", 999, "CANNOT CREATE CONTACT ROLE LIST", "EOF condition occured in rsLYNXrp recordset."
end if

set rsLYNXrp.ActiveConnection = nothing


%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft FrontPage 4.0">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>Managed Objects - Details</TITLE>
</HEAD>
<LINK href="stylesheets/styles.css" rel="stylesheet" type="text/css">
<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
<script type="text/javascript" SRC="AccessLevels.js"></script>
<SCRIPT type="text/javascript">
//set section title
setPageTitle("SMA - Managed Object");

var intAccessLevel=<%=CheckLogon(strConst_ManagedObjects)%>;
var intAccessLevelDetail=<%=CheckLogon(strConst_ManagedObjects)%>;
var intConst_MessageDisplay=<%=intConst_MessageDisplay%>
var bolSaveRequired = false;

function iFrame_display(){
	if ((intAccessLevelDetail & intConst_Access_ReadOnly) == intConst_Access_ReadOnly){
		document.frames("aifr").document.location.href = 'manobjalias.asp?ne_id=<%if not bolClone then response.write strNE_ID end if%>';
	}else{alert('You do not have access to name alias. Please contact your system administrator.')}
}

function btn_iFrmAdd(){
	//open a blank form
	if ((intAccessLevelDetail & intConst_Access_Create) != intConst_Access_Create){
		alert('Access denied. Please contact your system administrator.');
		return;
	}
	if (document.frmMODetails.txtObjID.value == "") {
		alert('At this time you cannot create a name alias. You must save the object first.');
		return;
	}
	var NewWin;
	var strMasterID = "<%=strNE_ID%>";
	NewWin=window.open("manobjaliasdetail.asp?action=new&masterID="+strMasterID ,"NewWin","toolbar=no,status=yes,width=700px,height=175px,left=150px,top=200,menubar=no,resize=no");
	NewWin.focus();
}

function btn_iFrmUpdate(){
	//open a detail form where the user can modify the alias
	if ((intAccessLevelDetail & intConst_Access_Update) != intConst_Access_Update){
		alert('Access denied. Please contact your system administrator.');
		return;
	}
	var NewWin;
	var strAliasID = document.frames("aifr").document.frmIFR.hdnNameAliasID.value;
	if (strAliasID == "") {
		alert("Please select an alias or click ADD NEW to create a new alias.");
		return;
	}
	var strMasterID = "<%=strNE_ID%>";
	NewWin=window.open("manobjaliasdetail.asp?action=update&aliasID="+strAliasID+"&masterID="+strMasterID ,"NewWin","toolbar=no,status=yes,width=700px,height=175px,left=150px,top=200,menubar=no,resise=no");
	NewWin.focus();
}

function btn_iFrmDelete(){
//delete selected row
	if ((intAccessLevelDetail & intConst_Access_Delete) != intConst_Access_Delete){
		alert('Access denied. Please contact your system administrator.');
		return;
	}
	var strAliasID = document.frames("aifr").document.frmIFR.hdnNameAliasID.value;
	if (strAliasID == "") {
		alert("Please select an alias or click ADD NEW to create a new alias.");
		return;
	}
	var strLastUpdate = document.frames("aifr").document.frmIFR.hdnLastUpdate.value;
	if (confirm("Do you want to delete this alias?")){
		document.frames("aifr").document.location.href = "ManObjAliasDetail.asp?action=delete&back=true&aliasID="+strAliasID+"&masterID=<%if not bolDelete then response.write strNE_ID end if%>&hdnLastUpdate="+strLastUpdate;
	}
}

//javascript code not related to iFrame functionality
var strWinMessage = "<%=strWinMessage%>";
var strOuterNameAliasHTML;

function iFrame2_display(){
	if ((intAccessLevelDetail & intConst_Access_ReadOnly) == intConst_Access_ReadOnly){
		document.frames("aifr2").document.location.href = 'ManObjPort.asp?ne_id=<%if not bolClone then response.write strNE_ID end if%>';
	}else{alert('You do not have access to Port Information. Please contact your system administrator.')}
}

function btn_iFrm2Add(){
	//open a blank form
	if ((intAccessLevelDetail & intConst_Access_Create) != intConst_Access_Create){
		alert('Access denied. Please contact your system administrator.');
		return;
	}
	if (document.frmMODetails.txtObjID.value == "") {
		alert('You cannot create a name Port Information record. You must save the Managed Object first.');
		return;
	}
	var NewWin;
	var strMasterID = "<%=strNE_ID%>";
	NewWin=window.open("ManObjPortDetail.asp?action=new&masterID="+strMasterID ,"NewWin","toolbar=no,status=yes,width=900px,height=275px,left=175px,top=200,menubar=no,resize=no");
	NewWin.focus();
}

function btn_iFrm2Update(){
	//open a detail form where the user can modify the Port Information
	if ((intAccessLevelDetail & intConst_Access_Update) != intConst_Access_Update){
		alert('Access denied. Please contact your system administrator.');
		return;
	}
	var NewWin;
	var strPortID = document.frames("aifr2").document.frmIFR2.hdnPortID.value;
	var strLastUpdate = document.frames("aifr2").document.frmIFR2.hdnLastUpdate.value;
	if (strPortID == "") {
		alert("Please select a Port Information record firt then click UPDATE to update the selected record.");
		return;
	}
	var strMasterID = "<%=strNE_ID%>";
	NewWin=window.open("ManObjPortDetail.asp?action=update&PortID="+strPortID+"&masterID="+strMasterID+"&hdnLastUpdate="+strLastUpdate,"NewWin","toolbar=no,status=yes,width=900px,height=275px,left=175px,top=200,menubar=no,resize=no");
	NewWin.focus();
}

function btn_iFrm2Clone(){
	//open a detail form where the user can modify the Port Information
	if ((intAccessLevelDetail & intConst_Access_Update) != intConst_Access_Update){
		alert('Access denied. Please contact your system administrator.');
		return;
	}
	var NewWin;
	var strPortID = document.frames("aifr2").document.frmIFR2.hdnPortID.value;
	if (strPortID == "") {
		alert("Please select a Port Information record first then click CLONE to create a new record using the selected record.");
		return;
	}
	var strMasterID = "<%=strNE_ID%>";
	NewWin=window.open("ManObjPortDetail.asp?action=clone&PortID="+strPortID+"&masterID="+strMasterID ,"NewWin","toolbar=no,status=yes,width=900px,height=275px,left=175px,top=200,menubar=no,resize=no");
	NewWin.focus();
}

function btn_iFrm2Delete(){
//delete selected row
	if ((intAccessLevelDetail & intConst_Access_Delete) != intConst_Access_Delete){
		alert('Access denied. Please contact your system administrator.');
		return;
	}
	var strPortID = document.frames("aifr2").document.frmIFR2.hdnPortID.value;
	if (strPortID == "") {
		alert("Please select a Port Information record first then click DELETE to delete the selected record.");
		return;
	}
	var strLastUpdate = document.frames("aifr2").document.frmIFR2.hdnLastUpdate.value;
	if (confirm("Do you want to delete this Port Information?")){
		document.frames("aifr2").document.location.href = "ManObjPortDetail.asp?action=delete&back=true&PortID="+strPortID+"&masterID=<%if not bolDelete then response.write strNE_ID end if%>&hdnLastUpdate="+strLastUpdate;
	}
}

//javascript code not related to iFrame functionality
var strWinMessage = "<%=strWinMessage%>";
var strOuterNameAliasHTML;

function fct_NewMO(){
	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create){
		alert('Access denied. Please contact your system administrator.');
		return;
	}
	document.location="manobjdet.asp?ne_id=";
}

function fct_onChange(){
	bolSaveRequired = true;
}

function btn_onSave(){

  
	if	((((intAccessLevel & intConst_Access_Create) == intConst_Access_Create) && (document.frmMODetails.txtObjID.value == "")) || (((intAccessLevel & intConst_Access_Update) == intConst_Access_Update) && (document.frmMODetails.txtObjID.value != ""))) {
    //check mandatory fields
		if (document.frmMODetails.txtObjName.value == "") {
			alert('Please enter a name for this object'); 
			document.frmMODetails.txtObjName.focus();
			return(false);
		}
		if (document.frmMODetails.selObjType.selectedIndex == 0) {
			alert('Please select a type for this object'); 
			document.frmMODetails.selObjType.focus();
			return(false);
		}
		if (document.frmMODetails.selSupportGroup.selectedIndex == 0) {
			alert('Please select a support group for this object'); 
			document.frmMODetails.selSupportGroup.focus();
			return(false);
		}
		if (document.frmMODetails.hdnCustomerID.value == "") {
			alert('Please select a customer for this object'); 
			document.frmMODetails.btnCustomer.focus();
			return(false);
		}
		if (document.frmMODetails.hdnServLocID.value == "") {
			alert('Please select a service location for this object'); 
			document.frmMODetails.btnServiceLocation.focus();
			return(false);
		}
		if (document.frmMODetails.hdnAssetCatalogueID.value == "") {
			if (document.frmMODetails.hdnAssetID.value == "") {
				document.frmMODetails.hdnAssetCatalogueID.value = 1879;
			}else{
				alert('Please select the asset catalogue details (make, model, part number)');
				document.frmMODetails.btnAssetCatalog.focus();
				return(false);
			}
		}

		var strComments = document.frmMODetails.txtComments.value ;
		if (strComments.length > 2000 ) {
			alert('The Comment can be at most 2000 characters.\n\nYou entered ' + strComments.length + ' character(s).');
			document.frmMODetails.txtComments.focus();
			return(false);  
		}
		
		var strMAC = document.frmMODetails.txtMACAddress.value;
		strMAC = strMAC.toUpperCase();
		if ((strMAC != "")&&(strMAC != "N/A")) {
			var re = /\b[0-F]{2}\.[0-F]{2}\.[0-F]{2}\.[0-F]{2}\.[0-F]{2}\.[0-F]{2}/;
			var rv = strMAC.search(re);
			if (rv != 0) {
				if (confirm('The standard format for a MAC Address is:\n\txx.xx.xx.xx.xx.xx\nwhere x is an hexadecimal character\nDo you still want to save it in the current format?')) {
				}else{
					document.frmMODetails.txtMACAddress.focus();
					document.frmMODetails.txtMACAddress.select();
					return(false);
				}
			};
		}else{document.frmMODetails.txtMACAddress.value = "N/A"}

		//Check if the IP Address is filled.
		var strIPAddr = document.frmMODetails.txtIPAddress.value;
		if (strIPAddr == "") {
			alert('IP Address is a mandatory field.  Please re-enter.'); 
			document.frmMODetails.txtIPAddress.focus();
			return(false);
		}
		
		//Check if IP address has the correct format.
		if (strIPAddr != ""){
			var re = /[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}/;
			var rv = strIPAddr.search(re);
			if (rv != 0) {
		        alert('The standard format for IP address is:\n\tnnn.nnn.nnn.nnn\nwhere n is a digit. Please re-enter.'); 
				document.frmMODetails.txtIPAddress.focus();
				return(false);
			}
		}
				
		//check if IP address has 4 segments.
		var IPOctet = strIPAddr.split(".",5);
		if (IPOctet.length != 4){
		   alert('The IP address should have 4 octets.\n\tPlease re-enter.');
			document.frmMODetails.txtIPAddress.focus();
			return(false);
		}

		//Check the value range of each segment.
		for (var i=0; i < IPOctet.length; i++) {
		    if (IPOctet[i] <0 || IPOctet[i] > 255){
		        alert('Invalid value in IP Address Octet #' + (i+1) + '.\nPlease re-enter.'); 
				document.frmMODetails.txtIPAddress.focus();
				return(false);
			}
		}
		
		//disable Save button
		bolSaveRequired = false;

		//submit the form
		document.frmMODetails.txtFrmAction.value = "SAVE";
		document.frmMODetails.submit();
		return(true);

	}else{
		alert('Access denied. Please contact your system administrator.');
		return(false);
	}
}

function btn_onDelete(){
	if ((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete){
		alert('Access denied. Please contact your system administrator.');
		return;
	}
	if (confirm('Do you really want to delete this object?')){
		//submit the form
		document.frmMODetails.txtFrmAction.value = "DELETE";
		document.frmMODetails.submit();
	}
}

function fct_onClone(){
	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create){
		alert('Access denied. Please contact your system administrator.');
		return;
	}
	document.frmMODetails.btnSave.focus();
	if (bolSaveRequired) {
		if	(((intAccessLevel & intConst_Access_Create == intConst_Access_Create) && (document.frmMODetails.txtObjID.value == "")) || ((intAccessLevel & intConst_Access_Update == intConst_Access_Update) && (document.frmMODetails.txtObjID.value != ""))) {
			if (!confirm("There is unsaved data on the screen.\nClick OK below to proceed without saving or click CANCEL to remain on this page.")) {
				return;
			}
		}
	}
	document.location = "manobjdet.asp?ne_id=<%=strNE_ID%>&txtFrmAction=CLONE";
}

function fct_onReset(){
	if(confirm('All changes will be lost. Do you really want to reset the page?')){
		bolSaveRequired = false;
		<%if not bolclone then%>
			document.location = "manobjdet.asp?ne_id=" + document.frmMODetails.txtObjID.value ;
		<%else%>
			document.location = "manobjdet.asp?ne_id=<%=strNE_ID%>&txtFrmAction=CLONE";
		<%end if%>
	}
}

function body_onBeforeUnload(){
	document.frmMODetails.btnSave.focus();
	if (bolSaveRequired) {
		if	(((intAccessLevel & intConst_Access_Create == intConst_Access_Create) && (document.frmMODetails.txtObjID.value == "")) || ((intAccessLevel & intConst_Access_Update == intConst_Access_Update) && (document.frmMODetails.txtObjID.value != ""))) {
			event.returnValue = "There is unsaved data on the screen. To save changes, click CANCEL below then click SAVE on the main form.";
		}
	}
}

function body_onUnload(){
}

function fct_lookupCustomer(){
	if (document.frmMODetails.txtCustomerName.value != "") {
		SetCookie("CustomerName", document.frmMODetails.txtCustomerName.value);
	}
	SetCookie("WinName", 'Popup');
	window.open('SearchFrame.asp?fraSrc=Cust', 'Popup', 'top=50, left=100, height=600, width=800' ) ;
}

function fct_lookupAsset(){
 	if (document.frmMODetails.hdnAssetID.value != "") {
 		SetCookie("AssetID", document.frmMODetails.hdnAssetID.value);
 	}
 	
 	if (document.frmMODetails.txtAssetName.value != "") {
 		SetCookie("AssetName", document.frmMODetails.txtAssetName.value);
 	}
	
	SetCookie("WinName", 'Popup');
	window.open('SearchFrame.asp?fraSrc=Asset', 'Popup', 'top=50, left=100, height=600, width=800' ) ;
}

function btnAssetClear_onclick(){
	document.frmMODetails.hdnAssetID.value = "";
	document.frmMODetails.txtAssetName.value = "";

	document.frmMODetails.hdnCity.value = "";
	document.frmMODetails.hdnStreetName.value = "";
	document.frmMODetails.hdnProvinceCode.value = "";

}

function fct_lookupAssetCatalog(){
	if (document.frmMODetails.hdnAssetCatalogueID.value != "") {
		SetCookie("AssetCatID", document.frmMODetails.hdnAssetCatalogueID.value);
	}
	if (document.frmMODetails.txtAssetMake.value != "") {
		SetCookie("AssetCatMake", document.frmMODetails.txtAssetMake.value);
	}
	if (document.frmMODetails.txtAssetModel.value != "") {
		SetCookie("AssetCatModel", document.frmMODetails.txtAssetModel.value);
	}
	if (document.frmMODetails.txtAssetPartNo.value != "") {
		SetCookie("AssetCatPartNumber", document.frmMODetails.txtAssetPartNo.value);
	}
	SetCookie("WinName", 'Popup');
	window.open('SearchFrame.asp?fraSrc=AssetCatalogue', 'Popup', 'top=50, left=100, height=600, width=800' ) ;
}

function fct_lookupServiceLocation(){

	var strCustomerName = document.frmMODetails.txtCustomerName.value ;
	var strCity = document.frmMODetails.hdnCity.value ;
	var strStreet = document.frmMODetails.hdnStreetName.value ;
	var strProvince = document.frmMODetails.hdnProvinceCode.value ;
	
	if (strCustomerName != ""){
		SetCookie("CustomerName", strCustomerName );
	}
	
	if (strCity != ""){
		SetCookie("CityName", strCity );
	}

	if (strProvince != ""){
		SetCookie("ProvinceName", strProvince );
	}

	if (strStreet != ""){
		SetCookie("Street", strStreet );
	}

	SetCookie("IncludeTelus", "yes");
	SetCookie("WinName", 'Popup');
	
	window.open('SearchFrame.asp?fraSrc=ServLoc','Popup','top=50, left=100, height=600, width=800') ;
}

function fct_clearStatus() {
	window.status = "";
}

function fct_displayStatus(strMessage){
	window.status = strMessage;
	setTimeout('fct_clearStatus()',intConst_MessageDisplay);
}

function body_onLoad(strWinStatus){
	fct_displayStatus(strWinStatus);
	asset_onLoad();
	iFrame_display();
	iFrame2_display();
}


function asset_onLoad(){
var strpart = GetCookie ("MoMakeModelPart");
var apart= strpart.split("?");
var strne_id='<%if not bolClone then response.write strNE_ID end if%>';
//This function is used to populate MO Fields when Navigating from the Asset Screen, and no Managed objects is linked to the asset
//check for cookie
  if ((GetCookie ("MoTacname")!="")&& (strne_id==""))
  {
	//populate empty fields
	document.frmMODetails.txtObjName.value = unescape(GetCookie ("MoTacname"));
	document.frmMODetails.txtAssetName.value = unescape(GetCookie ("MoTacname"));
	document.frmMODetails.hdnAssetCatalogueID.value = GetCookie ("MoAssetCatID");
	document.frmMODetails.hdnCustomerID.value = GetCookie ("MoCustID");
	document.frmMODetails.hdnServLocID.value = GetCookie ("MoServLocID");
	document.frmMODetails.txtAssetMake.value = apart[0];
	document.frmMODetails.txtAssetModel.value = apart[1];
	document.frmMODetails.txtAssetPartNo.value = apart[2];
	document.frmMODetails.txtSerialNumber.value = GetCookie ("MoSerial");
	document.frmMODetails.txtBarcode.value = GetCookie ("MoBarcode");
	document.frmMODetails.txtCustomerName.value = unescape(GetCookie ("MoCustomerName"));
	document.frmMODetails.txtCustomerShortName.value = unescape(GetCookie ("MoCustShortName"));
	document.frmMODetails.txtServLocName.value = unescape(GetCookie ("MoServLocName"));
	document.frmMODetails.hdnAssetID.value = GetCookie ("MoAssetID");
	document.frmMODetails.txtServLocAddress.value = unescape(GetCookie ("MoAddress"));

	//delete cookies
	DeleteCookie("MoTacname") ;
	DeleteCookie("MoAssetCatID");
	DeleteCookie("MoCustID");
	DeleteCookie("MoServLocID");
	DeleteCookie("MoMakeModelPart");
	DeleteCookie("MoSerial");
	DeleteCookie("MoBarcode");
	DeleteCookie("MoCustomerName");
	DeleteCookie("MoCustShortName");
	DeleteCookie("MoServLocName");
	DeleteCookie("MoAssetID");
	DeleteCookie("MoAddress");

  }

}

function btnReferences_onclick(){
var strOwner = 'CRP' ;				// owner name must be in Uppercase
var strTableName = 'NETWORK_ELEMENT' ;		// table name must be in Uppercase
var strRecordID = document.frmMODetails.txtObjID.value ;
var URL ;

	if (strRecordID != ""  ){
		URL ='Dependency.asp?Owner=' + strOwner + '&TableName=' + strTableName + '&RecordID='+ strRecordID   ;
		window.open(URL, 'Popup', 'top=100, left=100, WIDTH=500, HEIGHT=300'  ) ; }
	else
		{alert("No references. This is a new record."); }


}

function qlink_onChange(qlink){
switch (qlink) {
	case "Customer": {
			document.frmMODetails.selQuickLink.selectedIndex=0;
			self.location.href = "CustDetail.asp?CustomerID=" + document.frmMODetails.hdnCustomerID.value;
			break;}
	case "Service Location": {
			document.frmMODetails.selQuickLink.selectedIndex=0;
			self.location.href = "ServLocDetail.asp?ServLocID=" + document.frmMODetails.hdnServLocID.value;
			break;}
	case "Correlation": {
			document.frmMODetails.selQuickLink.selectedIndex=0;
			if (document.frmMODetails.txtObjName.value != '')
			{
				SetCookie("ObjectName", document.frmMODetails.txtObjName.value);
			}
			self.location.href = "searchFrame.asp?fraSrc=Correlation";
			break;}
	case "Asset": {
			document.frmMODetails.selQuickLink.selectedIndex=0;
			if (document.frmMODetails.hdnAssetID.value !="")
				self.location.href = "AssetDetail.asp?asset_id=" + document.frmMODetails.hdnAssetID.value;
			else
				alert('Unable to navigate to Asset null asset id');
			break;}
	default: return;
}
return;
}
</script>

<BODY onLoad="body_onLoad(strWinMessage);asset_onLoad();" onBeforeUnload="body_onBeforeUnload();" onUnload="body_onUnload();">
<FORM name="frmMODetails" action="manobjdet.asp" method="POST" onReset="fct_onReset();">
<INPUT type="hidden" name=txtObjID value="<%if not bolClone and (strNE_ID <> "") then Response.write rsNE("NETWORK_ELEMENT_ID") end if %>" onChange="fct_onChange();">
<INPUT type="hidden" name=hdnAssetID value=<% if strNE_ID <> "" then Response.write """"&rsNE("ASSET_ID")&"""" else Response.Write """""" end if %>>
<INPUT type="hidden" name=hdnAssetCatalogueID value="<%if strNE_ID <> "" then Response.write rsNE("ASSET_CATALOGUE_ID") else Response.Write "1879"%>">
<INPUT type="hidden" name=hdnCustomerID value="<% if strNE_ID <> "" then Response.Write rsNE("CUSTOMER_ID") end if %>">
<INPUT type="hidden" name=hdnServLocID value="<% if strNE_ID <> "" then Response.Write rsNE("SERVICE_LOCATION_ID") end if %>">
<INPUT type="hidden" name=hdnNEUpdateDateTime value="<% if strNE_ID <> "" then Response.Write  rsNE("UPDATE_DATE_TIME")  end if %>">

<INPUT type="hidden" name=hdnCity value="<% if strNE_ID <> "" then Response.write """"&rsNE("MUNICIPALITY_NAME")&"""" else Response.Write """""" end if %>">
<INPUT type="hidden" name=hdnStreetName value="<% if strNE_ID <> "" then Response.write """"&rsNE("STREET")&"""" else Response.Write """""" end if %>">
<INPUT type="hidden" name=hdnProvinceCode value="<% if strNE_ID <> "" then Response.write """"&rsNE("PROVINCE_STATE_CODE")&"""" else Response.Write """""" end if %>">

		
<INPUT type="hidden" name=hdnNameAlias value="">
<INPUT type="hidden" name=txtFrmAction value="">

<table width="100%" cols=4 border=0>
	<thead>
		<tr>
			<td colSpan="3">Managed Object - Details</td>
			<td align="right">
			  	<select name="selQuickLink" size="1" onChange="qlink_onChange(this.value);" <%if strNE_ID = "" then Response.Write "disabled" end if%>>
			  		<option value="">Quickly Goto...</option>
			  		<option value="Customer">Customer</option>
			  		<option value="Service Location">Service Location</option>
			  		<option value="Correlation">Correlation</option>
			  		<option value="Asset">Asset</option>
			  	</select>
			</td>
		</tr>
	</thead>
	<tbody>
	<tr>
		<td width="15%" align=right>Object Name<font color=red>*</font></td>
		<td width="35%"><INPUT name=txtObjName size=30 maxlength=30 value="<% if not bolClone and (strNE_ID <> "") then Response.write rsNE.Fields("NETWORK_ELEMENT_NAME") end if%>" onChange="fct_onChange();"></td>
		<td width="15%" rowSpan="5" valign="top" align="right">Name Alias</td>
		<td width="35%" rowSpan="5" valign="top">
			<iframe id=aifr width=100% height=100 src="" scrolling=yes marginheight=1 marginwidth=1></iframe>
			<br>
			<input type="button" value="Refresh" name="btn_iFrameRefresh" onClick="iFrame_display();" class=button>
			<input type="button" value="New"     name="btn_iFrameAdd"     onClick="btn_iFrmAdd();"    class=button>
			<input type="button" value="Update"  name="btn_iFrameUpdate"  onClick="btn_iFrmUpdate();" class=button>
			<input type="button" value="Delete"  name="btn_iFrameDelete"  onClick="btn_iFrmDelete();" class=button>
		</td>
	</tr>
	<tr>
		<td align=right>Object Description</td>
		<td><INPUT name=txtObjDesc size=40 maxlength=80 value="<% if strNE_ID <> "" then Response.write   rsNE.Fields("NETWORK_ELEMENT_DESC") end if%>" onChange="fct_onChange();"></td>
	</tr>
	<tr>
		<td align=right>Object Type<font color=red>*</font></td>
		<td><SELECT name=selObjType onChange="fct_onChange();"> 
			<OPTION></OPTION>
				<%
				while not rsNET.EOF 
					Response.Write "<OPTION"
					if strNE_ID <> "" then if rsNET("NETWORK_ELEMENT_TYPE_CODE") = rsNE("NETWORK_ELEMENT_TYPE_CODE") then Response.write " selected"
					Response.Write ">" & rsNET("NETWORK_ELEMENT_TYPE_CODE") & "</OPTION>"
					rsNET.MoveNext
				wend
				rsNET.Close
				%>
			</SELECT>
		</td>
	</tr>
	<tr>
		<td align=right>IP Address<font color=red>*</font></td>
		<td><INPUT name=txtIPAddress size=30 maxlength=30 value="<% if strNE_ID <> "" then Response.write   routineHtmlString(rsNE("MANAGED_IP_ADDRESS"))end if %>" onChange="fct_onChange();"></td>
	</tr>
	<tr>
		<td align=right>MAC Address</td>
		<td><INPUT name=txtMACAddress size=30 maxlength=17 value="<% if strNE_ID <> "" then Response.write  routineHtmlString(rsNE("TRUSTED_HOST_MAC_ADDRESS"))end if %>" onChange="fct_onChange();"></td>
	</tr>
	<tr>
		<td align=right>Asset</td>
		<td><INPUT disabled name=txtAssetName size=40 maxlength=80 value="<% if strNE_ID <> "" then Response.write  routineHtmlString(rsNE("ASSET_TAC_NAME")) end if%>" title="To change click on the attached button" onChange="fct_onChange();">
		<INPUT name="btnAsset" type="button" onClick="fct_lookupAsset();fct_onChange();" value="..." class=button title="Click here to edit asset information.">&nbsp;<INPUT name="btnAssetClear" type="button" onClick="btnAssetClear_onclick();" value="X" type=button title="Click here to clear the asset information."></td>
		<td valign="top" align="right" rowSpan="4">Comments</td>
		<td valign="top" rowSpan="4"><textarea style="width=100%" name="txtComments" onChange="fct_onChange();" rows=6><% if strNE_ID <> "" then Response.write  routineHtmlString(rsNE("COMMENTS")) end if%></textarea></td>
	</tr>
	<tr>
		<td align=right>Make/Model/Part #<font color=red>*</font></td>
		<td nowrap>
			<INPUT disabled name=txtAssetMake   size=12 value="<%if (strNE_ID <> "") AND (rsNE("ASSET_MAKE_DESC")    <> "<none>") then Response.write routineHtmlString(rsNE("ASSET_MAKE_DESC")) else Response.Write ""%>" title="To change click on the attached button" onChange="fct_onChange();">
			<INPUT disabled name=txtAssetModel  size=12 value="<%if (strNE_ID <> "") AND (rsNE("ASSET_MODEL_DESC")   <> "<none>") then Response.write routineHtmlString(rsNE("ASSET_MODEL_DESC")) else Response.Write ""%>">
			<INPUT disabled name=txtAssetPartNo size=12 value="<%if (strNE_ID <> "") AND (rsNE("ASSET_PART_NO_DESC") <> "<none>") then Response.write routineHtmlString(rsNE("ASSET_PART_NO_DESC")) else Response.Write ""%>">
			<INPUT name="btnAssetCatalog" type="button" onClick="fct_lookupAssetCatalog();fct_onChange();" value="..." class=button>
		</td>
	<tr>
		<td align=right>Serial</td>
		<td><INPUT name=txtSerialNumber size=40 onChange="fct_onChange();" value="<% if strNE_ID <> "" then Response.write  routineHtmlString(rsNE("SERIAL_NUMBER")) end if%>">
	</tr>
	<tr>
		<td align=right>Barcode</td>
		<td><INPUT name=txtBarcode size=40 onChange="fct_onChange();" value="<% if strNE_ID <> "" then Response.write  routineHtmlString(rsNE("BARCODE")) end if%>">
	</tr>
	<tr>
		<td align=right>Out of Band Dialup</td>
		<td><INPUT name=txtOBDialUp size=30 maxlength=30 value="<%if strNE_ID <> "" then Response.write   routineHtmlString(rsNE("OUT_OF_BAND_DIALUP")) end if%>" onChange="fct_onChange();"></td>
	<tr>
		<td align=right>Customer<font color=red>*</font></td>
		<td><INPUT disabled name=txtCustomerName size=40 maxlength=50 value="<%if strNE_ID <> "" then Response.write  routineHtmlString(rsNE("CUSTOMER_NAME"))end if%>" title="To change click the attached button" onChange="fct_onChange();">
		<INPUT name=btnCustomer type=button  value="..."  onClick="fct_lookupCustomer();fct_onChange();" class=button></td>
		<td align=right>Cust Short Name</td>
		<td><INPUT DISABLED name=txtCustomerShortName size=30 maxlength=15 value="<%if strNE_ID <> "" then Response.write   routineHtmlString(rsNE("CUSTOMER_SHORT_NAME"))end if%>" title="To change click the button attached to the customer name field" onChange="fct_onChange();"></td>
	</tr>
	<tr>
		<td align=right>Service Location<font color=red>*</font></td>
		<td><INPUT disabled name=txtServLocName size=40 value="<% if strNE_ID <> "" then Response.write  routineHtmlString(rsNE("SERVICE_LOCATION_NAME")) end if%>" title="To change click on the attached button" onChange="fct_onChange();">
		<INPUT name=btnServiceLocation type=button value="..." onClick="fct_lookupServiceLocation();fct_onChange();" class=button></td>
		<td align=right>Support Contact Role</td>
		<td><SELECT name=selSupportContactRole onChange="fct_onChange();"> 
			<OPTION></OPTION>
			<%
			while not rsSCR.EOF 
				Response.Write "<OPTION VALUE=""" & rsSCR("REMEDY_CONTACT_ROLE_ID") & """"
				if strNE_ID <> ""  then
					if not IsNull(rsNE("REMEDY_CONTACT_ROLE_ID")) then
						if strNE_ID <> "" then if CLng(rsNE("REMEDY_CONTACT_ROLE_ID")) = CLng(rsSCR("REMEDY_CONTACT_ROLE_ID")) then Response.write " selected"
					end if
				end if
				Response.Write ">" & routineHtmlString(rsSCR("CONTACT_ROLE_NAME")) & "</OPTION>" &vbCrLf
				rsSCR.MoveNext
			wend
			rsSCR.Close
			%>
			</SELECT>
		</td>
	</tr>
	<tr>
		<td align=right>Service Location Address</td>
		<td><TEXTAREA rows=3 style="WIDTH: 100%" id=txtServLocAddress name=txtServLocAddress disabled><%  if strNE_ID <> "" then Response.write  strServLocAddress end if%></TEXTAREA></td>
		<td align=right valign=top>Support Group<font color=red>*</font></td>
		<td valign=top><SELECT name=selSupportGroup onChange="fct_onChange();">
			<OPTION></OPTION>
			<%
			while not rsSG.EOF 
				Response.Write "<OPTION"
				if strNE_ID <> "" then if rsSG("REMEDY_SUPPORT_GROUP_ID") = rsNE("REMEDY_SUPPORT_GROUP_ID") then Response.write " selected"
					Response.Write " VALUE="& rsSG("REMEDY_SUPPORT_GROUP_ID") &">" & routineHtmlString(rsSG("GROUP_NAME")) & "</OPTION>" &vbCrLf
				rsSG.MoveNext
			wend
			rsSG.Close
			%>
			</SELECT>
		</td>
	</tr>
	<tr>
		<td width="50%" colspan="5" rowSpan="12" valign="top">
			<iframe id=aifr2 width=100% height=240 src="" scrolling=yes marginheight=1 marginwidth=1></iframe>
			<br>
			<input type="button" value="Refresh" name="btn_iFrame2Refresh" onClick="iFrame2_display();" class=button>
			<input type="button" value="New"     name="btn_iFrame2Add"     onClick="btn_iFrm2Add();"    class=button>
			<input type="button" value="Clone"   name="btn_iFrame2Clone"   onClick="btn_iFrm2Clone();"  class=button>
			<input type="button" value="Update"  name="btn_iFrame2Update"  onClick="btn_iFrm2Update();" class=button>
			<input type="button" value="Delete"  name="btn_iFrame2Delete"  onClick="btn_iFrm2Delete();" class=button>
		</td>
		<td align=right valign=top nowrap>Repair Priority</td>
		<td valign=top><SELECT name=selRepairPriority onChange="fct_onChange();">
			<OPTION></OPTION>
			<%
			while not rsLYNXrp.EOF 
				Response.Write "<OPTION"
				if strNE_ID <> "" then if CLng(rsNE("LYNX_DEF_SEV_LCODE")) = CLng(rsLYNXrp("LYNX_DEF_SEV_LCODE")) then Response.write " selected"
					Response.Write " VALUE="& rsLYNXrp("LYNX_DEF_SEV_LCODE") &">" & routineHtmlString(rsLYNXrp("LYNX_DEF_SEV_DESC")) & "</OPTION>" &vbCrLf
				rsLYNXrp.MoveNext
			wend
			rsLYNXrp.Close
			%>
			</SELECT>
		</td>
	</tr>

	</tbody>
	<tfoot>
	<tr>
		<td width="100%" colspan="4" align="right">
			<input name=btnReferences type=button style="width: 2.2cm"  value=References  tabindex=13 onclick="return btnReferences_onclick();">&nbsp;&nbsp;
			<INPUT name=btnDelete     type=button style="width: 2cm"    value=Delete      tabindex=13 onClick="btn_onDelete();"                > &nbsp;&nbsp;
			<INPUT name=btnReset      type=button style="width: 2cm"    value=Reset       tabindex=13 onClick="fct_onReset();"                 > &nbsp;&nbsp;
			<INPUT name=btnNew        type=button style="width: 2cm"    value=New         tabindex=13 onClick="fct_NewMO();"                   > &nbsp;&nbsp;
			<INPUT name=btnClone      type=button style="width: 2cm"    value=Clone       tabindex=13 onclick="fct_onClone();"                 > &nbsp;&nbsp;
			<INPUT name=btnSave       type=button style="width: 2cm"    value=Save        tabindex=13 onclick="btn_onSave();"                  > &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		</td>
	</tr>
	</tfoot>
</table>
	<FIELDSET>
	<%if bolClone then strNE_ID = ""%>
	<LEGEND align=right><B>Audit Information</B></LEGEND>
	<Div SIZE=8pt ALIGN=RIGHT>
		Record Status Indicator
		<INPUT align = left name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 18px"disabled value="<% if strNE_ID <> "" then Response.write  rsNE("RECORD_STATUS_IND") end if %>" >&nbsp;&nbsp;&nbsp;
		Create Date&nbsp;<INPUT align = center name=txtCreateDateTime type=text style="HEIGHT: 20px; WIDTH: 150px" disabled value="<% if strNE_ID <> "" then Response.write  rsNE("CREATE_DATE_TIME") end if %>" >&nbsp;
		Created By&nbsp; <INPUT align = right  name=txtCreateRealUser type=text style="HEIGHT: 20px; WIDTH: 200px" disabled value="<% if strNE_ID <> "" then Response.write  rsNE("CREATE_REAL_USERID") end if %>" ><BR>
		Update Date&nbsp;<INPUT align = center name=txtUpdateDateTime type=text style="HEIGHT: 20px; WIDTH: 150px" disabled value="<% if strNE_ID <> "" then Response.write  rsNE("UPDATE_DATE_TIME") end if %>" >
		Updated By&nbsp; <INPUT align = right  name=txtUpdateRealUser type=text style="HEIGHT: 20px; WIDTH: 200px" disabled value="<% if strNE_ID <> "" then Response.write  rsNE("UPDATE_REAL_USERID") end if %>" >
	</DIV>
	</FIELDSET>
</form>
</BODY>
</HTML>
<%
if strNE_ID <> "" then
	rsNE.Close
end if

%>
