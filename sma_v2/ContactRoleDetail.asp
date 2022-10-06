<%@ Language=VBScript %>
<% option explicit %>
<% Response.Buffer = true %>
<%on error resume next%>

<!--
********************************************************************************************
* Page name:	ContactRoleDetail.asp
* Purpose:		To display the detailed information about a customer contact role.
*				Customer chosen via ContactRoleList.asp
* Created by:	Nancy Mooney
* Updated by:	Shawn Myers
********************************************************************************************
        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
       22-Jan-01	 DTy		Increase contact priority from 10 to 30.
       19-Feb-02	 DTy		Provide extra space for email-address which had increased
                                  from 50 to 60 characters.
       03-Oct-07        ACheung		Add Area_of_Reponsibility (50 chars) field to the customer_contact table
********************************************************************************************
-->
<!--#include file="SmaConstants.inc"-->
<!--#include file="smaProcs.inc"-->
<!--#include file="databaseconnect.asp"-->

<%

'************ security ************************************

dim intAccessLevel
intAccessLevel = CInt(CheckLogon(strConst_ContactRole))


if intAccessLevel and intConst_Access_ReadOnly <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to managed object. Please contact your system administrator"
end if

dim lngCustomerContactID
dim strWinLocation
dim strWinMessage
dim strRealUserID

strWinLocation = "ContactRoleDetail.asp?CustomerID="& Request.Form("hdnCustomerID")
strRealUserID = Session("username")

'get the customer contact id
lngCustomerContactID = Request("hdnCustomerContactID")

select case Request("hdnFrmAction")

	case "SAVE"

'check to see if customer exists already in database by checking for a hidden customer contact id
	  if Request.Form("hdnCustomerContactID")  <> "" then  ' it is an existing record so save the changes

		if intAccessLevel and intConst_Access_Update <> intConst_Access_Update then

			DisplayError "REFRESH", strWinLocation, 0, "UPDATE DENIED", "You don't have access to update managed objects. Please contact your system administrator"
		end if

		dim cmdUpdateObj,arole

		set cmdUpdateObj = server.CreateObject("ADODB.Command")
		set cmdUpdateObj.ActiveConnection = objConn
		cmdUpdateObj.CommandType = adCmdStoredProc

		'get the contact_role_detail stored update procedure <schema.package.procedure>
		cmdUpdateObj.CommandText = "sma_sp_userid.spk_sma_cust_inter.sp_cont_role_update"

		'lngCustomerContactID = Request("hdnCustomerContactID")

		'create the associated parameters
		'"arole" is used to split value of selrole at delimiter
		'user id associated with time stamp

		arole = split(Request("selRole"),"¿")

		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_userid", adVarChar, adParamInput, 20,strRealUserID)
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_customer_contact_id", adNumeric, adParamInput,, Clng(Request("hdnCustomerContactID")))'the hidden customer contact id
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_contact_type",adVarChar, adParamInput, 8, arole(0))
        cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_customer_id", adNumeric, adParamInput,, Clng(Request("hdnCustomerID")))
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_contact_id", adNumeric, adParamInput,, Clng(Request("hdnContactID")))
        cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_contact_priority", adNumeric,adParamInput, 2, cint(Request("selPriority")))
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_update_dt", adDBTimeStamp, adParamInput, , CDate(Request("hdnUpdateDateTime")))'date means: update_date_time from contact record
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_AOR", adVarChar, adParamInput, 50, Request("AOR"))

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


		'get the contact_role_detail insert procedure <schema.package.procedure>

		cmdInsertObj.CommandText = "sma_sp_userid.spk_sma_cust_inter.sp_cont_role_insert"

		'create the insert parameters

		arole = split(Request("selRole"),"¿")

		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_userid", adVarChar, adParamInput, 20,strRealUserID)
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_customer_contact_id", adNumeric, adParamOutput,, null)
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_contact_type",adVarChar, adParamInput, 8, arole(0))
        cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_customer_id", adNumeric, adParamInput,, Clng(Request("hdnCustomerID")))
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_contact_id", adNumeric, adParamInput,, Clng(Request("hdnContactID")))
        cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_contact_priority", adNumeric,adParamInput, 2, Request("selPriority"))
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_AOR", adVarChar, adParamInput, 50,Request("AOR"))

		'execute the insert object

		cmdInsertObj.Execute

			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT CREATE NEW OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			else
				lngCustomerContactID = cmdInsertObj.Parameters("p_customer_contact_id").Value
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

			'get the contact_role_detail delete procedure <schema.package.procedure>
			cmdDeleteObj.CommandText = "sma_sp_userid.spk_sma_cust_inter.sp_cont_role_delete"

			'create the delete parameters
			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_customer_contact_id", adNumeric, adParamInput, , clng(lngCustomerContactID))
			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, ,Cdate(Request("UpdateDateTime")))

			'execute the delete object
			cmdDeleteObj.Execute

			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			end if

			lngCustomerContactID = 0
			strWinMessage = "Record deleted successfully."

end select


'********************get Contact Role information ****************************

Dim strSQL, strSelectClause, strFromClause, strWhereClause
Dim rsCustomerContact, rsRole, rsPriority

'get requested customer_contact_id

	if lngCustomerContactID <> 0 then

        'build query

		strSelectClause = "select " &_
					"cc.customer_contact_id, " & _
					"cc.customer_contact_type_lcode, " & _
					"cc.customer_id, " & _
					"cc.contact_id, " & _
					"cc.contact_priority, " & _
					"cct.customer_contact_type_desc, " & _
					"cust.customer_name, " & _
					"cont.contact_name, " & _
					"cont.work_for_customer_id, " & _
					"cont.last_name, " & _
					"cont.first_name, " & _
					"cont.work_number, " & _
					"cont.work_number_ext, " & _
					"cont.cell_number, " & _
					"cont.pager_number, " & _
					"cont.fax_number, " & _
					"cont.email_address, " & _
					"cont.position_title, " & _
					"ad.building_name, " & _
					"ad.street, " & _
					"ad.municipality_name, " & _
					"ad.province_state_lcode, " & _
					"ad.country_lcode, " & _
					"ad.postal_code_zip, " & _
					"cont.userid, " & _
					"work_for.customer_name as work_for_name, " & _
					"to_char(cc.create_date_time,'MON-DD-YYYY HH24:MI:SS') as create_date, " & _
					"sma_sp_userid.spk_sma_library.sf_get_full_username(cc.create_real_userid) as create_real_userid, " & _
					"to_char(cc.update_date_time,'MON-DD-YYYY HH24:MI:SS') as update_date, " & _
					"sma_sp_userid.spk_sma_library.sf_get_full_username(cc.update_real_userid) as update_real_userid, " & _
					"cc.update_date_time as last_update_date_time, " & _
					"cc.record_status_ind, " & _
					"cc.area_of_responsibility "

		strFromClause =	" from crp.customer_contact cc, " &_
					"crp.lcode_customer_contact_type cct, " & _
					"crp.customer cust, " & _
					"crp.contact cont, " & _
					"crp.v_address_consolidated_street ad, " & _
					"crp.customer work_for "

		 strWhereClause = " where " & _
					"cc.customer_contact_id = " & lngCustomerContactID & " and " & _
					"cc.customer_contact_type_lcode = cct.customer_contact_type_lcode and " & _
					"cc.customer_id = cust.customer_id and " & _
					"cc.contact_id = cont.contact_id and " & _
					"cont.address_id = ad.address_id(+) and " & _
					"cont.work_for_customer_id = work_for.customer_id "

		strSQL =  strSelectClause & strFromClause & strWhereClause

		'== show SQL for debugging if necessary

		'Response.Write strSQL


		'get the customer contact recordset

		set rsCustomerContact = Server.CreateObject("ADODB.Recordset")
		rsCustomerContact.CursorLocation = adUseClient
		rsCustomerContact.Open strSQL, objConn
		if err then
			DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
		end if
		if rsCustomerContact.EOF then
			DisplayError "BACK", "", 999, "CANNOT FIND REQUESTED OBJECT", "EOF condition occurred in rsCustomerContact recordset."
		end if
		set rsCustomerContact.ActiveConnection = nothing


		'set the information for the contact information text area (read-only)

		dim strContactInfo, strCustomerName ,strLName, strFName, lngContactID, strWorkFor, strContactName, strPosition, strWExt, strEmail, strBuilding, strStreet, strUserid, strAofR

		strContactName = rsCustomerContact("contact_name")
		strLName = rsCustomerContact("last_name")
		strFName = rsCustomerContact("first_name")
		strCustomerName = rsCustomerContact("customer_name")
		lngContactID = rsCustomerContact("contact_id")
		strWorkFor = rsCustomerContact("work_for_name")
		strPosition = rsCustomerContact("position_title")
		strWExt = rsCustomerContact("work_number_ext")
		strEmail = rsCustomerContact("email_address")
		strBuilding = rsCustomerContact("building_name")
		strStreet = rsCustomerContact("street")
		strAofR = rsCustomerContact("area_of_responsibility")

		'create CPC (City/Province/Country)
		dim strCPC, strCity, strProv, strCountry
		if rsCustomerContact("municipality_name") <> "" then
			strCity = rsCustomerContact("municipality_name") & " "
		end if
		if rsCustomerContact("province_state_lcode") <> "" then
			strProv = rsCustomerContact("province_state_lcode") & " "
		end if
		if rsCustomerContact("country_lcode") <> "" then
			strCountry = rsCustomerContact("country_lcode") & " "
		end if
		if rsCustomerContact("userid") <> "" then
			strUserID = rsCustomerContact("userid") & " "
		end if

		strCPC = strCity & strProv & strCountry

		'Parse out the phone number
		Dim strWPArea,strWPMid,strWPEnd,strWP
		 	strWP = rsCustomerContact("work_number")
		 	strWPArea = mid(strWP,1,3)
		 	strWPMid = mid(strWP,4,3)
		 	strWPEnd = mid(strWP,7,4)
		 	strWP = "(" & strWPArea & ") " & strWPMid & "-" & strWPEnd
		 	If strWP = "() -" then
		 		strWP = ""
		 	End If

		'Parse out the cell phone number
		Dim strCPArea,strCPMid,strCPEnd,strCP
		 	strCP = rsCustomerContact("cell_number")
		 	strCPArea = mid(strCP,1,3)
		 	strCPMid = mid(strCP,4,3)
		 	strCPEnd = mid(strCP,7,4)
		 	strCP = "(" & strCPArea & ") " & strCPMid & "-" & strCPEnd
		 	If strCP = "() -" then
		 		strCP = ""
		 	End If

		'Parse out the pager number
		Dim strPPArea,strPPMid,strPPEnd,strPP
		 	strPP = rsCustomerContact("pager_number")
		 	strPPArea = mid(strPP,1,3)
		 	strPPMid = mid(strPP,4,3)
		 	strPPEnd = mid(strPP,7,4)
		 	strPP = "(" & strPPArea & ") " & strPPMid & "-" & strPPEnd
		 	If strPP = "() -" then
		 		strPP = ""
		 	End If

 		'Parse out the fax number
 		Dim strFPArea,strFPMid,strFPEnd,strFP
		 	strFP = rsCustomerContact("fax_number")
		 	strFPArea = mid(strFP,1,3)
		 	strFPMid = mid(strFP,4,3)
		 	strFPEnd = mid(strFP,7,4)
		 	strFP = "(" & strFPArea & ") " & strFPMid & "-" & strFPEnd
		 	If strFP = "() -" then
		 		strFP = ""
		 	End If

 		'postal code
 		dim strPC
		strPC = rsCustomerContact("postal_code_zip")

		strContactInfo = "Works for:" & chr(9) & strWorkFor & chr(10) & "Position:" & chr(9) & strPosition & chr(10) & "UserID:" & chr(9) & strUserID & chr(10) & "Work # :" & chr(9) & strWP & " Ext: " & strWExt & chr(10) & "Cell # :" & chr(9) & strCP & chr(10) & "Pager # :" & chr(9) & strPP & chr(10) & "Fax # :" & chr(9) & strFP & chr(10)
		if len(strEmail) >40 then
		  strContactInfo = strContactInfo & "Email:" & chr(9) & mid(strEmail, 1, 40) & chr(10) & chr(9) & mid(strEmail, 41)
		else
		  strContactInfo = strContactInfo & "Email:" & chr(9) & strEmail
		end if
		strContactInfo = strContactInfo & chr(10) & "Building:" & chr(9) & strBuilding & chr(10) & "Address:" & chr(9) & strStreet & chr(10) & chr(9) & strCPC & chr(10) & chr(9) & strPC & chr(10)

	end if

   'get list items Note: Priority list is hard-coded, there is no corresponding table in the database.
   'get Role List
	strSQL = "select distinct customer_contact_type_lcode, customer_contact_type_desc" & _
			 " from crp.lcode_customer_contact_type" & _
			 " where record_status_ind = 'A'" & _
			 " order by Upper(customer_contact_type_lcode)"

	set rsRole = Server.CreateObject("ADODB.Recordset")
	rsRole.CursorLocation = adUseClient
	rsRole.Open strSQL, objConn
	if err then
		DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
	end if
	if rsRole.EOF then
		DisplayError "BACK", "", 999, "CANNOT CREATE OBJECT TYPE LIST", "EOF condition occurred in rsRole recorset."
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

	//**********************************Javascript Functions************************

setPageTitle("SMA - Contact Roles");

	//***********************************************

	function fct_selNavigate(){
	//***************************************************************************************************
	// Function:	fct_selNavigate															            *
	//																									*
	// Purpose:		To display the page selected by the user from Quick Navigation drop-down box. The	*
	//				function saves Customer Name and Contact Name in cookes, which may be retrieved     *
	//				by the called page(s).                                    							*
	//																									*
	// Created By:	Nancy Mooney 08/31/2000															    *
	//																									*																				*
	//***************************************************************************************************

	 var strPageName;
	 var lngContactID;
	 var lngCustomerID;

		strPageName = document.frmContactRoleDetail.selNavigate.item(document.frmContactRoleDetail.selNavigate.selectedIndex).value ;

		switch(strPageName){
			case 'Cust':
				document.frmContactRoleDetail.selNavigate.selectedIndex=0;
				lngCustomerID = document.frmContactRoleDetail.hdnCustomerID.value ;
				self.location.href = "CustDetail.asp?CustomerID=" + lngCustomerID;
				break;
			case 'Contact':
				document.frmContactRoleDetail.selNavigate.selectedIndex=0;
				lngContactID = document.frmContactRoleDetail.hdnContactID.value ;
				self.location.href = "ContactDetail.asp?ContactID=" + lngContactID;
				break;
			case 'DEFAULT':
				//do nothing
		}
	}

	//**********************************************************************

	function fct_NewContactRole(){

		if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create)
		{
			alert('Access denied. Please contact your system administrator.');
			return;
		}
		self.document.location.href = "ContactRoleDetail.asp?CustomerContactID=0";
	}

	//**********************************************************************

	function fct_onChange(){
		bolNeedToSave = true;
	}

	//*******************************************************************

	function fct_OnSave(){

		if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create)
		{
			alert('Access denied. Please contact your system administrator.');
			return (false);
		}
		else
		{
			if (document.frmContactRoleDetail.selRole.value == "" )
			{
				alert('Please select a customer role');
				document.frmContactRoleDetail.selRole.focus();
				return(false);
			}
			if (document.frmContactRoleDetail.txtCustomerName.value == "" )
			{
				alert('Please select a customer');
				document.frmContactRoleDetail.btnCustomerLookup.focus();
				return(false);
			}
			if (document.frmContactRoleDetail.txtContactName.value == "" )
			{
				alert('Please select a contact');
				document.frmContactRoleDetail.btnContactLookup.focus();
				return(false);
			}
			else
			{
				document.frmContactRoleDetail.hdnFrmAction.value = "SAVE";
				bolNeedToSave = false;
				document.frmContactRoleDetail.submit();
				return(true);
			}
		}
    }

	//*************************************************************

	function fct_onDelete() {

		if ((intAccessLevel & intConst_Access_Delete)!= intConst_Access_Delete)
		{
			alert('Access denied. Please contact your system administrator.');
			return;
		}

		var lngCustomerContactID = document.frmContactRoleDetail.hdnCustomerContactID.value ;
		var strUpdateDateTime = document.frmContactRoleDetail.hdnUpdateDateTime.value ;

		if (confirm('Do you really want to delete this object?'))
		{
			document.location = "ContactRoleDetail.asp?hdnFrmAction=DELETE&hdnCustomerContactID="+lngCustomerContactID+"&UpdateDateTime="+strUpdateDateTime ;
		}
	}

	//*************************************************************

	function fct_onReset() {
		if(confirm('All changes will be lost. Do you really want to reset the page?')){
			bolNeedToSave = false ;
			document.location = 'ContactRoleDetail.asp?hdnCustomerContactID='+ "<%=lngCustomerContactID%>" ;
		}
	}

	//*************************************************************

	function fct_onBeforeUnload()
	{
		document.frmContactRoleDetail.btnSave.focus();

		if ((intAccessLevel & intConst_Access_Update) == intConst_Access_Update || (intAccessLevel & intConst_Access_Create) == intConst_Access_Create)
		{
			if (bolNeedToSave == true)
			{
				event.returnValue = "There is unsaved data on the screen. To save changes, click CANCEL below then click SAVE on the main form.";
			}
		}

	}

	//*************************************************************

	function btnCustomerLookup_onclick(CustService)
	{
		if (document.frmContactRoleDetail.txtCustomerName.value != ""){
			 SetCookie("CustomerName", document.frmContactRoleDetail.txtCustomerName.value);
		}
		SetCookie("WinName", 'Popup');
		SetCookie("ServiceEnd", CustService);
		bolNeedToSave = true ;
		window.open('SearchFrame.asp?fraSrc=Cust', 'Popup', 'top=50, left=180, height=600, width=750' ) ;
		//enable Save button - may not need it but onChange event suck (is not fired when use lookup)
		document.frmContactRoleDetail.btnSave.disabled = false;
	}

	//*************************************************************

	function btnContactLookup_onclick(){

		if (document.frmContactRoleDetail.txtCustomerName.value != "") {
			SetCookie("WorkFor", document.frmContactRoleDetail.txtCustomerName.value);
		}
		if (document.frmContactRoleDetail.txtLName.value != ""){
			 SetCookie("LName", document.frmContactRoleDetail.txtLName.value);
		}
		if (document.frmContactRoleDetail.txtFName.value != ""){
			 SetCookie("FName", document.frmContactRoleDetail.txtFName.value);
		}
		SetCookie("WinName", 'Popup');
		bolNeedToSave = true ;
		window.open('SearchFrame.asp?fraSrc=Contact', 'Popup', 'top=50, left=130, height=600, width=870' ) ;
		//enable Save button - may not need it but onChange event suck (is not fired when use lookup)
		document.frmContactRoleDetail.btnSave.disabled = false;
	}

	//*************************************************************

	function fct_clearStatus() {
		window.status = "";
	}

	//*************************************************************

	function fct_onChangeRole() {

		var strWhole;
		var strRoleDesc, intStart, intIndex;

		intIndex = document.frmContactRoleDetail.selRole.selectedIndex;
		strWhole = document.frmContactRoleDetail.selRole.options[intIndex].value;
 		intStart = strWhole.indexOf('<%=strDelimiter%>');
		document.frmContactRoleDetail.txtRoleDesc.value = strWhole.substr(intStart+1);
		fct_onChange();
	}

	//*************************************************************

	function fct_DisplayStatus(strWindowStatus){

	window.status=strWindowStatus;
	setTimeout('fct_clearStatus()', '<%=intConst_MessageDisplay%>');

    }

    //*************************************************************

    function btnReferences_onclick() {

		var strOwner = 'CRP' ;
		var strTableName = 'Customer_Contact' ;
		var strRecordID = document.frmContactRoleDetail.hdnCustomerContactID.value ;
		var URL ;

		if ( lngCustomerContactID = 0)
		{
			alert("No references. This is a new record.");
		}
		else
		{
			URL ='Dependency.asp?Owner=' + strOwner + '&TableName=' + strTableName + '&RecordID='+ strRecordID   ;
			window.open(URL, 'Popup', 'top=100, left=100, WIDTH=500, HEIGHT=300'  ) ;
		}
	}
	//*************************************************************
	//-->
	</SCRIPT>

</HEAD>
<BODY onLoad="fct_DisplayStatus(strWinMessage);" onbeforeunload="fct_onBeforeUnload();">
<FORM name=frmContactRoleDetail action="ContactRoleDetail.asp"  method="POST">

	<!-- hidden variables -->

	<INPUT name=hdnCustomerContactID type=hidden value= <%if lngCustomerContactID <> 0 then Response.Write rsCustomerContact("customer_contact_id") else Response.Write """""" end if%>>
	<INPUT name=hdnCustomerID type=hidden value=<%if lngCustomerContactID <> 0 then Response.Write """"&rsCustomerContact("customer_id")&"""" else Response.Write """""" end if%>>
	<INPUT name=hdnContactID type=hidden value=<%if lngCustomerContactID <> 0 then Response.Write lngContactID else Response.Write """""" end if%>>
    <INPUT name=txtFName type=hidden value="<%if lngCustomerContactID <> 0 then Response.Write routineHTMLString(strFName) else Response.Write null end if%>">
	<INPUT name=txtLName type=hidden value="<%if lngCustomerContactID <> 0 then Response.Write routineHTMLString(strLName) else Response.Write null end if%>">
	<INPUT name=hdnUpdateDateTime type=hidden value=<%if lngCustomerContactID <> 0 then  Response.Write """"&rsCustomerContact("last_update_date_time")&"""" else Response.Write """""" end if%>>
	<INPUT name=hdnAreaofResponsibility type=hidden value="<%if lngCustomerContactID <> 0 then Response.Write routineHTMLString(strAofR) else Response.Write null end if%>">
    <INPUT id=hdnFrmAction name=hdnFrmAction type=hidden value= "">
    <!-- This is here since required for contact lookup to work properly-->
    <INPUT id=txtCustomerShortName name=txtCustomerShortName type=hidden value= "">

	<!-- user interface -->
	<TABLE border=0 width=100%>
	<thead>
		<tr>
			<td colspan=3 align=left>Contact Role Detail</td>
			<td align=right><SELECT <%if lngCustomerContactID = 0 then Response.Write "disabled" end if%> align=right valign=top name=selNavigate LANGUAGE=javascript onchange="fct_selNavigate();" tabindex=10 >
				<OPTION value="DEFAULT">Quickly Goto ...</OPTION>
				<OPTION value="Contact" >Contact</OPTION>
				<OPTION value="Cust" >Customer</OPTION></SELECT>
			</td>
	</thead>
	<TR>
		<TD align=right width=25%>Customer<font color=red>*</font></TD>
		<TD align=left width=50% colspan=2>
			<input name=txtCustomerName type=text disabled size=50 maxlength=50 value="<%if lngCustomerContactID <> 0 then Response.Write routineHTMLString(strCustomerName) else Response.Write null end if%>" onChange="fct_onChange();">
			<INPUT align=right type="button"  name=btnCustomerLookup   value="..." onclick="return btnCustomerLookup_onclick('C')" tabindex=1>
		</td>
		<td width=25% >&nbsp;</td>
	</TR>
	<TR>
		<TD align=right width=20%>Role<font color=red>*</font></TD>
			<TD align=left width=60% colspan=3>
				<SELECT name="selRole" tabindex=2 onChange="fct_onChangeRole();" tabindex=2>
					<option></option>
					<%
						dim strRoleDesc
						while not rsRole.EOF
							Response.write "<OPTION"
							if lngCustomerContactID <> 0 then
								if rsRole("customer_contact_type_lcode")= rsCustomerContact("customer_contact_type_lcode") then
									Response.write " selected"
									strRoleDesc = rsRole(1)
								end if
							end if
							Response.write " value=""" & rsRole(0)& strDelimiter & rsRole(1) & """>" & Ucase(routineHtmlString(rsRole(0))) & "</option>" & vbCrLf
							rsRole.MoveNext
						wend
						rsRole.Close
					%>
				</SELECT>
				<input type=text name=txtRoleDesc disabled size=50 maxlength=50 value="<%=strRoleDesc%>">
			</TD>
	</TR>
	<tr>
		<td width=25% align=right>Priority<font color=red>*</td>
		<TD align=left width =75% colspan=3>
			<select name=selPriority tabindex=3 onChange="fct_onChange();" tabindex=3>
				<%
				dim cnt
				for cnt = 1 to 99
					Response.Write "<OPTION"
					if lngCustomerContactID <> 0 then
						if cInt(rsCustomerContact("contact_priority"))= cInt(cnt) then
							Response.Write " selected "
						end if
					end if
					Response.Write " value=" & cnt & ">" & cnt & "</option>"&vbCrLf
				next
				%>
			</select>
		</TD>
	</tr>
		<TD align=right width=25%>Contact<font color=red>*</font></TD>
		<TD align=left width=50% colspan=2>
			<INPUT name=txtContactName disabled size=50 maxlength=50 onChange="fct_onChange();" value="<%if lngCustomerContactID <> 0 then Response.Write routineHTMLString(strContactName) else Response.Write null end if%>" >
			<INPUT name=btnContactLookup type=button value=... LANGUAGE=javascript onclick="btnContactLookup_onclick()" tabindex=4>
		</td>
		<td width=25%>&nbsp;</td>
	</tr>
	<TR>
		<TD align=right nowrap>Area of Responsibility</TD>
		<td><INPUT name=AOR size=50 maxlength=50 value="<% if lngCustomerContactID <> 0 then Response.Write routineHTMLString(strAofR) else Response.Write null end if%>"</td>
		<TD></TD>
		<TD></TD>
		<TD></TD>
	</TR>
        <tr>
		<td width=25%>&nbsp;</td>
		<td width=25%>&nbsp;</td>
		<td width=25%>&nbsp;</td>
		<td width=25%>&nbsp;</td>
	</tr><tr>
		<td align=right valign=top width=25%>Contact Information </td>
		<td align=left width=50% colspan=2 disabled><textarea name=txtContactInfo cols=85 style="HEIGHT: 200px"><%if lngCustomerContactID <> 0 then Response.Write routineHTMLString(strContactInfo) else Response.Write null end if%></textarea></td>
	</tr>

	<tfoot>
	<tr>
		<td width="100%" colspan="4" align="right">
			<INPUT name=btnReferences tabindex=5 type=button   style="width: 2.2cm" value=References  onclick="return btnReferences_onclick()">&nbsp;&nbsp;
			<INPUT name="btnDelete"   tabindex=6 type="button" style="width: 2cm"   value="Delete"    onClick="return fct_onDelete();">
			<INPUT name="btnReset"    tabindex=7 type="button" style="width: 2cm"   value="Reset"     onClick= "fct_onReset();">&nbsp;&nbsp;
			<INPUT name="btnNew"      tabindex=8 type="button" style="width: 2cm"   value="New"       onClick="return fct_NewContactRole();">&nbsp;&nbsp;
			<INPUT name="btnSave"     tabindex=9 type="button" style="width: 2cm"   value="Save"      onClick="return fct_OnSave();">
		</td>
	</tr>
	</tfoot>
</TABLE>
	<FIELDSET>
	<LEGEND align=right><B>Audit Information</B></LEGEND>
	<Div SIZE=8pt ALIGN=RIGHT>
		Record Status Indicator
		<INPUT align=left name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 18px"disabled value=<%if lngCustomerContactID <> 0 then Response.Write """"&rsCustomerContact("record_status_ind")&"""" else Response.Write """""" end if%> >&nbsp;&nbsp;&nbsp;
		Create Date
		<INPUT align=center name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 150px"disabled value=<%if lngCustomerContactID <> 0 then Response.Write """"&rsCustomerContact("create_date")&"""" else Response.Write """""" end if%>>&nbsp;
		Created By
		<INPUT align=right name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 200px"disabled value=<%if lngCustomerContactID <> 0 then Response.Write """"&rsCustomerContact("create_real_userid")&"""" else Response.Write """""" end if%> ><BR>
		Update Date
		<INPUT align=center name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 150px"disabled value=<%if lngCustomerContactID <> 0 then Response.Write """"&rsCustomerContact("update_date")&"""" else Response.Write """""" end if%> >
		Updated By
		<INPUT align=right name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 200px"disabled value=<%if lngCustomerContactID <> 0 then Response.Write """"&rsCustomerContact("update_real_userid")&"""" else Response.Write """""" end if%> >
	</DIV>
	</FIELDSET>
</FORM>



<%
    'release the active connection
	set rsRole.ActiveConnection = nothing

	'close the connection
	objConn.close
	set objConn = nothing


%>


</BODY>
</HTML>
