
<%@ Language=VBScript %>
<% option explicit %>
<% 'on error resume next
' Response.Buffer = true %>

<!--#include file="smaConstants.inc"-->
<!--#include file="smaProcs.inc"-->
<!--#include file="databaseconnect.asp" -->
<!--
****************************************************************************************************
* Page name:	AddressDetail.asp
* Purpose:		To display the detailed information about a Service Location.
*				Customer chosen via ServLocList.asp
*
* In Param:		This page reads Address ID from a query string.
*
* Out Param:	Sometimes this Page writes following cookeis
*				Cookie - AddressID
*				Cookie - CustomerName
*				Cookie - WinName
*
* Created by:	Sara Sangha	08/11/2000
*
******************************************************************************************************
        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
       24-Feb-02	  DTy		Simplified Detail Screen.
       02-May-08	ACheung		1. house_no is now varchar2(8);
					2. limit house_no to 6 characters from SMA2 screen
					3. amalgamate house_no and house_no suufix into one field
					4. reintroduce txtSuffix as a hidden field;
					5. pass null onto txtSuffix so that SP will always be entering null without SP errors

***************************************************************************************************
-->
<%

'check user's rights
dim  intAccessLevel, strSimple, CustomerID, CustomerName
Dim  logAddressID, datUpdateDateTime, strWinMessage
dim	 objRs, strSQL, strWhereClause, strServLocAddress
dim  strBillingFlag, strPrimaryFlag, strMailingFlag
dim  strWinLocation, strRealUserID, strReadOnly

intAccessLevel = CInt(CheckLogon(strConst_Address))
if (intAccessLevel and intConst_Access_ReadOnly) <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to Address. Please contact your system administrator"
end if

logAddressID = Request("AddressID")
datUpdateDateTime = Request("UpdateDateTime")
strSimple = Request.Cookies("strSimple")
CustomerID = Request.Cookies("CustomerID")
CustomerName = Request.Cookies("txtCustomerName")

'strWinLocation = "AddressDetail.asp?AddressID="&Request.Form("hdnAddressID")
strRealUserID = Session("username")

if err then
	'unexpected error
	DisplayError "BACK", "", 0, "UNEXPECTED ERROR", "Close alias window to return to managed objects form."
end if

select case Request("hdnFrmAction")
	case "SAVE"

		if Lcase(Request("chkPrimary")) = "on" then
			strPrimaryFlag = "Y"
		else
			strPrimaryFlag = "N"
		end if

		if lcase(Request("chkMailing")) = "on" then
			strMailingFlag = "Y"
		else
			strMailingFlag = "N"
		end if

		if lcase(Request("chkBilling")) = "on" then
			strBillingFlag = "Y"
		else
			strBillingFlag = "N"
		end if

	  if Request.Form("hdnAddressID")  <> "" then  ' it is an existing record so save the changes

		if (intAccessLevel and intConst_Access_Update) <> intConst_Access_Update then
				DisplayError "BACK", "", 0, "UPDATE DENIED", "You don't have access to update an address. Please contact your system administrator"
		end if

		dim cmdUpdateObj
		set cmdUpdateObj = server.CreateObject("ADODB.Command")
		set cmdUpdateObj.ActiveConnection = objConn
		cmdUpdateObj.CommandType = adCmdStoredProc
		cmdUpdateObj.CommandText = "sma_sp_userid.spk_sma_cust_inter.sp_address_update"

		logAddressID = Request("hdnAddressID")

		'create params
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_userid", adVarChar, adParamInput, 20, strRealUserID) 							'number(9)		means: Address Id
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_address_id", adNumeric, adParamInput,, Clng(Request("hdnAddressID"))) 			'number(9)		means: Address Id
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_customer_id", adNumeric, adParamInput,, Clng(Request("hdnCustomerID")))			'number(9)		means: Customer ID
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_billing_address", adChar, adParamInput , 1, strBillingFlag)						'varchar2(1)	means: Billing Address Flag
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_primary_address", adChar, adParamInput, 1, strPrimaryFlag)						'varchar2(1)	means: Primary Address Flag
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_mailing_address", adChar, adParamInput, 1, strMailingFlag)						'varchar2(1)	means: Mailing Address Flag
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_municipality_name", adVarChar,adParamInput, 50, Request("hdnCity"))				'varchar2(50)	means: Municipality name
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_country", adChar, adParamInput, 2, Request("hdnCountryCode"))					'varchar2(2)	means: Country Code
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_last_update", adDBTimeStamp, adParamInput, , CDate(Request("hdnUpdateDateTime")))'date			means: update_date_time from address record
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_province", adChar, adParamInput, 2, Request("hdnProvinceCode"))				'varchar2(2)	means: Province Code
		if Request("txtPostal") <> "" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_postal_code", adVarChar, adParamInput, 15, Request("txtPostal"))				'varchar2(15)	means: Postal Code
		else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_postal_code", adVarChar, adParamInput, 15, null)								'varchar2(15)	means: Postal Code
		end if

		if Request("txtBuildingName") <> "" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_building_name", adVarChar,adParamInput, 30, Request("txtBuildingName"))		'varchar2(30)	means: Building name
		else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_building_name", adVarChar,adParamInput, 30, null)
		end if


		if Request("txtApartment") <> "" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_apartment_no", adVarChar,adParamInput, 5, Request("txtApartment"))			'varchar2(5)	means: Apartment name
		else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_apartment_no", adVarChar,adParamInput, 5, null)
		end if


		if Request("txtHouse") <> "" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_house_no", adVarChar, adParamInput, 8, Request("txtHouse"))				'Number(9)		means: House Number
		else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_house_no", adVarChar, adParamInput, 8, NULL)
		end if

		if Request("txtSuffix") <> "" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_house_number_suff", adChar,adParamInput, 1, UCase(Request("txtSuffix")))			'varchar2(1)	means: House Number suffix
		else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_house_number_suff", adChar,adParamInput, 1, null)
		end if

		if Request("selVector") <> "" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_street_vect", adVarChar,adParamInput, 2, ucase(Request("selVector")))		'varchar2(2)	means: Street Vector
		else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_street_vect", adVarChar,adParamInput, 2, null)
		end if

		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_street", adVarChar,adParamInput, 75, Request("txtStreet"))						'varchar2(75)	means: Long Street Name


		'call the insert stored proc
  			'cmdUpdateObj.Parameters.Refresh

  		'	dim objparm
  		'	for each objparm in cmdUpdateObj.Parameters
  		'	  Response.Write "<b>" & objparm.name & "</b>"
  		'	  Response.Write " has size:  " & objparm.Size & " "
  		'	  Response.Write " and value:  " & objparm.value & " "
  		'	 Response.Write " and datatype:  " & objparm.Type & "<br> "
  		'  next

  		'   Response.Write "<b> count = " & cmdUpdateObj.Parameters.count & "<br>"
  		'	dim nx
  		'	 for nx=0 to cmdUpdateObj.Parameters.count-1
  		'	   Response.Write " parm value= " & cmdUpdateObj.Parameters.Item(nx).Value  & "<br>"
  		'	  next

		on error resume next
		cmdUpdateObj.Execute
			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT UPDATE Address", objConn.Errors(0).Description
				objConn.Errors.Clear
			end if
			strWinMessage = "Record saved successfully. You can now see the changes you made."

	  else 'create a new record
			if (intAccessLevel and intConst_Access_Create) <> intConst_Access_Create then
			strWinLocation = "CustServDetail.asp?CustServID=0"
			DisplayError "BACK", "", 0, "INSERT DENIED", "You don't have access to create an address. Please contact your system administrator"
		end if

		dim cmdInsertObj
		set cmdInsertObj = server.CreateObject("ADODB.Command")
		set cmdInsertObj.ActiveConnection = objConn
		cmdInsertObj.CommandType = adCmdStoredProc
		cmdInsertObj.CommandText = "sma_sp_userid.spk_sma_cust_inter.sp_address_insert"

		'create params
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_userid", adVarChar, adParamInput, 20, strRealUserID) 							'number(9)		means: Address Id
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_address_id", adNumeric, adParamOutput,,null)				 					'number(9)		means: Address Id
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_customer_id", adNumeric, adParamInput,, Clng(Request("hdnCustomerID")))			'number(9)		means: Customer ID
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_billing_address", adChar, adParamInput , 1, strBillingFlag)						'varchar2(1)	means: Billing Address Flag
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_primary_address", adChar, adParamInput, 1, strPrimaryFlag)						'varchar2(1)	means: Primary Address Flag
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_mailing_address", adChar, adParamInput, 1, strMailingFlag)						'varchar2(1)	means: Mailing Address Flag
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_municipality_name", adVarChar,adParamInput, 50, Request("hdnCity"))				'varchar2(50)	means: Municipality name
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_country", adChar, adParamInput, 2, Request("hdnCountryCode"))					'varchar2(2)	means: Country Code
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_province", adChar,adParamInput, 2, Request("hdnProvinceCode"))				'varchar2(2)	means: Province Code
		if Request("txtPostal") <> "" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_postal_code", adVarChar, adParamInput, 15, Request("txtPostal"))				'varchar2(15)	means: Postal Code
		else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_postal_code", adVarChar, adParamInput, 15, null)
		end if

		if Request("txtBuildingName") <> "" then

			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_building_name", adVarChar,adParamInput, 30, Request("txtBuildingName"))		'varchar2(30)	means: Building name
		else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_building_name", adVarChar,adParamInput, 30, null)
		end if


		if Request("txtApartment") <> "" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_apartment_no", adVarChar,adParamInput, 5, Request("txtApartment"))			'varchar2(5)	means: Apartment name
		else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_apartment_no", adVarChar,adParamInput, 5, null)
		end if


		if Request("txtHouse") <> "" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_house_no", adVarChar, adParamInput, 8, Request("txtHouse"))				'Number(9)		means: House Number
		else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_house_no", adVarChar, adParamInput, 8, NULL)
		end if

		if Request("txtSuffix") <> "" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_house_number_suff", adChar,adParamInput, 1, Request("txtSuffix"))			'varchar2(1)	means: House Number suffix
		else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_house_number_suff", adChar,adParamInput, 1, null)
		end if

		if Request("selVector") <> "" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_street_vect", adVarChar,adParamInput, 2, ucase(Request("selVector")))		'varchar2(2)	means: Street Vector
		else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_street_vect", adVarChar,adParamInput, 2, null)
		end if

		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_street", adVarChar,adParamInput, 75, Request("txtStreet"))						'varchar2(75)	means: Long Street Name


		'call the insert stored proc
  		'	cmdInsertObj.Parameters.Refresh

  		'dim objparm
  		'	for each objparm in cmdInsertObj.Parameters
  		'	  Response.Write "<b>" & objparm.name & "</b>"
  		'	  Response.Write " has size:  " & objparm.Size & " "
  		'	  Response.Write " and value:  " & objparm.value & " "
  		'	  Response.Write " and datatype:  " & objparm.Type & "<br> "
  		'  next

  		'  Response.Write "<b> count = " & cmdInsertObj.Parameters.count & "<br>"
  		'	dim nx
  		'	 for nx=0 to cmdInsertObj.Parameters.count-1
  		'	   Response.Write  " parm " & nx + 1 &  " value= " & cmdInsertObj.Parameters.Item(nx).Value  & "<br>"
  		 '   next
		on error resume next
		cmdInsertObj.Execute

			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT CREATE NEW ADDRESS", objConn.Errors(0).Description
				objConn.Errors.Clear
			else
				logAddressID = cmdInsertObj.Parameters("p_address_id").Value
			end if
			strWinMessage = "Record created successfully. You can now see the new record."

	  end if
	case "DELETE"
			if (intAccessLevel and intConst_Access_Delete) <> intConst_Access_Delete then
				DisplayError "BACK", "", 0, "DELETE DENIED", "You don't have access to delete Address. Please contact your system administrator"
			end if
			dim cmdDeleteObj
			set cmdDeleteObj = server.CreateObject("ADODB.Command")
			set cmdDeleteObj.ActiveConnection = objConn
			cmdDeleteObj.CommandType = adCmdStoredProc
			cmdDeleteObj.CommandText = "sma_sp_userid.spk_sma_cust_inter.sp_address_delete"
			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_address_id", adNumeric, adParamInput, , clng(logAddressID))					'number(9)
			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_last_update", adDBTimeStamp, adParamInput, ,Cdate(datUpdateDateTime))		'Date
            cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("P_USER_ID", adVarChar , adParamInput, 30, strRealUserID)

			'Response.Write "<b> count = " & cmdDeleteObj.Parameters.count & "<br>"
  			'dim nx
  			' for nx=0 to cmdDeleteObj.Parameters.count-1
  			'   Response.Write  " parm " & nx + 1 &  " value= " & cmdDeleteObj.Parameters.Item(nx).Value  & "<br>"
  			'  next
  			  on error resume next
			cmdDeleteObj.Execute
			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE ADDRESS", objConn.Errors(0).Description
				objConn.Errors.Clear
			end if

			logAddressID = 0
			strWinMessage = "Record deleted successfully."

end select

if logAddressID <> 0  then

		strSQL = "select c.customer_id, " &_
					"c.customer_name, " &_
	   				"a.address_id, " &_
	   				"a.building_name, " &_
	   				"a.apartment_number, " &_
	   				"a.house_number, " &_
	   				"a.house_number_suffix, " &_
	   				"a.street_vector, " &_
	   				"a.long_street_name, " &_
	   				"a.municipality_name, " &_
	   				"a.postal_code_zip, " &_
	   				"a.province_state_lcode, " &_
	   				"t1.province_state_name, " &_
	   				"a.country_lcode, " &_
					"t2.country_desc, " &_
	   				"a.billing_address_flag as billing, " &_
	   				"a.mailing_address_flag as mailing, " &_
	   				"a.primary_address_flag as primary, " &_
	   				"to_char(a.create_date_time, 'MON-DD-YYYY HH24:MI:SS') as create_date, " & _
					"sma_sp_userid.spk_sma_library.sf_get_full_username(a.create_real_userid) as create_real_userid, " & _
					"to_char(a.update_date_time, 'MON-DD-YYYY HH24:MI:SS') as update_date, " & _
					"sma_sp_userid.spk_sma_library.sf_get_full_username(a.update_real_userid) as update_real_userid, " & _
					"a.update_date_time as last_update_date_time, " & _
					"a.record_status_ind " & _
				"from crp.address a, " &_
					" crp.customer c, " &_
					" crp.lcode_province_state t1, " &_
					" crp.lcode_country t2 " &_
				"where a.customer_id = c.customer_id  " &_
				"and	  a.province_state_lcode = t1.province_state_lcode " &_
				"and	  a.country_lcode = t2.country_lcode  " &_
				"and      t1.country_lcode = t2.country_lcode " &_
				"and a.address_id = " & logAddressID

	'Response.Write (strSQL)

	set objRs = server.CreateObject("ADODB.Recordset")
		objRs.Open strSQL, objConn
end if

%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" type="text/css" href="stylesheets/styles.css">
<script type="text/javascript" SRC="AccessLevels.js"></script>
<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>

<SCRIPT LANGUAGE=javascript>
<!--

var strWinMessage = "<%=strWinMessage%>";
var intAccessLevel = "<%=intAccessLevel%>" ;
var bolNeedToSave = false ;

//************************************** Java Functions *************************************
//set section title
setPageTitle("SMA - Address");

function btnCustomerLookup_onclick(CustService)
{
//***************************************************************************************************
// Function:	btnCustomerLookup_onclick															*
//																									*
// Purpose:		To display Customer Search page with pre-populated customer name and to	indicate	*
//				that the search page is displayed in a popup window. (Note: search pages behave		*
//				differently when displayed in popup windows verses when displayed in the base window)*
//																									*
// Created By:	Sara Sangha Aug. 25th, 2000															*
//																									*
// Updated By:																						*
//***************************************************************************************************

	var strCustomerName ;
	strCustomerName = window.frmAddressDetail.txtCustomerName.value ;


	if (strCustomerName != "" ) {SetCookie("CustomerName", strCustomerName) ; }
	SetCookie("ServiceEnd", CustService);
	SetCookie("WinName", 'Popup');
	bolNeedToSave = true ;
	window.open('SearchFrame.asp?fraSrc=Cust', 'Popup', 'top=50, left=100, WIDTH=800, HEIGHT=600'  ) ;

}


function selNavigate_onchange(){
//**********************************************************************************************
// Function: selNavigate_onchange
//
// Purpose:		To display the page selected by the user from Quick Navigation drop-down box.	*
//																								*
// Created By:	Sara Sangha	Aug. 25th, 2000														*
//																								*
// Updated By:																					*
//***********************************************************************************************

 var strPageName = document.frmAddressDetail.selNavigate.item(document.frmAddressDetail.selNavigate.selectedIndex).value ;
 var strCustomerID = document.frmAddressDetail.hdnCustomerID.value ;
 var strCustomerName = document.frmAddressDetail.txtCustomerName.value ;

	switch ( strPageName ) {

	case 'Asset':
		document.frmAddressDetail.selNavigate.selectedIndex=0;
		if (strCustomerName != "") {SetCookie("CustomerName", strCustomerName)};
		self.location.href = "SearchFrame.asp?fraSrc=" + strPageName  ;
		break;

	case 'Cust' :
		document.frmAddressDetail.selNavigate.selectedIndex=0;
		self.location.href  = 'CustDetail.asp?CustomerID=' + strCustomerID ;
		break ;

	case 'DEFAULT' :
		// do nothing ;
	}
}

function btnDelete_onclick() {
//**********************************************************************************************
// Function:	btnDelte_onclick
//
// Purpose:		To delete the current record. The page is submitted with hdnFrmAction as DELETE.
//
// Created By:
//
// Updated By:
//***********************************************************************************************

var logAddressID = document.frmAddressDetail.hdnAddressID.value ;
var strUpdateDateTime = document.frmAddressDetail.hdnUpdateDateTime.value ;

	if ((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) {
		alert('Access denied. Please contact your system administrator.');
	return;
   }

	if (confirm('Do you really want to delete this object?')){
		document.location = "AddressDetail.asp?hdnFrmAction=DELETE&AddressID="+logAddressID+"&UpdateDateTime="+strUpdateDateTime ;
	}
}
 //**********************************************************************

function btnNew_onclick() {
	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
		alert('Access denied. Please contact your system administrator.');
		return;
	}
	self.document.location.href ="AddressDetail.asp?AddressID=0" ;
}

//**********************************************************************

function btnReset_onclick() {
	if(confirm('All changes will be lost. Do you really want to reset this page?')){
			bolNeedToSave = false;
			document.location = 'AddressDetail.asp?AddressID='+ "<%=logAddressID%>" ;
		}
}

//**********************************************************************

function btnMunicipalityLookup_onclick() {

var strCity ;
var strProvince ;
var	strCountry ;

	strCity = window.frmAddressDetail.txtCity.value ;
	strProvince = window.frmAddressDetail.txtProvince.value ;
	strCountry = window.frmAddressDetail.txtCountry.value ;

	if (strCity != "") {  SetCookie("CityName", strCity) ; }
	if (strProvince != "") {SetCookie("ProvinceName", strProvince); }
	if (strCountry != "" ) {SetCookie("CountryName", strCountry); }
	SetCookie("WinName", "Popup");
	bolNeedToSave = true ;
	window.open('SearchFrame.asp?fraSrc=Municipalities', 'Popup', 'top=50, left=100, WIDTH=800, HEIGHT=600'  ) ;
}

//**********************************************************************

function fct_onChange(){
	bolNeedToSave = true ;
}

//**********************************************************************

function btnReferences_onclick() {
var strOwner = 'CRP' ;				// owner name must be in Uppercase
var strTableName = 'ADDRESS' ;		// table name must be in Uppercase
var strRecordID = document.frmAddressDetail.hdnAddressID.value ;
var URL ;

	if (strRecordID != ""  ){
		URL ='Dependency.asp?Owner=' + strOwner + '&TableName=' + strTableName + '&RecordID='+ strRecordID   ;
		window.open(URL, 'Popup', 'top=100, left=100, WIDTH=500, HEIGHT=300'  ) ; }
	else
		{alert("No references. This is a new record."); }

}

//**********************************************************************

function form_onsubmit(){

  if (((intAccessLevel & intConst_Access_Create) == intConst_Access_Create) || ((intAccessLevel & intConst_Access_Update) == intConst_Access_Update) )
  {

	if (document.frmAddressDetail.txtCustomerName.value == "" ) {
		alert('Please select a customer using lookup function');
		document.frmAddressDetail.btnCustomerLookup.focus();
		return(false);}

/*remove on May 2, 2008
//	if (isNaN(document.frmAddressDetail.txtHouse.value)) {
//	    alert("House number must be a number");
//		document.frmAddressDetail.txtHouse.focus();
//		document.frmAddressDetail.txtHouse.select();
//		return(false) ;	 }

//	if ((document.frmAddressDetail.txtSuffix.value != "" ) && (document.frmAddressDetail.txtSuffix.value == 0) || (parseInt(document.frmAddressDetail.txtSuffix.value))) {
//	    alert("Suffix must be a letter");
//		document.frmAddressDetail.txtSuffix.focus();
//		document.frmAddressDetail.txtSuffix.select();
//		return(false) ;	 }
removed on May 2, 2008*/

	if (document.frmAddressDetail.txtCity.value == "" ) {
		alert('Please select a City using lookup function');
		document.frmAddressDetail.btnMunicipalityLookup.focus();
		return(false); }

	if (document.frmAddressDetail.txtStreet.value == "" ) {
		alert("Please enter the street address");
		document.frmAddressDetail.txtStreet.focus();
		return(false); }

	bolNeedToSave = false ;
	document.frmAddressDetail.hdnFrmAction.value = "SAVE" ;
	document.frmAddressDetail.submit();
	return(true);

  }
  else
  {
	alert('Access denied. Please contact your system administrator.');
  	return(false);
  }

}

//**********************************************************************

function body_onbeforeunload() {

	//must set focus to save button because is user has changed only one field and has not left it the on_change event will not have fired and the flag that //determines whether you need to save or not will be false
	document.frmAddressDetail.btnSave.focus();
	if  ( bolNeedToSave == true ) {
		if (((intAccessLevel & "<%=intConst_Access_Create%>") == "<%=intConst_Access_Create%>") || ((intAccessLevel & "<%=intConst_Access_Update%>") == "<%=intConst_Access_Update%>") ){
				event.returnValue = "There is unsaved data on the screen. To save changes, click CANCEL below then click SAVE on the main form.";
		}
	}
}

//**********************************************************************

function ClearStatus() {
	window.status = "";
}

//**********************************************************************

function DisplayStatus(strWindowStatus){
	window.status=strWindowStatus;
	setTimeout('ClearStatus()', "<%=intConst_MessageDisplay%>");
}


//******************************************** End of Java Functions *****************************
//-->
</SCRIPT>
</HEAD>
<BODY onLoad="DisplayStatus(strWinMessage);" onbeforeunload="body_onbeforeunload();">
<FORM name=frmAddressDetail action="AddressDetail.asp" method="POST" onsubmit="return form_onsubmit()">
	<!-- hidden variables -->
	<INPUT id=hdnAddressID name=hdnAddressID type=hidden value=<%if logAddressID <> 0 then  Response.Write """"&objRS("address_id")&"""" else Response.Write """""" end if%>  >
	<%if strSimple = "Simple" then%>
	   <input id=hdnCustomerID name=hdnCustomerID type=hidden value=<%Response.Write """"&CustomerID&""""%>>
	<%else%>
	   <input id=hdnCustomerID name=hdnCustomerID type=hidden value=<%if logAddressID <> 0 then  Response.Write """"&objRS("customer_id")&"""" else Response.Write """""" end if%> >
	<%end if%>
	<input id=hdnCity name=hdnCity type=hidden value=<%if logAddressID <> 0 then  Response.Write """"&objRS("municipality_name")&"""" else Response.Write """""" end if%>>
	<INPUT id=hdnProvinceCode name=hdnProvinceCode type=hidden value=<%if logAddressID <> 0 then  Response.Write """"&objRS("province_state_lcode")&"""" else Response.Write """""" end if%> >
	<INPUT id=hdnCountryCode name=hdnCountryCode type=hidden value=<%if logAddressID <> 0 then  Response.Write """"&objRS("country_lcode")&"""" else Response.Write """""" end if%> >
	<INPUT id=hdnFrmAction name=hdnFrmAction type=hidden value= "">
	<INPUT id=hdnUpdateDateTime name=hdnUpdateDateTime type=hidden value=<%if logAddressID <> 0 then  Response.Write """"&objRS("last_update_date_time")&"""" else Response.Write """""" end if%>>
	<INPUT id=txtSuffix name=txtSuffix type=hidden value= "">
<TABLE>
<thead>
	<tr>
		<td align=left colspan=3 >Address Detail</td>
		<td align=right><SELECT <%if logAddressID = 0 then  Response.Write "disabled" end if%>  ALIGN=right id=selNavigate name=selNavigate  onchange="return selNavigate_onchange()" tabindex=19>
			<OPTION value="DEFAULT">Quickly Goto ...</OPTION>
			<OPTION value="Asset">Asset</OPTION>
			<OPTION value="Cust" >Customer</OPTION>
		</td></TR></THEAD>
<TBODY>
    <TR>
        <td align=right>Customer Name<FONT COLOR=RED>*</FONT></td>
        <td align=left>

			<%if strSimple = "Simple" then%>
			   <input disabled id=txtCustomerName name=txtCustomerName size=50 maxlength=50 value=<%Response.Write """"&CustomerName&""""%>>
			<%else%>
			   <input disabled id=txtCustomerName name=txtCustomerName size=50 maxlength=50 tabindex=1 value="<%if logAddressID <> 0 then  Response.Write routineHTMLString(objRS("customer_name")) else Response.Write NULL end if%>"  >
			   <input id=btnCustomerLookup name=btnCustomerLookup type=button tabindex=2
			   value="..." LANGUAGE=javascript onclick="return btnCustomerLookup_onclick('C')" onchange ="fct_onChange();"></td>
			<%end if%>
        <td></td></tr>
        <td></td></tr>

    <%if strSimple = "Simple" then%>
		 <input id=txtSuffix name=txtSuffix type=hidden value=<%Response.Write null%>>
		 <input id=txtHouse  name=txtHouse  type=hidden value=<%Response.Write null%>>
    <%else%>
        <tr>
        <td align=right>Building Name</td>
        <td align=left><input id=txtBuildingName name=txtBuildingName size=30 maxlength=30 tabindex=3 value="<%if logAddressID <> 0 then  Response.Write routineHTMLString(objRS("building_name")) else Response.Write NULL end if%>" onchange ="fct_onChange();"></td>
        <td align=right>Primary Address</td>
        <td align=left><input id="chkPrimary" name="chkPrimary" type=checkbox tabindex=11
			<%if logAddressID <> 0 then IF objRs("primary") = "Y" THEN Response.Write  "CHECKED" end if end if%> onclick ="fct_onChange();" ></td></tr>
        <tr>
        <td align=right>Apartment</td>
        <td align=left>
			<input id=txtApartment name=txtApartment size=5 maxlength=5 tabindex=4 value= "<%if logAddressID <> 0 then  Response.Write routineHTMLString(objRS("apartment_number")) else Response.Write NULL end if%>" onchange ="fct_onChange();" >
			House # & Suffix
			<input id=txtHouse name=txtHouse size=6 maxlength=6 tabindex=5 value= "<%if logAddressID <> 0 then  Response.Write objRS("house_number") else Response.Write NULL end if%>" onchange ="fct_onChange();">
			Vector <SELECT id=selVector name=selVector tabindex=7 onchange ="fct_onChange();">
				<OPTION></OPTION>
				<OPTION VALUE="E"  <%If logAddressID <> 0 then if objRS("street_vector") = "E" Then Response.Write "selected" end if end if %> >E</OPTION>
				<OPTION VALUE="N"  <%If logAddressID <> 0 then if objRS("street_vector") = "N" Then Response.Write "selected" end if end if %>>N</OPTION>
				<OPTION VALUE="S"  <%If logAddressID <> 0 then if objRS("street_vector") = "S" Then Response.Write "selected" end if end if %>>S</OPTION>
				<OPTION VALUE="W"  <%If logAddressID <> 0 then if objRS("street_vector") = "W" Then Response.Write "selected" end if end if %>>W</OPTION>
				<OPTION VALUE="NE" <%If logAddressID <> 0 then if objRS("street_vector") = "NE" Then Response.Write "selected" end if end if %>>NE</OPTION>
				<OPTION VALUE="NW" <%If logAddressID <> 0 then if objRS("street_vector") = "NW" Then Response.Write "selected" end if end if %>>NW</OPTION>
				<OPTION VALUE="SE" <%If logAddressID <> 0 then if objRS("street_vector") = "SE" Then Response.Write "selected" end if end if %>>SE</OPTION>
				<OPTION VALUE="SW" <%If logAddressID <> 0 then if objRS("street_vector") = "SW" Then Response.Write "selected" end if end if %>>SW</OPTION>
			</SELECT></td>
        <td align=right>Mailing Address</td>
        <td align=left><input id=chkMailing name=chkMailing type=checkbox tabindex=12
			<%if logAddressID <> 0 then IF objRs("mailing")="Y" THEN Response.Write  "CHECKED" end if END IF%> onclick ="fct_onChange();" ></td></tr>

    <% end if %>

    <tr>
        <td align=right>Street Name<FONT COLOR=RED>*</FONT></td>
        <td align=left><input id=txtStreet name=txtStreet size=50 maxlength=75 tabindex=8 value="<%if logAddressID <> 0 then  Response.Write routineHTMLString(objRS("long_street_name")) else Response.Write NULL end if%>" onchange ="fct_onChange();"></td>

        <%if strSimple <> "Simple" then%>
              <td align=right>Billing Address</FONT></td>
              <td align=left><input id=chkBilling name=chkBilling type=checkbox tabindex=13
			  <%if logAddressID <> 0 then IF ObjRS("billing")="Y" THEN  Response.Write  "CHECKED"  end if END IF%> onclick ="fct_onChange();"></td>
	    <%end if%>
    <tr>
		<td align=right>City/Municipality Name<FONT COLOR=RED>*</FONT></td>
        <td align=left>
			<input disabled id=txtCity name=txtCity size=50 maxlength=50  value="<%if logAddressID <> 0 then  Response.Write routineHTMLString(objRS("municipality_name")) else Response.Write NULL end if%>" onchange ="fct_onChange();" >
			<input id=btnMunicipalityLookup name=btnMunicipalityLookup tabindex=9 type=button value=...  LANGUAGE=javascript onclick="return btnMunicipalityLookup_onclick()"></td>
        <td></td>
        <td></td></tr>

    <tr>
        <td align=right>Province</td>
        <td align=left>
			<INPUT id=txtProvince name=txtProvince size=30 maxlength=30 value="<%if logAddressID <> 0 then  Response.Write routineHTMLString(objRS("province_state_name")) else Response.Write NULL end if%>"  disabled >
        <td></td>
        <td></td></tr>
    <tr>
        <td align=right>Country</td>
        <td align=left>
			<INPUT id=txtCountry name=txtCountry size=30 maxlength=30 value="<%if logAddressID <> 0 then  Response.Write routineHTMLString(objRS("country_desc")) else Response.Write NULL end if%>"  disabled>
        <td></td>
        <td></td></tr>

    <%if strSimple <> "Simple" then%>
          <tr>
          <td align=right>Postal/Zip Code</td>
          <td align=left><input id=txtPostal name=txtPostal size=15 maxlength=15 tabindex=10  value="<%if logAddressID <> 0 then  Response.Write routineHTMLString(objRS("postal_code_zip")) else Response.Write NULL end if%>" onchange ="fct_onChange();"></td>
          <td></td>
          <td></td>
         </tr>
    <%end if%>
</tbody>
<tfoot>
    <tr>
        <td colSpan=4 align=right>
        <%if strSimple <> "Simple" then%>
			<input name=btnReferences type=button value=References  tabindex=14 style= "width: 2.2cm" LANGUAGE=javascript onclick="return btnReferences_onclick()">&nbsp;&nbsp;
        <% end if %>
            <input name=btnDelete type=button value=Delete tabindex=15 style= "width: 2cm" LANGUAGE=javascript onclick="return btnDelete_onclick()">&nbsp;&nbsp;
            <input name=btnReset type=button value=Reset tabindex=16 style= "width: 2cm" onclick="btnReset_onclick();" >&nbsp;&nbsp;
            <input name=btnNew type=button value=New tabindex=17  style= "width: 2cm" LANGUAGE=javascript onclick="return btnNew_onclick()">&nbsp;&nbsp;
            <input name=btnSave type=button value=Save tabindex=18 style= "width: 2cm" LANGUAGE=javascript onclick="return form_onsubmit()" >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
          </TR></TR>
</TFOOT>
</TABLE>
<FIELDSET>
	<LEGEND align=right><B>Audit Information</B></LEGEND>
	<Div SIZE=8pt ALIGN=RIGHT>
		Record Status Indicator
		<INPUT align=left name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 18px"disabled value="<%if logAddressID <> 0 then  Response.Write routineHTMLString(objRS("record_status_ind")) else Response.Write null end if%>" >&nbsp;&nbsp;&nbsp;
		Create Date
		<INPUT align=center name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 150px"disabled value="<%if logAddressID <> 0 then  Response.Write routineHTMLString(objRS("create_date")) else Response.Write null end if%>"  >&nbsp;
		Created By
		<INPUT align=right name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 200px"disabled value="<%if logAddressID <> 0 then  Response.Write routineHTMLString(objRS("create_real_userid")) else Response.Write null end if%>"  ><BR>
		Update Date
		<INPUT align=center name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 150px"disabled value="<%if logAddressID <> 0 then  Response.Write routineHTMLString(objRS("update_date")) else Response.Write null end if%>"  >
		Updated By
		<INPUT align=right name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 200px"disabled value="<%if logAddressID <> 0 then  Response.Write routineHTMLString(objRS("update_real_userid")) else Response.Write null end if%>" >
	</DIV>
</FIELDSET>
</FORM>
</BODY>
</html>
<% if logAddressID <> 0 then
	objRs.Close
    set objRS = nothing
    set objConn = nothing
end if %>
