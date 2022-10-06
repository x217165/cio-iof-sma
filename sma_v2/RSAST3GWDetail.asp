<%@ Language=VBScript %>
<% Option Explicit %>
<% Response.Buffer = true %>
<!--% on error resume next %-->
<!--#include file="smaConstants.inc"-->
<!--#include file="databaseconnect.asp"-->
<!--#include file="smaProcs.inc"-->
<!--
*************************************************************************************
* File Name:	RSAST3GWDetail.asp
*
* Purpose:	    Create/Update Gateway Detail
*
* In Param:
*
* Out Param:
*
* Created By:
* Edited by:
**************************************************************************************

        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
       10-22-01	     DTy		Change field names and variables.

                                Gateway IP         to GW Router T1 Serial IP Address

                                Gateway DLCI IP    to Gateway DLCI(IP)
                                txtGatewayDLCIIP   to txtDLCIIP
                                strGWDLCIIP        to strDLCIIP

                                PNG_IP_ID          to LAN_ID_IP
                                PNG_IP             to LAN_IP

                                Gateway DLCI POS   to Gateway DLCI(X25)
                                GATEWAY_DLCI_POS   to GATEWAY_DLCI_X25
                                txtGatewayDLCIPOS  to txtDLCIX25
                                strGWDLCIPOS       to strDLCIX25
                                p_gateway_dlci_pos to p_gateway_dlci_x25

                                Move Gateway DLCI (X25) before Gateway DLCI (IP)

                                Increase customer name field from 20 t0 50 characters.
                                Add a new field 'Gateway Circuit Number'.
                                Retrieve Customer ID and Address ID.
**************************************************************************************
                                Activate 'Close' button.

-->

<%

'check user's rights
dim intAccessLevel
intAccessLevel = CInt(CheckLogon(strConst_RSAS))
if intAccessLevel and intConst_Access_ReadOnly <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to POS PLUS Tier 3. Please contact your system administrator"
end if

dim sql, strWinMessage, rsGateway

dim strAction
strAction = Request("action")			'get the action code from caller
if strAction = "" then
	Response.write "No action requested"
	Response.End						'no action requested
end if

dim strGWID, intTCCount, lngCustID, lngAddrID
strGWID = Request("GWID")			'get gateway id
lngCustID = Request("CustID")		'get customer id
lngAddrID = Request("AddrID")		'get address id


dim strRealUserID
strRealUserID = Session("username")
if err then
	'unexpected error
	DisplayError "BACK", "", 0, "UNEXPECTED ERROR", "Close gateway window to return to tail circuit form."
end if
strLastUpdate = Request("hdnLastUpdate")

'Response.Write ("strGWID = " & strGWID & ", " & "strLastUpdate = " & strLastUpdate & " strAction = " & strAction)
'Response.end

'save changes?
if strAction = "save" then 'needs to be changed ADAM!!!
	dim strDLCIIP, strDLCIX25,  strLastUpdate, strIPAddress, strCustomerID, strGWCircuitNo
	strGWCircuitNo = Request("txtGWCircuitNo")
	strDLCIIP = Request("txtDLCIIP")
	strDLCIX25 = Request("txtDLCIX25")
	strIPAddress = Request("selGatewayIP")
	strCustomerID = Request("hdnCustomerID")
	'call stored proc to save the record

	if (strGWID <> "") and (intAccessLevel and intConst_Access_Update = intConst_Access_Update) then
		'create command object for update stored proc
		dim cmdUpdateObj
		set cmdUpdateObj = server.CreateObject("ADODB.Command")
		set cmdUpdateObj.ActiveConnection = objConn
		cmdUpdateObj.CommandType = adCmdStoredProc
		cmdUpdateObj.CommandText = "sma_sp_userid.spk_sma_rsas_inter.sp_gateway_update"
		'create params

		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_user_id", adVarChar , adParamInput, 30, strRealUserID)
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_gateway_id", adNumeric , adParamInput,, CLng(strGWID))
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , CDate(strLastUpdate))
		'optional parms
		if (strIPAddress <> "") then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_gateway_ip_id", adNumeric , adParamInput,, CLng(strIPAddress))
		else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_gateway_ip_id", adNumeric , adParamInput,, null)
		end if
		if (strDLCIIP <> "") then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_gateway_dlci_ip", adVarChar, adParamInput, 9, UCase(strDLCIIP))
		else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_gateway_dlci_ip", adVarChar, adParamInput, 9, null)
		end if
		if (strDLCIX25 <> "") then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_gateway_dlci_x25", adVarChar, adParamInput, 9, UCase(strDLCIX25))
		else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_gateway_dlci_x25", adVarChar, adParamInput, 9, null)
		end if
		if (strCustomerID <> "" ) then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_customer_id", adNumeric , adParamInput,, CLng(strCustomerID))
		else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_customer_id", adNumeric , adParamInput,, null)
		end if
		if (strGWCircuitNo <> "" ) then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_gateway_circuit_number", adVarChar , adParamInput, 19, strGWCircuitNo)
		else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_gateway_circuit_number", adVarChar , adParamInput, 19, null)
		end if
		'Response.Write "<b> count = " & cmdUpdateObj.Parameters.count & "<br>"
		'dim nx
		'for nx = 0 to cmdUpdateObj.Parameters.Count-1
	'		Response.Write cmdUpdateObj.Parameters.Item(nx).Name & " = " & cmdUpdateObj.Parameters.Item(nx).Value & " <br>"
	'	next
	'	Response.end
		'call the update stored proc
		cmdUpdateObj.Execute

		if err then
			DisplayError "BACK", "", err.Number, "CANNOT UPDATE OBJECT - PARAMETER ERROR", err.Description
			objConn.Errors.Clear
		end if
		if objConn.Errors.Count <> 0 then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT UPDATE OBJECT", objConn.Errors(0).Description
			objConn.Errors.Clear
		end if
		strWinMessage = "Record saved successfully."
	elseif (strGWID = "") and (intAccessLevel and intConst_Access_Create = intConst_Access_Create) then
		'create command object for insert stored proc
		dim cmdInsertObj
		set cmdInsertObj = server.CreateObject("ADODB.Command")
		set cmdInsertObj.ActiveConnection = objConn
		cmdInsertObj.CommandType = adCmdStoredProc
		cmdInsertObj.CommandText = "sma_sp_userid.spk_sma_rsas_inter.sp_gateway_insert"
		'create params
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_user_id", adVarChar , adParamInput, 30, strRealUserID)
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_gateway_id", adNumeric , adParamOutput)
		'optional parms
		if (strIPAddress <> "") then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_gateway_ip_id", adNumeric , adParamInput,, CLng(strIPAddress))
		else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_gateway_ip_id", adNumeric , adParamInput,,null)
		end if

		if (strDLCIIP <> "") then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_gateway_dlci_ip", adVarChar, adParamInput, 9, UCase(strDLCIIP))
		else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_gateway_dlci_ip", adVarChar, adParamInput, 9, null)
		end if

		if (strDLCIX25 <> "") then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_gateway_dlci_x25", adVarChar, adParamInput, 9, UCase(strDLCIX25))
		else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_gateway_dlci_x25", adVarChar, adParamInput, 9, null)
		end if
		if (strCustomerID <> "") then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_customer_id", adNumeric , adParamInput,, CLng(strCustomerID))
		else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_customer_id", adNumeric , adParamInput,, null)
		end if
		if (strGWCircuitNo <> "" ) then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_gateway_circuit_number", adVarChar , adParamInput, 19, strGWCircuitNo)
		else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_gateway_circuit_number", adVarChar , adParamInput, 19, null)
		end if

		'Response.Write "<b> count = " & cmdInsertObj.Parameters.count & "<br>"
		'dim nx2
		'for nx2 = 0 to cmdInsertObj.Parameters.Count-1
			'Response.Write cmdInsertObj.Parameters.Item(nx2).Name & " = " & cmdInsertObj.Parameters.Item(nx2).Value & " <br>"
		'next
		'Response.end

		if objConn.Errors.Count <> 0 then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT INSERT OBJECT - PARAMETER ERROR", objConn.Errors(0).Description
			objConn.Errors.Clear
		end if
		'call the update stored proc
		cmdInsertObj.Execute
		if objConn.Errors.Count <> 0 then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT INSERT OBJECT", objConn.Errors(0).Description
			objConn.Errors.Clear
		end if
		strGWID = cmdInsertObj.Parameters("p_gateway_id").Value	'set return parameter
		if strGWID = "" then
			DisplayError "BACK", "", 2100, "CANNOT DISPLAY NEW GATEWAY.", "?????Most probably the new alias has been saved successfully even if there was an error retrieving the new id. Close the alias window to return to the customer screen."
			objConn.Errors.Clear
		end if
		strWinMessage = "Record saved successfully."
	else
		DisplayError "BACK", "", err.Number, "ACCESS DENIED. PLEASE CONTACT YOUR SYSTEM ADMINISTRATOR.", err.Description
	end if
	if err then
		DisplayError "BACK", "", err.Number, "CANNOT UPDATE GATEWAY - TRY AGAIN", err.Description
	end if
end if

'delete gateway?

if strAction = "delete" then
	'call stor proc to delete current alias
	if intAccessLevel and intConst_Access_Delete = intConst_Access_Delete then
		'create command object for update stored proc
		dim cmdDeleteObj
		set cmdDeleteObj = server.CreateObject("ADODB.Command")
		set cmdDeleteObj.ActiveConnection = objConn
		cmdDeleteObj.CommandType = adCmdStoredProc
		cmdDeleteObj.CommandText = "sma_sp_userid.spk_sma_rsas_inter.sp_gateway_delete"
		'create params
		cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_gateway_id", adNumeric , adParamInput,, CLng(strGWID))
		cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_last_update", adDBTimeStamp, adParamInput,, CDate(strLastUpdate))
		'call the delete stored proc
		if objConn.Errors.Count <> 0 then
	'		DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT - PARAMETER ERROR", objConn.Errors(0).Description
			objConn.Errors.Clear
		end if

		'Response.Write "<b> count = " & cmdDeleteObj.Parameters.count & "<br>"
		'dim nx
		'for nx = 0 to cmdDeleteObj.Parameters.Count-1
		'	Response.Write cmdDeleteObj.Parameters.Item(nx).Name & " = " & cmdDeleteObj.Parameters.Item(nx).Value & " <br>"
		'next
		'Response.end

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
		DisplayError "BACK", "", err.Number, "CANNOT DELETE GATEWAY", err.Description
	end if
	'called from main form?
	'if Request("back") = "true" then
	'	Response.Redirect "CustAlias.asp?CustomerID="&strMasterID
	'end if
	'ready to enter a new alias?
	strGWID=""
	strAction="new"
end if ' if strAction = "delete" then

'display the gateway info
if strAction <> "new" then
	sql =	"SELECT " &_
				"GW.GATEWAY_ID, " &_
				"GW.GATEWAY_IP_ID, " &_
				"A.IP_ADDRESS, " &_
				"GW.GATEWAY_DLCI_IP, " &_
				"GW.GATEWAY_DLCI_X25, " &_
				"GW.CUSTOMER_ID, " &_
				"C.CUSTOMER_NAME, " &_
				"TO_CHAR(GW.CREATE_DATE_TIME,'MON-DD-YY HH24:MI:SS')AS CREATE_DATE_TIME, " &_
				"sma_sp_userid.spk_sma_library.sf_get_full_username(GW.CREATE_REAL_USERID) as create_real_userid, " &_
				"GW.UPDATE_DATE_TIME, " &_
				"TO_CHAR(GW.UPDATE_DATE_TIME,'MON-DD-YY HH24:MI:SS')AS LAST_UPDATE_DATE_TIME, " &_
				"sma_sp_userid.spk_sma_library.sf_get_full_username(GW.UPDATE_REAL_USERID) as update_real_userid, "&_
				"GW.RECORD_STATUS_IND, " &_
				"A.IP_ADDRESS_ID, " &_
				"GW.GATEWAY_CIRCUIT_NUMBER " &_
			"FROM CRP.RSAS_GATEWAY GW" &_
			", CRP.RSAS_IP_ADDRESS A" &_
			", CRP.CUSTOMER C " &_
			"WHERE GW.GATEWAY_ID = " & clng(strGWID) & " " &_
			" and GW.GATEWAY_IP_ID = A.IP_ADDRESS_ID(+) " &_
			" and GW.CUSTOMER_ID = C.CUSTOMER_ID(+) "

	'Response.Write (sql)
	'Response.end

	set rsGateway=server.CreateObject("ADODB.Recordset")
	rsGateway.CursorLocation = adUseClient
	rsGateway.Open sql, objConn
	if err then
		DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
	end if
	if rsGateway.EOF then
		DisplayError "BACK", "", err.Number, "CANNOT FIND GATEWAY", err.Description
	end if
	set rsGateway.ActiveConnection = nothing

	'Get the Tail Circuit Information
	dim rsTailCircuit, strSQL
		'strSQL= "Select customer_name_alias_id, customer_name_alias_upper from crp.customer_name_alias where customer_id = " & strCustomerID
		strSQL = "SELECT DISTINCT TC.TAIL_CIRCUIT_ID, " &_
			"WAN_IP.IP_ADDRESS, " &_
			"LAN_IP.IP_ADDRESS, " &_
			"TC.NODE_NAME, "&_
			"TC.TAIL_CIRCUIT_NUMBER, "&_
			"TC.WAN_IP_DLCI, "&_
			"TC.POS_IP_DLCI, "&_
					"NVL(AD.BUILDING_NAME,'<NO BUILDING SPECIFIED>') ||CHR(13)||CHR(10)|| " &_
					"decode(AD.APARTMENT_NUMBER, null, null, rtrim(AD.APARTMENT_NUMBER) || ' ') || " &_
					"decode(to_char(AD.HOUSE_NUMBER) || AD.HOUSE_NUMBER_SUFFIX, null, null, rtrim(to_char(AD.house_number) || AD.house_number_suffix)  || ' ') || " &_
					"decode(AD.STREET_VECTOR, null, null, rtrim(AD.STREET_VECTOR) || ' ') || " &_
					"NVL(AD.LONG_STREET_NAME,'NO STREET ADDRESS')||CHR(13)||CHR(10)|| " &_
					"NVL(AD.MUNICIPALITY_NAME,'NO MUNICIPALITY')||' '|| " &_
					"NVL(AD.PROVINCE_STATE_LCODE,'NO PROVINCE')||CHR(13)||CHR(10)|| " &_
					"NVL(AD.POSTAL_CODE_ZIP,'NO POSTAL CODE') ADDRESS, " &_
			"TC.ORDER_NUMBER " &_
			"FROM " &_
			"CRP.RSAS_TAIL_CIRCUIT	TC, "&_
			"CRP.RSAS_IP_ADDRESS	WAN_IP, "&_
			"CRP.RSAS_IP_ADDRESS	LAN_IP, "&_
			"CRP.CUSTOMER			CU, "&_
			"CRP.ADDRESS			AD "&_
			"WHERE " &_
		"TC.GATEWAY_ID= "& strGWID & " " &_
		"AND TC.WAN_IP_ID = WAN_IP.IP_ADDRESS_ID (+) "&_
		"AND TC.LAN_IP_ID = LAN_IP.IP_ADDRESS_ID (+) "&_
		"AND TC.SITE_ADDRESS_ID = AD.ADDRESS_ID (+) "

	'Response.Write (strSQL)
	'Response.end

		set rsTailCircuit=server.CreateObject("ADODB.Recordset")
		rsTailCircuit.CursorLocation = adUseClient
		rsTailCircuit.Open strSQL, objConn
		if err then
			DisplayError "BACK", "", err.Number, "Cannot create recordset rsTailCircuit.", err.Description
		end if
		set rsTailCircuit.ActiveConnection=nothing

		'create the innerValues for the iFrame
		dim intRowCount, intColCount, strInnerValues
			intRowCount = 0
			intColCount = 10
			strInnerValues = ""
		while not rsTailCircuit.EOF
			intRowCount = intRowCount + 1
			strInnerValues =strInnerValues & rsTailCircuit(0) & strDelimiter & rsTailCircuit(1) & strDelimiter &_
											rsTailCircuit(2) & strDelimiter & rsTailCircuit(3) & strDelimiter &_
											rsTailCircuit(4) & strDelimiter & rsTailCircuit(5) & strDelimiter &_
											rsTailCircuit(6) & strDelimiter & rsTailCircuit(7) & strDelimiter '&_
											'rsTailCircuit(8) & strDelimiter & rsTailCircuit(9) & strDelimiter
			rsTailCircuit.MoveNext
		wend
		intTCCount = intRowCount
		rsTailCircuit.Close
		set rsTailCircuit = nothing
end if

if strAction = "new" then
  intTCCount = 0
end if

dim objRsGatewayIP
   'get a list of available Gateway IPs


	if (strGWID <> "") and (intTCCount > 0) then
		strSQL = "SELECT  IPA.IP_ADDRESS_ID, IPA.IP_ADDRESS, " &_
			"IPA.SUBNET_MASK, IPA.LOCATION " &_
			"FROM CRP.RSAS_IP_ADDRESS IPA, CRP.RSAS_GATEWAY GW " &_
			"WHERE GW.GATEWAY_IP_ID = IPA.IP_ADDRESS_ID "&_
			"and GW.GATEWAY_ID = " &strGWID
			'Response.Write (strSQL)
		'Response.End
		'set objRsGatewayIP = objConn.Execute(strSQL)
		set objRsGatewayIP = objConn.Execute(strSQL)

	elseif (strGWID <> "") then
		strSQL = "SELECT  IPA.IP_ADDRESS_ID, IPA.IP_ADDRESS, " &_
			"IPA.SUBNET_MASK, IPA.LOCATION " &_
			"FROM CRP.RSAS_IP_ADDRESS IPA " &_
			"WHERE ((IPA.CODE='GWIP') AND (IPA.AVAILABLE='Y')) " &_
			" OR IPA.IP_ADDRESS_ID in " &_
				" (select GW.GATEWAY_IP_ID from crp.rsas_gateway GW " &_
				" where GW.GATEWAY_ID = " &strGWID & ")"

		'Response.Write (strSQL)
		'Response.End
		'set objRsGatewayIP = objConn.Execute(strSQL)
		set objRsGatewayIP = objConn.Execute(strSQL)
	else
		strSQL = "SELECT  IPA.IP_ADDRESS_ID, IPA.IP_ADDRESS, " &_
			"IPA.SUBNET_MASK, IPA.LOCATION " &_
			"FROM CRP.RSAS_IP_ADDRESS IPA " &_
			"WHERE ((IPA.CODE='GWIP') AND (IPA.AVAILABLE='Y')) "

		'Response.Write (strSQL)
		'Response.End
		'set objRsGatewayIP = objConn.Execute(strSQL)
		set objRsGatewayIP = objConn.Execute(strSQL)

	end if

%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK REL="stylesheet" TYPE="text/css" HREF="stylesheets/styles.css">
<TITLE>POS PLUS Gateway Detail</TITLE>
<SCRIPT type="text/javascript" language="javascript" src="AccessLevels.js"></SCRIPT>
<SCRIPT type="text/javascript" language="javascript" src="GeneralJavaFunctions.js"></SCRIPT>
<SCRIPT type="text/javascript">
<!--
//**************************************** JavaScipt Functions *********************************

var bolSaveRequired = false;
intAccessLevel=<%=intAccessLevel%>;
var intConst_MessageDisplay=<%=intConst_MessageDisplay%>;


	function iFrame_display(){
		if ((intAccessLevel & intConst_Access_ReadOnly) == intConst_Access_ReadOnly) {
			document.frames("tcifr").document.location.href = 'RSAST3GWTailCircuit.asp?GWID=<%=strGWID%>&CustID=<%=lngCustId%>&AddrID=<%=lngAddrID%>';
		}
		else {alert('Access Denied. You do not have access to tail circuit. Please contact your system administrator.')}
		//document.frmRSASGWDetail.btnSave.disabled = false;
	}

	function btn_iFrmTCAdd(){
	//open a blank form
		if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
			alert('Access Denied - can not tail circuit. Please contact your system administrator.');
			return (false);
		}
		if (document.frmRSASGWDetail.GWID.value == ""){
			alert('At this time you cannot create a tail circuit. You must save the gateway first.');
			return (false);
		}
		var NewWin;
		var strMasterID = "<%=strGWID%>";
		var lngCustomerID = "<%=lngCustID%>";
		var lngAddressID = "<%=lngAddrID%>";
		//document.frmRSASGWDetail.btnSave.disabled = true;
//		NewWin=window.open("RSAST3Detail.asp?action=new&hdnGatewayId="+strMasterID+"&hdnCustomerID="+lngCustomerID+"&hdnAddressID="+lngAddressID,"NewWin","toolbar=no,status=yes,width=800px,height=700px,left=100px,top=0,menubar=no,resize=no");

		NewWin=window.open("RSAST3Detail.asp?action=new&hdnGatewayId="+strMasterID+"&hdnCustomerID="+lngCustomerID+"&hdnAddressID=","NewWin","toolbar=no,status=yes,width=800px,height=700px,left=100px,top=0,menubar=no,resize=no");
		NewWin.focus();
	}

	function btn_iFrmTCUpdate(){
		//open a detail form where the user can modify the tail circuit
		if ((intAccessLevel & intConst_Access_Update) != intConst_Access_Update) {
			alert('Access Denied - can not update tail circuit. Please contact your system administrator.');
			return;
		}
		var NewWin;
		var strTCID = document.frames("tcifr").document.frmIFR.hdnTCID.value;
		if (strTCID == ""){
			alert("Please select an tail circuit or click NEW to create a new tail circuit.");
			return;
		}
		var strMasterID = "<%=strGWID%>";
		var lngCustomerID = "<%=lngCustID%>";
		var lngAddressID = "<%=lngAddrID%>";
		//document.frmRSASGWDetail.btnSave.disabled = true;
		NewWin=window.open("RSAST3Detail.asp?action=update&hdnTailCircuitID="+strTCID+"&hdnGatewayID="+strMasterID+"&hdnCustomerID="+lngCustomerID+"&hdnAddressID="+lngAddressID,"NewWin","toolbar=no,status=yes,width=800px,height=700px,left=100px,top=0,menubar=no,resize=no");
		NewWin.focus();
	}

	function btn_iFrmTCDelete(){
		//delete selected row
		if ((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) {
			alert('Access Denied - can not delete tail circuit. Please contact your system administrator.');
			return;
		}
		var strTCID = document.frames("tcifr").document.frmIFR.hdnTCID.value;
		if (strTCID == "") {
			alert("Please select an tail circuit or click ADD to create a new tail circuit.");
			return;
		}
		//document.frmRSASGWDetail.btnSave.disabled = true;
		var strLastUpdate = document.frames("tcifr").document.frmIFR.hdnLastUpdate.value;
		if (confirm("Are you sure you want to delete this tail circuit?")){
			document.frames("tcifr").document.location.href = "RSAST3Detail.asp?action=DELETE&back=true&hdnTailCircuitID="+strTCID+"&hdnGatewayId=<%=strGWID%>&hdnUpdateDateTime="+strLastUpdate;
		}
	}

	//-----------------end of iFrame Javascript-----------------------------------------------


function fct_onChange(){
	bolSaveRequired = true;
}

function btnNew_click(){
	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
		alert('Access denied. Please contact your system administrator.');
		return;
	}
	self.document.location.href ="RSAST3GWDetail.asp?action=new&GWID=";

}

function fct_lookupCustomer(CustService){

	strCustomerName = window.frmRSASGWDetail.txtCustomerName.value ;
	if (strCustomerName != "")
		{
			SetCookie("CustomerName", strCustomerName);
		}
	SetCookie("ServiceEnd", CustService);
	fct_onChange();
	SetCookie("WinName", 'Popup');
	window.open('SearchFrame.asp?fraSrc=Cust', 'Popup', 'top=50, left=100, height=600, width=800' ) ;

}
function fct_onDelete(){
	if (document.frmRSASGWDetail.GWID.value != '') {

		if ((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) {
			alert('Access denied. Please contact your system administrator.');
			return;
		}
		var intTCCount;

				intTCCount = document.frames("tcifr").document.frmIFR.hdnTCCount.value;;

				//if (<%=intTCCount%> > 0)
				if (intTCCount > 0)
				 {	alert('Tail Circuits exist for this gateway, please remove them before deleting the gateway.');
					return;
				}

   		if (confirm('Do you really want to delete this object?')){
			document.location = "RSAST3GWDetail.asp?action=delete&GWID="+document.frmRSASGWDetail.GWID.value+"&hdnLastUpdate="+escape(document.frmRSASGWDetail.hdnLastUpdate.value);
		}
	}
	else{fct_displayStatus('There is no need to delete an empty gateway.');}
}

function btnClose_onclick(){
window.close();
}

function frmRSASGWDetail_onsubmit() {
	if ((intAccessLevel & intConst_Access_Create) == intConst_Access_Create) {

		if (document.frmRSASGWDetail.txtCustomerName.value == "" ) {
				alert('Missing Required Field. Please enter a customer.');
				document.frmRSASGWDetail.btnCustomerLookup.focus();
				return(false);
			}
		if (document.frmRSASGWDetail.selGatewayIP.value == "" ) {
				alert('Missing Required Field. Please enter a valid Gateway IP.');
				document.frmRSASGWDetail.selGatewayIP.focus();
				return(false);
			}
		document.frmRSASGWDetail.action.value = "save";
		bolSaveRequired = false;
		document.frmRSASGWDetail.submit();
		return(true);

	} else {
		alert('Access denied. Please contact your system administrator.');
		return(false);
		}

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
	iFrame_display();
}

function body_onBeforeUnload(){
	document.frmRSASGWDetail.btnSave.focus();
	if (bolSaveRequired) {
		event.returnValue =
		"There is unsaved data on the screen. To save changes, click CANCEL below then click SAVE on the main form.";
	}
}

//***************************************** End of JavaScript Functions *************************
//-->
</SCRIPT>
</HEAD>

<BODY onLoad="body_onLoad();" onBeforeUnload="body_onBeforeUnload();">
<FORM name=frmRSASGWDetail LANGUAGE=javascript >
<INPUT type="hidden" name=action value="">
<INPUT type=hidden name=hdnLastUpdate value= <%if strGWID <> "" then Response.Write """"&routineHtmlString(rsGateway("UPDATE_DATE_TIME"))&"""" 	else Response.Write """""" end if%> >


<INPUT type=hidden name=GWID value=<%if strGWID <> "" then Response.Write rsGateway("GATEWAY_ID") end if %> >
<INPUT type=hidden name=hdnCustomerID value=<%if strGWID <> "" then Response.Write rsGateway("CUSTOMER_ID") end if%> >
<INPUT type=hidden name=hdnIPAddress value=<%if strGWID <> "" then Response.Write rsGateway("GATEWAY_IP_ID") end if%> >
<INPUT type=hidden name=hdnAddressID value=<%if strGWID <> "" then Response.Write Request("AddrID") end if%> >

<TABLE border=0 width=100%>
<THEAD>
	<TR ><TD colspan=5>POS PLUS Tier 3 Gateway Detail</td></tr>
</THEAD>

<TBODY>
<TR>
	<td align="right" >Gateway Circuit Number<font color="red"></font></td>
	<td align="left"  >
	<input name="txtGWCircuitNo" type="text" tabindex=1 size="19" maxlength="19" value="<%if strGWID <> "" then Response.write rsGateway("GATEWAY_CIRCUIT_NUMBER")%>" onChange="fct_onChange();"></td>
</TR>
<TR>
	<TD ALIGN="right">GW Router T1 Serial IP Address</TD>
	<!--TD align="left"><INPUT size=20 maxlength=20 disabled name=strGWRT1SIPAddr value="<%if strGWID <> "" then Response.write rsGateway("IP_ADDRESS")%>" onchange ="fct_onChange();">

	<INPUT  name=btnIPAddressLookup type=button tabindex=1 value=... LANGUAGE=javascript onclick="fct_lookupIP_Address('D')"></TD-->
	<TD align=left>
	<SELECT id=selGatewayIP name=selGatewayIP tabindex=2 style="HEIGHT: 22px; WIDTH: 425px" onchange ="fct_onChange();">
		<%if strGWID = "" then
			Response.Write "<OPTION></OPTION>"
		  end if %>
				<%Do while Not objRsGatewayIP.EOF
				Response.write "<OPTION "
					if strGWID <> ""  then
					 if  (CInt(objRsGatewayIP("IP_ADDRESS_ID")) = CInt(rsGateway("GATEWAY_IP_ID"))) then
						Response.Write " SELECTED "
					 end if
					end if
				Response.Write 	" VALUE=" &objRsGatewayIP(0)& ">" &objRsGatewayIP(1) &" subnet:" &objRsGatewayIP(2) &" location:" &objRsGatewayIP(3) & "</OPTION>" &vbCrLf
				objRsGatewayIP.MoveNext
				Loop
				%>
			</SELECT></TD>


			<td align="right" >Gateway DLCI (X25)<font color="red"></font></td>
			<td align="left"  >
				<input name="txtDLCIX25" type="text" tabindex=4 size="9" maxlength="9" value="<%if strGWID <> "" then Response.write rsGateway("GATEWAY_DLCI_X25")%>" onChange="fct_onChange();">
			</td>
</TR>
	<tr>
		<td align="right"  nowrap>Customer Name<font color="red">*</font></td>
		<td align="left">
			<input name="txtCustomerName" disabled type="text" size="50" maxlength="50" value="<%if strGWID <> "" then Response.write rsGateway("CUSTOMER_NAME")%>" onChange="fct_onChange();">
		<INPUT  name=btnCustomerLookup type=button tabindex=3 value=... LANGUAGE=javascript onclick="return fct_lookupCustomer('C')"></TD>

		<td align="right">Gateway DLCI (IP)<font color="red"></font></td>
		<td align="left" >
			<input name="txtDLCIIP" type="text" tabindex=5 size="9" maxlength="9" value="<%if strGWID <> "" then Response.write rsGateway("GATEWAY_DLCI_IP")%>" onChange="fct_onChange();">
		</td>

	</tr>
	<!--TR>
		<td align="right">Tail Circuit Count<font color="red"></font></td>
		<td align="left" >
			<input name="txtTailCircuitCount" disabled type="text"  size="5"  value="<%if strGWID <> "" then Response.write intTCCount%>" onChange="fct_onChange();">
		</td>

	</tr-->
	</table>

	<table>
	<thead><TR><TD align=left colspan=4>Tail Circuits</TD></TR></thead>
	<tbody>

		<td width=100% valign=top colspan= 4>
			<iframe tabindex=6 id=tcifr width=100% height=300 src="" scrolling=yes marginheight=1 marginwidth=1></iframe>
			<br>
			<input type="button" tabindex=7 style="width: 2cm" value="Delete" name="btn_iFrameDelete" onClick="btn_iFrmTCDelete();" class=button>&nbsp;
			<input type="button" tabindex=8 style="width: 2cm" value="Refresh" name="btn_iFrameRefresh" onClick="iFrame_display();" class=button>&nbsp;
			<input type="button" tabindex=9 style="width: 2cm" value="New" name="btn_iFrameAdd" onClick="btn_iFrmTCAdd();" class=button>&nbsp;
			<input type="button" tabindex=10 style="width: 2cm" value="Update" name="btn_iFrameUpdate" onCLick="btn_iFrmTCUpdate();" class=button>
		</td>

	</tr>
		</table>

	<table>
</TBODY>

<TFOOT>
	<TR><TD align=right colspan=5>
	  	<!--INPUT type=button tabindex=10 name=btnClose style= "width: 2cm" value="Close" onclick="return btnClose_onclick()">&nbsp;-->
	  	<INPUT type=button tabindex=11 name=btnDelete style= "width: 2cm" value="Delete" onclick="return fct_onDelete();">&nbsp;
	  	<INPUT type=reset  tabindex=12 name=btnReset style= "width: 2cm" value="Reset" style="HEIGHT: 24px; WIDTH: 51px">&nbsp;
	  	<INPUT type=button tabindex=13 name=btnAddNew style= "width: 2cm" value="New" onclick="return btnNew_click()">&nbsp;
	  	<INPUT type=button tabindex=14 name=btnSave style= "width: 2cm" value="Save" onclick="return frmRSASGWDetail_onsubmit()">&nbsp;&nbsp;
	</TD></TR>
</TFOOT>
</TABLE>

<FIELDSET>
<LEGEND align=right><B>Audit Information</B></LEGEND>
<Div SIZE=8pt ALIGN=RIGHT>
	Record Status Indicator
	<INPUT align = left name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 18px"disabled value="<%if strGWID <> "" then Response.write rsGateway("RECORD_STATUS_IND")%>" >&nbsp;&nbsp;&nbsp;
	Create Date&nbsp;<INPUT align=center name=txtCreateDateTime type=text style="HEIGHT: 20px; WIDTH: 150px" disabled value="<%if strGWID <> "" then Response.write rsGateway("CREATE_DATE_TIME")%>" >&nbsp;
	Created By&nbsp;<INPUT align=right name=txtCreateRealUser type=text style="HEIGHT: 20px; WIDTH: 200px" disabled value="<%if strGWID <> "" then Response.write rsGateway("CREATE_REAL_USERID")%>" ><BR>
	Update Date&nbsp;<INPUT align=center name=txtUpdateDateTime type=text style="HEIGHT: 20px; WIDTH: 150px" disabled value="<%if strGWID <> "" then Response.write rsGateway("LAST_UPDATE_DATE_TIME")%>" >
	Updated By&nbsp;<INPUT align=right name=txtUpdateRealUser type=text style="HEIGHT: 20px; WIDTH: 200px" disabled value="<%if strGWID <> "" then Response.write rsGateway("UPDATE_REAL_USERID")%>" >
</DIV>
</FIELDSET>
<%
'Clean up our ADO objects
if strGWID <> "" then
	rsGateway.Close
	set rsGateway = Nothing
	objConn.close
	set objConn = Nothing
end if
%>
</FORM>
</BODY>
</HTML>
