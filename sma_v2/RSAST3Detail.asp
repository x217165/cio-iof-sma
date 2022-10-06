<%@ Language=VBScript %>
<% option explicit %>
<% Response.Buffer = true %>

<!--% on error resume next %-->

<!--
*************************************************************************************
* File Name:	RSAST3Detail.asp
*
* Purpose:		To display the detailed information about a RSAS entry.
*				Entry is chosen via RSAST3List.asp*
*
* In Param:
*
* Out Param:
*
* Created By:	Shawn Meyers
**************************************************************************************

        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
       10-23-01	     DTy		Change field names and variables.

                                Gateway IP         to GW Router T1 Serial IP Address
                                GATEWAY_DLCI_POS   to GATEWAY_DLCI_X25
                                gateway_ip_dlci    to gateway_dlci_x25
                                strgatewaydlcipos  to strgatewaydlcix25
                                txtgatewaydlcipos  to txtgatewaydlcix25

								WAN IP Address     to WAN IP Port Address
								PNG IP Address     to LAN IP Port Address

								p_png_ip_id        to p_lan_ip_id
								selPNGIP           to selLANIP
								strPNGIPID         to strLANIPID
								objRSPNGIP         to objRSLANIP
								strPNGIPSQL        to strLANIPSQL
								hdnPNGIPID         to hdnLANIPID
								PNG_IP_ID          to LAN_IP_ID
								PNG_IP             to LAN_IP
								PNGIP              to LANIP

                                Gateway DLCI POS   to Gateway DLCI (X25)
                                Gateway DLCI IP    to Gateway DLCIs (IP)

                                Move Gateway DLCI (X25) before Gateway DLCI (IP)
                                Increase customer name field from floating t0 50 characters

								Remove screen fields: WAN IP DLCI, POS IP DLCI & Packet Size
								Remove variables: strPackSizeSQL
								Remove recordsets: objRsPacketSize

								Add Network Speeds entries and set default Network Speeds to 56K
								Add Port Speeds entries and set default Port Speeds to 9600
								Decrease Network and Port Speed screen fields size
								Re-arrange Tail Circuit Number position
								Add Total POS PLUS Sites count
								Adjust column sizes

								Fix bugs, bugs ...
								Disable 'Template' button.
								Re-order index.

								Remove <font color="red"></font>.
								Remove Width="??%".

								Retrieve/pass Municipality, Province, Country

======================================
selPacketSize
strX25SIPDLCI
strPOSIPDLCI
txtWANIPDLCI
strWANIPDLCI


**************************************************************************************
-->

<!--#include file="SmaConstants.inc"-->
<!--#include file="smaProcs.inc"-->
<!--#include file="databaseconnect.asp"-->

<%



'********************************
'check the present user's rights*
'********************************

dim intAccessLevel
dim intTotalNodes

intAccessLevel = CInt(CheckLogon(strConst_RSAS))


if intAccessLevel and intConst_Access_ReadOnly <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to managed objects. Please contact your system administrator."
end if



'****************************
'declare necessary variables*
'****************************
dim lngTCID   'variable for tail circuit ID being passed from list page
dim lngGWID   'variable for gateway ID being passed from list page
dim lngCustID 'variable for Customer ID being passed from list page

dim strWinLocation
dim strWinMessage
dim strRealUserID
dim strAction

strAction = Request("action")
if strAction = "" then
	Response.write "No action requested"
	Response.End						'no action requested
end if

intSiteAddressID = Request("hdnAddressID")

'get the hidden tail circuit id from string from list page
lngTCID = Request("hdnTailCircuitID")
if IsNumeric(lngTCID)  then
 lngTCID = Request("hdnTailCircuitID")
else
  lngTCID = 0
end if

'get the hidden Customer ID from string from list page
lngCustID = Request("hdnCustomerID")
if not IsNumeric(lngCustID)  or null then
  lngCustID = 0
end if

'get the hidden gateway ID from string from list page
lngGWID = Request("hdnGatewayId")
'Response.Write "hdn tail circuit id is equal to " & lngTCID
'Response.Write "hdn gateway id is equal to " & lngGWID
'Response.Write "<br> straction" & strAction
'Response.end

'get the hidden window location
strWinLocation = "RSAST3Detail.asp?RSASID="& Request("hdnTailCircuitID")

'set the variable for the UserInfo cookie
strRealUserID = Session("username")

'************************
'do save, insert, delete*
'************************

'Response.Write "hdnFormAction=" & Request("action")
select case strAction

	case "SAVE"

'check to see if tail circuit entry exists already in database by
'checking for the existence of the hidden tail circuit id

	  if lngTCID  <> 0 then  ' it is an existing record so save the changes

		if intAccessLevel and intConst_Access_Update <> intConst_Access_Update then

			DisplayError "REFRESH", strWinLocation, 0, "UPDATE DENIED", "You don't have access to update managed objects. Please contact your system administrator"

		end if

		dim cmdUpdateObj

		set cmdUpdateObj = server.CreateObject("ADODB.Command")
		set cmdUpdateObj.ActiveConnection = objConn
		cmdUpdateObj.CommandType = adCmdStoredProc

		'get the tail_circuit_detail stored update procedure <schema.package.procedure>
		cmdUpdateObj.CommandText = "sma_sp_userid.spk_sma_rsas_inter.sp_tail_circuit_update"


		'create the required parameters

		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_userid", adVarChar, adParamInput, 30,strRealUserID)
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_tail_circuit_id", adNumeric, adParamInput,, Clng(Request("hdnTailCircuitID")))
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_last_update_dt", adDBTimeStamp,adParamInput,, CDate(Request("hdnUpdateDateTime")))

		'create the optional parameters

		if Request("hdnGatewayId") <> "" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_gateway_id", adNumeric, adParamInput,, (Request("hdnGatewayId")))
		else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_gateway_id", adNumeric, adParamInput,, null)
		end if

		if Request("selWANIP") <> "" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_wan_ip_id", adNumeric, adParamInput,, (Request("selWANIP")))
		else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_wan_ip_id", adNumeric, adParamInput,, null)
		end if

		if Request("selLANIP") <> "" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_lan_ip_id", adNumeric, adParamInput,, (Request("selLANIP")))
		else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_lan_ip_id", adNumeric, adParamInput,, null)
		end if

		if Request("selNetworkSpeeds") <> "" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_network_speed", adVarChar, adParamInput, 6, (Request("selNetworkSpeeds")))
		else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_network_speed", adVarChar, adParamInput, 6, null)
		end if

		if Request("selPortSpeeds") <> "" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_port_speed", adVarChar, adParamInput, 6, (Request("selPortSpeeds")))
		else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_port_speed", adVarChar, adParamInput, 6, null)
		end if

		if Request("txtNodeName") <> "" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_node_name", adVarChar, adParamInput, 10, (Request("txtNodeName")))
		else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_node_name", adVarChar, adParamInput, 10, null)
		end if

		if Request("txtTailCircuitNumber") <> "" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_tail_circuit_number", adVarChar, adParamInput, 15, (Request("txtTailCircuitNumber")))
		else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_tail_circuit_number", adVarChar, adParamInput, 15, null)
		end if

		if Request("txtWANIPDLCI") <> "" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_wan_ip_dlci", adVarChar, adParamInput, 10, (Request("txtWANIPDLCI")))
		else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_wan_ip_dlci", adVarChar, adParamInput, 10, null)
		end if

		if Request("txtPOSIPDLCI") <> "" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_pos_ip_dlci", adVarChar, adParamInput, 10, (Request("txtPOSIPDLCI")))
		else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_pos_ip_dlci", adVarChar, adParamInput, 10, null)
		end if

		if Request("selPacketSize") <> "" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_packet_size", adVarChar, adParamInput, 5, (Request("selPacketSize")))
		else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_packet_size", adVarChar, adParamInput, 5, null)
		end if

		if Request("hdnAddressID") <> "" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_site_address_id", adNumeric, adParamInput,, (Request("hdnAddressID")))
		else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_site_address_id", adNumeric, adParamInput,, null)
		end if

		if Request("txtOrderNumber") <> "" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_order_number", adVarChar, adParamInput, 15, (Request("txtOrderNumber")))
		else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_order_number", adVarChar, adParamInput, 15, null)
		end if
		if Request("txtNodeNumber") <> "" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_node_number", adNumeric, adParamInput,, (Request("txtNodeNumber")))
		else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_node_number", adNumberic, adParamInput,, null)
		end if


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

		'get the tail_circuit_detail stored insert procedure <schema.package.procedure>

		cmdInsertObj.CommandText = "sma_sp_userid.spk_sma_rsas_inter.sp_tail_circuit_insert"


		'create the mandatory insert parameters

		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_userid", adVarChar, adParamInput, 30, strRealUserID)
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_tail_circuit_id", adNumeric, adParamOutput)

		'create the optional parameters

		if Request("hdnGatewayId") <> "" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_gateway_id", adNumeric, adParamInput,, (Request("hdnGatewayId")))
		else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_gateway_id", adNumeric, adParamInput,, null)
		end if


		if Request("selWANIP") <> "" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_wan_ip_id", adNumeric, adParamInput,, (Request("selWANIP")))
		else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_wan_ip_id", adNumeric, adParamInput,, null)
		end if

		if Request("selLANIP") <> "" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_lan_ip_id", adNumeric, adParamInput,, (Request("selLANIP")))
		else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_lan_ip_id", adNumeric, adParamInput,, null)
		end if

		if Request("selNetworkSpeeds") <> "" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_network_speed", adVarChar, adParamInput, 6, (Request("selNetworkSpeeds")))
		else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_network_speed", adVarChar, adParamInput, 6, null)
		end if

		if Request("selPortSpeeds") <> "" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_port_speed", adVarChar, adParamInput, 6, (Request("selPortSpeeds")))
		else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_port_speed", adVarChar, adParamInput, 6, null)
		end if

		if Request("txtNodeName") <> "" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_node_name", adVarChar, adParamInput, 10, (Request("txtNodeName")))
		else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_node_name", adVarChar, adParamInput, 10, null)
		end if

		if Request("txtTailCircuitNumber") <> "" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_tail_circuit_number", adVarChar, adParamInput, 15, (Request("txtTailCircuitNumber")))
		else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_tail_circuit_number", adVarChar, adParamInput, 15, null)
		end if

		if Request("txtWANIPDLCI") <> "" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_wan_ip_dlci", adVarChar, adParamInput, 10, (Request("txtWANIPDLCI")))
		else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_wan_ip_dlci", adVarChar, adParamInput, 10, null)
		end if

		if Request("txtPOSIPDLCI") <> "" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_pos_ip_dlci", adVarChar, adParamInput, 10, (Request("txtPOSIPDLCI")))
		else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_pos_ip_dlci", adVarChar, adParamInput, 10, null)
		end if

		if Request("selPacketSize") <> "" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_packet_size", adVarChar, adParamInput, 5, (Request("selPacketSize")))
		else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_packet_size", adVarChar, adParamInput, 5, null)
		end if

		if Request("hdnAddressID") <> "" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_site_address_id", adNumeric, adParamInput,, (Request("hdnAddressID")))
		else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_site_address_id", adNumeric, adParamInput,, null)
		end if

		if Request("txtOrderNumber") <> "" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_order_number", adVarChar, adParamInput, 15, (Request("txtOrderNumber")))
		else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_order_number", adVarChar, adParamInput, 15, null)
		end if
		if Request("txtNodeNumber") <> "" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_node_number", adNumeric, adParamInput,, (Request("txtNodeNumber")))
		else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_node_number", adNumberic, adParamInput,, null)
		end if


		'	TEST
			'dim objparm

  			'Response.Write "<BR>"

  			'for each objparm in cmdInsertObj.Parameters
  			 ' Response.Write "<b>" & objparm.name & "</b>"
  			  'Response.Write " has <b>size</b>:  " & objparm.Size & " "
  			  'Response.Write " and <b>value</b>:  " & objparm.value & " "
  			  'Response.Write " and <b>datatype</b>:  " & objparm.Type & "<br> "
  		    'next

  		   'Response.Write "<BR>"
  		   'Response.Write "<BR>"


  		   'Response.Write "<b> count = " & cmdInsertObj.Parameters.count & "<br>"
  		   'Response.Write "<BR>"
  		   'Response.Write "Parameter Values are as follows: " & "<BR>"

  			   'dim nx
  			   'for nx=0 to cmdInsertObj.Parameters.count-1
  			   'Response.Write " parm value= " & cmdInsertObj.Parameters.Item(nx).Value  & "<br>"
  			   'next

  		'	TEST
  		'Response.end


		' execute the insert object

		cmdInsertObj.Execute

			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT CREATE NEW OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			else
				 lngTCID = cmdInsertObj.Parameters("p_tail_circuit_id").Value
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


			'get the tail_circuit_detail stored insert procedure <schema.package.procedure>
			cmdDeleteObj.CommandText = "sma_sp_userid.spk_sma_rsas_inter.sp_tail_circuit_delete"

			'create the delete parameters
			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_tail_circuit_id", adNumeric, adParamInput, 22, CLng(lngTCID))
			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, ,CDate(Request("hdnUpdateDateTime")))

			'execute the delete object

			cmdDeleteObj.Execute

			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			end if
			if Request("back") = "true" then
				Response.Redirect "RSAST3GWTailCircuit.asp?action=update&GWID="&lngGWID&"&CustID="&lngCustID
			end if
			lngTCID = 0
			strWinMessage = "Record deleted successfully."

end select


'*************************
'end save, insert, delete*
'*************************


'ok, now go get the detailed Tail Circuit information

'declare the connection and sql variables
Dim strSQL
Dim strSelectClause
Dim strFromClause
Dim strWhereClause
Dim rsRSAS
DIM rsNode
DIM rsNodeNumber

'declare the detail variables which will be used to populate the
'displayed and hidden fields

dim strUpdateDateTime

'for Gateway Details <READ ONLY>

dim strGatewayCustomer
dim strGatewayIPID
dim strGatewayDLCIIP
dim strGatewayDLCIX25
dim strCustomerID
dim strGWCircuitNo

'for Tail Circuit Details

dim strTailCircuitID
dim strTCGatewayID
dim strWANIPID
dim strLANIPID
dim strNodeName
dim strNodeNumber
dim strTailCircuitNumber
dim strWANIPDLCI
dim strPOSIPDLCI
dim strOrderNumber

'for IP Address Details

dim strIPAddress
dim strIPAddressID
dim strSubnetMask
dim strGatewayIPAddress
dim strCode
dim strAvailable
dim strComments
dim strLocation
dim strStreet, strMunicipality, strProvince, strCountry, intSiteAddressID, StrSiteAddress

'for device details

dim strDeviceID
dim strDeviceTCID
dim local_dna
dim poll_code
dim host_dna_id


'connect to the database using databaseconnect inc/smaconstants inc connection string
'<<CONNECT>>

'use the sqlstring to extract the necessary information from the database

	'if  lngTCID <> 0 then
	if  strAction <> "new" and lngtcid <> 0 then
		strSelectClause = "SELECT " &_
					"T1.GATEWAY_ID, " & _
					"T1.GATEWAY_IP_ID, " & _
					"T1.GATEWAY_DLCI_IP, " & _
					"T1.GATEWAY_DLCI_X25, " & _
					"T1.CUSTOMER_ID, " & _
					"to_char(t3.create_date_time,'MON-DD-YYYY HH24:MI:SS') as create_date, " & _
					"sma_sp_userid.spk_sma_library.sf_get_full_username(t3.create_real_userid) as create_real_userid, " & _
					"to_char(t3.update_date_time,'MON-DD-YYYY HH24:MI:SS') as update_date, " & _
					"sma_sp_userid.spk_sma_library.sf_get_full_username(t3.update_real_userid) as update_real_userid, " & _
					"t3.update_date_time as last_update_date_time, " & _
					"t3.record_status_ind, " & _
					"T2.IP_ADDRESS_ID, " & _
					"T2.IP_ADDRESS, " & _
					"T2.SUBNET_MASK, " & _
					"T2.GATEWAY_IP_ADDRESS, " & _
					"T2.CODE, " & _
					"T2.LOCATION, " & _
					"T2.AVAILABLE, " & _
					"T2.COMMENTS, " & _
					"T3.TAIL_CIRCUIT_ID, " & _
					"T3.GATEWAY_ID, " & _
					"T3.WAN_IP_ID, " & _
					"T3.LAN_IP_ID, " & _
					"T3.NETWORK_SPEED, " & _
					"T3.PORT_SPEED, " & _
					"T3.NODE_NAME, " & _
					"T3.TAIL_CIRCUIT_NUMBER, " & _
					"T3.WAN_IP_DLCI, " & _
					"T3.POS_IP_DLCI, " & _
					"T3.PACKET_SIZE, " & _
					"T3.SITE_ADDRESS_ID, " & _
						"NVL(T4.BUILDING_NAME,'<NO BUILDING SPECIFIED>')||CHR(13)||CHR(10)|| " &_
						"NVL(T4.LONG_STREET_NAME,'NO STREET ADDRESS')||CHR(13)||CHR(10)|| " &_
						"NVL(T4.MUNICIPALITY_NAME,'NO MUNICIPALITY')||' '|| " &_
						"NVL(T4.PROVINCE_STATE_LCODE,'NO PROVINCE')||CHR(13)||CHR(10)|| " &_
						"NVL(T4.COUNTRY_LCODE,'NO COUNTRY') ADDRESS, " &_
					"T3.ORDER_NUMBER, " &_
					"T5.CUSTOMER_NAME, " &_
					"T1.GATEWAY_CIRCUIT_NUMBER, " &_
					"T3.NODE_NUMBER,  " &_
					"T4.MUNICIPALITY_NAME, T4.PROVINCE_STATE_LCODE, T4.COUNTRY_LCODE, t4.long_street_name, " &_
					"T2.SUBNET_MASK, T2.LOCATION "

		strFromClause =	" from CRP.RSAS_GATEWAY  T1, " &_
					"CRP.RSAS_IP_ADDRESS  T2, " & _
					"CRP.RSAS_TAIL_CIRCUIT  T3, " & _
					"CRP.ADDRESS  T4, " & _
					"CRP.CUSTOMER T5 "

		 strWhereClause = " where " & _
					"T3.TAIL_CIRCUIT_ID = " &  lngTCID & " AND " & _
					"T1.GATEWAY_ID = " &  lngGWID & " AND " & _
					"T1.GATEWAY_ID = T3.GATEWAY_ID AND " & _
					"T1.GATEWAY_IP_ID = T2.IP_ADDRESS_ID (+) AND " & _
					"T3.SITE_ADDRESS_ID = T4.ADDRESS_ID (+) AND " & _
					"T1.CUSTOMER_ID = T5.CUSTOMER_ID(+) "

		strSQL =  strSelectClause & strFromClause & strWhereClause
		set rsRSAS = Server.CreateObject("ADODB.Recordset")

		rsRSAS.CursorLocation = adUseClient
		rsRSAS.Open strSQL, objConn

		if err then
			DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
		end if
		if rsRSAS.EOF then
			DisplayError "BACK", "", 999, "CANNOT FIND REQUESTED OBJECT", "EOF condition occurred in rsRSAS recordset."
		end if
		set rsRSAS.ActiveConnection = nothing

		'Get the number of nodes
		strSQL = "SELECT DISTINCT TC.NODE_NAME " &_
			"FROM " &_
			"CRP.RSAS_TAIL_CIRCUIT	TC, "&_
			"CRP.RSAS_IP_ADDRESS	WAN_IP, "&_
			"CRP.RSAS_IP_ADDRESS	LAN_IP, "&_
			"CRP.CUSTOMER			CU, "&_
			"CRP.ADDRESS			AD "&_
			"WHERE " &_
		"TC.GATEWAY_ID= "& lngGWID & " " &_
		"AND TC.WAN_IP_ID = WAN_IP.IP_ADDRESS_ID (+) "&_
		"AND TC.LAN_IP_ID = LAN_IP.IP_ADDRESS_ID (+) "&_
		"AND TC.SITE_ADDRESS_ID = AD.ADDRESS_ID (+) "

		'set and open the tail circuit recordset and database connection
		set rsNode = Server.CreateObject("ADODB.Recordset")

		rsNode.CursorLocation = adUseClient
		rsNode.Open strSQL, objConn
		if err then
			DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
		end if

		intTotalNodes =  rsNode.RecordCount
		set rsNode.ActiveConnection = nothing
	else

		strSQL = "SELECT DISTINCT " &_
					"T1.GATEWAY_ID, " & _
					"T1.GATEWAY_IP_ID, " & _
					"T1.GATEWAY_DLCI_IP, " & _
					"T1.GATEWAY_DLCI_X25, " & _
					"T1.CUSTOMER_ID, " & _
					"T4.address_id AS site_address_id, " & _
						"NVL(T4.BUILDING_NAME,'<NO BUILDING SPECIFIED>')||CHR(13)||CHR(10)|| " &_
						"NVL(T4.LONG_STREET_NAME,'NO STREET ADDRESS')||CHR(13)||CHR(10)|| " &_
						"NVL(T4.MUNICIPALITY_NAME,'NO MUNICIPALITY')||' '|| " &_
						"NVL(T4.PROVINCE_STATE_LCODE,'NO PROVINCE')||CHR(13)||CHR(10)|| " &_
						"NVL(T4.COUNTRY_LCODE,'NO COUNTRY') ADDRESS, " &_
					"T4.MUNICIPALITY_NAME, T4.PROVINCE_STATE_LCODE, T4.COUNTRY_LCODE, T4.LONG_STREET_NAME, " &_
					"T2.GATEWAY_IP_ADDRESS, " & _
					"T5.CUSTOMER_NAME, " &_
					"T1.GATEWAY_CIRCUIT_NUMBER, " &_
					"T2.SUBNET_MASK, T2.LOCATION " &_
				"FROM CRP.RSAS_GATEWAY  T1, " &_
					"CRP.RSAS_IP_ADDRESS  T2, " & _
					"CRP.ADDRESS  T4, " & _
					"CRP.CUSTOMER T5 " &_
				"WHERE " & _
					"T1.GATEWAY_ID = " &  lngGWID & " AND " & _
					"T1.GATEWAY_IP_ID = T2.IP_ADDRESS_ID (+) AND " & _
					"T1.CUSTOMER_ID = T5.CUSTOMER_ID(+) AND " &_
					"T1.CUSTOMER_ID = T4.ADDRESS_ID (+)"

		'show SQL for debugging if necessary by using>>
		'Response.Write "Gateway SQL is <BR>" & strSQL	 & "<br>"
		'Response.end
		'set and open the tail circuit recordset and database connection
		set rsRSAS = Server.CreateObject("ADODB.Recordset")

		rsRSAS.CursorLocation = adUseClient
		rsRSAS.Open strSQL, objConn
		if err then
			DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
		end if

		set rsRSAS.ActiveConnection = nothing


		intTotalNodes = 0

        'Get the next node number for this customer
		strSQL = "SELECT DECODE(MAX(TC.NODE_NUMBER), null, 1, MAX(TC.NODE_NUMBER) + 1) AS NEXT_NODE_NUMBER " &_
			"FROM " &_
			"CRP.RSAS_GATEWAY GW, "&_
			"CRP.RSAS_TAIL_CIRCUIT TC " &_
			"WHERE " &_
		    "GW.CUSTOMER_ID= " & lngCustID & " " &_
		    "AND GW.GATEWAY_ID = TC.GATEWAY_ID "

		'set and open the recordsets and database connection
		set rsNodeNumber = Server.CreateObject("ADODB.Recordset")
		rsNodeNumber.Open strSQL, objConn
		if err then
			DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
		end if

	    strNodeNumber = rsNodeNumber("next_node_number")

	end if ' end if goes here Adam

'************************************************
'************************************************

'extract the info for the WANIP Dropdown


dim objRsWANIP
dim strWANIPSQL

   'get a list of available WAN IPs for dropdown

	  strWANIPSQL = "SELECT " & _
					  "WAN_IP.IP_ADDRESS_ID, " & _
					  "WAN_IP.IP_ADDRESS, " & _
					  "WAN_IP.SUBNET_MASK " & _
					"FROM " & _
					  "crp.rsas_IP_address WAN_IP, " & _
					  "crp.rsas_gateway GW, " & _
					  "crp.rsas_IP_address GW_IP " & _
					"WHERE " & _
					  "GW.gateway_ip_id = GW_IP.IP_ADDRESS_ID AND " & _
					  "WAN_IP.GATEWAY_IP_ADDRESS = GW_IP.IP_ADDRESS AND " & _
					  "GW.gateway_id = " & lngGWID & " AND " & _
					  "((WAN_IP.CODE = 'WANIP' AND " & _
					  "WAN_IP.AVAILABLE = 'Y') "

	if (lngTCID <> 0) then

					 strWANIPSQL = strWANIPSQL & " OR " & _
					  "(WAN_IP.IP_ADDRESS_ID IN " & _
	  				    "(SELECT TC.WAN_IP_ID " & _
					     "FROM crp.RSAS_TAIL_CIRCUIT TC " & _
					     "WHERE TC.tail_circuit_id = " & lngTCID & ")" & _
					     "))"
	else
		  strWANIPSQL = strWANIPSQL & " )"

	end if

'Response.Write "wanip sql is <BR>" & (strWANIPSQL)
'Response.End

	set objRsWANIP = objConn.Execute(strWANIPSQL)

'************************************************
'************************************************

'extract the info for the LANIP Dropdown


dim objRsLANIP
dim strLANIPSQL

   'get a list of available LAN IPs for dropdown
  			'"FROM " & _
			'		  "crp.rsas_IP_address LAN_IP, " & _
			'		  "crp.rsas_gateway GW, " & _
			'		  "crp.rsas_IP_address GW_IP " & _
			'		"WHERE " & _
			'		  "GW.gateway_id = " & lngGWID & " AND " & _
			'		  "GW.gateway_ip_id = GW_IP.IP_ADDRESS_ID AND " & _
			'		  "LAN_IP.GATEWAY_IP_ADDRESS= GW_IP.IP_ADDRESS AND " & _
			'


		  '****this sql was being used from here

		  'strLANIPSQL = "SELECT " & _
			'		  "LAN_IP.IP_ADDRESS_ID, " & _
			'		  "LAN_IP.IP_ADDRESS, " & _
			'		  "LAN_IP.SUBNET_MASK, " & _
			'		  "LAN_IP.LOCATION "&_
			'		"FROM " & _
			'		  "crp.rsas_IP_address LAN_IP " & _
			'		"WHERE " & _
			'		  "((LAN_IP.CODE= 'LANIP' AND " & _
			'		  "LAN_IP.AVAILABLE='Y')"

			'***to here


	'strLANIPSQL = "SELECT " & _
	'				  "LAN_IP.IP_ADDRESS_ID, " & _
	'				  "LAN_IP.IP_ADDRESS, " & _
	'				  "LAN_IP.SUBNET_MASK, " & _
	'				  "LAN_IP.LOCATION "&_
	'				"FROM " & _
	'				  "crp.rsas_IP_address LAN_IP, " & _
	'				  "crp.rsas_gateway GW, " & _
	'				  "crp.rsas_IP_address GW_IP " & _
	'				"WHERE " & _
	'				  "GW.gateway_ip_id = GW_IP.IP_ADDRESS_ID AND " & _
	'				  "LAN_IP.GATEWAY_IP_ADDRESS = GW_IP.IP_ADDRESS AND " & _
	'				  "GW.gateway_id = " & lngGWID & " AND " & _
	'				  "((LAN_IP.CODE = 'LANIP' AND " & _
	'				  "LAN_IP.AVAILABLE = 'Y') "
	strLANIPSQL = "SELECT " & _
					  "LAN_IP.IP_ADDRESS_ID, " & _
					  "LAN_IP.IP_ADDRESS, " & _
					  "LAN_IP.SUBNET_MASK, " & _
					  "LAN_IP.LOCATION "&_
					"FROM " & _
					  "crp.rsas_IP_address LAN_IP " & _
					"WHERE " & _
					  "((LAN_IP.CODE = 'LANIP' AND " & _
					  "LAN_IP.AVAILABLE = 'Y') "

	if (lngTCID <> 0) then

	  strLANIPSQL = strLANIPSQL & " OR " & _
					  "(LAN_IP.IP_ADDRESS_ID IN " & _
	  				    "(SELECT TC.LAN_IP_ID " & _
					     "FROM crp.RSAS_TAIL_CIRCUIT TC " & _
					     "WHERE TC.tail_circuit_id = " & lngTCID & ")" & _
					     "))"
	else
	  strLANIPSQL = strLANIPSQL & ") "
	end if


	'Response.Write (strLANIPSQL)
	'Response.End
	set objRsLANIP = objConn.Execute(strLANIPSQL)



'************************************************
'************************************************

'extract the info for the Port Speeds Dropdown


dim objRsPortSpeed
dim strPortSpeedSQL

   'get a list of available Port Speeds for dropdown

	'if (lngTCID <> 0) then


	strPortSpeedSQL = "SELECT " & _
						"C.CODE_ID, " & _
						"C.CODE_TYPE_CODE, " & _
						"C.CODE_DESC, " & _
						"C.CODE_ORDER " & _
					  "FROM " & _
						"CRP.RSAS_CODE C" & _
					  " WHERE " & _
					  "(C.CODE_TYPE_CODE='PS') " & _
					  "ORDER BY " & _
					     "C.CODE_ORDER "

	'end if
	'Response.Write (strPortSpeedSQL)
	'Response.End
	set objRsPortSpeed = objConn.Execute(strPortSpeedSQL)



'************************************************
'************************************************

'extract the info for the Network Speeds Dropdown


dim objRsNetworkSpeed
dim strNWSpeedSQL

   'get a list of available Network Speeds for dropdown

	'if (lngTCID <> 0) then


	strNWSpeedSQL = "SELECT " & _
						"C.CODE_ID, " & _
						"C.CODE_TYPE_CODE, " & _
						"C.CODE_DESC, " & _
						"C.CODE_ORDER " & _
					"FROM " & _
						"CRP.RSAS_CODE C" & _
					" WHERE " & _
						"(C.CODE_TYPE_CODE='NS') " & _
					"ORDER BY " & _
						"C.CODE_ORDER "

	'end if
	'Response.Write (strNWSpeedSQL)
	'Response.End
	set objRsNetworkSpeed = objConn.Execute(strNWSpeedSQL)

'************************************************
'fill the detail variables with values from the main recordset



'for Gateway Details <READ ONLY>

'response.write "lngtcid = " & lngTCID


if lngTCID <> 0 then


		strGatewayIPID       = rsRSAS("gateway_ip_id")
		strGatewayDLCIIP     = rsRSAS("gateway_dlci_ip")
		strGatewayDLCIX25    = rsRSAS("gateway_dlci_x25")
		strCustomerID        = rsRSAS("customer_id")
		strGatewayCustomer   = rsRSAS("customer_name")
		strGWCircuitNo       = rsRSAS("gateway_circuit_number")

		strStreet            = rsRSAS("long_street_name")
		strMunicipality      = rsRSAS("municipality_name")
		strProvince          = rsRSAS("province_state_lcode")
		strCountry           = rsRSAS("country_lcode")

		'for Tail Circuit Details
		strWANIPID           = rsRSAS("wan_ip_id")
		strLANIPID           = rsRSAS("lan_ip_id")
		strNodeName          = rsRSAS("node_name")
		strTailCircuitNumber = rsRSAS("tail_circuit_number")
		strWANIPDLCI         = rsRSAS("wan_ip_dlci")
		strPOSIPDLCI         = rsRSAS("pos_ip_dlci")
		intSiteAddressID     = rsRSAS("site_address_id")
		strSiteAddress       = rsRSAS("address")
		strOrderNumber       = rsRSAS("order_number")
		strNodeNumber        = rsRSAS("node_number")

		'for IP Address Details
		strIPAddress         = rsRSAS("ip_address")
		strSubnetMask        = rsRSAS("subnet_mask")
		strGatewayIPAddress  = rsRSAS("gateway_ip_address")
		strCode              = rsRSAS("code")
		strAvailable         = rsRSAS("available")
		strComments          = rsRSAS("comments")
		strLocation          = rsRSAS("location")

else
		strGatewayIPAddress  = rsRSAS("gateway_ip_address")
		strSubnetMask        = rsRSAS("subnet_mask")
		strLocation          = rsRSAS("location")

		strCustomerID        = rsRSAS("customer_id")
		strGatewayCustomer   = rsRSAS("customer_name")
		intSiteAddressID     = rsRSAS("site_address_id")
		strSiteAddress       = rsRSAS("address")

		strStreet            = rsRSAS("long_street_name")
		strMunicipality      = rsRSAS("municipality_name")
		strProvince          = rsRSAS("province_state_lcode")
		strCountry           = rsRSAS("country_lcode")

		strGatewayIPID       = rsRSAS("gateway_ip_id")
		strGatewayDLCIIP     = rsRSAS("gateway_dlci_ip")
		strGatewayDLCIX25    = rsRSAS("gateway_dlci_x25")
		strGWCircuitNo       = rsRSAS("gateway_circuit_number")

end if

%>


<html>


<head>

	<META HTTP-EQUIV="Pragma" CONTENT="no-cache">

	<link rel="stylesheet" type="text/css" href="stylesheets/styles.css">
<TITLE>POS PLUS Tail Circuit Detail</TITLE>
	<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
	<script type="text/javascript" SRC="AccessLevels.js"></script>

	<script LANGUAGE="JavaScript">

	var strWinMessage = '<%=strWinMessage%>';
    var intAccessLevel = '<%=intAccessLevel%>';
    var bolNeedToSave = false ;


setPageTitle("SMA - POS PLUS Detail");


    function window_onload() {

			iFrame_display();

			}


	//-----------------beginning of iFrame Javascript-----------------------------------------------


	function iFrame_display(){

		//called whenever a refresh of the iFrame is needed
		//loads iFrame at onload

		if ((intAccessLevel & intConst_Access_ReadOnly) == intConst_Access_ReadOnly) {
			document.frames("aifr").document.location.href = 'RSAST3DevList.asp?TailCircuitID=<%response.write lngTCID%>&hdnCustomerID=<%response.write lngCustID%>';
		}
		else {alert('Access Denied. You do not have access to device list. Please contact your system administrator.')}
	}

	function btn_iFrmAdd(){
	//open a blank form
		if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
			alert('Access Denied - can not access device. Please contact your system administrator.');
			return (false);
		}
		if (document.frmRSASDetail.hdnTailCircuitID.value == ""){
			alert('At this time you cannot create a device. You must save the tail circuit first.');
			return (false);
		}
		var NewDevWin;
		var strMasterID = "<%=lngTCID%>";
		var strCustomerID = "<%=lngCustID%>";
		NewDevWin=window.open("RSAST3DevDetail.asp?action=new&hdnDeviceID=0&masterID="+strMasterID+"&hdnCustomerID="+strCustomerID,"NewGWWin","toolbar=no,status=yes,width=800px,height=250px,left=100px,top=200,menubar=no,resize=no");
		NewDevWin.focus();
	}

	function btn_iFrmUpdate(){
		//open a detail form where the user can modify the alias
		if ((intAccessLevel & intConst_Access_Update) != intConst_Access_Update) {
			alert('Access Denied - can not update device. Please contact your system administrator.');
			return;
		}
		var NewWin;
		var strDeviceID = document.frames("aifr").document.frmIFR.hdnDeviceID.value;
		if (strDeviceID == ""){
			alert("Please select a device or click NEW to create a new device.");
			return;
		}
		var strMasterID = "<%=lngTCID%>";
		var strCustomerID = "<%=lngCustID%>";
		NewWin=window.open("RSAST3DevDetail.asp?action=update&hdnDeviceID="+strDeviceID+"&masterId="+strMasterID+"&hdnCustomerID="+strCustomerID,"NewGWWin","toolbar=no,status=yes,width=800px,height=250px,left=100px,top=200,menubar=no,resize=no");
		NewWin.focus();
	}

	function btn_iFrmDelete(){
		//delete selected row
		if ((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) {
			alert('Access Denied - can not delete device. Please contact your system administrator.');
			return;
		}
		var strDeviceID = document.frames("aifr").document.frmIFR.hdnDeviceID.value;
		if (strDeviceID == "") {
			alert("Please select a device or click ADD to create a new device.");
			return;
		}
		var strLastUpdate = document.frames("aifr").document.frmIFR.hdnUpdateDateTime.value;
		if (confirm("Are you sure you want to delete this device?")){
			document.frames("aifr").document.location.href = "RSAST3DevDetail.asp?action=delete&back=true&hdnDeviceID="+strDeviceID+"&masterID=<%=lngTCID%>&hdnLastUpdate="+strLastUpdate;
		}
	}
	//-----------------end of iFrame Javascript-----------------------------------------------


	//DONE


	function fct_NewTailCircuitEntry(){

	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create)

			{

			alert('Access denied. Please contact your system administrator.');
			return (false);

			}

			self.document.location.href = "RSAST3Detail.asp?action=new&hdnGatewayId=<%=lngGWID%>&hdnTailCircuitID=0&hdnCustomerID=<%=lngCustID%>";

	}



	//fields required for validation!


	//WAN IP
	//LAN IP
	//Tail Circuit NUmber
	//Site Address
	//Order Number


	function fct_OnSave(){

	//var strComments = document.frmRSASDetail.txtComments.value;

	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create)

			{
				alert('Access denied. Please contact your system administrator.');
				return (false);
			}

			else
			{


				if (document.frmRSASDetail.txtTailCircuitNumber.value == "" )
					{
						alert('Please select a Tail Circuit Number');
						//document.frmRSASDetail.btnPartNumLookup.focus();
						return(false);

					}

				if (document.frmRSASDetail.textAddress.value == "" )
					{
						alert('Please enter a site address');
						//document.frmRSASDetail.btnPartNumLookup.focus();
						return(false);

					}

				if (document.frmRSASDetail.txtOrderNumber.value == "" )
					{
						alert('Please enter an Order Number');
						//document.frmRSASDetail.btnPartNumLookup.focus();
						return(false);

					}

				/*
				if (strComments.length > 255)
					{
						alert('Comments can be at most 255 characters.\n\nYou entered ' + strComments.length + ' character(s).');
						document.frmRSASDetail.txtComments.focus();
						return false;
					}
				**/




					document.frmRSASDetail.action.value = "SAVE";
					document.frmRSASDetail.hdnGatewayID.value = "<%=lngGWID%>";
					document.frmRSASDetail.hdnTailCircuitID.value = "<%=lngTCID%>";
					bolNeedToSave = false;
					document.frmRSASDetail.submit();
					return(true);

			}


    }


	//OK

	function fct_onDelete() {

	var  lngTCID = document.frmRSASDetail.hdnTailCircuitID.value;
	var strUpdateDate = document.frmRSASDetail.hdnUpdateDateTime.value;
	var intDevCount;

	if ((intAccessLevel && intConst_Access_Delete)!= intConst_Access_Delete)
		{
			alert('Access denied. Please contact your system administrator.');
			return (false);
		}



	intDevCount = document.frames("aifr").document.frmIFR.hdnDevCount.value;


	if (intDevCount > 0)
		{	alert('Devices exist for this tail circuit, please remove them before deleting the tail circuit.');
					return;
		}
	if (confirm('Do you really want to delete this object?'))
		{
			document.location = "RSAST3Detail.asp?action=DELETE&back=false&hdnGatewayID="+ <%=lngGWID%>+"&hdnTailCircuitID="+ <%=lngTCID%>+"&hdnUpdateDateTime="+strUpdateDate ;
		}
	}




	//OK

	function fct_onReset() {
		if(confirm('All changes will be lost. Do you really want to reset the page?')){
			bolNeedToSave = false ;
			document.location = 'RSAST3Detail.asp?hdnTailCircuitID='+ '<%=lngTCID%>' ;
		}
	}

	//OK

    function fct_onChange(){

		bolNeedToSave = true;
	}


	//OK


	function fct_onBeforeUnload()

	{


		document.frmRSASDetail.btnSave.focus();

		if ((intAccessLevel & intConst_Access_Update) == intConst_Access_Update || (intAccessLevel & intConst_Access_Create) == intConst_Access_Create)
		{
			if (bolNeedToSave == true)
			{
				event.returnValue = "There is unsaved data on the screen. To save changes, click CANCEL below then click SAVE on the main form.";
			}
		}

	}

	function fct_onClose()
	{

		window.close();
	}

	/************************************
	   *BEGIN LOOKUP BUTTON FUNCTIONS*
	*************************************/

	//OK

	function btnCustomerLookup_onclick(CustService) {

	var strCustomerName = window.frmRSASDetail.txtCustomerName.value ;

	if (strCustomerName != "" )
		{
			SetCookie("CustomerName", strCustomerName);

		}

	SetCookie("ServiceEnd", CustService);
	SetCookie("WinName", 'Popup');
	fct_onChange();
	//opens CustomerCriteria.asp for search frame, CustomerCriteriaList for list
	window.open('SearchFrame.asp?fraSrc=Cust', 'Popup', 'top=50, left=100, WIDTH=840, HEIGHT=600'  ) ;

	}


	function btnAddressLookup_onclick() {

	var lngCustomerID    = window.frmRSASDetail.hdnCustomerID.value ;
	var strCustomerName  = window.frmRSASDetail.txtCustomerName.value ;
	var intSiteAddressID = window.frmRSASDetail.hdnAddressID.value ;
	var strStreet        = window.frmRSASDetail.hdnStreet.value ;
//	var strMunicipality  = Window.frmRSASDetail.hdnMunicipality.value ;
	var strProvince      = window.frmRSASDetail.hdnProvince.value ;
	var strCountry       = window.frmRSASDetail.hdnCountry.value ;
	var strSiteAddress   = window.frmRSASDetail.textAddress.value;

	if (strCustomerName != "" )
	   {
		SetCookie("CustomerID", lngCustomerID) ;
		SetCookie("CustomerName", strCustomerName) ;
		SetCookie("WinName", "Popup") ;
	    SetCookie("Street", strStreet) ;
		SetCookie("Municipality", "") ;
		SetCookie("Province", strProvince) ;
		SetCookie("Country", strCountry);
		SetCookie("SiteAddressID", intSiteAddressID);
		SetCookie("SiteAddress", strSiteAddress);
	   }
	SetCookie("WinName", "Simple") ;
 	window.open('SearchFrame.asp?fraSrc=Address', 'Simple', 'top=50, left=75, WIDTH=850, HEIGHT=650' ) ;
	}
// 	window.open('SearchFrame.asp?fraSrc=RSAST3Addr', 'Popup', 'top=50, left=75, WIDTH=850, HEIGHT=650' ) ;

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
    /*
	function btnTemplate_onclick() {

				{
					//alert("");
				}

		}

	*/

	function body_onUnload(){
	 opener.document.frmRSASGWDetail.btn_iFrameRefresh.click();
	}


	</script>

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function btnSave_onafterupdate() {

}

//-->
</SCRIPT>
</head>

<body onLoad="fct_DisplayStatus(strWinMessage);window_onload();" onbeforeunload="fct_onBeforeUnload();" onUnload="body_onUnload();">


<form name="frmRSASDetail" action="RSAST3Detail.asp" method="POST">

    <input id="action" name="action" type="hidden" value="">

	<INPUT type="hidden" name=hdnTailCircuitID value="<% if lngTCID <> 0 then Response.Write rsRSAS("tail_circuit_id") end if %>">
	<INPUT type="hidden" name=hdnGatewayID value="<% if lngTCID <> 0 then Response.Write rsRSAS("gateway_id") end if %>">

	<INPUT type="hidden" name=hdnGatewayIPID value="<% if lngTCID <> 0 then Response.Write rsRSAS("tail_circuit_id") end if %>">
	<INPUT type="hidden" name=hdnCustomerID value="<% if lngTCID <> 0 then Response.Write rsRSAS("customer_id") end if %>">

	<!--INPUT type="hidden" name=hdnWANIPID value="<% if lngTCID <> 0 then Response.Write rsRSAS("wan_ip_id") end if %>"-->
	<INPUT type="hidden" name=hdnLANIPID value="<% if lngTCID <> 0 then Response.Write rsRSAS("lan_ip_id") end if %>">
	<!--INPUT type="text" name=hdnAddressID value="<% if lngTCID <> 0 then Response.Write rsRSAS("site_address_id") end if %>"-->

	<!--for address cookie-->
	<INPUT type="hidden" name=hdnAddressID    value="<% if lngTCID <> 0 then Response.Write rsRSAS("site_address_id") end if %>">
	<INPUT type="hidden" name=hdnStreet       value="<% if lngTCID <> 0 then Response.Write rsRSAS("long_street_name") end if%>">
	<INPUT type="hidden" name=hdnMunicipality value="<% if lngTCID <> 0 then Response.Write rsRSAS("municipality_name") end if%>">
	<INPUT type="hidden" name=hdnProvince     value="<% if lngTCID <> 0 then Response.Write rsRSAS("province_state_lcode") end if%>">
	<INPUT type="hidden" name=hdnCountry      value="<% if lngTCID <> 0 then Response.Write rsRSAS("country_lcode") end if%>">
	<!--/cookie-->

	<input name="hdnUpdateDateTime" type="hidden" value="<%if  lngTCID <> 0 then  Response.Write rsRSAS("last_update_date_time") else Response.Write """""" end if%>">


	<!-- user interface -->

	<table border="0" width="100%" cols="6">

	<thead>
		<tr>
			<td colspan = "6" align="left"><strong>Gateway Detail</strong></td>
		</tr>
	</thead>


	<tbody>

	<tr>
		<td align="right" width="40%">GW Circuit Number</td>
		<td align="left" width="25%" >
			<input name="txtGWCircuitNo" disabled  type="text" value="<%=strGWCircuitNo%>" onChange="fct_onChange();">
		</td>
	</tr>

	<tr>
		<td align="right" width="40%">GW Router T1 Serial IP Address</td>
		<td align="left" width="25%" >
			<input name="txtGatewayIP" disabled  type="text" size="50" value="<%=strGatewayIPAddress%> subnet: <%=strSubnetMask%> location: <%=strLocation%>" onChange="fct_onChange();">
		</td>

		<td align="right" width="30%">Gateway DLCI(X25)</td>
		<td align="left" width="25%" >
			<input name="txtGatewayDLCIX25" disabled type="text"  size="20" maxlength="20" value="<%=strGatewayDLCIX25%>" onChange="fct_onChange();">
		</td>

	</tr>
	<tr>
		<td align="right" width="40%">Gateway Customer</td>
		<td align="left" width="20%" >
			<input name="txtGatewayCustomer" disabled type="text" size="50" value="<%=strGatewayCustomer%>" onChange="fct_onChange();">
		</td>

		<td align="right" width="50%">Gateway DLCI(IP)</td>
		<td align="left" width="25%" >
			<input name="txtGatewayDLCIIP" disabled type="text" size="20" maxlength="20" value="<%=strGatewayDLCIIP%>" onChange="fct_onChange();">
		</td>
	</tr>

	</tr>

	</tbody>

</table>

	<table border="0" cols="4">

	<thead>
	<tr>
 	  <td align="Left"  colspan="2"><strong>Tail Circuit Detail</strong></td>
	  <td align="Right" colspan="2"><Strong><%if intTotalNodes <> 0 then Response.write "Total POS PLUS Sites: " & intTotalNodes%></strong></td>
	</tr>
	</thead>

	<TR>
	<td align="right" >Tail Circuit Number<font color="red">*</font></td>
	<td align="left"  >
		<input name="txtTailCircuitNumber" type="text" tabindex=1 size="19" maxlength="19" value="<%=strTailCircuitNumber%>" onChange="fct_onChange();">
	</TR>

	<TR>
	<TD ALIGN="right"  NOWRAP>Network Speeds</TD>

	<TD align=left>
	<SELECT id=selNetworkSpeeds name=selNetworkSpeeds tabindex=2 onchange ="fct_onChange();">
		<OPTION></OPTION>

				<%Do while Not objRsNetworkSpeed.EOF
				Response.write "<OPTION "
					if lngTCID <> 0 then
					  if rsRSAS("NETWORK_SPEED")<> "" then
						if CInt(objRsNetworkSpeed("CODE_ID")) = CInt(rsRSAS("NETWORK_SPEED")) then
							Response.Write " SELECTED "
						end if
					  end if
					else
					  if trim(objRsNetworkSpeed("CODE_DESC")) = "56K" then
					     Response.Write " SELECTED "
					  end if
					end if
				Response.Write 	" VALUE=" &objRsNetworkSpeed("CODE_ID")& ">" &objRsNetworkSpeed("CODE_DESC") & "</OPTION>" &vbCrLf
				objRsNetworkSpeed.MoveNext
				Loop%>

			</SELECT></TD>

	<TD ALIGN="right"  NOWRAP>Port Speeds</TD>

	<TD align=left>
	<SELECT id=selPortSpeeds name=selPortSpeeds tabindex=3 onchange ="fct_onChange();">
		<OPTION></OPTION>

				<%Do while Not objRsPortSpeed.EOF
				Response.write "<OPTION "
					if lngTCID <> 0 then
					  if rsRSAS("PORT_SPEED")<> "" then
					  	if CInt(objRsPortSpeed("CODE_ID")) = CInt(rsRSAS("PORT_SPEED")) then
							Response.Write " SELECTED "
						end if
					  end if
					else
					  if trim(objRsPortSpeed("CODE_DESC")) = "9600" then
					     Response.Write " SELECTED "
					  end if
					end if
				Response.Write 	" VALUE=" &objRsPortSpeed("CODE_ID")& ">" &objRsPortSpeed("CODE_DESC") & "</OPTION>" &vbCrLf
				'code, description,

				objRsPortSpeed.MoveNext
				Loop%>

			</SELECT></TD>
	</TD>
	</TR>

	<tr>
		<td align="right" >Node Name</td>
		<td align="left">
			<input name="txtNodeName" type="text" tabindex=4 size="10" maxlength="10" value="<%=strNodeName%>" onChange="fct_onChange();"></td>
		<td align="right" colspan="1">Node Number</td>
		<td align="left" colspan="1">
			<input name="txtNodeNumber" type="text" tabindex=5 size="10" maxlength="10" value="<%=strNodeNumber%>" onChange="fct_onChange();"></td>
	</tr>

	<TR>
	<TD ALIGN="right"  NOWRAP>WAN IP Port Address</TD>

	<TD align=left colspan="2">
	<SELECT id=selWANIP name=selWANIP tabindex=6 style="HEIGHT: 22px; WIDTH: 300px" onchange ="fct_onChange();">
		<OPTION></OPTION>
				<%Do while Not objRsWANIP.EOF ' and objRsWANIP.BOF
				Response.write "<OPTION "
					if lngTCID <> 0 then
					  if rsRSAS("WAN_IP_ID") <> "" then

						if  CInt(objRsWANIP("IP_ADDRESS_ID")) = CInt(rsRSAS("WAN_IP_ID")) then
							Response.Write " SELECTED "
						end if

					  end if
					end if

				Response.Write 	" VALUE=" &objRsWANIP(0)& ">" &objRsWANIP(1) &" subnet:" &objRsWANIP(2) & "</OPTION>" &vbCrLf
				objRsWANIP.MoveNext
				Loop%>

			</SELECT></TD>

	</TR>

	<TR>
	<TD ALIGN="right"  NOWRAP>LAN IP Port Address</TD>

	<TD align=left colspan="2">
	<SELECT id=selLANIP name=selLANIP tabindex=7 style="HEIGHT: 22px; WIDTH: 350px" onchange ="fct_onChange();">
			<OPTION></OPTION>
				<%Do while Not objRsLANIP.EOF
				Response.write "<OPTION "
					if lngTCID <> 0 then
					  if rsRSAS("LAN_IP_ID")<>"" then

						if  CInt(objRsLANIP("IP_ADDRESS_ID")) = CInt(rsRSAS("LAN_IP_ID")) then
							Response.Write " SELECTED "
						end if

					  end if
					end if

				Response.Write 	" VALUE=" &objRsLANIP("IP_ADDRESS_ID")& ">" &objRsLANIP(1) &" subnet:" &objRsLANIP(2) &" location:" &objRsLANIP(3) & "</OPTION>" &vbCrLf
				objRsLANIP.MoveNext
				Loop%>

		</SELECT></TD>

	</TR>

	<tr>
		<td align="right">Customer<font color="red">*</font></td>
		<td align="left" colspan="2">
			<input name="txtCustomerName" type="text"  size="50" disabled value = "<%=strGatewayCustomer%>" onChange="fct_onChange();">
			<INPUT align=right type="button"  tabindex=8 name=btnCustomerLookup  value="..." onclick="return btnCustomerLookup_onclick();">
		</td>
	</tr>

	<tr>
		<td align="right">Site Address<font color="red">*</font></td>
		<td colspan=2>
			<TEXTAREA align=left rows=3 cols=50 id=textAddress disabled name=textAddress><%Response.write routineHTMLString(strSiteAddress)%></TEXTAREA>
			<INPUT align=right type="button" tabindex=9 name=btnAddressLookup  value="..." onclick="return btnAddressLookup_onclick();fct_onChange();">
		</td>
	</tr>

	<tr>
		<td align="right">Order Number<font color="red">*</font></td>
		<td align="left" colspan="2">
			<input name="txtOrderNumber" type="text" tabindex=10 size="15" maxlength="15" value="<%=strOrderNumber%>" onChange="fct_onChange();">
		</td>
	</tr>

	<td valign="top" align="right">Devices</td>
	<td rowSpan="10" colspan="3" valign="top">
				<iframe id=aifr width=90% height=100 tabindex=11 src="" scrolling=yes marginheight=1 marginwidth=1></iframe>
				<br>
				<input type="button" value="Refresh" tabindex=12 name="btn_iFrameRefresh" onClick="iFrame_display();">
				<input type="button" value="New"    tabindex=13 name="btn_iFrameAdd"     onClick="btn_iFrmAdd();">
				<input type="button" value="Update" tabindex=14 name="btn_iFrameUpdate"  onClick="btn_iFrmUpdate();">
				<input type="button" value="Delete" tabindex=15 name="btn_iFrameDelete"  onClick="btn_iFrmDelete();">
	</td>


	<hr>

	<tr>
		<td width="25%">&nbsp;</td>
	</tr>

	<tfoot>

		<table border="1" align="right">
		<tr>
			<td width="100%" colspan="4" align="right" bordercolor=Black>
			<input name="btnReset" type="button" value="Close" tabindex=18 style="width: 2cm" onClick="return fct_onClose();">
			<input name="btnDelete" type="button" value="Delete" tabindex=19 style="width: 2cm" onClick="return fct_onDelete();">
			<input name="btnNew" type="button" value="New" tabindex=20 style="width: 2cm" onClick="return fct_NewTailCircuitEntry();">
			<input id="btnSave" name="btnSave" type="button" tabindex=21 value="Save" style="width: 2cm" onClick="return fct_OnSave();" onafterupdate="return btnSave_onafterupdate()">

			</td>
		</tr>
		</table>
	</tfoot>

</table>

	<br>
	<br>
	<br>
	<fieldset>
	<legend align="right"><b>Audit Information</b></legend>
	<div SIZE="8pt" ALIGN="RIGHT">
		Record Status Indicator
		<input align="left" name="txtRecordStatusInd" type="text" style="HEIGHT: 20px; WIDTH: 18px" disabled value="<%if  lngTCID <> 0 then Response.Write rsRSAS("record_status_ind") else Response.Write """""" end if%>">&nbsp;&nbsp;&nbsp;
		Create Date
		<input align="center" name="txtRecordStatusInd" type="text" style="HEIGHT: 20px; WIDTH: 140px" disabled value="<%if  lngTCID <> 0 then Response.Write rsRSAS("create_date") else Response.Write """""" end if%>">&nbsp;
		Created By
		<input align="right" name="txtRecordStatusInd" type="text" style="HEIGHT: 20px; WIDTH: 100px" disabled value="<%if  lngTCID <> 0 then Response.Write rsRSAS("create_real_userid") else Response.Write """""" end if%>"><br>
		Update Date
		<input align="center" name="txtRecordStatusInd" type="text" style="HEIGHT: 20px; WIDTH: 140px" disabled value="<%if  lngTCID <> 0 then Response.Write rsRSAS("update_date") else Response.Write """""" end if%>">
		Updated By
		<input align="right" name="txtRecordStatusInd" type="text" style="HEIGHT: 20px; WIDTH: 100px" disabled value="<%if  lngTCID <> 0 then Response.Write rsRSAS("update_real_userid") else Response.Write """""" end if%>">
	</div>
	</fieldset>

</form>

<%

	if  lngGWID <> 0 then

		rsRSAS.close
		set rsRSAS = nothing

	end if

	objConn.close
	set objConn = nothing

%>


</body>
</html>

