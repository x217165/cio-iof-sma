<%@  language="VBScript" %>
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
       20-Oct-03   DTy          Revise Port Information layout.asd;
								Expand screen, add 'Clone' button.
	13-Nov-03   DTy		Make 'IP Address' mandatory.
	16-Aug-04   ACheung	Add LYNX repair priority
	18-Apr-05    MWong  Grey out NetCracker items.
	19-Feb-06   ACheung Enable Port Information, Repair Priority and Comments.
			    Make other 15+ field readonly or disabled.
	29-Sep-06   ACheung Distinguish CIU MOs. Reenable customer and service location for non-CIU MOs.
	07-May-08   ACheung Add CLLI Code (Geocode) as part of the Service Location selection
**************************************************************************************
-->

<%
'check user's rights
    
dim intAccessLevel
intAccessLevel = CInt(CheckLogon(strConst_ManagedObjects))

   
if intAccessLevel and intConst_Access_ReadOnly <> intConst_Access_ReadOnly then
     
 DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to managed object. Please contact your system administrator"
end if
    
dim sql, ne_id, strNE_ID, strWinMessage, strServLocAddress, bolClone , displayAdditionalInfo, rsNE

dim intAccessLevelForSNMP_read
intAccessLevelForSNMP_read = CInt(CheckLogon(strConst_SNMP))
    displayAdditionalInfo = false

    
dim intAccessLevelForSNMP_write
intAccessLevelForSNMP_write = CInt(CheckLogon(strConst_SNMP_write))
   
    diM  CanDisplayButton 
    CanDisplayButton = (intAccessLevelForSNMP_read > 0 ) or (intAccessLevelForSNMP_write > 0)
    
    
'=== Script Fields relating to table record fields:
Dim network_element_name, network_element_type_code, network_element_desc
Dim managed_ip_address, trusted_host_mac_address, out_of_band_dialup, support_group, serial_no, Status4
Dim barcode, remedy_contact

bolClone = false

'get requested network element id
strNE_ID = Request("ne_id")
'if strNE_ID = "" then
'	strNE_ID = Request.Cookies("txtNE_ID")
'end if

dim strRealUserID
strRealUserID =  Session("username")  'To be removed and set back

'field disabler default to not disabled
    
dim strDisable, strDisble_customer_service_location
dim strReadonly, strReadonlystyle
dim strcurrent_NEType, strcurrent_contactrole, strcurrent_supportgroup
dim satelliteflag, visibleflag



strDisable = ""
strDisble_customer_service_location = ""
strReadonly = ""
strReadonlystyle = ""

strcurrent_NEType  = ""
strcurrent_contactrole = ""
strcurrent_supportgroup = ""

Dim satelliteWrite, objsatellite
Dim uservisibilityWrite, objuserview

dim smaroles
smaroles = session("SMARoles")
if instr(smaroles, "SMA2 - Design Engineer")>0 or instr(smaroles, "SMA2 - Operations")>0 or instr(smaroles, "SMA2 - Super User")>0 then
   satelliteWrite = "Y"
   uservisibilityWrite = "Y"
else
   satelliteWrite = "N"
   uservisibilityWrite = "N"
end if 
     
select case Request("txtFrmAction")
	case "SAVE"
      
    dim MgmtSpace,MGMT_SPACE_ID

     MgmtSpace = Request("selMGMT_SPACE_NAME")
    MGMT_SPACE_ID = Request("selMGMT_SPACE_NAME")
'sql = "SELECT  MGMT_SPACE_ID from CRP.LCODE_MGMT_SPACE where MGMT_SPACE_NAME='"+Request("txtMGMT_SPACE_NAME")+"'"
'set MgmtSpace=server.CreateObject("ADODB.Recordset")
'MgmtSpace.CursorLocation = adUseClient
'MgmtSpace.Open sql, objConn
'if err then
'     
'	DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET 1", err.Description
'end if
'
'if MgmtSpace.EOF then
'     
'	DisplayError "BACK", "", 999, "CANNOT CREATE OBJECT TYPE LIST", "EOF condition occured in rsNET recordset."
'end if
'MGMT_SPACE_ID=MgmtSpace("MGMT_SPACE_ID").Value
'set MgmtSpace.ActiveConnection = nothing

   
'sql = "SELECT  MGMT_SYSTEM_ID from CRP.LCODE_MGMT_SYSTEMS where MGMT_SYSTEM_NAME='"+Request("txtMGMT_SYSTEMS_NAME")+"'"
'set MgmtSystems=server.CreateObject("ADODB.Recordset")
'MgmtSystems.CursorLocation = adUseClient
'MgmtSystems.Open sql, objConn
'if err then
     
'	DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET 1", err.Description
'end if

'if MgmtSystems.EOF then
     
	'DisplayError "BACK", "", 999, "CANNOT CREATE OBJECT TYPE LIST", "EOF condition occured in rsNET recordset."
'end if
'MGMT_Systems=MgmtSystems("MGMT_SYSTEM_ID")
'set MgmtSystems.ActiveConnection = nothing
dim Tenent,TenentID
  Tenent =  Request("selTENANT_NAME")
    TenentID =  Request("selTENANT_NAME")
'sql = "SELECT TENANT_ID from CRP.LCODE_TENANT_CODE where TENANT_NAME='"+Request("txtTENANT_NAME")+"'"
'set Tenent=server.CreateObject("ADODB.Recordset")
'Tenent.CursorLocation = adUseClient
'Tenent.Open sql, objConn
'if err then
'     
'	DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET 1", err.Description
'end if
'
'if Tenent.EOF then
'     
'	DisplayError "BACK", "", 999, "CANNOT CREATE OBJECT TYPE LIST", "EOF condition occured in rsNET recordset."
'end if
'TenentID=Tenent("TENANT_ID").Value
'set Tenent.ActiveConnection = nothing

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
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_network_element_name", adVarChar,adParamInput, 30, Request.form("txtObjName"))					'varchar2(30)	means: Managed Object Name
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_network_element_type_code", adVarChar,adParamInput, 6, Request.form("selObjType"))				'varchar2(6)	means: Managed Object Type
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_customer_id", adNumeric, adParamInput,, CLng(Request("hdnCustomerID")))						'number(9)		means: Customer (id)
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_service_location", adNumeric, adParamInput,, CLng(Request("hdnServLocID")))					'number(9)		means: Service Location (id)
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_asset_catalogue_id", adNumeric, adParamInput,, CLng(Request("hdnAssetCatalogueID")))			'number(9)		means: Id of Make/Model/Port reference
		 cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_network_element_desc", adVarChar,adParamInput, 80, Request("txtObjDesc"))					'varchar2(80)	means: Managed Object Description
			'if Request("hdnAssetID") <> "" then
    '
'  dim assetId 
'   assetId= Request.Form("hdnAssetID")(1)
'  
'   assetId = Right(assetId,len(assetId)-1)
'  assetId = Left(assetId,len(assetId)-1)
'			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_asset_id", adNumeric, adParamInput,, clng(assetId)) 
'		else
'			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_asset_id", adNumeric, adParamInput,, 0)
		'	end if

    
    if Request("hdnAssetID") <> "" then
                 dim assetId
                 assetId= Request.Form("hdnAssetID")(1)
                  assetId = Right(assetId,len(assetId)-1)
                 assetId = Left(assetId,len(assetId)-1)
                 if assetId <> "" then
	               cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_asset_id", adNumeric, adParamInput,, CLng(assetId))
                 else
                  cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_asset_id", adNumeric, adParamInput,, null)
                 end if
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
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_remedy_contact", adVarChar, adParamInput,15, "LYNX")	'20121012
			end if
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , CDate(Request("hdnNEUpdateDateTime")))		'date			means: update_date_time from Network_Element record
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_repair_priority", adVarChar, adParamInput, 30, Request("selRepairPriority")	)       'LYNX repair priority

			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_satellit_flag", adVarChar, adParamInput, 30, Request("selsatellitflag"))
		    cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_user_visible_flag", adVarChar, adParamInput, 30, Request("selvisibleflag"))
     		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_ipv6_address", adVarChar,adParamInput,50,"10.10.12.15")
      if MGMT_SPACE_ID <> "" then
            cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_MGMT_SPACE_ID", adNumeric, adParamInput, 10, cint(MGMT_SPACE_ID ))
    else
    cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_MGMT_SPACE_ID", adNumeric, adParamInput, 10, null)
    end if

    
          '  cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_MGMT_SYSTEMS_ID",  adVarChar, adParamInput, 30, MGMT_Systems)
    if TenentID <> "" then
            cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_TENANT_ID", adNumeric, adParamInput, 10,cint( TenentID))
    else 
    cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_TENANT_ID", adNumeric, adParamInput, 10, null)
    end if 
            cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_nc_ne_role_lcode", adNumeric, adParamInput, 4, cint( Request("selNC_NE_ROLE_LCODE")))
            ' cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_DE_DESIGN_STATUS", adNumeric, adParamInput, 20, cint(Request("selDesignStatus")))
            
     		'call the insert stored proc
  			'dim objparm
  			'for each objparm in cmdUpdateObj.Parameters
  			'Response.Write "<b>" & objparm.name & "</b>"
  			'Response.Write " has size:  " & objparm.Size & " "
  			'Response.Write " and value:  " & objparm.value & " "
  			'Response.Write " and datatype:  " & objparm.Type & "<br> "
  		  'next

  		   'Response.Write "<b> count = " & cmdUpdateObj.Parameters.count & "<br>"
  		  'dim nx
  			' for nx=0 to cmdUpdateObj.Parameters.count-1
  			'   Response.Write " parm value= " & cmdUpdateObj.Parameters.Item(nx).Value  & "<br>"
  			'  next
		'response.end
           

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
			if Request("hdnAssetID") <> "" then
                 dim assetIdInsert
                 assetIdInsert= Request.Form("hdnAssetID")(1)
                  assetIdInsert = Right(assetIdInsert,len(assetIdInsert)-1)
                 assetIdInsert = Left(assetIdInsert,len(assetIdInsert)-1)
                 if assetIdInsert <> "" then
	               cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_asset_id", adNumeric, adParamInput,, CLng(assetIdInsert))
                 else
                  cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_asset_id", adNumeric, adParamInput,, null)
                 end if
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
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_remedy_contact", adVarChar, adParamInput, 15, "LYNX")
			end if
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_repair_priority", adVarChar, adParamInput, 30, Request("selRepairPriority"))	        'LYNX repair priority
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_satellit_flag", adVarChar, adParamInput, 30, Request("selsatellitflag"))
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_user_visible_flag", adVarChar, adParamInput, 30, Request("selvisibleflag"))
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_ipv6_address", adVarChar,adParamInput,50,null)
            if MGMT_SPACE_ID <> "" then
            cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_MGMT_SPACE_ID", adNumeric, adParamInput, 10, cint(MGMT_SPACE_ID ))
    else
      cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_MGMT_SPACE_ID", adNumeric, adParamInput, 10, null)
    end if
            'cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_MGMT_SYSTEM_ID ", adVarChar, adParamInput, 30, MGMT_Systems )
    if TenentID <> "" then
            cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_TENANT_ID", adNumeric, adParamInput, 10, cint(TenentID))
    else
    cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_TENANT_ID", adNumeric, adParamInput, 10, null)
    end if
    'cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_DE_DESIGN_STATUS", adNumeric, adParamInput, 20, cint(Request("selDesignStatus")))

            'cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_Usage Type", adVarChar, adParamInput, 10, strNC_NE_ROLE_LCODE)
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
			 'if strWarning <> "" then
			'	strWinLocation = "manobjdet.asp?ne_id=" & strNE_ID
			'	DisplayError "REFRESH", strWinLocation, "-20040", "OBJECT INSERTED", strWarning
             '  end if
   
    
			'on error goto 0

			if objConn.Errors.Count <> 0 then
     
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT CREATE NEW OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
        End If  


  
				'end if
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
			cmdDeleteObj.CommandText = "sma_sp_userid.spk_sma_mo_inter.sp_mo_delete" 'set record = 'D'
			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_network_element_id", adNumeric, adParamInput, ,CLng(Request("txtObjID")))					'number(9)		means: Managed Object Id
			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , CDate(Request("hdnNEUpdateDateTime")))		'date			means: update_date_time from Network_Element record
            cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("P_USER_ID", adVarChar , adParamInput, 30, strRealUserID)
			cmdDeleteObj.Execute
			if objConn.Errors.Count <> 0 then
     
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			end if

		cmdDeleteObj.CommandType = adCmdText
		'cmdDeleteObj.CommandText = "sma_sp_userid.spk_sma_mo_inter.sp_mo_port_delete"
         cmdDeleteObj.CommandText = "update CRP.NETWORK_ELEMENT_PORT set RECORD_STATUS_IND ='D',CI_STATUS_ID = 6 where NETWORK_ELEMENT_ID = " & CLng(Request("txtObjID"))

            cmdDeleteObj.Execute
			if objConn.Errors.Count <> 0 then
     
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE ports", objConn.Errors(0).Description
				objConn.Errors.Clear
			end if
			set strNE_ID =Request("txtObjID")
			strWinMessage = "Record deleted successfully."
    case "UNDELETE"
		'delete record
		if (intAccessLevel and intConst_Access_Delete <> intConst_Access_Delete) then
     
			DisplayError "BACK", "", 0, "DELETE DENIED", "You don't have access to delete managed objects. Please contact your system administrator"
		end if
			dim cmdUnDeleteObj
			set cmdUnDeleteObj = server.CreateObject("ADODB.Command")
			set cmdUnDeleteObj.ActiveConnection = objConn
			cmdUnDeleteObj.CommandType = adCmdText
			cmdUnDeleteObj.CommandText = " update  crp.network_element set RECORD_STATUS_IND = 'A'  WHERE network_element_id = " &  CLng(Request("txtObjID"))
			
			cmdUnDeleteObj.Execute
			if objConn.Errors.Count <> 0 then
     
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT Un DELETE OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			end if

    stop

    set rsNE=server.CreateObject("ADODB.Recordset")
	rsNE.CursorLocation = adUseClient
	rsNE.Open "select * from CRP.NETWORK_ELEMENT_PORT  where NETWORK_ELEMENT_ID = " & CLng(Request("txtObjID")), objConn

  
if err then
	DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
end if


if not rsNE.EOF then
	Response.Write("<script>alert('Reminder: Please ensure that the CI Status for the rows in the Port Information Table are correct as all rows have been moved to an Active status.');</script>")
end if    

        cmdUnDeleteObj.CommandType = adCmdText
		'cmdDeleteObj.CommandText = "sma_sp_userid.spk_sma_mo_inter.sp_mo_port_delete"
         cmdUnDeleteObj.CommandText = "update CRP.NETWORK_ELEMENT_PORT set RECORD_STATUS_IND ='A',CI_STATUS_ID = 3 where NETWORK_ELEMENT_ID = " & CLng(Request("txtObjID"))

            cmdUnDeleteObj.Execute
			if objConn.Errors.Count <> 0 then
     
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT activate ports", objConn.Errors(0).Description
				objConn.Errors.Clear
			end if
			set strNE_ID =Request("txtObjID")


			strWinMessage = "Record activated successfully."
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
			"NE.OWNED_BY_NC, " &_
			"NE.NC_NE_ROLE_LCODE, " &_
			"SLA.BUILDING_NAME, "&_
			"SLA.STREET, "&_
			"SLA.MUNICIPALITY_NAME, "&_
			"SLA.PROVINCE_STATE_LCODE, "&_
			"SLA.POSTAL_CODE_ZIP, "&_
			"TO_CHAR(NE.CREATE_DATE_TIME,'MON-DD-YYYY HH24:MI:SS') CREATE_DATE_TIME, "&_
			"sma_sp_userid.spk_sma_library.sf_get_full_username(NE.CREATE_REAL_USERID) as create_real_userid, "&_
			"TO_CHAR(NE.UPDATE_DATE_TIME,'MON-DD-YYYY HH24:MI:SS') UPDATE_DATE_TIME, "&_
			"sma_sp_userid.spk_sma_library.sf_get_full_username(NE.UPDATE_REAL_USERID) as update_real_userid, "&_
			"NE.RECORD_STATUS_IND RECORD_STATUS_IND, "&_
			"g.clli_code AS full_clli_code, "&_
			"NE.SERVICE_BY_SATELLITE, "&_
            "NE.CUSTOMER_VISIBILITY_FLAG ,"&_
             "LMS.MGMT_SPACE_NAME as mgmtSpaeName , "&_
          
             "LTC.TENANT_NAME as TenantName,LMS.mgmt_space_id, LTC.Tenant_id "&_
            
 
		"FROM " &_
			"CRP.NETWORK_ELEMENT				NE, "&_
			"CRP.ASSET							AST, "&_
			"CRP.ASSET_CATALOGUE				ACAT, "&_
			"CRP.MAKE							AMAK, "&_
    "CRP. LCODE_MGMT_SPACE							LMS, "&_ 
 
    "CRP.LCODE_TENANT_CODE                       LTC, "&_
			"CRP.MODEL							AMOD, "&_
			"CRP.PART_NUMBER					APN, "&_
			"CRP.CUSTOMER						CUS, "&_
			"CRP.SERVICE_LOCATION				SL, "&_
			"crp.service_location_geocode        slg, "&_
			"crp.lcode_geocodeid                   g, "&_
			"CRP.V_ADDRESS_CONSOLIDATED_STREET	SLA "&_
           
		"WHERE " &_
			"NE.CUSTOMER_ID = CUS.CUSTOMER_ID "&_
			"AND NE.ASSET_ID = AST.ASSET_ID (+) "&_
			"AND NE.ASSET_CATALOGUE_ID = ACAT.ASSET_CATALOGUE_ID "&_
			"AND ACAT.MAKE_ID = AMAK.MAKE_ID "&_
			"AND ACAT.MODEL_ID = AMOD.MODEL_ID "&_
			"AND ACAT.PART_NUMBER_ID = APN.PART_NUMBER_ID "&_
			"AND NE.SERVICE_LOCATION_ID = SL.SERVICE_LOCATION_ID "&_
			"AND sl.service_location_id = slg.service_location_id(+) "&_
			"AND slg.geocodeid_lcode = g.geocodeid_lcode(+) "&_
			"AND SL.ADDRESS_ID = SLA.ADDRESS_ID "&_
    "AND NE.MGMT_SPACE_ID = LMS.MGMT_SPACE_ID(+) "&_
    
     "AND NE.TENANT_ID = LTC.TENANT_ID(+) "&_
			"AND NE.NETWORK_ELEMENT_ID = " & strNE_ID


	'get the network element recordset
     
	if err then
     
		DisplayError "BACK", "", err.Number, "UNEXPECTED ERROR", err.Description
	end if
	set rsNE=server.CreateObject("ADODB.Recordset")
	rsNE.CursorLocation = adUseClient
	rsNE.Open sql, objConn
	if err then
     
		DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET 7", err.Description
	end if

	if rsNE.EOF then
     
		DisplayError "BACK", "", 999, "CANNOT FIND REQUESTED OBJECT", "EOF condition occured in rsNE recordset."
	else
	'preform at Service Location Address
    
   
    Request("status")=rsNE("RECORD_STATUS_IND").Value
    status4=rsNE("RECORD_STATUS_IND").Value
		if len(rsNE("building_name") ) > 0 then
			strServLocAddress = rsNE("building_name") & vbNewLine & rsNE("street") & vbNewLine&_
			rsNE("municipality_name") & " " & rsNE("province_state_lcode")
		else
			strServLocAddress = rsNE("street") & rsNE("municipality_name") & " " & rsNE("province_state_lcode")
		end if
	end if

    if(rsNE("nc_ne_role_lcode") = strConst_MO_LCODE) then
        displayAdditionalInfo = true
    end if
    dim showAdditionalDetails1 
  showAdditionalDetails1  = "visible"
    'check if NC ; MO with nc_ne_role_lcode=4 is CIU
	If  (rsNE("owned_by_nc").Value = "Y" and rsNE("nc_ne_role_lcode").Value = "4")  Then
		strDisble_customer_service_location = "DISABLED"
            showAdditionalDetails = "visible"
	End If
	If  rsNE("owned_by_nc") = "Y"  Then
		strDisable = "DISABLED"
		strReadonly = "READONLY"
	End If
	If  strReadonly = "READONLY"  Then
		strReadonlystyle = " style=color:silver"
	else
		strReadonlystyle = " style=color:black"
	End If


	satelliteflag = rsNE("SERVICE_BY_SATELLITE")


  'response.write "before reset satelliteflag = " &satelliteflag
  ' reset satelliteflag so it is N when it is N or null
  if satelliteflag="Y" then
     satelliteflag="Y"
  else
     satelliteflag="N"
  end if
 'response.write " after reset satelliteflag = " &satelliteflag

	visibleflag = rsNE("CUSTOMER_VISIBILITY_FLAG")

	 ' reset visibleflag so it is N when it is N or null
    if visibleflag="Y" then
       visibleflag="Y"
    else
       visibleflag="N"
    end if



	network_element_name = routineHtmlString(rsNE("NETWORK_ELEMENT_NAME").value)
	network_element_type_code = rsNE("NETWORK_ELEMENT_TYPE_CODE").value
	network_element_desc = rsNE("NETWORK_ELEMENT_DESC").value
	managed_ip_address = rsNE("MANAGED_IP_ADDRESS").value
	trusted_host_mac_address = rsNE("TRUSTED_HOST_MAC_ADDRESS").value
	out_of_band_dialup = rsNE("OUT_OF_BAND_DIALUP").value
	support_group = rsNE("REMEDY_SUPPORT_GROUP_ID").value
	serial_no = rsNE("SERIAL_NUMBER").value
	barcode = rsNE("BARCODE").value
	remedy_contact = rsNE("REMEDY_CONTACT_ROLE_ID").value

	if strDisable <> "" Then
		strcurrent_NEType = rsNE("NETWORK_ELEMENT_TYPE_CODE").value
		strcurrent_supportgroup = rsNE("REMEDY_SUPPORT_GROUP_ID").value
	end if
	strcurrent_contactrole =  rsNE("REMEDY_CONTACT_ROLE_ID").value

	'LC added for ITSM
	if strcurrent_contactrole = "ITSM" then
		strcurrent_contactrole = "ITSM"
	else
	 	strcurrent_contactrole = "LYNX"
	end if

	set rsNE.ActiveConnection = nothing
end if


'if strDisable <> "" Then
'	network_element_name = Request.form("txtObjName")
'	network_element_type_code = Request("selObjType")
'	network_element_desc = Request("txtObjDesc")
'	managed_ip_address = Request("txtIPAddress")
'	trusted_host_mac_address = Request("txtMACAddress")
'	out_of_band_dialup = Request("txtOBDialUp")
'	support_group = Request("selSupportGroup")
'	serial_no = Request("txtSerialNumber")
'	barcode = Request("txtBarcode")
'	remedy_contact = Request("selSupportContactRole")
'End If

'get the network element type recordset
dim rsNET
sql = "SELECT NETWORK_ELEMENT_TYPE_CODE FROM CRP.NETWORK_ELEMENT_TYPE WHERE RECORD_STATUS_IND='A' ORDER BY NETWORK_ELEMENT_TYPE_CODE"
set rsNET=server.CreateObject("ADODB.Recordset")
rsNET.CursorLocation = adUseClient
rsNET.Open sql, objConn
if err then
     
	DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET 1", err.Description
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
     
	DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET 2", err.Description
end if

if rsSG.EOF then
     
	DisplayError "BACK", "", 999, "CANNOT CREATE SUPPORT GROUP LIST", "EOF condition occured in rsSG recordset."
end if

set rsSG.ActiveConnection = nothing


'get the support contact role recordset
'2015March no need for contact_role, remove all related code lines 
'dim rsSCR
'sql = "SELECT REMEDY_CONTACT_ROLE_ID, CONTACT_ROLE_NAME  FROM CRP.V_REMEDY_CONTACT_ROLE  ORDER BY CONTACT_ROLE_NAME"
'set rsSCR=server.CreateObject("ADODB.Recordset")
'rsSCR.CursorLocation = adUseClient
'rsSCR.Open sql, objConn
'if err then
'	DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET 3", err.Description
'end if

'if rsSCR.EOF then
'	DisplayError "BACK", "", 999, "CANNOT CREATE CONTACT ROLE LIST", "EOF condition occured in rsSCR recordset."
'end if

'set rsSCR.ActiveConnection = nothing

'get current the support contact role


'LC below commented for ITSM:
'dim rsSCRcurrent, strcurrent_contactname
''sql = "SELECT REMEDY_CONTACT_ROLE_ID, CONTACT_ROLE_NAME  FROM CRP.V_REMEDY_CONTACT_ROLE  ORDER BY CONTACT_ROLE_NAME"
'if strcurrent_contactrole <> "" then
'	sql = "SELECT REMEDY_CONTACT_ROLE_ID, "&_
'      		"CONTACT_ROLE_NAME "&_
'      		"FROM CRP.V_REMEDY_CONTACT_ROLE "&_
'      		"WHERE REMEDY_CONTACT_ROLE_ID = '" & strcurrent_contactrole & "'"

'	if err then
'		DisplayError "BACK", "", err.Number, "UNEXPECTED ERROR", err.Description
'	end if
'	set rsSCRcurrent=server.CreateObject("ADODB.Recordset")
'	rsSCRcurrent.CursorLocation = adUseClient
'	rsSCRcurrent.Open sql, objConn
'	if err then
'		DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET 4", err.Description
'	end if

'	if rsSCRcurrent.EOF then
'		DisplayError "BACK", "", 999, "CANNOT CREATE CONTACT ROLE LIST", "EOF condition occured in rsSCRcurrent recordset."
'	end if
'	strcurrent_contactname = rsSCRcurrent("CONTACT_ROLE_NAME").value
'       set rsSCRcurrent.ActiveConnection = nothing
'else
'	strcurrent_contactname = ""
'end if
'LC above are commented for ITSM:



'get the current support group recordset
dim rsSGcurrent, strcurrent_supportgroupname

if strcurrent_supportgroup <> "" then
	'sql = "SELECT REMEDY_SUPPORT_GROUP_ID, GROUP_NAME FROM CRP.V_REMEDY_SUPPORT_GROUP ORDER BY GROUP_NAME"
	sql = "SELECT REMEDY_SUPPORT_GROUP_ID, "&_
	      "GROUP_NAME "&_
	      "FROM CRP.V_REMEDY_SUPPORT_GROUP "&_
	      "WHERE REMEDY_SUPPORT_GROUP_ID = '" & strcurrent_supportgroup & "'"

	if err then
     
		 DisplayError "BACK", "", err.Number, "UNEXPECTED ERROR", err.Description
	end if
	set rsSGcurrent=server.CreateObject("ADODB.Recordset")
	rsSGcurrent.CursorLocation = adUseClient
	rsSGcurrent.Open sql, objConn
	if err then
		DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET 5", err.Description
	end if

	if rsSGcurrent.EOF then
		DisplayError "BACK", "", 999, "CANNOT CREATE SUPPORT GROUP LIST", "EOF condition occured in rsSCRCurrent recordset."
	end if
	strcurrent_supportgroupname = rsSGcurrent("GROUP_NAME").value
	set rsSGcurrent.ActiveConnection = nothing
else
	strcurrent_supportgroupname = ""
end if

'Response.Write sql
'Response.End

'get the LYNX default repair priority
    dim rsLYNXrp,rsMgmtSpaceName,rsTenantName,rsUsageType,sqlMgmtSpaceName,sqlMgmtSystemName,sqlTenantName,sqlUsageType
    
    sql = "select LYNX_DEF_SEV_DESC, LYNX_DEF_SEV_LCODE from CRP.LCODE_LYNX_DEF_SEV where RECORD_STATUS_IND='A' ORDER BY LYNX_DEF_SEV_LCODE"
    sqlMgmtSpaceName = "select mgmt_space_id,mgmt_space_name from CRP.LCODE_MGMT_SPACE"
    sqlMgmtSystemName ="select mgmt_system_id,mgmt_system_name from CRP.LCODE_MGMT_SYSTEMS"
    sqlTenantName = "select Tenant_id,Tenant_Name from CRP.LCODE_TENANT_CODE "
    sqlUsageType  ="select NC_NE_ROLE_LCODE,NC_NE_ROLE_DESC from CRP.LCODE_NC_NE_ROLE"
    set rsLYNXrp=server.CreateObject("ADODB.Recordset")
    set rsMgmtSpaceName=server.CreateObject("ADODB.Recordset")
   ' set rsMgmtSystem=server.CreateObject("ADODB.Recordset")
    set rsTenantName =server.CreateObject("ADODB.Recordset")
    SET rsUsageType = server.CreateObject("ADODB.Recordset")

    rsUsageType.CursorLocation = adUseClient
    rsLYNXrp.CursorLocation = adUseClient
   ' rsMgmtSystem.CursorLocation = adUseClient
    rsMgmtSpaceName.CursorLocation = adUseClient
   ' rsMgmtSystem.Open sqlMgmtSystemName, objConn
    rsMgmtSpaceName.Open sqlMgmtSpaceName, objConn
    rsTenantName.Open sqlTenantName, objConn

     
    rsLYNXrp.Open sql, objConn
    rsUsageType.Open sqlUsageType, objConn
if err then
	DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET 6", err.Description
end if

if rsLYNXrp.EOF then
	DisplayError "BACK", "", 999, "CANNOT CREATE CONTACT ROLE LIST", "EOF condition occured in rsLYNXrp recordset."
end if

set rsLYNXrp.ActiveConnection = nothing

'Response.Write "<b>" & network_element_name & "</b>"
'Response.Write "<b>" & network_element_type_code & "</b>"
'Response.Write "<b>" & network_element_desc & "</b>"
'Response.Write "<b>" & managed_ip_address & "</b>"
'Response.Write "<b>" & trusted_host_mac_address & "</b>"
'Response.Write "<b>" & out_of_band_dialup & "</b>"
'Response.Write "<b>" & support_group & "</b>"
'Response.Write "<b>" & serial_no & "</b>"
'Response.Write "<b>" & barcode & "</b>"
'Response.Write "<b>" & remedy_contact & "</b>"
'Response.Write "<b>" & strcurrent_NEType  & "</b>"
'Response.Write "<b>" & strcurrent_contactname  & "</b>"
'Response.Write "<b>" & strcurrent_supportgroupname  & "</b>"

if strDisable <> "" Then
Response.Write "<B>NetCracker-controlled MO. Updates are limited to 4 fields: Name Alias, Port Information, Repair Priority and Comments.</B>"
End If
if (strDisable <> "" and strDisble_customer_service_location <> "DISABLED") Then
Response.Write "<P><B>NetCracker-controlled, non-CIU MO. Updates are allowed on 2 additional fields: Customer and Service Location.</B></P>"
End If

           
%>
<html>
<head>
    <meta name="GENERATOR" content="Microsoft FrontPage 12.0">
    <meta http-equiv="Pragma" content="no-cache">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>Managed Objects - Details</title>
</head>
<link href="stylesheets/styles.css" rel="stylesheet" type="text/css">
<script type="text/javascript" src="GeneralJavaFunctions.js"></script>
<script type="text/javascript" src="AccessLevels.js"></script>
<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.8.3/jquery.min.js"></script>

<script type="text/javascript">
    //set section title
    setPageTitle("SMA - Managed Object");

    var intAccessLevel=<%=CheckLogon(strConst_ManagedObjects)%>;
    var intAccessLevelDetail=<%=CheckLogon(strConst_ManagedObjects)%>;
    var intConst_MessageDisplay=<%=intConst_MessageDisplay%> ;
    var bolSaveRequired = false;

    function iFrame_display(){
        if ((intAccessLevelDetail & intConst_Access_ReadOnly) == intConst_Access_ReadOnly){
            
            document.getElementById("aifr").src = 'manobjalias.asp?ne_id=<%=strNE_ID %>' ;
            //document.location.href.replace("manobjdet", "manobjalias") ;
            // document.frames("aifr").document.location.href = 'manobjalias.asp?ne_id=<%if not bolClone then response.write strNE_ID end if%>';
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

        var doc;
        var iframeObject = document.getElementById('aifr'); // MUST have an ID
        if (iframeObject.contentDocument) { // DOM
            doc = iframeObject.contentDocument;
        } 
        else if (iframeObject.contentWindow) { // IE win
            doc = iframeObject.contentWindow.document;
        }


        var NewWin;
        var strAliasID =  doc.getElementsByName("hdnNameAliasID")[0].value; // document.frames("aifr").document.frmIFR.hdnNameAliasID.value;// hdnLastUpdate
        if (strAliasID == "") {
            alert("Please select an alias or click ADD NEW to create a new alias.");
            return;
        }
        var strMasterID = "<%=strNE_ID%>";
        NewWin=window.open("manobjaliasdetail.asp?action=update&aliasID="+strAliasID+"&masterID="+strMasterID ,"NewWin","toolbar=no,status=yes,width=1000px,height=350px,left=350px,top=400,menubar=no");
        NewWin.focus();
    }

    function btn_iFrmDelete(){
        //delete selected row

        
        if ((intAccessLevelDetail & intConst_Access_Delete) != intConst_Access_Delete){
            alert('Access denied. Please contact your system administrator.');
            return;
        }

        var doc;
        var iframeObject = document.getElementById('aifr'); // MUST have an ID
        if (iframeObject.contentDocument) { // DOM
            doc = iframeObject.contentDocument;
        } 
        else if (iframeObject.contentWindow) { // IE win
            doc = iframeObject.contentWindow.document;
        }


        var NewWin;
        var strAliasID =  doc.getElementsByName("hdnNameAliasID")[0].value; //document.frames("aifr").document.frmIFR.hdnNameAliasID.value;
        if (strAliasID == "") {
            alert("Please select an alias or click ADD NEW to create a new alias.");
            return;
        }
        var strLastUpdate = doc.getElementsByName("hdnLastUpdate")[0].value ;//document.frames("aifr").document.frmIFR.hdnLastUpdate.value; //document.getElementById("aifr2").contentDocument.getElementsByName("hdnLastUpdate")[0].value
        if (confirm("Do you want to delete this alias?")){
            // document.frames("aifr").document.location.href = "ManObjAliasDetail.asp?action=delete&back=true&aliasID="+strAliasID+"&masterID=<%if not bolDelete then response.write strNE_ID end if%>&hdnLastUpdate="+strLastUpdate;
            document.getElementById("aifr").src = "ManObjAliasDetail.asp?action=delete&back=true&aliasID="+strAliasID+"&masterID=<%if not bolDelete then response.write strNE_ID end if%>&hdnLastUpdate="+strLastUpdate;
        }
       
    }

   
    //javascript code not related to iFrame functionality
    var strWinMessage = "<%=strWinMessage%>";
    var strOuterNameAliasHTML;

    function iFrame2_display(){
        
        if ((intAccessLevelDetail & intConst_Access_ReadOnly) == intConst_Access_ReadOnly){
            if(document.location.href.indexOf("ne_id") > -1)
            {
                //  document.getElementById("aifr2").src = "";
                //var url = document.getElementById("aifr2").src ;

                //if(url)
                //{
                //    url += "&dt="+ new Date();
                //}
                // debugger;
                if(document.location.href.indexOf("manobjdet") <0)
                {
                    document.getElementById("aifr2").src = document.location.href.replace("manobjdet", "ManObjPort") + "&dt="+ new Date();;
                }
                else{
                    document.getElementById("aifr2").src = document.location.href.replace("manobjdet", "ManObjPort") ;
                }
                //  document.getElementById("aifr2").contentWindow.location.reload()
            }
            else{
                document.getElementById("aifr2").src = document.location.href.replace("manobjdet", "ManObjPort") + "?ne_id="+ "<%=strNE_ID %>";
            }
            //document.frames("aifr2").document.location.href = 
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
        NewWin=window.open("ManObjPortDetail.asp?action=new&masterID="+strMasterID +"&CustId=" + document.frmMODetails.hdnCustomerID.value ,"NewWin","toolbar=no,status=yes,width=1200px,height=600px,left=175px,top=200");
        NewWin.focus();
    }

    function btn_iFrm2Update(){
        //open a detail form where the user can modify the Port Information
        if ((intAccessLevelDetail & intConst_Access_Update) != intConst_Access_Update){
            alert('Access denied. Please contact your system administrator.');
            return;
        }
       
        var NewWin;
        var doc;
        var iframeObject = document.getElementById('aifr2'); // MUST have an ID
        if (iframeObject.contentDocument) { // DOM
            doc = iframeObject.contentDocument;
        } 
        else if (iframeObject.contentWindow) { // IE win
            doc = iframeObject.contentWindow.document;
        }

        var iFrame =  document.getElementById("aifr2");
        var strPortID = doc.getElementsByName("hdnPortID")[0].value ;//document.frames("aifr2").document.frmIFR2.hdnPortID.value;
        var strLastUpdate = doc.getElementsByName("hdnLastUpdate")[0].value; //document.frames("aifr2").document.frmIFR2.hdnLastUpdate.value;
        if (strPortID == "") {
            alert("Please select a Port Information record firt then click UPDATE to update the selected record.");
            return;
        }
        var strMasterID = "<%=strNE_ID%>";
        NewWin=window.open("ManObjPortDetail.asp?action=update&PortID="+strPortID+"&masterID="+strMasterID+"&hdnLastUpdate="+strLastUpdate +"&CustId=" + document.frmMODetails.hdnCustomerID.value,"NewWin","toolbar=no,status=yes,width=1200px,height=600px,left=175px,top=200");
        NewWin.focus();
    }

    function btn_iFrm2Clone(){
        //open a detail form where the user can modify the Port Information
        if ((intAccessLevelDetail & intConst_Access_Update) != intConst_Access_Update){
            alert('Access denied. Please contact your system administrator.');
            return;
        }
        var NewWin;
        var doc;
        var iframeObject = document.getElementById('aifr2'); // MUST have an ID
        if (iframeObject.contentDocument) { // DOM
            doc = iframeObject.contentDocument;
        } 
        else if (iframeObject.contentWindow) { // IE win
            doc = iframeObject.contentWindow.document;
        }

        var strPortID =  doc.getElementsByName("hdnPortID")[0].value;//document.frames("aifr2").document.frmIFR2.hdnPortID.value;
        if (strPortID == "") {
            alert("Please select a Port Information record first then click CLONE to create a new record using the selected record.");
            return;
        }
        var strMasterID = "<%=strNE_ID%>";
        NewWin=window.open("ManObjPortDetail.asp?action=clone&PortID="+strPortID+"&masterID="+strMasterID +"&CustId=" + document.frmMODetails.hdnCustomerID.value ,"NewWin","toolbar=no,status=yes,width=1400px,height=550px,left=175px,top=200");
        NewWin.focus();
    }

    function btn_iFrm2Delete(){
        //delete selected row
        if ((intAccessLevelDetail & intConst_Access_Delete) != intConst_Access_Delete){
            alert('Access denied. Please contact your system administrator.');
            return;
        }
        var doc;
        var iframeObject = document.getElementById('aifr2'); // MUST have an ID
        if (iframeObject.contentDocument) { // DOM
            doc = iframeObject.contentDocument;
        } 
        else if (iframeObject.contentWindow) { // IE win
            doc = iframeObject.contentWindow.document;
        }
        var strPortID = doc.getElementsByName("hdnPortID")[0].value;//document.frames("aifr2").document.frmIFR2.hdnPortID.value;
        if (strPortID == "") {
            alert("Please select a Port Information record first then click DELETE to delete the selected record.");
            return;
        }
        var strLastUpdate = doc.getElementsByName("hdnLastUpdate")[0].value; //document.frames("aifr2").document.frmIFR2.hdnLastUpdate.value;
        if (confirm("Do you want to delete this Port Information?")){
            //document.frames("aifr2").document.location.href = "ManObjPortDetail.asp?action=delete&back=true&PortID="+strPortID+"&masterID=<%if not bolDelete then response.write strNE_ID end if%>&hdnLastUpdate="+strLastUpdate;
            if( document.getElementsByName("btn_iFrame2Delete")[0].value == "Delete" )
            {
                document.getElementById("aifr2").src = "ManObjPortDetail.asp?action=delete&back=true&PortID="+strPortID+"&masterID=<%if not bolDelete then response.write strNE_ID end if%>&hdnLastUpdate="+strLastUpdate  +"&CustId=" + document.frmMODetails.hdnCustomerID.value;
            }
            else{
                document.getElementById("aifr2").src = "ManObjPortDetail.asp?action=undelete&back=true&PortID="+strPortID+"&masterID=<%if not bolDelete then response.write strNE_ID end if%>&hdnLastUpdate="+strLastUpdate  +"&CustId=" + document.frmMODetails.hdnCustomerID.value;
            }
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
        self.document.location.href = "manobjdet.asp?ne_id=";
        //	document.location="manobjdet.asp?ne_id=";

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

    function btn_onDelete(status){

        if ((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete){
            alert('Access denied. Please contact your system administrator.');
            return;
        }
        if( status == 'A')
        {
            
            if (confirm('Do you really want to delete this object?')){
                //submit the form
                document.frmMODetails.txtFrmAction.value = "DELETE";
                document.frmMODetails.submit();
            }
        }
        else{

            if (confirm('Do you really want to un delete this object?')){
                //submit the form
                document.frmMODetails.txtFrmAction.value = "UNDELETE";
                document.frmMODetails.submit();
            }
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
        var apart= strpart.split("Ã‚Â¿");
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

    function btnSNMP_onclick(event){
        var strOwner = 'CRP' ;				// owner name must be in Uppercase
        var strTableName = 'NETWORK_ELEMENT' ;		// table name must be in Uppercase
        var strRecordID = document.frmMODetails.txtObjID.value ;
        var URL ;
       
        if (strRecordID != ""  ){
            URL ='SNMPview.asp?NEId='+ strRecordID   ;
            window.open(URL, 'Popup', 'top=100, left=100, WIDTH=1248, HEIGHT=500,resizable=yes'  ) ;
            
            return false;
        }
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

<body onload="body_onLoad(strWinMessage);asset_onLoad();" onbeforeunload="body_onBeforeUnload();" onunload="body_onUnload();">
    <form name="frmMODetails" action="manobjdet.asp" method="POST" onreset="fct_onReset();">
        <input type="hidden" name="txtObjID" value="<%if not bolClone and (strNE_ID <> "") then Response.write rsNE("NETWORK_ELEMENT_ID") end if %>" onchange="fct_onChange();">
        <input type="hidden" name="hdnAssetID" value='<% if strNE_ID <> "" then Response.write """"&rsNE("ASSET_ID")&"""" else Response.Write """""" end if %>'>
        <input type="hidden" name="hdnAssetCatalogueID" value="<%if strNE_ID <> "" then Response.write rsNE("ASSET_CATALOGUE_ID") else Response.Write "1879"%>">
        <input type="hidden" name="hdnCustomerID" value="<% if strNE_ID <> "" then Response.Write rsNE("CUSTOMER_ID") end if %>">
        <input type="hidden" name="hdnServLocID" value="<% if strNE_ID <> "" then Response.Write rsNE("SERVICE_LOCATION_ID") end if %>">
        <input type="hidden" name="hdnNEUpdateDateTime" value="<% if strNE_ID <> "" then Response.Write  rsNE("UPDATE_DATE_TIME")  end if %>">

        <input type="hidden" name="hdnCity" value="<% if strNE_ID <> "" then Response.write """"&rsNE("MUNICIPALITY_NAME")&"""" else Response.Write """""" end if %>">
        <input type="hidden" name="hdnStreetName" value="<% if strNE_ID <> "" then Response.write """"&rsNE("STREET")&"""" else Response.Write """""" end if %>">
        <input type="hidden" name="hdnProvinceCode" value="<% if strNE_ID <> "" then Response.write """"&rsNE("PROVINCE_STATE_CODE")&"""" else Response.Write """""" end if %>">

        <input type="hidden" name="hdnNameAlias" value="">
        <input type="hidden" name="txtFrmAction" value="">
        <input type="hidden" name="status" value="">

        <table width="100%" cols="4" border="0">
            <thead>
                <tr>
                    <td colspan="3">Managed Object - Details</td>
                    <td align="right">
                        <select name="selQuickLink" size="1" onchange="qlink_onChange(this.value);" <%if strNE_ID = "" then Response.Write "disabled" end if%>>
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
                    <td width="15%" align="right">Object Name<font color="red">*</font></td>
                    <td width="35%">
                        <input <%=strReadonly%> <%=strReadonlystyle%> name="txtObjName" size="30" maxlength="30" value="<% if not bolClone and (strNE_ID <> "") then Response.write rsNE.Fields("NETWORK_ELEMENT_NAME") end if%>" onchange="fct_onChange();"></td>
                    <td width="15%" rowspan="5" valign="top" align="right">Name Alias</td>
                    <td width="35%" rowspan="5" valign="top">
                        <iframe id="aifr" width="100%" height="100" src="" scrolling="yes" marginheight="1" marginwidth="1"></iframe>
                        <br>
                        <input type="button" value="Refresh" name="btn_iFrameRefresh" onclick="iFrame_display();" class="button">
                        <input type="button" value="New" name="btn_iFrameAdd" onclick="btn_iFrmAdd();" class="button">
                        <input type="button" value="Update" name="btn_iFrameUpdate" onclick="btn_iFrmUpdate();" class="button">
                        <input type="button" value="Delete" name="btn_iFrameDelete" onclick="btn_iFrmDelete();" class="button">
                    </td>
                </tr>
                <tr>
                    <td align="right">Object Description</td>
                    <td>
                        <input <%=strReadonly%> <%=strReadonlystyle%> name="txtObjDesc" size="40" maxlength="80" value="<% if strNE_ID <> "" then Response.write   rsNE.Fields("NETWORK_ELEMENT_DESC") end if%>" onchange="fct_onChange();"></td>
                </tr>
                <tr>
                    <td align="right">Object Type<font color="red">*</font></td>
                    <td>
                        <select <%=strReadonlystyle%> name="selObjType" onchange="fct_onChange();">
                            <option></option>
                            <%
				if strDisable <> "" then
					Response.Write "<OPTION"
					Response.write " selected"
					Response.Write ">" & strcurrent_NEType & "</OPTION>"
				else
				while not rsNET.EOF
					Response.Write "<OPTION"
					if strNE_ID <> "" then if rsNET("NETWORK_ELEMENT_TYPE_CODE") = rsNE("NETWORK_ELEMENT_TYPE_CODE") then Response.write " selected"
					Response.Write ">" & rsNET("NETWORK_ELEMENT_TYPE_CODE") & "</OPTION>"
					rsNET.MoveNext
				wend
				rsNET.Close
				end if
                            %>
                        </select>
                    </td>
                </tr>
                <tr>
                    <td align="right">IP Address<font color="red">*</font></td>
                    <td>
                        <input name="txtIPAddress" <%=strReadonly%> <%=strReadonlystyle%> size="30" maxlength="30" value="<% if strNE_ID <> "" then Response.write   routineHtmlString(rsNE("MANAGED_IP_ADDRESS"))end if %>" onchange="fct_onChange();"></td>
                </tr>
                <tr>
                    <td align="right">MAC Address</td>
                    <td>
                        <input name="txtMACAddress" <%=strReadonly%> <%=strReadonlystyle%> size="30" maxlength="17" value="<% if strNE_ID <> "" then Response.write  routineHtmlString(rsNE("TRUSTED_HOST_MAC_ADDRESS"))end if %>" onchange="fct_onChange();"></td>
                </tr>
                <tr>
                    <td align="right">Asset</td>
                    <td>
                        <input disabled name="txtAssetName" size="40" maxlength="80" value="<% if strNE_ID <> "" then Response.write  routineHtmlString(rsNE("ASSET_TAC_NAME")) end if%>" title="To change click on the attached button" onchange="fct_onChange();">
                        <input <%=strDisable%> name="btnAsset" type="button" onclick="fct_lookupAsset();fct_onChange();" value="..." class="button" title="Click here to edit asset information.">&nbsp;<input name="btnAssetClear" type="button" onclick="    btnAssetClear_onclick();" value="X" type="button" title="Click here to clear the asset information."></td>
                    <td valign="top" align="right" rowspan="4">Comments</td>
                    <td valign="top" rowspan="4">
                        <textarea style="width=100%" name="txtComments" onchange="fct_onChange();" rows="6"><% if strNE_ID <> "" then Response.write  routineHtmlString(rsNE("COMMENTS")) end if%></textarea></td>
                </tr>
                <tr>
                    <td align="right">Make/Model/Part #<font color="red">*</font></td>
                    <td nowrap>
                        <input disabled name="txtAssetMake" size="12" value="<%if (strNE_ID <> "") AND (rsNE("ASSET_MAKE_DESC")    <> "<none>") then Response.write routineHtmlString(rsNE("ASSET_MAKE_DESC")) else Response.Write ""%>" title="To change click on the attached button" onchange="fct_onChange();">
                        <input disabled name="txtAssetModel" size="12" value="<%if (strNE_ID <> "") AND (rsNE("ASSET_MODEL_DESC")   <> "<none>") then Response.write routineHtmlString(rsNE("ASSET_MODEL_DESC")) else Response.Write ""%>">
                        <input disabled name="txtAssetPartNo" size="12" value="<%if (strNE_ID <> "") AND (rsNE("ASSET_PART_NO_DESC") <> "<none>") then Response.write routineHtmlString(rsNE("ASSET_PART_NO_DESC")) else Response.Write ""%>">
                        <input <%=strDisable%> name="btnAssetCatalog" type="button" onclick="fct_lookupAssetCatalog();fct_onChange();" value="..." class="button">
                    </td>
                    <tr>
                        <td align="right">Serial</td>
                        <td>
                            <input name="txtSerialNumber" <%=strReadonly%> <%=strReadonlystyle%> size="40" onchange="fct_onChange();" value="<% if strNE_ID <> "" then Response.write  routineHtmlString(rsNE("SERIAL_NUMBER")) end if%>">
                    </tr>
                <tr>
                    <td align="right">Barcode</td>
                    <td>
                        <input name="txtBarcode" <%=strReadonly%> <%=strReadonlystyle%> size="40" onchange="fct_onChange();" value="<% if strNE_ID <> "" then Response.write  routineHtmlString(rsNE("BARCODE")) end if%>">
                </tr>
                <tr>
                    <td align="right">Out of Band Dialup</td>
                    <td>
                        <input name="txtOBDialUp" <%=strReadonly%> <%=strReadonlystyle%> size="30" maxlength="30" value="<%if strNE_ID <> "" then Response.write   routineHtmlString(rsNE("OUT_OF_BAND_DIALUP")) end if%>" onchange="fct_onChange();"></td>
                    <tr>
                        <td align="right">Customer<font color="red">*</font></td>
                        <td>
                            <input disabled name="txtCustomerName" size="40" maxlength="50" value="<%if strNE_ID <> "" then Response.write  routineHtmlString(rsNE("CUSTOMER_NAME"))end if%>" title="To change click the attached button" onchange="fct_onChange();">
                            <input <%=strDisble_customer_service_location%> name="btnCustomer" type="button" value="..." onclick="fct_lookupCustomer();fct_onChange();" class="button"></td>
                        <td align="right">Cust Short Name</td>
                        <td>
                            <input disabled name="txtCustomerShortName" size="30" maxlength="15" value="<%if strNE_ID <> "" then Response.write   routineHtmlString(rsNE("CUSTOMER_SHORT_NAME"))end if%>" title="To change click the button attached to the customer name field" onchange="fct_onChange();"></td>
                    </tr>
                <tr>
                    <td align="right">Service Location<font color="red">*</font></td>
                    <td>
                        <input disabled name="txtServLocName" size="40" value="<% if strNE_ID <> "" then Response.write  routineHtmlString(rsNE("SERVICE_LOCATION_NAME")) end if%>" title="To change click on the attached button" onchange="fct_onChange();">
                        <input <%=strDisble_customer_service_location%> name="btnServiceLocation" type="button" value="..." onclick="fct_lookupServiceLocation();fct_onChange();" class="button"></td>
                    <td align="right">Support Contact Role</td>
                    <td>
                        <select name="selSupportContactRole" onchange="fct_onChange();" style="width: 72px">
                            <%
				if (strNE_ID <> "") then
					Response.Write "<OPTION VALUE=""" & strcurrent_contactrole & """"
					Response.write " selected"
					if strcurrent_contactrole ="ITSM" then
						response.write ">ITSM</option>"
						response.write "<option></option>"
					else
						response.write "></option>"
						response.write "<option>ITSM</option>"
					end if
			     else
			     	response.write "<option></option>"
			     	response.write "<option>ITSM</option>"
			     end if



                            %>
                        </select>
                    </td>
                </tr>
                <tr>
                    <td align="right">Service Location Address</td>
                    <td>
                        <textarea rows="3" style="width: 100%" id="txtServLocAddress" name="txtServLocAddress" disabled><%  if strNE_ID <> "" then Response.write  strServLocAddress end if%></textarea></td>
                    <td align="right" valign="top">Support Group<font color="red">*</font></td>
                    <td valign="top">
                        <select <%=strReadonlystyle%> name="selSupportGroup" onchange="fct_onChange();">
                            <option></option>
                            <%
			if strDisable <> "" then
				Response.Write "<OPTION VALUE=""" & strcurrent_supportgroup & """"
				Response.write " selected"
			        Response.Write ">" & routineHtmlString(strcurrent_supportgroupname) & "</OPTION>"
			else
				while not rsSG.EOF
					Response.Write "<OPTION"
					if strNE_ID <> "" then if rsSG("REMEDY_SUPPORT_GROUP_ID") = rsNE("REMEDY_SUPPORT_GROUP_ID") then Response.write " selected"
						Response.Write " VALUE="& rsSG("REMEDY_SUPPORT_GROUP_ID") &">" & routineHtmlString(rsSG("GROUP_NAME")) & "</OPTION>" &vbCrLf
					rsSG.MoveNext
				wend
				rsSG.Close
			end if
                            %>
                        </select>
                    </td>
                </tr>
                <% 
                           
                    
                    if CanDisplayButton = true then
                       Response.Write "<tr>   <td colspan='3'></td>  <td>  <button type='button'  onclick='btnSNMP_onclick();' >SNMP Device Information</button>   </td>    </tr>"
                     end if
                %>

                <tr style="visibility: '<% showAdditionalDetails %>'">


                    <td align="RIGHT" width="20%" nowrap>MGMT SPACE NAME</td>
                    <td width="80%">
                        <!--   <input size='40' maxlength='30' name='txtMGMT_SPACE_NAME' value='<%  rsNE("mgmtSpaeName") %>'>-->

                        <select name="selMGMT_SPACE_NAME" onchange="fct_onChange();" style="width: 72px">
                            <option selected value=""></option>
                            <%
				while not rsMgmtSpaceName.EOF
				Response.Write "<OPTION"
				if strNE_ID <> "" then 
                                  if IsNull( rsNE("mgmt_space_id").Value) <>  true and IsEmpty(rsNE("mgmt_space_id").Value) <> true then
                                  if CLng(rsNE("mgmt_space_id").Value) = CLng(rsMgmtSpaceName("mgmt_space_id")) then Response.write " selected" else Response.write " " end if
                                end if
                                end if
					Response.Write " VALUE="& rsMgmtSpaceName("mgmt_space_id") &">" & routineHtmlString(rsMgmtSpaceName("mgmt_space_name")) & "</OPTION>" &vbCrLf
				rsMgmtSpaceName.MoveNext
			wend
			rsMgmtSpaceName.Close



                            %>
                        </select>
                    </td>
                    <!--<td align="RIGHT" width="20%" nowrap>Design Status</td>
                    <td width="80%">
                        <select name="selDesignStatus">

                            <option <% if IsEmpty(rsNE("DE_DESIGN_STATUS").value) or rsNE("DE_DESIGN_STATUS").value = "1" then Response.Write "selected" end if %> value="1">Not Started</option>
                            <option  <% if rsNE("DE_DESIGN_STATUS").value = "2" then Response.Write"selected" end if %> value="2">In Progress</option>
                            <option  <% if rsNE("DE_DESIGN_STATUS").value = "3" then Response.Write"selected" end if %> value="3">Design Complete</option>
                            <option  <% if rsNE("DE_DESIGN_STATUS").value = "4" then Response.Write"selected" end if %> value="4">PoC Only</option>
                        </select>

                    </td>-->

                </tr>

                <tr style="visibility: '<% showAdditionalDetails %>'">


                    <td align="RIGHT" width="20%" nowrap>TENANT NAME</td>
                    <td width="80%">
                        <!-- <input size='40' maxlength='30' name='selTENANT_NAME' value='<%   rsNE("Tenant_id")  %>'>-->
                        <select name="selTENANT_NAME" onchange="fct_onChange();" style="width: 72px">
                            <option selected value=""></option>
                            <%
				while not rsTenantName.EOF
				Response.Write "<OPTION"
				if strNE_ID <> "" then 
                                
                                if IsNull( rsNE("Tenant_id").Value) <>  true and IsEmpty(rsNE("Tenant_id").Value) <> true then
                                  if CLng(rsNE("Tenant_id").Value) = CLng(rsTenantName("Tenant_id")) then Response.write " selected" else Response.write " " end if
                                end if
                                end if
					Response.Write " VALUE="& rsTenantName("Tenant_id") &">" & routineHtmlString(rsTenantName("Tenant_Name")) & "</OPTION>" &vbCrLf
				rsTenantName.MoveNext
			wend
			rsTenantName.Close



                            %>
                        </select>
                    </td>

                </tr>
                <tr style="visibility: '<% showAdditionalDetails %>'">
                    <td align="RIGHT" width="20%" nowrap>Usage Type</td>
                    <td width="80%">
                        <!--<input size='40' maxlength='30' name='txtUsage_Type' value='<% rsAlias("strNC_NE_ROLE_LCODE")  %>'>-->
                        <select name="selNC_NE_ROLE_LCODE" onchange="fct_onChange();" style="width: 72px">
                            <%
				while not rsUsageType.EOF
				Response.Write "<OPTION"
				if strNE_ID <> "" then if CLng(rsNE("NC_NE_ROLE_LCODE")) = CLng(rsUsageType("NC_NE_ROLE_LCODE")) then Response.write " selected"
					Response.Write " VALUE="& rsUsageType("NC_NE_ROLE_LCODE") &">" & routineHtmlString(rsUsageType("NC_NE_ROLE_DESC")) & "</OPTION>" &vbCrLf
				rsUsageType.MoveNext
			wend
			rsUsageType.Close



                            %>
                        </select>
                    </td>
                </tr>
                <tr>
                    <td align="right">CLLI Code (Geocode)</td>
                    <td>
                        <input name="txtGeocode" disabled size="18" maxlength="18" value="<%if strNE_ID <> "" then Response.write   routineHtmlString(rsNE("FULL_CLLI_CODE")) end if%>" onchange="fct_onChange();"></td>
                    <td align="right" valign="top" nowrap>Repair Priority</td>
                    <td valign="top">
                        <select id="selRepairPriority" name="selRepairPriority" onchange="fct_onChange();">
                            <%
			while not rsLYNXrp.EOF
				Response.Write "<OPTION"
				if strNE_ID <> "" then if CLng(rsNE("LYNX_DEF_SEV_LCODE")) = CLng(rsLYNXrp("LYNX_DEF_SEV_LCODE")) then Response.write " selected"
					Response.Write " VALUE="& rsLYNXrp("LYNX_DEF_SEV_LCODE") &">" & routineHtmlString(rsLYNXrp("LYNX_DEF_SEV_DESC")) & "</OPTION>" &vbCrLf
				rsLYNXrp.MoveNext
			wend
			rsLYNXrp.Close
                            %>
                        </select>
                    </td>
                </tr>
                <tr>
                    <td align="right">Serviced by Satellite </td>
                    <td valign="middle">
                        <select id="selsatellitflag" name="selsatellitflag" onchange="fct_onChange();">
                            <%if (strNE_ID <> "") then
					Response.Write "<OPTION VALUE=""" & satelliteflag & """"

					    Response.write " selected"

					if satelliteflag ="Y" then
						response.write ">Y</option>"
						if satelliteWrite = "N" then
							response.write "<option disabled>N</option>"
						else
							response.write "<option>N</option>"
						end if
					else
						response.write ">N</option>"
						if satelliteWrite = "N" then
							response.write "<option  disabled>Y</option>"
						else
							response.write "<option>Y</option>"
						end if
					end if
			  else
			  	    if satelliteWrite = "N" then
				     	response.write "<option  disabled>N</option>"
				     	response.write "<option  disabled>Y</option>"
				    else
				    	response.write "<option>N</option>"
				     	response.write "<option>Y</option>"

				    end if
			  end if
                            %>
                    </td>
                    <td align="right">Viewable by Customer Visibility Tools</td>
                    <td valign="middle">
                        <select id="selvisibleflag" name="selvisibleflag" onchange="fct_onChange();">
                            <%if (strNE_ID <> "") then
					Response.Write "<OPTION VALUE=""" & visibleflag & """"

					    Response.write " selected"

					if visibleflag ="Y" then
						response.write ">Y</option>"
						if uservisibilityWrite = "N" then
							response.write "<option disabled>N</option>"
						else
							response.write "<option>N</option>"
						end if
					else
						response.write ">N</option>"
						if uservisibilityWrite = "N" then
							response.write "<option  disabled>Y</option>"
						else
							response.write "<option>Y</option>"
						end if
					end if
			  else
			  	    if uservisibilityWrite = "N" then
				     	response.write "<option  disabled>N</option>"
				     	response.write "<option  disabled>Y</option>"
				    else
				    	response.write "<option>N</option>"
				     	response.write "<option>Y</option>"

				    end if
			  end if
                            %>
                    </td>
                </tr>
                <tr>
                    <td width="70%" colspan="4" rowspan="12" valign="top">
                        <iframe id="aifr2" width="100%" height="240" src="" scrolling="yes" marginheight="1" marginwidth="1"></iframe>
                        <br>
                        <input type="button" value="Refresh" name="btn_iFrame2Refresh" onclick="iFrame2_display();" class="button">
                        <input type="button" value="New" name="btn_iFrame2Add" onclick="btn_iFrm2Add();" class="button">
                        <input type="button" value="Clone" name="btn_iFrame2Clone" onclick="btn_iFrm2Clone();" class="button">
                        <input type="button" value="Update" name="btn_iFrame2Update" onclick="btn_iFrm2Update();" class="button">
                        <input disabled type="button" value="Delete" name="btn_iFrame2Delete" onclick="btn_iFrm2Delete();" class="button">
                    </td>
                </tr>

            </tbody>
            <tfoot>
                <tr>
                    <td width="100%" colspan="4" align="right">
                        <input name="btnReferences" type="button" style="width: 2.2cm" value="References" tabindex="13" onclick="return btnReferences_onclick();">&nbsp;&nbsp;

			<!--<input name="btnDelete" type="button" style="width: 2cm" value="Delete" tabindex="13" onclick="btn_onDelete();">-->

                        <input name="btnDelete" type="button" style="width: 2cm" onclick="<% if (status4 = "A") then Response.write "btn_onDelete('A');" else Response.write "btn_onDelete('D');" end if %>"
                            value="<% if (rsNE("RECORD_STATUS_IND") = "A") then Response.write "Delete" else Response.write "UnDelete" end if  %>">
                        &nbsp;&nbsp;
			<input name="btnReset" type="button" style="width: 2cm" value="Reset" tabindex="13" onclick="fct_onReset();">&nbsp;&nbsp;
			<input name="btnNew" type="button" style="width: 2cm" value="New" tabindex="13" onclick="fct_NewMO();">
                        &nbsp;&nbsp;
			<input name="btnClone" type="button" style="width: 2cm" value="Clone" tabindex="13" onclick="fct_onClone();">
                        &nbsp;&nbsp;
			<input name="btnSave" type="button" style="width: 2cm" value="Save" tabindex="13" onclick="btn_onSave();">
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    </td>
                </tr>
            </tfoot>
        </table>
        <%
if strDisable <> "" Then
//Response.Write "<B>NetCracker Item. Update is limited to 4 areas: Name Alias, Port Information, Repair Priority and Comments.</B>"
End If
if strDisble_customer_service_location = "" Then
//Response.Write "<P><B> non-CIU MO. Update is opened to 2 more areas: Customer and Service location.</B></P>"
End If
        %>
        <fieldset>
            <!-- <%if bolClone then strNE_ID = ""%>-->
            <legend align="right"><b>Audit Information</b></legend>
            <div size="8pt" align="RIGHT">
                Record Status Indicator
		<input align="left" name="txtRecordStatusInd" type="text" style="height: 20px; width: 18px" disabled value="<% if strNE_ID <> "" then Response.write  rsNE("RECORD_STATUS_IND") end if %>">&nbsp;&nbsp;&nbsp;
		Create Date&nbsp;<input align="center" name="txtCreateDateTime" type="text" style="height: 20px; width: 150px" disabled value="<% if strNE_ID <> "" then Response.write  rsNE("CREATE_DATE_TIME") end if %>">&nbsp;
		Created By&nbsp;
                <input align="right" name="txtCreateRealUser" type="text" style="height: 20px; width: 200px" disabled value="<% if strNE_ID <> "" then Response.write  rsNE("CREATE_REAL_USERID") end if %>"><br>
                Update Date&nbsp;<input align="center" name="txtUpdateDateTime" type="text" style="height: 20px; width: 150px" disabled value="<% if strNE_ID <> "" then Response.write  rsNE("UPDATE_DATE_TIME") end if %>">
                Updated By&nbsp;
                <input align="right" name="txtUpdateRealUser" type="text" style="height: 20px; width: 200px" disabled value="<% if strNE_ID <> "" then Response.write  rsNE("UPDATE_REAL_USERID") end if %>">
            </div>
        </fieldset>
    </form>
</body>
</html>
<%
if strNE_ID <> "" then
	rsNE.Close
end if
   
%>