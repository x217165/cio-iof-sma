<%@  language="VBScript" %>
<% Option Explicit
   on error resume next
%>
<% Response.Buffer = true %>
<!--#include file="SmaConstants.inc"-->
<!--#include file="databaseconnect.asp"-->
<!--#include file="smaProcs.inc"-->
<!--
*************************************************************************************
* File Name:	AssetDetail.asp
*
* Purpose:
*
* In Param:
*
* Out Param:
*
* Created By:
* Edited by:    Adam Haydey Mar 2, 2001
*               CR 1550 ... leave Make/Model/Part No. blank instead of with "<none>"
**************************************************************************************

        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
	   20-Jul-01	 DTy  		Do not allow partial entry for date received,
	                              date installed, SAP capitalisation, date disposed,
                                  and next scheduled maintenance date.
**************************************************************************************
-->
<%
'check user's rights
dim intAccessLevel,strWinMessage,assettypeid,intAccessLevel2
dim strPoDefault,strPartDefault
strPoDefault = "Not Applicable"
'strPartDefault = "<none>"
strPartDefault = ""

intAccessLevel = CInt(CheckLogon(strConst_Asset))
intAccessLevel2 = CInt(CheckLogon(strConst_AssetAdditionalCosts))

if ((intAccessLevel and intConst_Access_ReadOnly) <> intConst_Access_ReadOnly) then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to the Asset SCREEN. Please contact your system administrator"
end if

Dim StrAssetID,StrSql,strWhereClause,objRsAsset,objRsOwnerStatus
Dim objRsDeployStatus,objRsDepartment,objRsVendor,objRsFloor,objRsRack,objRsFinance,strNew,bolClone,strTmpAssetID


StrAssetID = Request("asset_id")
strTmpAssetID = StrAssetID
bolClone = false


 'Response.Write "ASSET_ID=" & StrAssetID
 'Response.End

 'Response.Write "ACCESS_LEVEL=" & intAccessLevel

strNew =Request("NewFacility")
dim strRealUserID,strIncludepp
strRealUserID = Session("username")

if  strNew = "NEW" THEN
  StrAssetID = "0"
  strNew = ""
END IF


select case Request("txtFrmAction")
	case "SAVE"
		if (Request("txtassetid") <> "") then
		  if ((intAccessLevel and intConst_Access_Update) <> intConst_Access_Update) then
			 DisplayError "BACK", "", 0, "UPDATE DENIED", "You don't have access to update Assets. Please contact your system administrator"
		  end if
			'Response.Write "updating.." & "<br>"
			'GET THE ASSET TYPE ID FROM STRING
			assettypeid = split(Request("selassettype"),"¿")
			StrAssetID = Request("txtassetid")
			'create command object for update stored proc
			dim cmdUpdateObj
			set cmdUpdateObj = server.CreateObject("ADODB.Command")
			set cmdUpdateObj.ActiveConnection = objConn
			cmdUpdateObj.CommandType = adCmdStoredProc
			cmdUpdateObj.CommandText = "sma_sp_userid.spk_sma_asset_inter.sp_asset_update"
			'create parameters
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_user_id", adVarChar , adParamInput, 20, strRealUserID)
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_asset_id",adNumeric , adParamInput,, Clng(Request("txtassetid")))
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_asset_type_id",adNumeric , adParamInput,, Clng(assettypeid(0)))
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_ownership_status",adNumeric , adParamInput,, Clng(Request("selownerstat")))
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_deployment_status", adNumeric,adParamInput,, Clng(Request("seldeploystat")))
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_department_status",adNumeric,adParamInput,, Clng(Request("seldeptment")))


			if Request("hdnStaffID") <>"" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_staff_id",adNumeric,adParamInput,, Clng(Request("hdnStaffID")))
			else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_staff_id",adNumeric,adParamInput,, null)
			end if

			if Request("hdnReceivedDt") <>"" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_received_dt",adVarChar,adParamInput,20 , Request("hdnReceivedDt"))
			else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_received_dt",adVarChar,adParamInput,20 , null)
			end if

			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_purchase_order_no", adVarChar,adParamInput, 20, Request("txtpo"))

			IF Request("chkinclparent") = "on" then
			    strIncludepp = "Y"
			ELSE
			    strIncludepp = "N"
			END IF

			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_include_in_pp", adChar, adParamInput, 1, strIncludepp)
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_finance_type", adVarChar,adParamInput, 20, Request("selfinancetype"))
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , CDate(Request("hdnUpdateDateTime")))

			if Request("hdnAssetCatalogueID") <>"" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_asset_catalogue_id",adNumeric,adParamInput,, Clng(Request("hdnAssetCatalogueID")))
			else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_asset_catalogue_id",adNumeric,adParamInput,, null)
			end if

			if Request("selvendor") <>"" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_vendor_id",adNumeric,adParamInput,, Clng(Request("selvendor")))
			else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_vendor_id",adNumeric,adParamInput,, null)
			end if

			if Request("txtserial") <>"" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_serial_no",adVarChar,adParamInput,30 , Request("txtserial"))
			else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_serial_no",adVarChar,adParamInput,30 , null)
			end if

			if Request("txtbarcode") <>"" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_bar_code",adVarChar,adParamInput,30 , Request("txtbarcode"))
			else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_bar_code",adVarChar,adParamInput,30 , null)
			end if


           if Request("txtacomments") <>"" then
  	        cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_comments", adVarChar, adParamInput, 2000 ,Request("txtacomments"))
  	       else
  	       cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_comments", adVarChar, adParamInput, 2000 ,null)
  	       end if


  	       if Request("hdnCustomerID") <>"" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_customer_id",adNumeric,adParamInput,, Clng(Request("hdnCustomerID")))
		   else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_customer_id",adNumeric,adParamInput,, null)
		   end if

		   if Request("hdnAddressID") <>"" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_address_id",adNumeric,adParamInput,, Clng(Request("hdnAddressID")))
		   else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_address_id",adNumeric,adParamInput,, null)
		   end if

		   if Request("txtdetloc") <>"" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_specific_location",adVarChar,adParamInput,50 , Request("txtdetloc"))
			else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_specific_location",adVarChar,adParamInput,50 , null)
		   end if

		   if Request("txtlocbarcode") <>"" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_location_barcode",adVarChar,adParamInput,50 , Request("txtlocbarcode"))
			else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_location_barcode",adVarChar,adParamInput,50 , null)
		   end if

		   if Request("selfloor") <>"" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_floor",adVarChar,adParamInput,15 , Request("selfloor"))
			else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_floor",adVarChar,adParamInput,15 , null)
		   end if

		   if Request("selrack") <>"" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_rack",adVarChar,adParamInput,20 , Request("selrack"))
			else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_rack",adVarChar,adParamInput,20 , null)
		   end if

		   if Request("txtslotcount") <>"" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_port_count",adNumeric,adParamInput,, Clng(Request("txtslotcount")))
		   else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_port_count",adNumeric,adParamInput,, null)
		   end if

		   if Request("txttacname") <>"" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_tac_name",adVarChar,adParamInput,80 , Request("txttacname"))
			else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_tac_name",adVarChar,adParamInput,80 , null)
		   end if

		   if Request("hdnDateInstalled") <>"" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_installed_dt",adVarChar,adParamInput,20 , Request("hdnDateInstalled"))
			else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_installed_dt",adVarChar,adParamInput,20 , null)
			end if

			if Request("hdndatedisposed") <>"" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_disposed_dt",adVarChar,adParamInput,20 , Request("hdndatedisposed"))
			else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_disposed_dt",adVarChar,adParamInput,20 , null)
			end if

		   if Request("txtpurprice") <>"" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_purchase_price",adVarChar,adParamInput,50, Request("txtpurprice"))
		   else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_purchase_price",adVarChar,adParamInput,50, null)
		   end if

		   if Request("hdnsapCapDt") <>"" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_sapcapitalization_dt",adVarChar,adParamInput,20 , Request("hdnsapCapDt"))
			else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_sapcapitalization_dt",adVarChar,adParamInput,20 , null)
		   end if

		   if Request("txtsapwbsno") <>"" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_sap_wbs_no",adVarChar,adParamInput,50 , Request("txtsapwbsno"))
			else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_sap_wbs_no",adVarChar,adParamInput,50 , null)
			end if

			if Request("txtmasterrecno") <>"" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_sap_asset_master_no",adVarChar,adParamInput,50 , Request("txtmasterrecno"))
			else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_sap_asset_master_no",adVarChar,adParamInput,50 , null)
			end if

		   if Request("txtsalvageval") <>"" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_salvage_value",adVarChar,adParamInput,50, Request("txtsalvageval"))
		   else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_salvage_value",adVarChar,adParamInput,50, null)
		   end if

		   if Request("txthwversion") <>"" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_hardware_revision",adVarChar,adParamInput,30, Request("txthwversion"))
		   else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_hardware_revision",adVarChar,adParamInput,30, null)
		   end if

		   if Request("txtswversion") <>"" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_software_revision",adVarChar,adParamInput,30, Request("txtswversion"))
		   else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_software_revision",adVarChar,adParamInput,30, null)
		   end if

		   if Request("txtfwversion") <>"" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_freeware_revision",adVarChar,adParamInput,30, Request("txtfwversion"))
		   else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_freeware_revision",adVarChar,adParamInput,30, null)
		   end if


		   if Request("hdnschedulemaintdt") <>"" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_next_ched_maint",adVarChar,adParamInput,20, Request("hdnschedulemaintdt"))
		   else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_next_ched_maint",adVarChar,adParamInput,20, null)
		   end if

		   if Request("txtwarranty") <>"" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_warranty_period",adVarChar,adParamInput,30, Request("txtwarranty"))
		   else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_warranty_period",adVarChar,adParamInput,30, null)
		   end if

		   if Request("txtcllicode") <>"" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_clli_code",adVarChar,adParamInput,11, Request("txtcllicode"))
		   else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_clli_code",adVarChar,adParamInput,11, null)
		   end if

			'call the insert stored proc
  			'cmdUpdateObj.Parameters.Refresh
  			'Response.Write "Updating..."

  			'dim objparm
  			'for each objparm in cmdUpdateObj.Parameters
  			  'Response.Write "<b>" & objparm.name & "</b>"
  			 ' Response.Write " has size:  " & objparm.Size & " "
  			 ' Response.Write " and value:  " & objparm.value & " "
  			'Response.Write " and datatype:  " & objparm.Type & "<br> "
  		  ' next

  		   'Response.Write "<b> count = " & cmdUpdateObj.Parameters.count & "<br>"
  			'dim nx
  			 'for nx=0 to cmdUpdateObj.Parameters.count-1
  			  ' Response.Write " parm value= " & cmdUpdateObj.Parameters.Item(nx) & "<br>"
  			' next

  			'call the update stored proc

  			'if objConn.Errors.Count <> 0 then
			'DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT UPDATE ASSET - PARAMETER ERROR", objConn.Errors(0).Description
			'objConn.Errors.Clear
		   ' end if

			cmdUpdateObj.Execute
			dim strWinLocation
			if objConn.Errors.Count <> 0 then
			  if instr(1, objConn.Errors(0).Description, "ORA-20041",1 ) or instr(1, objConn.Errors(0).Description, "ORA-20042",1 )then

					strWinLocation = "AssetDetail.asp?asset_id="&Request("txtassetid")
					DisplayError "REFRESH", strWinLocation, objConn.Errors(0).NativeError, "ASSET UPDATED", objConn.Errors(0).Description
				else
					DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT UPDATE ASSET", objConn.Errors(0).Description
				end if
				objConn.Errors.Clear
			else
				strWinMessage = "Record saved successfully. You can now see the changes you made."
			end if

			'strWinMessage = "Record saved successfully. You can now see the changes you made."

		 else
		   'if (Request("txtassetid") = "")
		    if ((intAccessLevel and intConst_Access_Create) <> intConst_Access_Create) then
		      DisplayError "BACK", "", 0, "INSERT DENIED", "You don't have access to create Assets. Please contact your system administrator"
			end if

			'Response.Write "INSERTING.." & "<br>"
			dim cmdInsertObj

			IF Request("chkinclparent") = "on" then
			    strIncludepp = "Y"
			ELSE
			    strIncludepp = "N"
			END IF

			assettypeid = split(Request("selassettype"),"¿")

			set cmdInsertObj = server.CreateObject("ADODB.Command")
			set cmdInsertObj.ActiveConnection = objConn
			cmdInsertObj.CommandType = adCmdStoredProc
			cmdInsertObj.CommandText = "sma_sp_userid.spk_sma_asset_inter.sp_asset_insert"

			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_user_id", adVarChar , adParamInput, 20, strRealUserID)
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_asset_id",adNumeric , adParamOutput,, null)
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_asset_type_id",adNumeric , adParamInput,, Clng(assettypeid(0)))
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_ownership_status",adNumeric , adParamInput,, Clng(Request("selownerstat")))
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_deployment_status", adNumeric,adParamInput,, Clng(Request("seldeploystat")))
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_department_status",adNumeric,adParamInput,, Clng(Request("seldeptment")))
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_staff_id",adNumeric,adParamInput,, Clng(Request("hdnStaffID")))

			if Request("hdnReceivedDt") <>"" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_received_dt",adVarChar,adParamInput,20 , Request("hdnReceivedDt"))
			else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_received_dt",adVarChar,adParamInput,20 , null)
			end if

			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_purchase_order_no", adVarChar,adParamInput, 20, Request("txtpo"))

			IF Request("chkinclparent") = "on" then
			    strIncludepp = "Y"
			ELSE
			    strIncludepp = "N"
			END IF

			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_include_in_pp", adChar, adParamInput, 1, strIncludepp)
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_finance_type", adVarChar,adParamInput, 20, Request("selfinancetype"))


			if Request("hdnAssetCatalogueID") <>"" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_asset_catalogue_id",adNumeric,adParamInput,, Clng(Request("hdnAssetCatalogueID")))
			else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_asset_catalogue_id",adNumeric,adParamInput,, null)
			end if

			if Request("selvendor") <>"" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_vendor_id",adNumeric,adParamInput,, Clng(Request("selvendor")))
			else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_vendor_id",adNumeric,adParamInput,, null)
			end if

			if Request("txtserial") <>"" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_serial_no",adVarChar,adParamInput,30 , Request("txtserial"))
			else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_serial_no",adVarChar,adParamInput,30 , null)
			end if

			if Request("txtbarcode") <>"" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_bar_code",adVarChar,adParamInput,30 , Request("txtbarcode"))
			else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_bar_code",adVarChar,adParamInput,30 , null)
			end if


           if Request("txtacomments") <>"" then
  	        cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_comments", adVarChar, adParamInput, 2000 ,Request("txtacomments"))
  	       else
  	       cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_comments", adVarChar, adParamInput, 2000 ,null)
  	       end if


  	       if Request("hdnCustomerID") <>"" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_customer_id",adNumeric,adParamInput,, Clng(Request("hdnCustomerID")))
		   else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_customer_id",adNumeric,adParamInput,, null)
		   end if

		   if Request("hdnAddressID") <>"" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_address_id",adNumeric,adParamInput,, Clng(Request("hdnAddressID")))
		   else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_address_id",adNumeric,adParamInput,, null)
		   end if

		   if Request("txtdetloc") <>"" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_specific_location",adVarChar,adParamInput,50 , Request("txtdetloc"))
			else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_specific_location",adVarChar,adParamInput,50 , null)
		   end if

		   if Request("txtlocbarcode") <>"" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_location_barcode",adVarChar,adParamInput,50 , Request("txtlocbarcode"))
			else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_location_barcode",adVarChar,adParamInput,50 , null)
		   end if

		   if Request("selfloor") <>"" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_floor",adVarChar,adParamInput,15 , Request("selfloor"))
			else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_floor",adVarChar,adParamInput,15 , null)
		   end if

		   if Request("selrack") <>"" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_rack",adVarChar,adParamInput,20 , Request("selrack"))
			else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_rack",adVarChar,adParamInput,20 , null)
		   end if

		   if Request("txtslotcount") <>"" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_port_count",adNumeric,adParamInput,, Clng(Request("txtslotcount")))
		   else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_port_count",adNumeric,adParamInput,, null)
		   end if

		   if Request("txttacname") <>"" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_tac_name",adVarChar,adParamInput,80 , Request("txttacname"))
			else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_tac_name",adVarChar,adParamInput,80 , null)
		   end if

		   if Request("hdnDateInstalled") <>"" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_installed_dt",adVarChar,adParamInput,20 , Request("hdnDateInstalled"))
			else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_installed_dt",adVarChar,adParamInput,20 , null)
			end if

			if Request("hdndatedisposed") <>"" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_disposed_dt",adVarChar,adParamInput,20 , Request("hdndatedisposed"))
			else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_disposed_dt",adVarChar,adParamInput,20 , null)
			end if

		   if Request("txtpurprice") <>"" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_purchase_price",adVarChar,adParamInput,50, Request("txtpurprice"))
		   else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_purchase_price",adVarChar,adParamInput,50, null)
		   end if

		   if Request("hdnsapCapDt") <>"" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_sapcapitalization_dt",adVarChar,adParamInput,20 , Request("hdnsapCapDt"))
			else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_sapcapitalization_dt",adVarChar,adParamInput,20 , null)
		   end if

		   if Request("txtsapwbsno") <>"" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_sap_wbs_no",adVarChar,adParamInput,50 , Request("txtsapwbsno"))
			else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_sap_wbs_no",adVarChar,adParamInput,50 , null)
			end if

			if Request("txtmasterrecno") <>"" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_sap_asset_master_no",adVarChar,adParamInput,50 , Request("txtmasterrecno"))
			else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_sap_asset_master_no",adVarChar,adParamInput,50 , null)
			end if

		   if Request("txtsalvageval") <>"" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_salvage_value",adVarChar,adParamInput,50, Request("txtsalvageval"))
		   else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_salvage_value",adVarChar,adParamInput,50, null)
		   end if

		   if Request("txthwversion") <>"" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_hardware_revision",adVarChar,adParamInput,30, Request("txthwversion"))
		   else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_hardware_revision",adVarChar,adParamInput,30, null)
		   end if

		   if Request("txtswversion") <>"" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_software_revision",adVarChar,adParamInput,30, Request("txtswversion"))
		   else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_software_revision",adVarChar,adParamInput,30, null)
		   end if

		   if Request("txtfwversion") <>"" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_freeware_revision",adVarChar,adParamInput,30, Request("txtfwversion"))
		   else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_freeware_revision",adVarChar,adParamInput,30, null)
		   end if


		   if Request("hdnschedulemaintdt") <>"" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_next_ched_maint",adVarChar,adParamInput,20, Request("hdnschedulemaintdt"))
		   else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_next_ched_maint",adVarChar,adParamInput,20, null)
		   end if

		   if Request("txtwarranty") <>"" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_warranty_period",adVarChar,adParamInput,30, Request("txtwarranty"))
		   else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_warranty_period",adVarChar,adParamInput,30, null)
		   end if

		   if Request("txtcllicode") <>"" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_clli_code",adVarChar,adParamInput,11, Request("txtcllicode"))
		   else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_clli_code",adVarChar,adParamInput,11, null)
		   end if
		   cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_warning_message", adVarChar, adParamOutput, 200, null)

		'call the insert stored proc
  			'cmdInsertObj.Parameters.Refresh

  			'Response.Write "inserting.."

  			'dim objparm
  			'for each objparm in cmdInsertObj.Parameters-1
  			  'Response.Write "<b>" & objparm.name & "</b>"
  			 ' Response.Write " has size:  " & objparm.Size & " "
  			  'Response.Write " and value:  " & objparm.value & " "
  			 'Response.Write " and datatype:  " & objparm.Type & "<br> "
  		   'next

  			'Response.Write "<b> count = " & cmdInsertObj.Parameters.count & "<br>"
  			'dim nx
  			 'for nx=0 to cmdInsertObj.Parameters.count
  			  ' Response.Write " parm value= " & cmdInsertObj.Parameters.Item(nx) & "<br>"
  			  'next

  		 'if objConn.Errors.Count <> 0 then
			'DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT CREATE ASSET - PARAMETER ERROR", objConn.Errors(0).Description
			'objConn.Errors.Clear
		 'end if

  			cmdInsertObj.Execute

			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT CREATE NEW ASSET", objConn.Errors(0).Description
				objConn.Errors.Clear
			else
			dim strWarning
				StrAssetID = cmdInsertObj.Parameters("p_asset_id").Value
				strNew=""
				strWarning = cmdInsertObj.Parameters("p_warning_message").Value
				if strWarning <> "" then
					strWinLocation = "AssetDetail.asp?asset_id=" & StrAssetID
					DisplayError "REFRESH", strWinLocation, "-20042", "ASSET INSERTED", strWarning
				end if

			end if
			strWinMessage = "Record created successfully. You can now see the new record."
		'else
		  ' DisplayError "BACK", "", err.Number, "ACCESS DENIED. PLEASE CONTACT YOUR SYSTEM ADMINISTRATOR.", err.Description
	    end if

	     'if err then
		  'DisplayError "BACK", "", err.Number, "CANNOT CREATE ASSET - TRY AGAIN", err.Description
	    'end if
		'end if
		case "DELETE"
		'delete record
         if ((intAccessLevel and intConst_Access_Delete) <> intConst_Access_Delete) then
            DisplayError "BACK", "", 0, "DELETE DENIED", "You don't have access to delete Assets. Please contact your system administrator"
		  end if
			dim cmdDeleteObj
			set cmdDeleteObj = server.CreateObject("ADODB.Command")
			set cmdDeleteObj.ActiveConnection = objConn
			cmdDeleteObj.CommandType = adCmdStoredProc
			cmdDeleteObj.CommandText = "sma_sp_userid.spk_sma_asset_inter.sp_asset_delete"

			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_asset_id", adNumeric, adParamInput, ,Clng(Request("asset_id")))
			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , CDate(Request("UpdateDateTime")))
             cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("P_USER_ID", adVarChar , adParamInput, 30, strRealUserID)

			'Response.write "asseti_id= " & Request("asset_id")
			'Response.write "date= " & Request("UpdateDateTime")

			'dim objparm
  			'for each objparm in cmdDeleteObj.Parameters-1
  			  'Response.Write "<b>" & objparm.name & "</b>"
  			 ' Response.Write " has size:  " & objparm.Size & " "
  			  'Response.Write " and value:  " & objparm.value & " "
  			 'Response.Write " and datatype:  " & objparm.Type & "<br> "
  		    'next

			cmdDeleteObj.Execute
			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE ASSET", objConn.Errors(0).Description
				objConn.Errors.Clear
			end if
			StrAssetID = 0
			strWinMessage = "Record deleted successfully."

			'else
		      'DisplayError "BACK", "", err.Number, "ACCESS DENIED. PLEASE CONTACT YOUR SYSTEM ADMINISTRATOR.", err.Description
	     'end if

     end select



 StrSql = "SELECT A.ASSET_TYPE_ID,A.ASSET_TYPE_DESC,B.ASSET_SUB_CLASS_DESC,C.ASSET_CLASS_DESC  " &_
          "FROM CRP.ASSET_TYPE A,CRP.ASSET_SUB_CLASS B,CRP.ASSET_CLASS C " &_
          "WHERE A.ASSET_SUB_CLASS_ID = B.ASSET_SUB_CLASS_ID AND "&_
          "B.ASSET_CLASS_ID = C.ASSET_CLASS_ID AND  " &_
          "A.RECORD_STATUS_IND = 'A' ORDER BY A.ASSET_TYPE_DESC"

     'Create Recordset object
 set objRsAsset = objConn.Execute(StrSql)

 if err then
	  DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
  end if


 StrSql = "SELECT OWNERSHIPSTATUSID,STATUS FROM MSACCESS.TLKPOWNERSHIPSTATUS WHERE RECORD_STATUS_IND = 'A' ORDER BY STATUS"

     'Create Recordset object
 set objRsOwnerStatus = objConn.Execute(StrSql)

 if err then
	  DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
  end if


 StrSql = "SELECT DEPLOYMENTSTATUSID,STATUS FROM MSACCESS.TLKPDEPLOYMENTSTATUS WHERE RECORD_STATUS_IND = 'A' ORDER BY STATUS"

     'Create Recordset object
 set objRsDeployStatus = objConn.Execute(StrSql)

 if err then
	  DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
  end if



 StrSql = "SELECT DEPARTMENT_ID,DEPARTMENT_DESC FROM CRP.DEPARTMENT_LOOKUP WHERE RECORD_STATUS_IND = 'A' ORDER BY DEPARTMENT_DESC"

     'Create Recordset object
 set objRsDepartment = objConn.Execute(StrSql)

 if err then
	  DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
  end if



 StrSql = "SELECT VENDOR_ID,VENDOR_NAME FROM CRP.VENDOR WHERE RECORD_STATUS_IND = 'A' ORDER BY VENDOR_NAME"

     'Create Recordset object
 set objRsVendor = objConn.Execute(StrSql)

 if err then
	  DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
  end if

  StrSql = "SELECT FLOOR FROM MSACCESS.TLKPFLOOR WHERE RECORD_STATUS_IND = 'A' ORDER BY FLOOR"

     'Create Recordset object
 set objRsFloor = objConn.Execute(StrSql)

 if err then
	  DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
  end if


  StrSql = "SELECT RACK FROM MSACCESS.TLKPRACK WHERE RECORD_STATUS_IND = 'A' ORDER BY RACK"

     'Create Recordset object
 set objRsRack = objConn.Execute(StrSql)

 if err then
	  DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
  end if

 StrSql =  "SELECT distinct finance_type from crp.asset WHERE RECORD_STATUS_IND = 'A'order by finance_type"

   'Create Recordset object
 set objRsFinance = objConn.Execute(StrSql)

 if err then
	  DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
  end if



 if  StrAssetID = "0" and strNew <> "CLONED" THEN
 dim newobjRs
 StrSql = "select J.CUSTOMER_NAME,J.CUSTOMER_SHORT_NAME,J.CUSTOMER_ID,K.ADDRESS_ID," &_
          "LTRIM(NVL(K.BUILDING_NAME,'')) ||' '||NVL(K.STREET,'')||' '||NVL(K.MUNICIPALITY_NAME,'')||' '" &_
	      "||NVL(K.PROVINCE_STATE_LCODE,'') ||' '||NVL(K.POSTAL_CODE_ZIP,'') ADDRESS_A  " &_
          "FROM  " &_
          "CRP.CUSTOMER J," &_
          "CRP.V_ADDRESS_CONSOLIDATED_STREET K  " &_
          "WHERE  " &_
          "J.CUSTOMER_ID = K.CUSTOMER_ID AND " &_
          "J.CUSTOMER_ID = 38 AND " &_
          "K.ADDRESS_ID = 2"

    set newobjRs =  objConn.Execute(StrSql)

    if err then
	  DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET newobjRs", err.Description
  end if

 END IF

   'Response.Write "asset id just before select=" & StrAssetID
   'Response.Write "status=" & strNew

 if  StrAssetID <> "0" or StrAssetID ="" THEN
   StrSql = "SELECT A.ASSET_ID,A.ASSET_TYPE_ID,A.OWNERSHIP_STATUS_ID,A.DEPLOYMENT_STATUS_ID," &_
          "A.DEPARTMENT_ID,A.STAFF_ID,A.ASSET_CATALOGUE_ID,A.VENDOR_ID,A.SERIAL_NUMBER," &_
          "A.ASSET_BARCODE,TO_CHAR(A.DATE_RECEIVED,'MON-DD-YYYY') DATE_RECEIVED,D.MAKE_DESC,E.MODEL_DESC,F.PART_NUMBER_DESC," &_
          "G.ASSET_SUB_CLASS_DESC,H.ASSET_CLASS_DESC,A.COMMENTS,I.CONTACT_NAME,A.LOCATION_BARCODE," &_
          "J.CUSTOMER_NAME,J.CUSTOMER_SHORT_NAME,A.CUSTOMER_ID,A.ADDRESS_ID," &_
          "LTRIM(NVL(K.BUILDING_NAME,'')) ||' '||NVL(K.STREET,'')||' '" &_
          "||NVL(K.MUNICIPALITY_NAME,'')||' '||NVL(K.PROVINCE_STATE_LCODE,'')||' '" &_
          "||NVL(K.POSTAL_CODE_ZIP,'') ADDRESS_A,A.SPECIFIC_LOCATION,A.FLOOR,A.RACK,A.TAC_NAME," &_
          "A.PORT_COUNT,A.PARENT_ASSET_ID,A.LOCATION_WITHIN_PARENT,A.PURCHASE_PRICE," &_
          "TO_CHAR(A.SAP_CAPITALIZATION_DATE,'MON-DD-YYYY') SAP_CAPITALIZATION_DATE,A.PURCHASE_ORDER_NUMBER,A.SAP_WBS_NUMBER," &_
          "A.SAP_ASSET_MASTER_RECORD_NUMBER,A.FINANCE_TYPE,A.PP_INCLUDED_IN_PARENT_FLAG," &_
          "A.SALVAGE_VALUE,TO_CHAR(A.DATE_SOLD,'MON-DD-YYYY') DATE_SOLD,A.HW_REVISION,A.SW_REVISION,A.FW_REVISION," &_
          "A.NEXT_SCHEDULE_MAINTENANCE,A.WARRANTY_PERIOD,TO_CHAR(A.DATE_INSTALL,'MON-DD-YYYY') DATE_INSTALL,TO_CHAR(A.CREATE_DATE_TIME,'MON-DD-YYYY HH24:MI:SS') CREATE_DATE_TIME," &_
          "sma_sp_userid.spk_sma_library.sf_get_full_username(A.CREATE_REAL_USERID) CREATE_REAL_USERID,TO_CHAR(A.UPDATE_DATE_TIME,'MON-DD-YYYY HH24:MI:SS') UPDATE_DATE_TIME_CONV,A.UPDATE_DATE_TIME,sma_sp_userid.spk_sma_library.sf_get_full_username(A.UPDATE_REAL_USERID) UPDATE_REAL_USERID,A.CLLI_CODE,A.RECORD_STATUS_IND,L.SERVICE_LOCATION_ID,L.SERVICE_LOCATION_NAME " &_
          " FROM CRP.ASSET A," &_
          "CRP.ASSET_TYPE B," &_
          "CRP.ASSET_CATALOGUE C," &_
          "CRP.MAKE D," &_
          "CRP.MODEL E," &_
          "CRP.PART_NUMBER F," &_
          "CRP.ASSET_SUB_CLASS G," &_
          "CRP.ASSET_CLASS H," &_
          "CRP.CONTACT I," &_
          "CRP.CUSTOMER J," &_
          "CRP.V_ADDRESS_CONSOLIDATED_STREET K, " &_
          "CRP.SERVICE_LOCATION L"



  strWhereClause = " WHERE " &_
                   "A.ASSET_TYPE_ID = B.ASSET_TYPE_ID(+) AND " &_
                   "A.ASSET_CATALOGUE_ID = C.ASSET_CATALOGUE_ID(+) AND " &_
                   "C.MAKE_ID = D.MAKE_ID(+) AND " &_
                   "C.MODEL_ID =E.MODEL_ID(+) AND " &_
                   "C.PART_NUMBER_ID = F.PART_NUMBER_ID(+) AND " &_
                   "B.ASSET_SUB_CLASS_ID = G.ASSET_SUB_CLASS_ID(+) AND " &_
                   "G.ASSET_CLASS_ID = H.ASSET_CLASS_ID(+) AND " &_
                   "A.STAFF_ID = I.CONTACT_ID(+) AND " &_
                   "A.CUSTOMER_ID = J.CUSTOMER_ID(+) AND " &_
                   "A.ADDRESS_ID = K.ADDRESS_ID(+) AND " &_
                   "A.ADDRESS_ID = L.ADDRESS_ID(+) AND " &_
                   "A.ASSET_ID = " & StrAssetID



  StrSql =  StrSql & " "& strWhereClause

   'Response.Write "SQL STATEMENT WIH WHERE=" & StrSql & "<p>"
   'Response.end
   'Create the command object


     'Create Recordset object

   Dim  objRs
   set objRS = objConn.Execute(StrSql)

   if strNew = "CLONED" then
       StrAssetID=""
       bolClone = true
       strNew=""
   end if

   if err then
	  DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
  end if

 END IF




  'Do while Not objRS.EOF
  'dim dblAssetVal
 'if StrAssetID <> 0 then
 'dblAssetVal = dblAssetVal+cdbl(objRS("PURCHASE_PRICE"))
 'end if


%>
<html>
<head>
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
    <meta http-equiv="Pragma" content="no-cache">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <link rel="stylesheet" type="text/css" href="stylesheets/styles.css">
    <title></title>
    <script type="text/javascript" src="GeneralJavaFunctions.js"></script>
    <script type="text/javascript" src="AccessLevels.js"></script>
    <script language="javascript">
<!--
    //set the heading
    setPageTitle("SMA - Asset");

    //javascript code related to iFrame functionality

    var strDelimiter='<%=strDelimiter%>';
    var intAssetID= '<%=StrAssetID%>';
    var bolSaveRequired = false;
    var intAccessLevel=<%=intAccessLevel%>;
    var intAccessLevel2=<%=intAccessLevel2%>;
    var intConst_MessageDisplay=<%=intConst_MessageDisplay%>;


    function iFrame_display(){
        //called whenever a refresh of the iFrame is needed
        // document.frames("aifr").document.location.href ='AssetAlias.asp?AssetID=' + intAssetID;
        document.getElementById("aifr").src = 'AssetAlias.asp?AssetID=' + intAssetID;

    }


    function btn_iFrmAdd(){

        var NewWin;
        var strSource;
        if ((intAccessLevel2 & intConst_Access_Create) != intConst_Access_Create) {alert('Access denied. Please contact your system administrator.'); return;}
        strSource="AssetAliasDetail.asp?NewFacility=NEW&AssetID="+document.fmAssetDetail.txtassetid.value;
        NewWin=window.open(strSource,"NewWin","toolbar=no,status=no,width=700,height=300,menubar=no resize=no");
        NewWin.focus();
    }


    function btn_iFrmUpdate(){
        var NewWin;
        if ((intAccessLevel2 & intConst_Access_Update) != intConst_Access_Update) {alert('Access denied. Please contact your system administrator.'); return;}

        var doc;
        var iframeObject = document.getElementById('aifr'); // MUST have an ID
        if (iframeObject.contentDocument) { // DOM
            doc = iframeObject.contentDocument;
        } 
        else if (iframeObject.contentWindow) { // IE win
            doc = iframeObject.contentWindow.document;
        }

        var txtAliasID = doc.getElementsByName("txtAliasID")[0].value ;

        if (txtAliasID !="")
        {
            var strSource ="AssetAliasDetail.asp?AliasID="+txtAliasID;
            NewWin=window.open(strSource ,"NewWin","toolbar=no,status=no,width=700,height=300,menubar=no resize=no");
            NewWin.focus();
        }
        else
        {
            alert('You must select a record to update!');
        }

    }

    function btn_iFrmDelete()
    {
        if ((intAccessLevel2 & intConst_Access_Delete) != intConst_Access_Delete)
        {
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

        var txtAliasID = doc.getElementsByName("txtAliasID")[0].value ; 
        var hdnUpdatedDate =doc.getElementsByName("hdnUpdateDateTime")[0].value ;
        var txtAssetID = doc.getElementsByName("hdnUpdateDtxtAssetIDateTime")[0].value ;
        if (txtAliasID !="")
        {
            if (confirm('Do you really want to delete this Alias?')){
                document.getElementById('aifr').src = "AssetAlias.asp?txtFrmAction=DELETE&AliasID="+txtAliasID+"&hdnUpdateDateTime="+ hdnUpdatedDate +"&AssetID="+txtAssetID;

            }
        }
        else
        {
            alert('You must select a record to delete!');
        }

        function body_onLoad(){
            iFrame_display();
        }


        function fct_onUnload(){
            //if (document.fmAssetDetail.btnSave.disabled == false) {
            //if (confirm('Do you want to save changes?')) {
            //alert('saved!');
            //document.fmAssetDetail.submit();

            //}
        }
    }


    function body_onBeforeUnload(){
        document.fmAssetDetail.btnSave.focus();
        if (bolSaveRequired) {
            if	(((intAccessLevel & intConst_Access_Create == intConst_Access_Create) && (document.fmAssetDetail.txtassetid.value == "")) || ((intAccessLevel & intConst_Access_Update == intConst_Access_Update) && (document.fmAssetDetail.txtassetid.value != ""))) {
                event.returnValue = "There is unsaved data on the screen. To save changes, click CANCEL below then click SAVE on the main form.";
            }
        }
    }

    function fct_onClone(){
        if (document.fmAssetDetail.txtassetid.value != "")
        {
            if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {alert('Access denied. Please contact your system administrator.'); return;}

            document.location = "AssetDetail.asp?NewFacility=CLONED&asset_id=<%=StrAssetID %>";
            alert("Record Cloned. Please make changes then save!");
        }
        else
        {
            alert("Unable to clone this may be a new Record!");
        }
    }

    function selNavigate_onchange(){
        //***************************************************************************************************	**
        // Function:	selNavigate_onchange																*
        //																									*
        // Purpose:		To display the page selected by the user from Quick Navigation drop-down box. The	*
        //				fucntion saves Customer Name in a cookie, which is retervied by the called page.	*
        //																									*
        // Created By:	Sara Sangha Aug. 25th, 2000															*
        //																									*
        // Update By:																						*
        //***************************************************************************************************

        var strPageName ;
        var strCustomerIdA,strAddressId,strAssetId,strTacName;



        strCustomerIdA = document.fmAssetDetail.hdnCustomerID.value;
        strAddressId = document.fmAssetDetail.hdnAddressID.value;
        strAssetId = document.fmAssetDetail.txtassetid.value;
        strTacName = document.fmAssetDetail.txttacname.value;

        strPageName = document.fmAssetDetail.selNavigate.item(document.fmAssetDetail.selNavigate.selectedIndex).value ;
        document.fmAssetDetail.selNavigate.SelectedIndex=0;

        // from Customer Detail Page, user will always navigate to lists i.e. search pages.
        switch (strPageName) {

            case 'Cust':
                if (strCustomerIdA !="")
                {
                    document.fmAssetDetail.selNavigate.selectedIndex=0;
                    self.location.href = "CustDetail.asp?CustomerId=" + strCustomerIdA;
                    return true;
                }
                else{
                    alert("Unable to Navigate to Customer Detail as Customer does not exist");
                    return false;
                }
                break;
            case 'ADDR':
                if (strAddressId !="")
                {
                    document.fmAssetDetail.selNavigate.selectedIndex=0;
                    self.location.href = "AddressDetail.asp?AddressID=" + strAddressId;
                    return true;
                }
                else{
                    alert("Unable to Navigate to Address as AddressID does not exist");
                    return false;
                }
                break;
            case 'MO':
                if ((strAssetId !="") && (strTacName !=""))
                {
                    document.fmAssetDetail.selNavigate.selectedIndex=0;
                    SetCookie("MoTacname", escape(document.fmAssetDetail.txttacname.value));
                    //SetCookie("ServLocName", document.fmAssetDetail.hdnServiceLocationName.value);
                    if (document.fmAssetDetail.hdnServiceLocationName.value != ""){SetCookie("MoServLocName", escape(document.fmAssetDetail.hdnServiceLocationName.value))};
                    if (document.fmAssetDetail.hdnServiceLocationID.value != ""){SetCookie("MoServLocID", document.fmAssetDetail.hdnServiceLocationID.value)};
                    //SetCookie("CustomerName", document.fmAssetDetail.txtCustomerName.value);
                    if (document.fmAssetDetail.txtCustomerName.value != ""){SetCookie("MoCustomerName", escape(document.fmAssetDetail.txtCustomerName.value))};
                    if (document.fmAssetDetail.hdnCustomerID.value != ""){SetCookie("MoCustID", document.fmAssetDetail.hdnCustomerID.value)};
                    if (document.fmAssetDetail.txtserial.value != ""){SetCookie("MoSerial", document.fmAssetDetail.txtserial.value)};
                    if (document.fmAssetDetail.txtbarcode.value != ""){SetCookie("MoBarcode", document.fmAssetDetail.txtbarcode.value)};
                    if (document.fmAssetDetail.hdnAssetCatalogueID.value != ""){SetCookie("MoAssetCatID", document.fmAssetDetail.hdnAssetCatalogueID.value)};
                    //SetCookie("MoMake", document.fmAssetDetail.txtAssetMake.value);
                    //SetCookie("MoModel", document.fmAssetDetail.txtAssetModel.value);
                    //SetCookie("MoPartNo", document.fmAssetDetail.txtAssetPartNo.value);
                    //no need to check for empty string since will always contain delimiters ¿
                    SetCookie("MoMakeModelPart", document.fmAssetDetail.txtAssetMake.value+'¿'+document.fmAssetDetail.txtAssetModel.value+'¿'+document.fmAssetDetail.txtAssetPartNo.value);
                    if (document.fmAssetDetail.txtCustomerShortName.value != ""){SetCookie("MoCustShortName", escape(document.fmAssetDetail.txtCustomerShortName.value))};
                    if (document.fmAssetDetail.textAddress.value != ""){SetCookie("MoAddress", escape(document.fmAssetDetail.textAddress.value))};
                    SetCookie("AssetID", strAssetId);
                    SetCookie("MoAssetID", strAssetId);
                    self.location.href = "SearchFrame.asp?fraSrc=" + 'ManagedObjects'  ;
                    return true;
                }
                else{
                    alert("Unable to Navigate to Managed Objects missing AssetId / TacName");
                    return false;
                }
                break;

        }

    }

    function fct_lookupAssetCatalog(){
        if (document.fmAssetDetail.hdnAssetCatalogueID.value != "") {SetCookie("AssetCatID", document.fmAssetDetail.hdnAssetCatalogueID.value);}
        if (document.fmAssetDetail.txtAssetMake.value != "") {SetCookie("AssetCatMake", document.fmAssetDetail.txtAssetMake.value);}
        if (document.fmAssetDetail.txtAssetModel.value != "") {SetCookie("AssetCatModel", document.fmAssetDetail.txtAssetModel.value);}
        if (document.fmAssetDetail.txtAssetPartNo.value != "") {SetCookie("AssetCatPartNumber", document.fmAssetDetail.txtAssetPartNo.value);}

        SetCookie("WinName", 'Popup');
        window.open('SearchFrame.asp?fraSrc=AssetCatalogue', 'Popup', 'top=50, left=100, height=600, width=800' ) ;
        //enable Save button - may not need it but onChange event suck (is not fired when use lookup)
        document.fmAssetDetail.btnSave.disabled = false;
    }

    function fct_lookupCustomer(){
        if (document.fmAssetDetail.txtCustomerName.value != ""){SetCookie("CustomerName", document.fmAssetDetail.txtCustomerName.value);}
        SetCookie("WinName", 'Popup');
        window.open('SearchFrame.asp?fraSrc=Cust', 'Popup', 'top=50, left=100, height=600, width=800' ) ;
        //enable Save button - may not need it but onChange event suck (is not fired when use lookup)
        document.fmAssetDetail.btnSave.disabled = false;
    }

    function fct_lookupContact(){
        var strContact,aNames;

        strContact = document.fmAssetDetail.txtcustodian.value;
        SetCookie("TelusOnly", "yes");
        if (strContact != "")
        {
            aNames = strContact.split(",");

            SetCookie("LName", aNames[0]);
            SetCookie("FName", aNames[1]);


        }
        SetCookie("WinName", 'Popup');
        window.open('SearchFrame.asp?fraSrc=Contact', 'Popup', 'top=50, left=100, height=600, width=800' ) ;
        //enable Save button - may not need it but onChange event suck (is not fired when use lookup)
        //document.fmAssetDetail.btnSave.disabled = false;
    }


    function btnCalendar_onclick(intDateFieldNo) {
        var NewWin;
        if (intDateFieldNo != ""){SetCookie("Field",intDateFieldNo)};
        NewWin=window.open("calendar.asp","NewWin","toolbar=no,status=no,width=260,height=225,menubar=no resize=no");
        //NewWin.creator=self;
        NewWin.focus();
    }

    function btnAddressLookup_onclick() {
        //***************************************************************************************************
        // Function:	btnAddressLookup_onclick															*
        //																									*
        // Purpose:		To display Address Search page with pre-populated customer name and to indicate		*
        //				that the search page is displayed in a popup window. (Note: search pages behave		*
        //				differently when displayed in popup windows verses when displayed in the base window)
        //																									*
        // Created By:	Sara Sangha Aug. 25th, 2000															*
        //																									*
        // Updated By:																						*
        //***************************************************************************************************

        var strCustomerName  = window.fmAssetDetail.txtCustomerName.value ;

        if (strCustomerName != "" ) {
            SetCookie("CustomerName", strCustomerName) ;
            SetCookie("WinName", 'Popup');

            window.open('SearchFrame.asp?fraSrc=Address', 'Popup', 'top=50, left=100, WIDTH=800, HEIGHT=600' ) ;
        }
        else
        {
            alert("Unexpected Error: \nDo not have enough information to move forward");
        }

    }

    function setClass_Sub_Class()
    {
        var strDel ='<%=strDelimiter%>';
        var aValues;
        var strOption = unescape(document.fmAssetDetail.selassettype.item(document.fmAssetDetail.selassettype.selectedIndex).value);
        aValues = strOption.split(strDel);
        document.fmAssetDetail.txtassetsubclass.value = aValues[1];
        document.fmAssetDetail.txtassetclass.value = aValues[2];

    }

    //-->
    </script>

    <script id="clientEventHandlersJS" language="javascript">
<!--

    function fmAssetDetail_OnSubmit()
    {

        var strMonth,strDay,strYear,strDate;

        if	((((intAccessLevel & intConst_Access_Create) == intConst_Access_Create) && (document.fmAssetDetail.txtassetid.value == "")) || ((intAccessLevel & intConst_Access_Update) == intConst_Access_Update) && (document.fmAssetDetail.txtassetid.value != ""))
        {
            if (isWhitespace(document.fmAssetDetail.selownerstat.item(document.fmAssetDetail.selownerstat.selectedIndex).value)) {
                alert('Please enter Ownership Status');
                document.fmAssetDetail.selownerstat.focus();
                return(false);
            }

            if (isWhitespace(document.fmAssetDetail.seldeploystat.item(document.fmAssetDetail.seldeploystat.selectedIndex).value)) {
                alert('Please enter Deployment Status');
                document.fmAssetDetail.seldeploystat.focus();
                return(false);
            }

            if (isWhitespace(document.fmAssetDetail.seldeptment.item(document.fmAssetDetail.seldeptment.selectedIndex).value)) {
                alert('Please enter a Department');
                document.fmAssetDetail.seldeptment.focus();
                return(false);
            }

            if (isWhitespace(document.fmAssetDetail.hdnStaffID.value)) {
                alert('Please enter Requestor/Custodian');
                document.fmAssetDetail.btncustodianlookup.focus();
                return(false);
            }


            if (isWhitespace(document.fmAssetDetail.txtpo.value)) {
                alert('Please enter the Purchase order number');
                document.fmAssetDetail.txtpo.focus();
                return(false);
            }

            if (isWhitespace(document.fmAssetDetail.selassettype.item(document.fmAssetDetail.selassettype.selectedIndex).value)) {
                alert('Please enter the Asset type');
                document.fmAssetDetail.selassettype.focus();
                return(false);
            }

            if (isWhitespace(document.fmAssetDetail.selfinancetype.item(document.fmAssetDetail.selfinancetype.selectedIndex).value)) {
                alert('Please enter the finance type');
                document.fmAssetDetail.selfinancetype.focus();
                return(false);
            }

            if (isWhitespace(document.fmAssetDetail.txtmasterrecno.value) && (document.fmAssetDetail.selmonth4.item(document.fmAssetDetail.selmonth4.selectedIndex).value !="")) {
                alert('Please enter the SAP Asset Master Record Number');
                document.fmAssetDetail.txtmasterrecno.focus();
                return(false);
            }


            if (isWhitespace(document.fmAssetDetail.txttacname.value) && (document.fmAssetDetail.seldeploystat.item(document.fmAssetDetail.seldeploystat.selectedIndex).value =='4')) {
                alert('Please enter a Tacname');
                document.fmAssetDetail.txttacname.focus();
                return(false);
            }


            if (isWhitespace(document.fmAssetDetail.hdnCustomerID.value)) {
                alert('Please enter a Customer ');
                document.fmAssetDetail.btncustomerlookup.focus();
                return(false);
            }


            if (isWhitespace(document.fmAssetDetail.hdnAddressID.value)) {
                alert('Please enter Address ');
                document.fmAssetDetail.btnAddressLookup.focus();
                return(false);
            }


            if (isWhitespace(document.fmAssetDetail.hdnAssetCatalogueID.value)) {
                alert('Please enter make, model and Part ');
                document.fmAssetDetail.btnMakelookup.focus();
                return(false);
            }


            if (isNaN(document.fmAssetDetail.txtpurprice.value))
            {
                alert('Please enter a valid Purchase Price');
                document.fmAssetDetail.txtpurprice.focus();
                return(false);
            }


            if (isNaN(document.fmAssetDetail.txtsalvageval.value))
            {
                alert('Please enter a valid Salvage Value');
                document.fmAssetDetail.txtsalvageval.focus();
                return(false);
            }


            if (isNaN(document.fmAssetDetail.txtslotcount.value))
            {
                alert('Please enter a valid Port/slotcount');
                document.fmAssetDetail.txtslotcount.focus();
                return(false);
            }

            //Date Received
            strMonth = document.fmAssetDetail.selmonth.item(document.fmAssetDetail.selmonth.selectedIndex).value;
            strDay = document.fmAssetDetail.selday.item(document.fmAssetDetail.selday.selectedIndex).value;
            strYear = document.fmAssetDetail.selyear.item(document.fmAssetDetail.selyear.selectedIndex).value;

            if ((strMonth != "") & (strDay !="") & (strYear !=""))
            {
                strDate = strMonth + "/" + strDay + "/" + strYear;
                document.fmAssetDetail.hdnReceivedDt.value = strDate;
            }
            else
                if ((strMonth != "")||(strDay != "" || strYear != ""  )) {
                    alert('Please enter a valid Date Received');
                    document.fmAssetDetail.selmonth.focus();
                    return(false);
                }
                else
                {
                    strDate = "";
                    document.fmAssetDetail.hdnReceivedDt.value = strDate;
                }

            //Date Installed
            strMonth = document.fmAssetDetail.selmonth2.item(document.fmAssetDetail.selmonth2.selectedIndex).value;
            strDay = document.fmAssetDetail.selday2.item(document.fmAssetDetail.selday2.selectedIndex).value;
            strYear = document.fmAssetDetail.selyear2.item(document.fmAssetDetail.selyear2.selectedIndex).value;

            if ((strMonth != "") & (strDay !="") & (strYear !=""))
            {
                strDate = strMonth + "/" + strDay + "/" + strYear;
                document.fmAssetDetail.hdnDateInstalled.value = strDate;
            }
            else
                if ((strMonth != "")||(strDay != "" || strYear != ""  ))
                {alert('Please enter a valid Date Installed');
                    document.fmAssetDetail.selmonth2.focus();
                    return(false);
                }
                else
                {
                    strDate = "";
                    document.fmAssetDetail.hdnDateInstalled.value = strDate;
                }

            //Sap capitalization date
            strMonth = document.fmAssetDetail.selmonth4.item(document.fmAssetDetail.selmonth4.selectedIndex).value;
            strDay = document.fmAssetDetail.selday4.item(document.fmAssetDetail.selday4.selectedIndex).value;
            strYear = document.fmAssetDetail.selyear4.item(document.fmAssetDetail.selyear4.selectedIndex).value;

            if ((strMonth != "") & (strDay !="") & (strYear !=""))
            {
                strDate = strMonth + "/" + strDay + "/" + strYear;
                document.fmAssetDetail.hdnsapCapDt.value = strDate;
            }
            else
                if ((strMonth != "")||(strDay != "" || strYear != ""  )) {
                    alert('Please enter a valid SAP Capitalization date');
                    document.fmAssetDetail.selmonth4.focus();
                    return(false);
                }
                else
                {
                    strDate = "";
                    document.fmAssetDetail.hdnsapCapDt.value = strDate;
                }

            //Date disposed
            strMonth = document.fmAssetDetail.selmonth3.item(document.fmAssetDetail.selmonth3.selectedIndex).value;
            strDay = document.fmAssetDetail.selday3.item(document.fmAssetDetail.selday3.selectedIndex).value;
            strYear = document.fmAssetDetail.selyear3.item(document.fmAssetDetail.selyear3.selectedIndex).value;

            if ((strMonth != "") & (strDay !="") & (strYear !=""))
            {
                strDate = strMonth + "/" + strDay + "/" + strYear;
                document.fmAssetDetail.hdndatedisposed.value = strDate;
            }
            else
                if ((strMonth != "")||(strDay != "" || strYear != ""  )) {
                    alert('Please enter a valid Date disposed');
                    document.fmAssetDetail.selmonth3.focus();
                    return(false);
                }
                else
                {
                    strDate = "";
                    document.fmAssetDetail.hdndatedisposed.value = strDate;
                }

            //Next scheduled maintenance date
            strMonth = document.fmAssetDetail.selmonth5.item(document.fmAssetDetail.selmonth5.selectedIndex).value;
            strDay = document.fmAssetDetail.selday5.item(document.fmAssetDetail.selday5.selectedIndex).value;
            strYear = document.fmAssetDetail.selyear5.item(document.fmAssetDetail.selyear5.selectedIndex).value;

            if ((strMonth != "") & (strDay !="") & (strYear !=""))
            {
                strDate = strMonth + "/" + strDay + "/" + strYear;
                document.fmAssetDetail.hdnschedulemaintdt.value = strDate;
            }
            else
                if ((strMonth != "")||(strDay != "" || strYear != ""  )) {
                    alert('Please enter a valid Next Sheduled Maintenance Date');
                    document.fmAssetDetail.selmonth5.focus();
                    return(false);
                }
                else
                {
                    strDate = "";
                    document.fmAssetDetail.hdnschedulemaintdt.value = strDate;
                }

            bolSaveRequired = false;

            //submit the form
            document.fmAssetDetail.txtFrmAction.value = "SAVE";
            return(true);
        } //end if intAccessLevel >= intConst_Access_Create
        else {
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


    function window_onload() {
        iFrame_display();

    }

    function btnNew_click(){

        if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {alert('Access denied. Please contact your system administrator.'); return;}
        self.document.location.href ="AssetDetail.asp?NewFacility=NEW";

    }


    function fct_onChange(){
        if (intAccessLevel >= intConst_Access_Create) {
            if (document.fmAssetDetail.txtassetid.value != "") {bolSaveRequired = true}
        }
    }


    function fct_onDelete(){
        if (((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) || (document.fmAssetDetail.txtRecordStatusInd.value == "D"))
        {
            alert('Access denied. Please contact your system administrator.');
            return;
        }
        if (confirm('Do you really want to delete this object?')){
            document.location = "AssetDetail.asp?txtFrmAction=DELETE&asset_id="+document.fmAssetDetail.txtassetid.value+"&UpdateDateTime="+document.fmAssetDetail.hdnUpdateDateTime.value;
        }
    }

    function fct_onReset(){
        if(confirm('All changes will be lost. Do you really want to reset the page?')){
            bolSaveRequired = false;
            <%if not bolclone then%>
                document.location = "AssetDetail.asp?asset_id=<%=StrAssetID %>";
            <%else%>
                document.location = "AssetDetail.asp?asset_id=<%=StrTmpAssetID %>&NewFacility=CLONED";
            <%end if%>
            }
    }


    function btnReferences_onclick() {
        var strOwner = 'CRP' ;
        var strTableName = 'ASSET' ;
        var strRecordID = document.fmAssetDetail.txtassetid.value ;
        var URL ;
        if (document.fmAssetDetail.txtassetid.value ==""){
            alert('No references. This is a new record.');
            return false;
        }
        else
        {
            URL ='Dependency.asp?Owner=' + strOwner + '&TableName=' + strTableName + '&RecordID='+ strRecordID   ;
            window.open(URL, 'Popup', 'top=100, left=100, WIDTH=500, HEIGHT=300'  ) ;
        }
    }



    function round_value(val)
    {
        if (val==1){
            if (!(isNaN(document.fmAssetDetail.txtpurprice.value)))
            {
                document.fmAssetDetail.txtpurprice.value =Math.round(document.fmAssetDetail.txtpurprice.value*100)/100;
            }
            else
            {
                alert("Enter a Purchase Price!");
                document.fmAssetDetail.txtpurprice.value ="";
                document.fmAssetDetail.txtpurprice.focus();
            }
        }

        if (val==2){

            if (!(isNaN(document.fmAssetDetail.txtsalvageval.value)))
            {
                document.fmAssetDetail.txtsalvageval.value =Math.round(document.fmAssetDetail.txtsalvageval.value*100)/100;
            }
            else
            {
                alert("Enter a Salvage Value!");
                document.fmAssetDetail.txtsalvageval.value ="";
                document.fmAssetDetail.txtsalvageval.focus();

            }

        }

    }


    function btnSave_onclick()
    {
        var bolretval
        bolretval=fmAssetDetail_OnSubmit();
        if(bolretval)
            document.fmAssetDetail.submit();
    }

    //-->
    </script>
</head>
<body language="javascript" onbeforeunload="body_onBeforeUnload();" onload="return window_onload()">
    <form name="fmAssetDetail" method="POST" action="" onsubmit="">
        <input type="hidden" name="hdnReceivedDt" value="">
        <input type="hidden" name="hdnDateInstalled" value="">
        <input type="hidden" name="hdndatedisposed" value="">
        <input type="hidden" name="hdnsapCapDt" value="">
        <input type="hidden" name="hdnschedulemaintdt" value="">
        <input type="hidden" name="txtFrmAction" value="">
        <input name="hdnStaffID" type="hidden" style="height: 24px; width: 300px" value='<%if StrAssetID <> "0" then  Response.Write ""&objRS("STAFF_ID")&"" else Response.Write "" end if%>'>
        <input name="hdnServiceLocationID" type="hidden" style="height: 24px; width: 300px" value='<%if StrAssetID <> "0" then  Response.Write ""&objRS("SERVICE_LOCATION_ID")&"" else Response.Write "" end if%>'>
        <input name="hdnServiceLocationName" type="hidden" style="height: 24px; width: 300px" value='<%if StrAssetID <> "0" then  Response.Write ""&objRS("SERVICE_LOCATION_NAME")&"" else Response.Write "" end if%>'>

        <table width="100%" border="0">
            <thead>
                <tr>
                    <td width="25%" align="left" colspan="3">Asset Detail</td>
                    <td width="25%">
                        <select align="RIGHT" valign="top" id="selNavigate" name="selNavigate" language="javascript" onchange="return selNavigate_onchange()" <%if StrAssetID = "0" then Response.Write "disabled" end if%>>
                            <option value="DEFAULT">Quickly Goto ...</option>
                            <option value="Cust">Customer</option>
                            <option value="ADDR">Address</option>
                            <option value="MO">Managed Objects</option>
                        </select></td>
                </tr>
                <tr>
                    <td width="25%" colspan="4">Identification</td>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td align="RIGHT" nowrap width="20%">Ownership Status<font color="red">*</font></td>
                    <td width="30%">
                        <select name="selownerstat" style="height: 20px; width: 200px" onchange="fct_onChange();">
                            <option></option>
                            <% Do while Not objRsOwnerStatus.EOF
			   Response.write "<OPTION "
		   if StrAssetID <> "0" then
			 if CInt(objRsOwnerStatus("OWNERSHIPSTATUSID")) = CInt(objRs("OWNERSHIP_STATUS_ID")) then
			   Response.Write " selected "
		     end if
		   end if
			   Response.write "VALUE ="& objRsOwnerStatus("OWNERSHIPSTATUSID") & ">" & routineHtmlString(objRsOwnerStatus("STATUS")) & "</OPTION>"
			   objRsOwnerStatus.MoveNext
			 Loop
                            %>
                        </select>
                    </td>
                    <td align="RIGHT" nowrap width="20%">Deployment Status<font color="red">*</font></td>
                    <td width="30%">
                        <select name="seldeploystat" style="height: 20px; width: 200px" onchange="fct_onChange();">
                            <option></option>
                            <%Do while Not objRsDeployStatus.EOF
				 Response.write "<OPTION "
			if StrAssetID <> "0" then
				 if CInt(objRsDeployStatus("DEPLOYMENTSTATUSID")) = CInt(objRs("DEPLOYMENT_STATUS_ID")) then
				   Response.Write " selected "
				 end if
			end if
				   Response.write "VALUE ="& objRsDeployStatus("DEPLOYMENTSTATUSID") & ">" & routineHtmlString(objRsDeployStatus("STATUS")) & "</OPTION>"
				   objRsDeployStatus.MoveNext
				 Loop
                            %>
                        </select>
                    </td>
                </tr>
                <tr>
                    <td align="RIGHT" nowrap width="15%">Department<font color="red">*</font></td>
                    <td width="25%">
                        <select name="seldeptment" style="height: 20px; width: 200px" onchange="fct_onChange();">
                            <option></option>
                            <%Do while Not objRsDepartment.EOF
					Response.write "<OPTION "
				if StrAssetID <> "0" then
					if Cint(objRsDepartment("DEPARTMENT_ID")) = Cint(objRs("DEPARTMENT_ID")) then
					  Response.Write " selected "
					end if
				end if
					  Response.write "VALUE ="& objRsDepartment("DEPARTMENT_ID") & ">" & routineHtmlString(objRsDepartment("DEPARTMENT_DESC")) & "</OPTION>"
					  objRsDepartment.MoveNext
				Loop
                            %>
                        </select>
                    </td>
                    <td align="RIGHT" nowrap>Requestor/Custodian<font color="red">*</font></td>
                    <td>
                        <input disabled name="txtcustodian" style="height: 24px; width: 250px" value='<%if StrAssetID <> "0" then  Response.Write ""& routineHtmlString(objRS("CONTACT_NAME"))&"" else Response.Write "" end if%>' onchange="fct_onChange();">
                        <input id="btncustodianlookup" name="btncustodianlookup" type="button" value="..." language="javascript" onclick="return fct_lookupContact()"></td>
                    </TD>
                </tr>
                <tr>
                    <td align="RIGHT" nowrap>Date Received<font color="red">*</font></td>
                    <td width="25%">
                        <select name="selmonth" style="height: 20px; width: 70px" onchange="fct_onChange();">
                            <option></option>
                            <%
 dim k

 for k = 1 to 12
   Response.Write "<option "
 if StrAssetID <> "0" then
  if k = month(objRS("DATE_RECEIVED")) then
    Response.Write " selected "
  end if
 end if
  if k < 10 then
  k="0"&k
  end if
  Response.write " VALUE ="& k & ">" &ucase(monthName(k,true)) & "</OPTION>"
  next
                            %>
                        </select>

                        <select name="selday" style="height: 20px; width: 60px" onchange="fct_onChange();">
                            <option></option>
                            <%

 for k = 1 to 31
   Response.Write "<option "
 if StrAssetID <> "0" then
  if k = day(objRS("DATE_RECEIVED")) then
    Response.Write " selected "
  end if
 end if
  if k < 10 then
  k="0"&k
  end if
  Response.write " VALUE ="& k & ">" &k & "</OPTION>"
  next
                            %>
                        </select>
                        <select name="selyear" style="height: 20px; width: 60px" onchange="fct_onChange();">
                            <option></option>
                            <%
 dim i,baseYear
 baseYear = 1994
 for i = 0 to 30
   Response.Write "<option "
 if StrAssetID <> "0" then
  if (baseYear+i) = year(objRS("DATE_RECEIVED")) then
    Response.Write " selected "
  end if
 end if
  Response.write " VALUE ="& baseYear+i & ">" &baseYear+i & "</OPTION>"
  next
                            %>
                        </select>
                        <input type="button" value="..." id="btnCalendar" name="btnCalendar" language="javascript" onclick="return btnCalendar_onclick(1);fct_onChange();">
                    </td>

                    <td align="RIGHT" nowrap>Purchase Order<font color="red">*</font></td>
                    <td>
                        <input name="txtpo" style="height: 24px; width: 150px" value='<%if StrAssetID <> "0" then  Response.Write ""&routineHtmlString(objRS("PURCHASE_ORDER_NUMBER"))&"" else Response.Write ""&strPoDefault&"" end if%>' onchange="fct_onChange();"></td>
                </tr>
                <tr>
                    <td align="RIGHT" nowrap>Asset ID</td>
                    <td>
                        <input readonly style="color: silver" name="txtassetid" style="height: 24px; width: 300px" value='<%if StrAssetID <> "0" then  Response.Write ""&StrAssetID&"" else Response.Write "" end if%>' onchange="fct_onChange();"></td>
                    <td align="RIGHT" nowrap>Asset Type<font color="red">*</font></td>
                    <td>
                        <select name="selassettype" style="height: 20px; width: 200px" onchange="setClass_Sub_Class();">
                            <option></option>
                            <%Do while Not objRsAsset.EOF
				 Response.write "<OPTION "
				if StrAssetID <> "0" then
				 if Cint(objRsAsset("ASSET_TYPE_ID")) = Cint(objRs("ASSET_TYPE_ID")) then
				   Response.Write " selected "
				 end if
				end if
				   Response.write "VALUE ="& objRsAsset("ASSET_TYPE_ID")&strDelimiter&escape(objRsAsset("ASSET_SUB_CLASS_DESC"))&strDelimiter&escape(objRsAsset("ASSET_CLASS_DESC")) & ">" & objRsAsset("ASSET_TYPE_DESC") & "</OPTION>"
				   objRsAsset.MoveNext
				 Loop
                            %>
                        </select>
                    </td>
                </tr>
                <tr>
                    <td width="20%" align="RIGHT" nowrap>Vendor&nbsp;</td>
                    <td width="30%">
                        <select name="selvendor" onchange="fct_onChange();">
                            <option></option>
                            <%Do while Not objRsVendor.EOF
				 Response.write "<OPTION "
			if StrAssetID <> "0" then
			  if not isNull(objRs("VENDOR_ID")) then
				 if cint(objRsVendor("VENDOR_ID")) = cint(objRs("VENDOR_ID")) then
				   Response.Write " selected "
				 end if
				end if
			end if
				   Response.write "VALUE ="& objRsVendor("VENDOR_ID") & ">" & routineHtmlString(objRsVendor("VENDOR_NAME")) & "</OPTION>"
				   objRsVendor.MoveNext

				 Loop
                            %>
                        </select>
                    </td>
                    <td width="20%" align="RIGHT" nowrap>Asset Sub Class&nbsp;</td>
                    <td width="30%">
                        <input name="txtassetsubclass" disabled style="height: 24px; width: 300px" value='<%if StrAssetID <> "0" then  Response.Write ""&routineHtmlString(objRS("ASSET_SUB_CLASS_DESC"))&"" else Response.Write "" end if%>' onchange="fct_onChange();"></td>

                </tr>
                <tr>
                    <td align="RIGHT" nowrap>Barcode&nbsp;</td>
                    <td>
                        <input name="txtbarcode" style="height: 24px; width: 150px" value='<%if StrAssetID <> "0" then  Response.Write ""&routineHtmlString(objRS("ASSET_BARCODE"))&"" else Response.Write "" end if%>' onchange="fct_onChange();"></td>
                    <td align="RIGHT" nowrap>Class&nbsp;</td>
                    <td>
                        <input name="txtassetclass" disabled style="height: 24px; width: 300px" value='<%if StrAssetID <> "0" then  Response.Write ""&routineHtmlString(objRS("ASSET_CLASS_DESC"))&"" else Response.Write "" end if%>' onchange="fct_onChange();">
                        <input type="hidden" name="hdnAssetCatalogueID" style="height: 24px; width: 150px" value='<%if StrAssetID <> "0" then  Response.Write ""&objRS("ASSET_CATALOGUE_ID")&"" else Response.Write "1879" end if%>' onchange="fct_onChange();">
                    </td>
                </tr>
                <tr>
                    <td align="RIGHT" nowrap>Make/Model/Part<font color="red">*</font></td>
                    <td colspan="2">
                        <input name="txtAssetMake" disabled style="height: 24px; width: 150px" value='<%if (StrAssetID <> "0") AND (objRS("MAKE_DESC") <> "<none>") then  Response.Write ""&routineHtmlString(objRS("MAKE_DESC"))&"" else Response.Write ""&strPartDefault &""  end if%>' onchange="fct_onChange();">
                        <input name="txtAssetModel" disabled style="height: 24px; width: 100px" value='<%if (StrAssetID <> "0") AND (objRS("MODEL_DESC") <> "<none>") then  Response.Write ""&routineHtmlString(objRS("MODEL_DESC"))&"" else Response.Write ""&strPartDefault &"" end if%>' onchange="fct_onChange();">
                        <input name="txtAssetPartNo" disabled style="height: 24px; width: 100px" value='<%if (StrAssetID <> "0") AND (objRS("PART_NUMBER_DESC") <> "<none>") then  Response.Write ""&routineHtmlString(objRS("PART_NUMBER_DESC"))&"" else Response.Write ""&strPartDefault &"" end if%>' onchange="fct_onChange();">
                        <input id="btnMakelookup" name="btnMakelookup" type="button" value="..." language="javascript" onclick="return fct_lookupAssetCatalog();fct_onChange();">
                    </td>
                </tr>
                <tr>
                    <td align="RIGHT" nowrap>Serial Number&nbsp;</td>
                    <td>
                        <input name="txtserial" style="height: 24px; width: 150px" value='<%if StrAssetID <> "0" then  Response.Write ""&routineHtmlString(objRS("SERIAL_NUMBER"))&"" else Response.Write "" end if%>' onchange="fct_onChange();"></td>
                </tr>
                <tr>
                    <td align="RIGHT" nowrap rowspan="2" valign="TOP">Comments&nbsp;</td>
                    <td align="LEFT" nowrap colspan="3" rowspan="2">
                        <textarea id="txtacomments" name="txtacomments" rows="3" style="width: 100%" onchange="fct_onChange();"><%if StrAssetID <> "0" then  Response.Write routineHtmlString(objRS("COMMENTS")) else Response.Write null end if%></textarea></td>
                </tr>
                </TR>
            </tbody>
            <tfoot>
                <tr></tr>
            </tfoot>
        </table>
        <table width="100%" border="0">
            <thead>
                <tr>
                    <td width="25%" colspan="4">Deployment</td>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td align="RIGHT" nowrap width="20%">Location Barcode&nbsp;</td>
                    <td width="30%">
                        <input name="txtlocbarcode" style="height: 24px; width: 300px" value='<%if StrAssetID <> "0" then  Response.Write ""&routineHtmlString(objRS("LOCATION_BARCODE"))&"" else Response.Write "" end if%>' onchange="fct_onChange();"></td>
                    <td align="RIGHT" nowrap width="20%">Date Installed&nbsp;</td>

                    <td width="30%">
                        <select name="selmonth2" style="height: 20px; width: 70px" onchange="fct_onChange();">
                            <option></option>
                            <%

 for k = 1 to 12
   Response.Write "<option "
 if StrAssetID <> "0" then
  if k = month(objRS("DATE_INSTALL")) then
    Response.Write " selected "
  end if
 end if
  if k < 10 then
  k="0"&k
  end if
  Response.write " VALUE ="& k & ">" &ucase(monthName(k,true)) & "</OPTION>"
  next
                            %>
                        </select>

                        <select name="selday2" style="height: 20px; width: 60px" onchange="fct_onChange();">
                            <option></option>
                            <%

 for k = 1 to 31
   Response.Write "<option "
 if StrAssetID <> "0" then
  if k = day(objRS("DATE_INSTALL")) then
    Response.Write " selected "
  end if
 end if
  if k < 10 then
  k="0"&k
  end if
  Response.write " VALUE ="& k & ">" &k & "</OPTION>"
  next
                            %>
                        </select>
                        <select name="selyear2" style="height: 20px; width: 60px" onchange="fct_onChange();">
                            <option></option>
                            <%
 baseYear = 1994
 for i = 0 to 30
   Response.Write "<option "
 if StrAssetID <> "0" then
  if (baseYear+i) = year(objRS("DATE_INSTALL")) then
    Response.Write " selected "
  end if
 end if
  Response.write " VALUE ="& baseYear+i & ">" &baseYear+i & "</OPTION>"
  next
                            %>
                        </select>
                        <input type="button" value="..." id="btnCalendar" name="btnCalendar" language="javascript" onclick="return btnCalendar_onclick(2);fct_onChange();">
                    </td>
                </tr>
                <tr>
                    <td align="RIGHT" nowrap width="20%">Customer<font color="red">*</font></td>
                    <td width="30%">
                        <input disabled name="txtCustomerName" style="height: 24px; width: 250px" value='<%if StrAssetID <> "0" then  Response.Write ""&routineHtmlString(objRS("CUSTOMER_NAME"))&"" else Response.Write ""&routineHtmlString(newobjRs("CUSTOMER_NAME"))&"" end if%>' onchange="fct_onChange();">
                        <input id="btnlookup" name="btncustomerlookup" type="button" value="..." language="javascript" onclick="return fct_lookupCustomer();fct_onChange();">
                    </td>

                    <td align="RIGHT" nowrap width="20%">Customer Short Name&nbsp;</td>
                    <td width="30%">
                        <input disabled name="txtCustomerShortName" style="height: 24px; width: 300px" value='<%if StrAssetID <> "0" then  Response.Write ""&routineHtmlString(objRS("CUSTOMER_SHORT_NAME"))&"" else Response.Write ""&routineHtmlString(newobjRs("CUSTOMER_SHORT_NAME"))&"" end if%>' onchange="fct_onChange();"></td>
                </tr>
                <tr>
                    <td align="RIGHT" nowrap width="15%">Address<font color="red">*</font></td>
                    <td width="25%" colspan="3">
                        <input disabled name="textAddress" style="height: 24px; width: 450px" value='<%if StrAssetID <> "0" then  Response.Write ""&routineHtmlString(objRS("ADDRESS_A"))&"" else Response.Write ""&routineHtmlString(newobjRs("ADDRESS_A"))&"" end if%>' onchange="fct_onChange();">
                        <input id="btnAddressLookup" name="btnAddressLookup" style="height: 23px; width: 19px" type="button" value="..." language="javascript" onclick="return btnAddressLookup_onclick();fct_onChange();">
                    </td>
                </tr>
                <tr>
                    <td align="RIGHT" nowrap width="15%">CLLI CODE&nbsp;</td>
                    <td width="25%">
                        <input name="txtcllicode" style="height: 24px; width: 300px" value='<%if StrAssetID <> "0" then  Response.Write ""&routineHtmlString(objRS("CLLI_CODE"))&"" else Response.Write "" end if%>' onchange="fct_onChange();"></td>
                </tr>
                <tr>
                    <td align="RIGHT" nowrap width="15%">Detailed Location&nbsp;</td>
                    <td width="25%" colspan="3">
                        <input name="txtdetloc" style="height: 24px; width: 450px" value='<%if StrAssetID <> "0" then  Response.Write ""&routineHtmlString(objRS("SPECIFIC_LOCATION"))&"" else Response.Write "" end if%>' onchange="fct_onChange();"></td>
                </tr>
                <tr>
                    <td align="RIGHT" nowrap width="15%">Floor&nbsp;</td>
                    <td width="25%">
                        <select name="selfloor" style="height: 20px; width: 200px" onchange="fct_onChange();">
                            <option></option>
                            <% Do while Not objRsFloor.EOF
			   Response.write "<OPTION "
			if StrAssetID <> "0" then
			 if objRsFloor("FLOOR") = objRs("FLOOR") then
			   Response.Write " selected "
		     end if
		    end if
			   Response.write "VALUE ="& routineHtmlString(objRsFloor("FLOOR")) & ">" & routineHtmlString(objRsFloor("FLOOR")) & "</OPTION>"
			   objRsFloor.MoveNext
			 Loop
                            %>
                        </select>
                    </td>
                    <td align="RIGHT" nowrap width="15%">Rack&nbsp;</td>
                    <td width="25%">
                        <select name="selrack" style="height: 20px; width: 200px" onchange="fct_onChange();">
                            <option></option>
                            <%Do while Not objRsRack.EOF
				 Response.write "<OPTION "
				if StrAssetID <> "0" then
				 if objRsRack("RACK") = objRs("RACK") then
				   Response.Write " selected "
				 end if
				end if
				   Response.write "VALUE ="& routineHtmlString(objRsRack("RACK")) & ">" & routineHtmlString(objRsRack("RACK")) & "</OPTION>"
				   objRsRack.MoveNext
				 Loop
                            %>
                        </select>
                    </td>
                </tr>
                <tr>
                    <td align="RIGHT" nowrap>Unique TAC Name&nbsp;</td>
                    <td>
                        <input name="txttacname" style="height: 24px; width: 300px" value='<%if StrAssetID <> "0" then  Response.Write ""&routineHtmlString(objRS("TAC_NAME"))&"" else Response.Write "" end if%>' onchange="fct_onChange();"></td>
                    <td align="RIGHT" nowrap width="15%">Port/Slot count&nbsp;</td>
                    <td width="25%">
                        <input name="txtslotcount" style="height: 24px; width: 150px" value='<%if StrAssetID <> "0" then  Response.Write ""&objRS("PORT_COUNT")&"" else Response.Write "" end if%>' onchange="fct_onChange();"></td>
                </tr>
            </tbody>

        </table>
        <table width="100%" border="0">
            <thead>
                <tr>
                    <td align="left" colspan="4">Financial</td>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td align="RIGHT" nowrap width="15%">SAP Asset Master Record#&nbsp;</td>
                    <td width="25%">
                        <input name="txtmasterrecno" style="height: 24px; width: 300px" value='<%if StrAssetID <> "0" then  Response.Write ""&routineHtmlString(objRS("SAP_ASSET_MASTER_RECORD_NUMBER"))&"" else Response.Write "" end if%>' onchange="fct_onChange();"></td>
                    <td align="RIGHT" nowrap width="15%">SAP WBS Number&nbsp;</td>
                    <td width="25%">
                        <input name="txtsapwbsno" style="height: 24px; width: 150px" value='<%if StrAssetID <> "0" then  Response.Write ""&routineHtmlString(objRS("SAP_WBS_NUMBER"))&"" else Response.Write "" end if%>' onchange="fct_onChange();"></td>
                </tr>
                <tr>
                    <td align="RIGHT" nowrap width="15%">SAP Capitalization Date&nbsp;</td>
                    <td width="25%">
                        <select name="selmonth4" style="height: 20px; width: 70px" onchange="fct_onChange();">
                            <option></option>
                            <%

 for k = 1 to 12
   Response.Write "<option "
 if StrAssetID <> "0" then
  if k = month(objRS("SAP_CAPITALIZATION_DATE")) then
    Response.Write " selected "
  end if
 end if
  if k < 10 then
  k="0"&k
  end if
  Response.write " VALUE ="& k & ">" &ucase(monthName(k,true)) & "</OPTION>"
  next
                            %>
                        </select>

                        <select name="selday4" style="height: 20px; width: 60px" onchange="fct_onChange();">
                            <option></option>
                            <%

 for k = 1 to 31
   Response.Write "<option "
 if StrAssetID <> "0" then
  if k = day(objRS("SAP_CAPITALIZATION_DATE")) then
    Response.Write " selected "
  end if
 end if
  if k < 10 then
  k="0"&k
  end if
  Response.write " VALUE ="& k & ">" &k & "</OPTION>"
  next
                            %>
                        </select>
                        <select name="selyear4" style="height: 20px; width: 60px" onchange="fct_onChange();">
                            <option></option>
                            <%

 baseYear = 1994
 for i = 0 to 30
   Response.Write "<option "
 if StrAssetID <> "0" then
  if (baseYear+i) = year(objRS("SAP_CAPITALIZATION_DATE")) then
    Response.Write " selected "
  end if
 end if
  Response.write " VALUE ="& baseYear+i & ">" &baseYear+i & "</OPTION>"
  next
                            %>
                        </select>
                        <input type="button" value="..." id="btnCalendar" name="btnCalendar" language="javascript" onclick="return btnCalendar_onclick(4);fct_onChange();">
                    </td>
                    <td align="RIGHT" nowrap width="15%">Purchase Price&nbsp;</td>
                    <td width="25%">
                        <input name="txtpurprice" style="height: 24px; width: 150px" onchange="fct_onChange();round_value(1)" value='<%if (StrAssetID <> "0" )  then if objRS("PURCHASE_PRICE")<> "" then  Response.Write FormatNumber(objRS("PURCHASE_PRICE"),-1,-2,-2,0)end if else Response.Write "" end if%>'></td>
                </tr>
                <tr>
                    <%
      Response.Write "<TD ALIGN=RIGHT NOWRAP>Included in Parent Price<font color=red>*</font></TD><TD><INPUT TYPE=CHECKBOX NAME=""chkinclparent"" "
    if StrAssetID <> "0" then
     if objRs("PP_INCLUDED_IN_PARENT_FLAG") = "Y" then
      Response.Write " CHECKED "
      end if
    end if
      Response.write " onclick =""fct_onChange();"" ></TD> "
                    %>
                    <td align="RIGHT" nowrap width="15%">Finance Type<font color="red">*</font></td>
                    <td width="25%">
                        <select name="selfinancetype" style="height: 20px; width: 200px" onchange="fct_onChange();">
                            <option></option>
                            <%Do while Not objRsFinance.EOF
				 Response.write "<OPTION "
			if StrAssetID <> "0" then
				 if objRsFinance("FINANCE_TYPE") = objRs("FINANCE_TYPE") then
				   Response.Write " selected "
				 end if
			end if

		   if StrAssetID =  "0" then
			if objRsFinance("FINANCE_TYPE") = "Capital" then
				   Response.Write " selected "
				 end if
			end if
				   Response.write "VALUE ="& routineHtmlString(objRsFinance("FINANCE_TYPE")) & ">" & routineHtmlString(objRsFinance("FINANCE_TYPE")) & "</OPTION>"
				   objRsFinance.MoveNext
				 Loop
                            %>
                        </select>
                    </td>
                </tr>
                <tr>
                    <td align="RIGHT" nowrap width="15%">Date Disposed&nbsp;</td>
                    <td width="25%">
                        <select name="selmonth3" style="height: 20px; width: 70px" onchange="fct_onChange();">
                            <option></option>
                            <%

 for k = 1 to 12
   Response.Write "<option "
 if StrAssetID <> "0" then
  if k = month(objRS("DATE_SOLD")) then
    Response.Write " selected "
  end if
 end if
  if k < 10 then
  k="0"&k
  end if
  Response.write " VALUE ="& k & ">" &ucase(monthName(k,true)) & "</OPTION>"
  next
                            %>
                        </select>

                        <select name="selday3" style="height: 20px; width: 60px" onchange="fct_onChange();">
                            <option></option>
                            <%

 for k = 1 to 31
   Response.Write "<option "
 if StrAssetID <> "0" then
  if k = day(objRS("DATE_SOLD")) then
    Response.Write " selected "
  end if
 end if
  if k < 10 then
  k="0"&k
  end if
  Response.write " VALUE ="& k & ">" &k & "</OPTION>"
  next
                            %>
                        </select>
                        <select name="selyear3" style="height: 20px; width: 60px" onchange="fct_onChange();">
                            <option></option>
                            <%
 baseYear = 1994
 for i = 0 to 30
   Response.Write "<option "
 if StrAssetID <> "0" then
  if (baseYear+i) = year(objRS("DATE_SOLD")) then
    Response.Write " selected "
  end if
 end if
  Response.write " VALUE ="& baseYear+i & ">" &baseYear+i & "</OPTION>"
  next
                            %>
                        </select>
                        <input type="button" value="..." id="btnCalendar" name="btnCalendar" language="javascript" onclick="return btnCalendar_onclick(3);fct_onChange();">
                    </td>
                    <td align="RIGHT" nowrap width="15%">Salvage Value&nbsp;</td>
                    <td width="25%">
                        <input name="txtsalvageval" style="height: 24px; width: 150px" onchange="fct_onChange();round_value(2)" value='<%if (StrAssetID <> "0" )  then if objRS("SALVAGE_VALUE")<> "" then  Response.Write FormatNumber(objRS("SALVAGE_VALUE"),-1,-2,-2,0)end if else Response.Write "" end if%>'></td>
                </tr>
                <tr>
                    <td valign="top" align="right">Asset Additional Cost&nbsp;</td>
                    <td width="35%" rowspan="5" colspan="2" valign="top">
                        <iframe id="aifr" width="100%" height="100" src="" scrolling="yes" marginheight="1" marginwidth="1"></iframe>
                        <br>
                        <div size="8pt" align="right">
                            Asset Value&nbsp;<input name="txtassetval" disabled style="height: 24px; width: 120px" value="" onchange="fct_onChange();">
                        </div>
                        <br>
                        <input type="button" value="Delete" style="width: 2cm" <%if StrAssetID="" or StrAssetID="0" then Response.Write "DISABLED"  end if%> name="btn_iFrameDelete" onclick="btn_iFrmDelete();">&nbsp;&nbsp;
			<input type="button" value="Refresh" style="width: 2cm" <%if StrAssetID="" or StrAssetID="0" then Response.Write "DISABLED" end if%> name="btn_iFrameRefresh" onclick="iFrame_display();">&nbsp;&nbsp;
			<input type="button" value="New" style="width: 2cm" <%if StrAssetID="" or StrAssetID="0" then Response.Write "DISABLED" end if%> name="btn_iFrameAdd" onclick="btn_iFrmAdd();">&nbsp;&nbsp;
			<input type="button" value="Update" style="width: 2cm" <%if StrAssetID="" or StrAssetID="0" then Response.Write "DISABLED" end if%> name="btn_iFrameupdate" onclick="btn_iFrmUpdate();">
                    </td>
                </tr>

            </tbody>

        </table>


        <!--
<table width="100%" BORDER=0>
	<thead>
		<tr><td colSpan=5 align=left>Configuration</td></tr>
	</thead>
	<tbody>
		<tr>
			<TD ALIGN=RIGHT NOWRAP WIDTH="15%">Parent Asset Id:</TD>
			<TD WIDTH="25%"><INPUT name=txtparentasset style="HEIGHT: 24px; WIDTH: 150px" value= <%if StrAssetID <> "0" then  Response.Write ""&objRS("PARENT_ASSET_ID")&"" else Response.Write "" end if%> onchange ="fct_onChange();"></TD>
			<TD ALIGN=left colspan=2><INPUT name=btnParent type=button style="HEIGHT: 24px; WIDTH: 200px" value=" Show the Parent of this Asset "  LANGUAGE=javascript onclick="">	</TD>
		</tr>
		<tr>
			<TD ALIGN=RIGHT NOWRAP WIDTH="15%">Location Within Parent:</TD>
			<TD WIDTH="25%"><INPUT name=txtslotcount style="HEIGHT: 24px; WIDTH: 150px" value= <%if StrAssetID <> "0" then  Response.Write ""&objRS("LOCATION_WITHIN_PARENT")&"" else Response.Write "" end if%> onchange ="fct_onChange();"></TD>
			<TD ALIGN=left colspan=2><INPUT name=btnChildren type=button style="HEIGHT: 24px; WIDTH: 200px" value="Show the Children of this Asset"  LANGUAGE=javascript onclick=""></TD>
		</TR>
		<TR>
		   <td width=15%>&nbsp;</td>
		   <td width=25%>&nbsp;</td>
		   <TD ALIGN=left colspan=2><INPUT name=btnSibling  type=button style="HEIGHT: 24px; WIDTH: 200px" value="Show the Sibling of this Asset "  LANGUAGE=javascript onclick="">	</TD>
		</TR>
	</tbody>
</table>
	-->
        <table width="100%" border="0">
            <thead>
                <tr>
                    <td colspan="4" align="left">Maintenance</td>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td align="RIGHT" nowrap width="15%">Warranty Information&nbsp;</td>
                    <td width="25%">
                        <input name="txtwarranty" style="height: 24px; width: 300px" value='<%if StrAssetID <> "0" then  Response.Write ""&routineHtmlString(objRS("WARRANTY_PERIOD"))&"" else Response.Write "" end if%>' onchange="fct_onChange();"></td>
                    <td align="RIGHT" nowrap width="15%">Next Scheduled Maintenance&nbsp;</td>
                    <td width="25%">
                        <select name="selmonth5" style="height: 20px; width: 70px" onchange="fct_onChange();">
                            <option></option>
                            <%
 dim q

 for q = 1 to 12
   Response.Write "<option "
 if StrAssetID <> "0" then
  if q = month(objRS("NEXT_SCHEDULE_MAINTENANCE")) then
    Response.Write " selected "
  end if
 end if
  if q < 10 then
  q="0"&q
  end if
  Response.write " VALUE ="& q & ">" &ucase(monthName(q,true)) & "</OPTION>"
  next
                            %>
                        </select>

                        <select name="selday5" style="height: 20px; width: 60px" onchange="fct_onChange();">
                            <option></option>
                            <%

 for q = 1 to 31
   Response.Write "<option "
 if StrAssetID <> "0" then
  if q = day(objRS("NEXT_SCHEDULE_MAINTENANCE")) then
    Response.Write " selected "
  end if
 end if
  if q < 10 then
  q="0"&q
  end if
  Response.write " VALUE ="& q & ">" &q & "</OPTION>"
  next
                            %>
                        </select>
                        <select name="selyear5" style="height: 20px; width: 60px" onchange="fct_onChange();">
                            <option></option>
                            <%
 dim h,baseYear5
 baseYear5 = 1994
 for h = 0 to 30
   Response.Write "<option "
 if StrAssetID <> "0" then
  if (baseYear5+h) = year(objRS("NEXT_SCHEDULE_MAINTENANCE")) then
    Response.Write " selected "
  end if
 end if
  Response.write " VALUE ="& baseYear5+h & ">" &baseYear5+h & "</OPTION>"
  next
                            %>
                        </select>
                        <input type="button" value="..." id="btnCalendar" name="btnCalendar" language="javascript" onclick="return btnCalendar_onclick(5);fct_onChange();">
                    </td>
                </tr>
                <tr>
                    <td align="RIGHT" nowrap width="15%">Hardware Version&nbsp;</td>
                    <td width="25%">
                        <input name="txthwversion" style="height: 24px; width: 300px" value='<%if StrAssetID <> "0" then  Response.Write ""&routineHtmlString(objRS("HW_REVISION"))&"" else Response.Write "" end if%>' onchange="fct_onChange();"></td>
                    <td align="RIGHT" nowrap width="15%">Software Version&nbsp;</td>
                    <td width="25%">
                        <input name="txtswversion" style="height: 24px; width: 300px" value='<%if StrAssetID <> "0" then  Response.Write ""&routineHtmlString(objRS("SW_REVISION"))&"" else Response.Write "" end if%>' onchange="fct_onChange();"></td>
                </tr>
                <tr>
                    <td align="RIGHT" nowrap width="15%">Firmware Version&nbsp;</td>
                    <td width="25%">
                        <input name="txtfwversion" style="height: 24px; width: 300px" value='<%if StrAssetID <> "0" then  Response.Write ""&routineHtmlString(objRS("FW_REVISION"))&"" else Response.Write "" end if%>' onchange="fct_onChange();"></td>
                </tr>
            </tbody>

        </table>
        <table width="100%" border="0">
            <tr>

                <td width="25%">
                    <input name="hdnCustomerID" type="hidden" style="height: 24px; width: 300px" value='<%if StrAssetID <> "0" then  Response.Write ""&objRS("CUSTOMER_ID")&"" else Response.Write ""&newobjRs("CUSTOMER_ID")&"" end if%>'></td>
                <td width="25%">
                    <input name="hdnAddressID" type="hidden" style="height: 24px; width: 300px" value='<%if StrAssetID <> "0" then  Response.Write ""&objRS("ADDRESS_ID")&"" else Response.Write ""&newobjRs("ADDRESS_ID")&"" end if%>'></td>
                <td width="25%">
                    <input name="hdnUpdateDateTime" type="hidden" style="height: 24px; width: 300px" value='<%if StrAssetID <> "0" then  Response.Write ""&objRS("UPDATE_DATE_TIME")&"" else Response.Write "" end if%>'></td>
            </tr>
            <tfoot>
                <tr>
                    <td align="right" colspan="5">
                        <input name="btnReferences" type="button" value="References" style="width: 2.2cm" language="javascript" onclick="return btnReferences_onclick()">&nbsp;&nbsp;
			<input id="btnDelete" name="btnDelete" type="button" value="Delete" style="width: 2cm" language="javascript" onclick="return fct_onDelete();">&nbsp;&nbsp;
			<input id="btnReset" name="btnReset" type="reset" value="Reset" onclick="fct_onReset();" style="width: 2cm">&nbsp;&nbsp;
			<input id="btnAddNew" name="btnAddNew" type="button" value="New" style="width: 2cm" language="javascript" onclick="return btnNew_click();">&nbsp;&nbsp;
			<input name="btnClone" type="button" value="Clone" style="width: 2cm" onclick="fct_onClone();">&nbsp;&nbsp;
			<input id="btnSave" name="btnSave" type="button" value="Save" style="width: 2cm" onclick="btnSave_onclick();">&nbsp;&nbsp;
                    </td>
                </tr>
            </tfoot>
        </table>
        <fieldset>
            <%if bolClone then StrAssetID = "0"%>
            <legend align="RIGHT"><b>Audit Information</b></legend>
            <div size="8pt" align="right">
                Record Status Indicator&nbsp;
		<input align="left" name="txtRecordStatusInd" type="text" style="height: 20px; width: 18px" disabled value='<%if StrAssetID <> "0" then  Response.Write ""&objRS("RECORD_STATUS_IND")&"" else Response.Write "" end if%>'>
                Create Date&nbsp;
		<input align="center" name="txtcrdate" type="text" style="height: 20px; width: 140px" disabled value='<%if StrAssetID <> "0" then  Response.Write ""&objRS("CREATE_DATE_TIME")&"" else Response.Write "" end if%>'>
                &nbsp;Created By&nbsp;
		<input align="right" name="txtcrby" type="text" style="height: 20px; width: 100px" disabled value='<%if StrAssetID <> "0" then  Response.Write ""&routineHtmlString(objRS("CREATE_REAL_USERID"))&"" else Response.Write "" end if%>'><br>
                Update Date&nbsp;
		<input align="center" name="txtupdate" type="text" style="height: 20px; width: 140px" disabled value='<%if StrAssetID <> "0" then  Response.Write ""&objRS("UPDATE_DATE_TIME_CONV")&"" else Response.Write "" end if%>'>
                Updated By&nbsp;
		<input align="right" name="txtupby" type="text" style="height: 20px; width: 100px" disabled value='<%if StrAssetID <> "0" then  Response.Write ""&routineHtmlString(objRS("UPDATE_REAL_USERID"))&"" else Response.Write "" end if%>'>
            </div>
        </fieldset>


    </form>
    <%
   'Move to the next row in the Friends table
   ' objRS.MoveNext


 'Loop


 'Clean up our ADO objects
  if  StrAssetID <> "0"  THEN
    objRS.close
    set objRS = Nothing
  END IF

  if  StrAssetID = "0" and not bolClone THEN
    newobjRs.close
    set newobjRs = Nothing
  END IF

    objRsAsset.close
    set objRsAsset = Nothing

     objRsOwnerStatus.close
     set objRsOwnerStatus = Nothing

     objRsDeployStatus.close
     set objRsDeployStatus = Nothing

     objRsDepartment.close
     set objRsDepartment = Nothing

     objRsVendor.close
    set objRsVendor = Nothing

    objRsRack.close
    set objRsRack = Nothing

    objRsFloor.close
    set objRsFloor = Nothing

    objRsFinance.close
    set objRsFinance = Nothing




    %>
</body>
</html>
