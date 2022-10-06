<%@ Language=VBScript %>
<% Option Explicit
  on error resume next
%>
<% Response.Buffer = true %>
<!--#include file="SmaConstants.inc"-->
<!--#include file="databaseconnect.asp"-->
<!--#include file="smaProcs.inc"-->
<!--

********************************************************************************************
* Page name:	FacilitDetail.asp
* Purpose:
*
* In Param:
*
* Created by:
***************************************************************************************************

		 Date		Author			Changes/enhancements made
		12-Mar-01	A Haydey		From the PVC detail, the file will now open the CustServPVCCriteria
									screen.  This will allow for a search based on a managed object names
									with that particular customer service.
		19-Mar-01	A Haydey		Took out the ability to save a PVC without a side A or B.
		11-Apr-01	DTy  		    Do not allow partial date entry on ADSL due date and facility
                                        start date.
		23-Jan-02	DTy  		   Enhance Facility Provider drop down list.
		08-Mar-02   DTy			   Change CInt to CLng.
***************************************************************************************************
-->
<%

Dim StrCircuitID,StrSql,strWhereClause,StrCircuitTyp,objRsFacTyp,objRsFacStat
Dim objRsAdslTyp,objRsRegion,objRsManEms,objRsCktProv,StrSql2,objRsBilTyp,objUsgCalc,objRsFacilityAlias
dim strReadOnly,strWinMessage,strNew,strCpeOwnership,bolClone,strTmpCircuitID

 strNew =Request("NewFacility")
 StrCircuitID = Request("CircuitID")
 StrCircuitTyp = Request("CircuitTyp")

bolClone = false
strTmpCircuitID= StrCircuitID

dim intAccessLevel

IF StrCircuitTyp = "ATMPVC" THEN
intAccessLevel = CInt(CheckLogon(strConst_PVC))
ELSE
 intAccessLevel = CInt(CheckLogon(strConst_Facilities))
END IF



if ((intAccessLevel and intConst_Access_ReadOnly) <> intConst_Access_ReadOnly) then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to PVC/Facilities SCREEN. Please contact your system administrator"
end if

 dim strRealUserID
 strRealUserID = Session("username")
 
 if  strNew = "NEW" THEN
  StrCircuitID = 0
  strNew =""
 END IF
 
 select case Request("txtFrmAction")
	case "SAVE"
	  if (Request("CircuitID") <> "") then
		 if ((intAccessLevel and intConst_Access_Update) <> intConst_Access_Update) then
		   DisplayError "BACK", "", 0, "UPDATE DENIED", "You don't have access to update Facilities/PVC. Please contact your system administrator"
		 end if

		StrCircuitID = Request("CircuitID")

			'create command object for update stored proc
			dim cmdUpdateObj
			set cmdUpdateObj = server.CreateObject("ADODB.Command")
			set cmdUpdateObj.ActiveConnection = objConn
			cmdUpdateObj.CommandType = adCmdStoredProc
			cmdUpdateObj.CommandText = "sma_sp_userid.spk_sma_fac_inter.sp_fac_update"
			'create parameters

			IF Request("chkadslcpe") = "on" then
			    strCpeOwnership = "Y"
			ELSE
			    strCpeOwnership = "N"
			END IF
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_user_id", adVarChar , adParamInput, 20, strRealUserID)
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_circuit_id",adNumeric , adParamInput,, Clng(Request("CircuitID")))
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_circuit_number", adVarChar,adParamInput, 50, Request("txtcktnum"))
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_circuit_name", adVarChar,adParamInput, 65, Request("txtcktname"))
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_circuit_type", adVarChar,adParamInput, 6, Request("selfactyp"))
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_circuit_provider", adVarChar,adParamInput, 6, Request("selcktprov"))

			if Request("selmgtbyems") <>"" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_circuit_ems", adVarChar,adParamInput, 10, Request("selmgtbyems"))
			else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_circuit_ems", adVarChar,adParamInput, 10, null)
			end if

			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_noc_region", adVarChar,adParamInput, 8, Request("selregion"))
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , CDate(Request("hdnUpdateDateTime")))
		    cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_circuit_status", adVarChar, adParamInput, 6,Request("selfacstat"))

		    if Request("chkadslcpe") <>"" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_adsl_cpe_flag", adChar, adParamInput, 1, strCpeOwnership)
			else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_adsl_cpe_flag", adChar, adParamInput, 1, "N")
			end if

			if Request("hdnCustomerServIDA") <>"" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_customer_service_id_a",adNumeric , adParamInput,, Clng(Request("hdnCustomerServIDA")))
			else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_customer_service_id_a",adNumeric , adParamInput,,null)
			end if

			if Request("hdnCustomerIdA") <>"" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_billing_customer_id_a",adNumeric , adParamInput,, Clng(Request("hdnCustomerIdA")))
			else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_billing_customer_id_a",adNumeric , adParamInput,,null)
			end if

			if Request("hdnServiceLocIdA") <>"" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_service_location_id_a",adNumeric , adParamInput,, Clng(Request("hdnServiceLocIdA")))
			else
		     cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_service_location_id_a",adNumeric , adParamInput,, null)
			end if

			if Request("hdnCustomerServIDB") <>"" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_customer_service_id_b",adNumeric , adParamInput,, Clng(Request("hdnCustomerServIDB")))
		    else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_customer_service_id_b",adNumeric , adParamInput,,null)
			end if

			if Request("hdnCustomerIdB") <>"" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_billing_customer_id_b",adNumeric , adParamInput,, Clng(Request("hdnCustomerIdB")))
			else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_billing_customer_id_b",adNumeric , adParamInput,, null)
			end if

			if Request("hdnServiceLocIdB") <>"" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_service_location_id_b",adNumeric , adParamInput,, Clng(Request("hdnServiceLocIdB")))
			else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_service_location_id_b",adNumeric , adParamInput,, null)
		    end if

		    if Request("hdnCircuitStartDt") <>"" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_circuit_start_dt",adVarChar,adParamInput,20 , Request("hdnCircuitStartDt"))
			else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_circuit_start_dt",adVarChar,adParamInput,20 , null)
			end if

			if Request("selbilltype") <>"" then
			 cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_billing_type", adVarChar, adParamInput, 10, Request("selbilltype"))
			else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_billing_type", adVarChar, adParamInput, 10, null)
			end if

			if Request("selusgcalc") <>"" then
		      cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_usage_calculation_type", adVarChar, adParamInput, 6, Request("selusgcalc"))
		    else
		      cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_usage_calculation_type", adVarChar, adParamInput, 6, null)
		    end if

		    if Request("txtfrdlcifrom") <>"" then
			 cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_fr_dlci_from", adVarChar, adParamInput, 30 ,Request("txtfrdlcifrom"))
			else
			  cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_fr_dlci_from", adVarChar, adParamInput, 30 ,null)
			end if

			if Request("txtfrdlcito") <>"" then
  			  cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_fr_dlci_to", adVarChar, adParamInput, 30, Request("txtfrdlcito"))
  			else
  			  cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_fr_dlci_to", adVarChar, adParamInput, 30, null)
  			end if

  			if Request("txtfibreord") <>"" then
  			  cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_fibre_order_no", adVarChar, adParamInput, 30,Request("txtfibreord"))
  			else
  			   cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_fibre_order_no", adVarChar, adParamInput, 30,null)
  			end if

  		if Request("txtfibrechk") <>"" then
  		   cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_fibre_check_no", adVarChar, adParamInput, 30, Request("txtfibrechk"))
  		else
  		   cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_fibre_check_no", adVarChar, adParamInput, 30, null)
  		end if

  	  if Request("txtadslshelf") <>"" then
  	    cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_adsl_shelf_no", adVarChar, adParamInput, 30, Request("txtadslshelf"))
  	  else
  	    cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_adsl_shelf_no", adVarChar, adParamInput, 30, null)
  	  end if

  	 if Request("txtadslslot") <>"" then
  	    cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_adsl_slot_no", adVarChar, adParamInput, 30, Request("txtadslslot"))
  	 else
  	    cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_adsl_slot_no", adVarChar, adParamInput, 30,null)
  	 end if

  	 if Request("txtadslldb") <>"" then
  		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_adsl_loop_loss", adVarChar, adParamInput, 30 ,Request("txtadslldb"))
  	 else
  		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_adsl_loop_loss", adVarChar, adParamInput, 30 ,null)
  	 end if

    if Request("txtadsltsp") <>"" then
  	  cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_adsl_trained_speed", adVarChar, adParamInput, 30 ,Request("txtadsltsp"))
  	else
  	 cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_adsl_trained_speed", adVarChar, adParamInput, 30 ,null)
  	end if

  	if Request("txtadsldisbl") <>"" then
  	   cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_adsl_dist_block", adVarChar, adParamInput, 30, Request("txtadsldisbl"))
  	else
  	   cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_adsl_dist_block", adVarChar, adParamInput, 30, null)
  	end if

  	if Request("selfacadsltyp") <>"" then
  	   cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_adsl_type", adVarChar, adParamInput, 6, Request("selfacadsltyp"))
  	else
  	   cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_adsl_type", adVarChar, adParamInput, 6, null)
  	end if

  	if Request("hdnAdslDueDt") <>"" then
  	  cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_adsl_due_dt", adVarChar, adParamInput,20 , Request("hdnAdslDueDt"))
  	else
  	cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_adsl_due_dt", adVarChar, adParamInput,20 , null)
  	end if

  	if Request("txtadslorder") <>"" then
  	  cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_adsl_order_no", adVarChar, adParamInput,20 ,Request("txtadslorder"))
  	else
  	   cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_adsl_order_no", adVarChar, adParamInput,20 ,null)
  	end if

    if Request("txtacomments") <>"" then
  	   cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_comments", adVarChar, adParamInput, 2000 ,Request("txtacomments"))
  	else
  	  cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_comments", adVarChar, adParamInput, 2000 ,null)
  	end if

  if StrCircuitTyp = "FIBRE" then
  	if Request("txtfibrelength") <>"" then
	  cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_fibre_length",adNumeric , adParamInput,, Clng(Request("txtfibrelength")))
	else
	cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_fibre_length",adNumeric , adParamInput,,null)
	end if

	if Request("fibreinstallationcost") <>"" then
	  cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_fibre_installation_cost",adNumeric , adParamInput,, Cdbl(Request("fibreinstallationcost")))
	else
	cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_fibre_installation_cost",adNumeric , adParamInput,,null)
	end if

	if Request("txtfibrecost") <>"" then
	  cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_fibre_riser_cost",adNumeric , adParamInput,, Cdbl(Request("txtfibrecost")))
	else
	cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_fibre_riser_cost",adNumeric , adParamInput,,null)
	end if

	if Request("txtfloora") <>"" then
  	   cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_fibre_floor_a", adVarChar, adParamInput, 10 ,Request("txtfloora"))
  	else
  	  cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_fibre_floor_a", adVarChar, adParamInput, 10 ,null)
  	end if

  	if Request("txtbaya") <>"" then
  	   cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_fibre_bay_a", adVarChar, adParamInput, 10 ,Request("txtbaya"))
  	else
  	  cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_fibre_bay_a", adVarChar, adParamInput, 10 ,null)
  	end if


  	if Request("txtshellmodulea") <>"" then
  	   cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_fibre_shelfmod_a", adVarChar, adParamInput, 10 ,Request("txtshellmodulea"))
  	else
  	  cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_fibre_shelfmod_a", adVarChar, adParamInput, 10 ,null)
  	end if

  	if Request("txtpositiona") <>"" then
  	   cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_fibre_position_a", adVarChar, adParamInput, 10 ,Request("txtpositiona"))
  	else
  	  cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_fibre_position_a", adVarChar, adParamInput, 10 ,null)
  	end if

  	if Request("txtcablenumbera") <>"" then
  	   cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_fibre_cable_a", adVarChar, adParamInput, 10 ,Request("txtcablenumbera"))
  	else
  	  cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_fibre_cable_a", adVarChar, adParamInput, 10 ,null)
  	end if

  	if Request("txtpairsa") <>"" then
  	   cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_fibre_pair_a", adVarChar, adParamInput, 10 ,Request("txtpairsa"))
  	else
  	  cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_fibre_pair_a", adVarChar, adParamInput, 10 ,null)
  	end if


  	if Request("txtfloorb") <>"" then
  	   cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_fibre_floor_b", adVarChar, adParamInput, 10 ,Request("txtfloorb"))
  	else
  	  cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_fibre_floor_b", adVarChar, adParamInput, 10 ,null)
  	end if

  	if Request("txtbayb") <>"" then
  	   cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_fibre_bay_b", adVarChar, adParamInput, 10 ,Request("txtbayb"))
  	else
  	  cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_fibre_bay_b", adVarChar, adParamInput, 10 ,null)
  	end if


  	if Request("txtshellmoduleb") <>"" then
  	   cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_fibre_shelfmod_b", adVarChar, adParamInput, 10 ,Request("txtshellmoduleb"))
  	else
  	  cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_fibre_shelfmod_b", adVarChar, adParamInput, 10 ,null)
  	end if

  	if Request("txtpositionb") <>"" then
  	   cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_fibre_position_b", adVarChar, adParamInput, 10 ,Request("txtpositionb"))
  	else
  	  cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_fibre_position_b", adVarChar, adParamInput, 10 ,null)
  	end if

  	if Request("txtcablenumberb") <>"" then
  	   cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_fibre_cable_b", adVarChar, adParamInput, 10 ,Request("txtcablenumberb"))
  	else
  	  cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_fibre_cable_b", adVarChar, adParamInput, 10 ,null)
  	end if

  	if Request("txtpairsb") <>"" then
  	   cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_fibre_pair_b", adVarChar, adParamInput, 10 ,Request("txtpairsb"))
  	else
  	  cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_fibre_pair_b", adVarChar, adParamInput, 10 ,null)
  	end if

 end if 'type=fibre

  			'call the insert stored proc
  			'cmdUpdateObj.Parameters.Refresh
  			'Response.Write "Updating..."

  			'dim objparm
  			'for each objparm in cmdUpdateObj.Parameters
  			 ' Response.Write "<b>" & objparm.name & "</b>"
  			 ' Response.Write " has size:  " & objparm.Size & " "
  			 ' Response.Write " and value:  " & objparm.value & " "
  			 'Response.Write " and datatype:  " & objparm.Type & "<br> "
  		   'next

  		   'Response.Write "<b> count = " & cmdUpdateObj.Parameters.count & "<br>"
  			'dim nx
  			 'for nx=0 to cmdUpdateObj.Parameters.count-1
  			  ' Response.Write " parm value= " & cmdUpdateObj.Parameters.Item(nx) & "<br>"
  			' next

  			'call the update stored proc
  		'if objConn.Errors.Count <> 0 then
			'DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT UPDATE FACILITY/PVC - PARAMETER ERROR", objConn.Errors(0).Description
			'objConn.Errors.Clear
		'end if

			cmdUpdateObj.Execute

	   if objConn.Errors.Count <> 0 then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT UPDATE FACILITY/PVC", objConn.Errors(0).Description
			objConn.Errors.Clear
	   end if

			strWinMessage = "Record saved successfully. You can now see the changes you made."

			'create a new record
		else
		'(Request("CircuitID") = "") and
		if ((intAccessLevel and intConst_Access_Create) <> intConst_Access_Create) then
		   DisplayError "BACK", "", 0, "INSERT DENIED", "You don't have access to create FACILITY/PVC. Please contact your system administrator"
		end if

			dim cmdInsertObj

			IF Request("chkadslcpe") = "on" then
			    strCpeOwnership = "Y"
			ELSE
			    strCpeOwnership = "N"
			END IF

			set cmdInsertObj = server.CreateObject("ADODB.Command")
			set cmdInsertObj.ActiveConnection = objConn
			cmdInsertObj.CommandType = adCmdStoredProc
			cmdInsertObj.CommandText = "sma_sp_userid.spk_sma_fac_inter.sp_fac_insert"
			'create parameters
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_user_id", adVarChar , adParamInput, 30, strRealUserID)
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_circuit_id",adNumeric , adParamOutput,,null) 					'number(9)		means: Managed Object Id
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_circuit_number", adVarChar,adParamInput, 50, Request("txtcktnum"))
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_circuit_name", adVarChar,adParamInput, 65, Request("txtcktname"))
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_circuit_type", adVarChar,adParamInput, 6, Request("selfactyp"))
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_circuit_provider", adVarChar,adParamInput, 6, Request("selcktprov"))

			if Request("selmgtbyems") <>"" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_circuit_ems", adVarChar,adParamInput, 10, Request("selmgtbyems"))
			else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_circuit_ems", adVarChar,adParamInput, 10, null)
			end if

			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_noc_region", adVarChar,adParamInput, 8, Request("selregion"))
		    cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_circuit_status", adVarChar, adParamInput, 6,Request("selfacstat"))

		    if Request("chkadslcpe") <>"" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_adsl_cpe_flag", adChar, adParamInput, 1,  strCpeOwnership)
			else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_adsl_cpe_flag", adChar, adParamInput, 1, null)
			end if

			if Request("hdnCustomerServIDA") <>"" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_customer_service_id_a",adNumeric , adParamInput,, Clng(Request("hdnCustomerServIDA")))
			else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_customer_service_id_a",adNumeric , adParamInput,,null)
			end if

			if Request("hdnCustomerIdA") <>"" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_billing_customer_id_a",adNumeric , adParamInput,, Clng(Request("hdnCustomerIdA")))
			else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_billing_customer_id_a",adNumeric , adParamInput,,null)
			end if

			if Request("hdnServiceLocIdA") <>"" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_service_location_id_a",adNumeric , adParamInput,, Clng(Request("hdnServiceLocIdA")))
			else
		     cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_service_location_id_a",adNumeric , adParamInput,, null)
			end if

			if Request("hdnCustomerServIDB") <>"" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_customer_service_id_b",adNumeric , adParamInput,, Clng(Request("hdnCustomerServIDB")))
		    else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_customer_service_id_b",adNumeric , adParamInput,,null)
			end if

			if Request("hdnCustomerIdB") <>"" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_billing_customer_id_b",adNumeric , adParamInput,, Clng(Request("hdnCustomerIdB")))
			else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_billing_customer_id_b",adNumeric , adParamInput,, null)
			end if

			if Request("hdnServiceLocIdB") <>"" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_service_location_id_b",adNumeric , adParamInput,, Clng(Request("hdnServiceLocIdB")))
			else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_service_location_id_b",adNumeric , adParamInput,, null)
		    end if

		    if Request("hdnCircuitStartDt") <>"" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_circuit_start_dt",adVarChar,adParamInput,20 , Request("hdnCircuitStartDt"))
			else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_circuit_start_dt",adVarChar,adParamInput,20 , null)
			end if

			if Request("selbilltype") <>"" then
			 cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_billing_type", adVarChar, adParamInput, 10, Request("selbilltype"))
			else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_billing_type", adVarChar, adParamInput, 10, null)
			end if

			if Request("selusgcalc") <>"" then
		      cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_usage_calculation_type", adVarChar, adParamInput, 6, Request("selusgcalc"))
		    else
		      cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_usage_calculation_type", adVarChar, adParamInput, 6, null)
		    end if

		    if Request("txtfrdlcifrom") <>"" then
			 cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_fr_dlci_from", adVarChar, adParamInput, 30 ,Request("txtfrdlcifrom"))
			else
			  cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_fr_dlci_from", adVarChar, adParamInput, 30 ,null)
			end if

			if Request("txtfrdlcito") <>"" then
  			  cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_fr_dlci_to", adVarChar, adParamInput, 30, Request("txtfrdlcito"))
  			else
  			  cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_fr_dlci_to", adVarChar, adParamInput, 30, null)
  			end if

  			if Request("txtfibreord") <>"" then
  			  cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_fibre_order_no", adVarChar, adParamInput, 30,Request("txtfibreord"))
  			else
  			   cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_fibre_order_no", adVarChar, adParamInput, 30,null)
  			end if

  		if Request("txtfibrechk") <>"" then
  		   cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_fibre_check_no", adVarChar, adParamInput, 30, Request("txtfibrechk"))
  		else
  		   cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_fibre_check_no", adVarChar, adParamInput, 30, null)
  		end if

  	  if Request("txtadslshelf") <>"" then
  	    cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_adsl_shelf_no", adVarChar, adParamInput, 30, Request("txtadslshelf"))
  	  else
  	    cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_adsl_shelf_no", adVarChar, adParamInput, 30, null)
  	  end if

  	 if Request("txtadslslot") <>"" then
  	    cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_adsl_slot_no", adVarChar, adParamInput, 30, Request("txtadslslot"))
  	 else
  	    cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_adsl_slot_no", adVarChar, adParamInput, 30,null)
  	 end if

  	 if Request("txtadslldb") <>"" then
  		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_adsl_loop_loss", adVarChar, adParamInput, 30 ,Request("txtadslldb"))
  	 else
  		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_adsl_loop_loss", adVarChar, adParamInput, 30 ,null)
  	 end if

    if Request("txtadsltsp") <>"" then
  	  cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_adsl_trained_speed", adVarChar, adParamInput, 30 ,Request("txtadsltsp"))
  	else
  	  cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_adsl_trained_speed", adVarChar, adParamInput, 30 ,null)
  	end if

  	if Request("txtadsldisbl") <>"" then
  	   cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_adsl_dist_block", adVarChar, adParamInput, 30, Request("txtadsldisbl"))
  	else
  	   cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_adsl_dist_block", adVarChar, adParamInput, 30, null)
  	end if

  	if Request("selfacadsltyp") <>"" then
  	   cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_adsl_type", adVarChar, adParamInput, 6, Request("selfacadsltyp"))
  	else
  	   cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_adsl_type", adVarChar, adParamInput, 6, null)
  	end if

  	if Request("hdnAdslDueDt") <>"" then
  	  cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_adsl_due_dt", adVarChar, adParamInput,20 , Request("hdnAdslDueDt"))
  	else
  	cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_adsl_due_dt", adVarChar, adParamInput,20 , null)
  	end if

  	if Request("txtadslorder") <>"" then
  	  cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_adsl_order_no", adVarChar, adParamInput,20 ,Request("txtadslorder"))
  	else
  	   cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_adsl_order_no", adVarChar, adParamInput,20 ,null)
  	end if

    if Request("txtacomments") <>"" then
  	   cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_comments", adVarChar, adParamInput, 2000 ,Request("txtacomments"))
  	else
  	  cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_comments", adVarChar, adParamInput, 2000 ,null)
  	end if

  if StrCircuitTyp = "FIBRE" then
  	if Request("txtfibrelength") <>"" then
	  cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_fibre_length",adNumeric , adParamInput,, Clng(Request("txtfibrelength")))
	else
	cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_fibre_length",adNumeric , adParamInput,,null)
	end if

	if Request("fibreinstallationcost") <>"" then
	  cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_fibre_installation_cost",adNumeric , adParamInput,, Cdbl(Request("fibreinstallationcost")))
	else
	cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_fibre_installation_cost",adNumeric , adParamInput,,null)
	end if

	if Request("txtfibrecost") <>"" then
	  cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_fibre_riser_cost",adNumeric , adParamInput,, Cdbl(Request("txtfibrecost")))
	else
	cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_fibre_riser_cost",adNumeric , adParamInput,,null)
	end if

	if Request("txtfloora") <>"" then
  	   cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_fibre_floor_a", adVarChar, adParamInput, 10 ,Request("txtfloora"))
  	else
  	  cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_fibre_floor_a", adVarChar, adParamInput, 10 ,null)
  	end if

  	if Request("txtbaya") <>"" then
  	   cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_fibre_bay_a", adVarChar, adParamInput, 10 ,Request("txtbaya"))
  	else
  	  cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_fibre_bay_a", adVarChar, adParamInput, 10 ,null)
  	end if


  	if Request("txtshellmodulea") <>"" then
  	   cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_fibre_shelfmod_a", adVarChar, adParamInput, 10 ,Request("txtshellmodulea"))
  	else
  	  cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_fibre_shelfmod_a", adVarChar, adParamInput, 10 ,null)
  	end if

  	if Request("txtpositiona") <>"" then
  	   cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_fibre_position_a", adVarChar, adParamInput, 10 ,Request("txtpositiona"))
  	else
  	  cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_fibre_position_a", adVarChar, adParamInput, 10 ,null)
  	end if

  	if Request("txtcablenumbera") <>"" then
  	   cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_fibre_cable_a", adVarChar, adParamInput, 10 ,Request("txtcablenumbera"))
  	else
  	  cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_fibre_cable_a", adVarChar, adParamInput, 10 ,null)
  	end if

  	if Request("txtpairsa") <>"" then
  	   cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_fibre_pair_a", adVarChar, adParamInput, 10 ,Request("txtpairsa"))
  	else
  	  cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_fibre_pair_a", adVarChar, adParamInput, 10 ,null)
  	end if


  	if Request("txtfloorb") <>"" then
  	   cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_fibre_floor_b", adVarChar, adParamInput, 10 ,Request("txtfloorb"))
  	else
  	  cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_fibre_floor_b", adVarChar, adParamInput, 10 ,null)
  	end if

  	if Request("txtbayb") <>"" then
  	   cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_fibre_bay_b", adVarChar, adParamInput, 10 ,Request("txtbayb"))
  	else
  	  cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_fibre_bay_b", adVarChar, adParamInput, 10 ,null)
  	end if


  	if Request("txtshellmoduleb") <>"" then
  	   cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_fibre_shelfmod_b", adVarChar, adParamInput, 10 ,Request("txtshellmoduleb"))
  	else
  	  cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_fibre_shelfmod_b", adVarChar, adParamInput, 10 ,null)
  	end if

  	if Request("txtpositionb") <>"" then
  	   cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_fibre_position_b", adVarChar, adParamInput, 10 ,Request("txtpositionb"))
  	else
  	  cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_fibre_position_b", adVarChar, adParamInput, 10 ,null)
  	end if

  	if Request("txtcablenumberb") <>"" then
  	   cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_fibre_cable_b", adVarChar, adParamInput, 10 ,Request("txtcablenumberb"))
  	else
  	  cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_fibre_cable_b", adVarChar, adParamInput, 10 ,null)
  	end if

  	if Request("txtpairsb") <>"" then
  	   cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_fibre_pair_b", adVarChar, adParamInput, 10 ,Request("txtpairsb"))
  	else
  	  cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_fibre_pair_b", adVarChar, adParamInput, 10 ,null)
  	end if

 end if 'type=fibre

  			'call the insert stored proc
  			'cmdInsertObj.Parameters.Refresh

  			'Response.Write "inserting.."

  			'dim objparm
  			'for each objparm in cmdInsertObj.Parameters-1
  			 ' Response.Write "<b>" & objparm.name & "</b>"
  			 ' Response.Write " has size:  " & objparm.Size & " "
  			 ' Response.Write " and value:  " & objparm.value & " "
  			' Response.Write " and datatype:  " & objparm.Type & "<br> "
  		  'next

  			'Response.Write "<b> count = " & cmdInsertObj.Parameters.count & "<br>"
  			'dim nx
  			 'for nx=0 to cmdInsertObj.Parameters.count
  			  ' Response.Write " parm value= " & cmdInsertObj.Parameters.Item(nx) & "<br>"
  			  'next

  		 'if objConn.Errors.Count <> 0 then
			'DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT CREATE FACILITY/PVC - PARAMETER ERROR", objConn.Errors(0).Description
			'objConn.Errors.Clear
		 'end if

  			cmdInsertObj.Execute

			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT CREATE NEW FACILITY", objConn.Errors(0).Description
				objConn.Errors.Clear
			else
				StrCircuitID = cmdInsertObj.Parameters("p_circuit_id").Value
			end if
			strWinMessage = "Record created successfully. You can now see the new record."
		'else
		    'DisplayError "BACK", "", err.Number, "ACCESS DENIED. PLEASE CONTACT YOUR SYSTEM ADMINISTRATOR.", err.Description
	    end if

	     if err then
		  DisplayError "BACK", "", err.Number, "CANNOT CREATE FACILITY - TRY AGAIN", err.Description
	    end if
		'end if
		case "DELETE"
		'delete record
         if ((intAccessLevel and intConst_Access_Delete) <> intConst_Access_Delete) then
          DisplayError "BACK", "", 0, "DELETE DENIED", "You don't have access to delete FACILITY/PVC. Please contact your system administrator"
		 end if
			dim cmdDeleteObj
			set cmdDeleteObj = server.CreateObject("ADODB.Command")
			set cmdDeleteObj.ActiveConnection = objConn
			cmdDeleteObj.CommandType = adCmdStoredProc
			cmdDeleteObj.CommandText = "sma_sp_userid.spk_sma_fac_inter.sp_fac_delete"
			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_circuit_id", adNumeric, adParamInput, ,CLng(Request("CircuitID")))
			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , CDate(Request("UpdateDateTime")))
			cmdDeleteObj.Execute
			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE FACILITY", objConn.Errors(0).Description
				objConn.Errors.Clear
			end if
			StrCircuitID = 0
			strWinMessage = "Record deleted successfully."

			'else
		      'DisplayError "BACK", "", err.Number, "ACCESS DENIED. PLEASE CONTACT YOUR SYSTEM ADMINISTRATOR.", err.Description
	     'end if
	    'if err then
		   'DisplayError "BACK", "", err.Number, "CANNOT DELETE FACILITY", err.Description
	    'end if
     end select


  StrSql = "SELECT NOC_REGION_LCODE,NOC_REGION_DESC FROM CRP.LCODE_NOC_REGION WHERE RECORD_STATUS_IND = 'A' ORDER BY  NOC_REGION_LCODE"

     'Create Recordset object
   set objRsRegion = objConn.Execute(StrSql)


   StrSql = "SELECT CIRCUIT_PROVIDER_CODE, CIRCUIT_PROVIDER_NAME, DECODE(IS_ON_NET, 'Y', ' (ON NET)', '') AS ""IS_ON_NET"" FROM CRP.CIRCUIT_PROVIDER WHERE RECORD_STATUS_IND = 'A' ORDER BY CIRCUIT_PROVIDER_NAME"

     'Create Recordset object
   set objRsCktProv = objConn.Execute(StrSql)


  StrSql = "SELECT ELEMENT_MANAGEMENT_SYSTEM_CODE FROM CRP.ELEMENT_MANAGEMENT_SYSTEM WHERE RECORD_STATUS_IND = 'A' ORDER BY  ELEMENT_MANAGEMENT_SYSTEM_CODE"

     'Create Recordset object
   set objRsManEms = objConn.Execute(StrSql)


   StrSql = "SELECT ADSL_TYPE_CODE,ADSL_TYPE_DESC FROM CRP.ADSL_TYPE WHERE RECORD_STATUS_IND = 'A' ORDER BY ADSL_TYPE_CODE"

     'Create Recordset object
   set objRsAdslTyp = objConn.Execute(StrSql)

   StrSql = "SELECT CIRCUIT_STATUS_CODE FROM CRP.CIRCUIT_STATUS WHERE RECORD_STATUS_IND = 'A' ORDER BY CIRCUIT_STATUS_CODE"

   set objRsFacStat = objConn.Execute(StrSql)

   IF StrCircuitTyp <> "ATMPVC" THEN
   StrSql = "SELECT CIRCUIT_TYPE_CODE FROM CRP.CIRCUIT_TYPE WHERE CIRCUIT_TYPE_CODE NOT LIKE '%PVC%' AND RECORD_STATUS_IND = 'A' ORDER BY CIRCUIT_TYPE_CODE"
   ELSE
    StrSql = "SELECT CIRCUIT_TYPE_CODE FROM CRP.CIRCUIT_TYPE WHERE CIRCUIT_TYPE_CODE LIKE '%PVC%' AND RECORD_STATUS_IND = 'A' ORDER BY CIRCUIT_TYPE_CODE"
    StrSql2 ="SELECT BILLING_TYPE_CODE FROM CRP.BILLING_TYPE WHERE RECORD_STATUS_IND = 'A' ORDER BY BILLING_TYPE_CODE"
    set objRsBilTyp = objConn.Execute(StrSql2)
    StrSql2 ="SELECT USAGE_CALCULATION_TYPE_CODE FROM CRP.USAGE_CALCULATION_TYPE WHERE RECORD_STATUS_IND = 'A' ORDER BY USAGE_CALCULATION_TYPE_CODE"
    set objUsgCalc = objConn.Execute(StrSql2)
   END IF

   set objRsFacTyp = objConn.Execute(StrSql)



 'StrSql = "SELECT CIRCUIT_NUMBER_ALIAS_ID,CIRCUIT_NUMBER_ALIAS,CIRCUIT_PROVIDER_CODE FROM CRP.CIRCUIT_NUMBER_ALIAS WHERE RECORD_STATUS_IND = 'A' AND CIRCUIT_ID =" & StrCircuitID


  'set objRsFacilityAlias = objConn.Execute(StrSql)

dim intRowCount, intColCount,strInnerValues
intRowCount = 0
intColCount = 3



strInnerValues = ""
'while not objRsFacilityAlias.EOF
	'intRowCount = intRowCount + 1
	'strInnerValues = strInnerValues & objRsFacilityAlias(0) & strDelimiter & routineHtmlString(objRsFacilityAlias(1)) & strDelimiter & objRsFacilityAlias(2) & strDelimiter
	'strInnerValues = strInnerValues & objRsFacilityAlias(0) & strDelimiter & escape(objRsFacilityAlias(1)) & strDelimiter & escape(objRsFacilityAlias(2)) & strDelimiter
	'objRsFacilityAlias.MoveNext

'wend

'Response.write strInnerValues

'objRsFacilityAlias.Close

  if StrCircuitID <> 0 or StrCircuitID ="" then

 StrSql ="SELECT A.CIRCUIT_ID,A.CIRCUIT_NAME,A.CIRCUIT_NUMBER,A.CIRCUIT_TYPE_CODE," &_
         "A.ADSL_SHELF_NUMBER,A.ADSL_LOOP_LOSS_DECIBEL,A.ADSL_TYPE_CODE," &_
         "A.BILLING_CUSTOMER_ID_A,A.BILLING_CUSTOMER_ID_B,A.SERVICE_LOCATION_ID_A,A.SERVICE_LOCATION_ID_B,A.CUSTOMER_SERVICE_ID_A,A.CUSTOMER_SERVICE_ID_B," &_
         "A.FIBRE_ORDER_NUMBER,A.FIBRE_CHECK_NUMBER,A.FRAME_RELAY_DLCI_FROM,A.FRAME_RELAY_DLCI_TO," &_
         "A.ADSL_SLOT_NUMBER,A.ADSL_TRAINED_SPEED,A.ADSL_DISTRIBUTION_BLOCK_NUMBER," &_
         "A.BILLING_TYPE_CODE,A.USAGE_CALCULATION_TYPE_CODE,H.CUSTOMER_NAME CUSTOMER_B,I.CUSTOMER_SERVICE_DESC SERVICE_B," &_
         "TO_CHAR(A.ADSL_DUE_DATE,'MON-DD-YYYY') ADSL_DUE_DATE,A.ADSL_ORDER_NUMBER,A.ADSL_CPE_OWNERSHIP_FLAG, " &_
         "A.NOC_REGION_LCODE,TO_CHAR(A.CIRCUIT_START_DATE,'MON-DD-YYYY') CIRCUIT_START_DATE,A.CIRCUIT_STATUS_CODE," &_
         "A.CIRCUIT_PROVIDER_CODE,A.MANAGED_BY_EMS_CODE,B.CUSTOMER_SERVICE_DESC," &_
         "E.CUSTOMER_NAME,C.SERVICE_LOCATION_NAME LOCATION_A,A.COMMENTS," &_
         "D.SERVICE_LOCATION_NAME LOCATION_B,sma_sp_userid.spk_sma_library.sf_get_full_username(A.CREATE_REAL_USERID) CREATE_REAL_USERID,A.RECORD_STATUS_IND," &_
         "TO_CHAR(A.CREATE_DATE_TIME,'MON-DD-YYYY HH:MI:SS') CREATE_DATE_TIME,sma_sp_userid.spk_sma_library.sf_get_full_username(A.UPDATE_REAL_USERID)UPDATE_REAL_USERID,TO_CHAR(A.UPDATE_DATE_TIME,'MON-DD-YYYY HH:MI:SS') UPDATE_DATE_TIME_CONV,A.UPDATE_DATE_TIME," &_
         "NVL(F.BUILDING_NAME,'NO BUILDING NAME') ||CHR(10)||NVL(F.STREET,'NO STREET ADDRESS')||CHR(10)||NVL(F.MUNICIPALITY_NAME,'NO MUNICIPALITY')||' '||" &_
         "NVL(F.PROVINCE_STATE_LCODE,'NO PROVINCE')||CHR(10)||NVL(F.POSTAL_CODE_ZIP,'NO POSTAL CODE') ADDRESS_A," &_
         "NVL(G.BUILDING_NAME,'NO BUILDING NAME') ||CHR(10)||NVL(G.STREET,'NO STREET ADDRESS')||CHR(10)||NVL(G.MUNICIPALITY_NAME,'NO MUNICIPALITY')||' '||" &_
         "NVL(G.PROVINCE_STATE_LCODE,'NO PROVINCE')||CHR(10)||NVL(G.POSTAL_CODE_ZIP,'NO POSTAL CODE') ADDRESS_B,A.FIBRE_LENGTH_M,A.FIBRE_INSTALLATION_COST,A.FIBRE_RISER_COST, " &_
         "A.FIBRE_FLOOR_A,A.FIBRE_BAY_A,A.FIBRE_SHELF_MODULE_A,A.FIBRE_POSITION_A,A.FIBRE_CABLE_A,A.FIBRE_PAIRS_A," &_
         "A.FIBRE_FLOOR_B,A.FIBRE_BAY_B,A.FIBRE_SHELF_MODULE_B,A.FIBRE_POSITION_B,A.FIBRE_CABLE_B,A.FIBRE_PAIRS_B " &_
         " FROM " &_
         " CRP.CIRCUIT A," &_
         "CRP.CUSTOMER_SERVICE B," &_
         "CRP.SERVICE_LOCATION C," &_
         "CRP.SERVICE_LOCATION D," &_
         "CRP.CUSTOMER E,"&_
         "crp.v_address_consolidated_street F," &_
         "crp.v_address_consolidated_street G," &_
         "CRP.CUSTOMER H," &_
         "CRP.CUSTOMER_SERVICE I "

      strWhereClause =  " WHERE " &_
                        "A.CUSTOMER_SERVICE_ID_A = B.CUSTOMER_SERVICE_ID(+) AND" &_
                        " A.SERVICE_LOCATION_ID_A = C.SERVICE_LOCATION_ID(+) AND" &_
                        " A.SERVICE_LOCATION_ID_B = D.SERVICE_LOCATION_ID(+) AND" &_
                        " A.BILLING_CUSTOMER_ID_A = E.CUSTOMER_ID(+) AND " &_
                        " C.ADDRESS_ID = F.ADDRESS_ID(+) AND " &_
                        " D.ADDRESS_ID = G.ADDRESS_ID(+) AND " &_
                        " A.CUSTOMER_SERVICE_ID_B = I.CUSTOMER_SERVICE_ID(+) AND" &_
                        " A.BILLING_CUSTOMER_ID_B = H.CUSTOMER_ID(+) AND " &_
                        " A.CIRCUIT_ID =" & StrCircuitID

      StrSql =  StrSql & " "& strWhereClause

  ' Response.Write "SQL STATEMENT WIH WHERE=" & StrSql & "<p>"

   'Create the command object

     'Create Recordset object

   'set cookies for circuitID to be used by facilityAliasDetail

   Dim  objRs
   set objRS = objConn.Execute(StrSql)

  if strNew <> "CLONED" then
   Response.Cookies("ParentCircuitID") = objRs("CIRCUIT_ID")
  end if

   if strNew = "CLONED" then
       StrCircuitID=""
       bolClone = true
   end if

  end if
   'IF StrCircuitID <> 0 THEN
   'Do while Not objRS.EOF OR  StrCircuitID =0
   'END IF
 

%>


<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<LINK REL="stylesheet" TYPE="text/css" HREF="stylesheets/styles.css">
<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
<script type="text/javascript" SRC="AccessLevels.js"></script>
<script type="text/javascript">
//set the heading
if (('<%=StrCircuitTyp%>' == 'ATMPVC') || ('<%=StrCircuitTyp%>' == 'PVC')) {
	setPageTitle("SMA - PVC");
}
else {
	setPageTitle("SMA - Facility");
}

//javascript code related to iFrame functionality

var strDelimiter='<%=strDelimiter%>';
var intCircuitID='<%=StrCircuitID%>';
var strCircuitType='<%=StrCircuitTyp%>';
var bolSaveRequired = false;
var intAccessLevel=<%=intAccessLevel%>;
var intConst_MessageDisplay=<%=intConst_MessageDisplay%>;


function iFrame_display(){
//called whenever a refresh of the iFrame is needed
//alert('CircuitID=' + intCircuitID);
	window.frames["aifr"].src = 'FacilityAlias.asp?CircuitID=' + intCircuitID+'&FacType='+document.fmfacDetail.selfactyp.item(document.fmfacDetail.selfactyp.selectedIndex).value;
}



function btn_iFrmAdd(){

var NewWin;
if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {alert('Access denied. Please contact your system administrator.'); return;}
NewWin=window.open("FacilityAliasDetail.asp?NewFacility=NEW" ,"NewWin","toolbar=no,status=no,width=700,height=250,menubar=no resize=no");
NewWin.focus();
}


function btn_iFrmUpdate(){
var NewWin;
if ((intAccessLevel & intConst_Access_Update) != intConst_Access_Update) {alert('Access denied. Please contact your system administrator.'); return;}
if (window.frames["aifr"].contentDocument.frmIFR.txtCircuitID.value !=="")
{
var strSource ="FacilityAliasDetail.asp?AliasID="+window.frames["aifr"].contentDocument.frmIFR.txtAliasID.value
				+'&CircuitID='+window.frames["aifr"].contentDocument.frmIFR.txtCircuitID.value
				+'&FacType='+document.getElementById("selfactyp").item(document.getElementById("selfactyp").selectedIndex).value;
NewWin=window.open(strSource ,"NewWin","toolbar=no,status=no,width=700,height=250,menubar=no resize=no");
NewWin.focus();
}
else
{
 alert('You must select a record to update!');
}

}

function btn_iFrmDelete()
{

if ((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete)
  {
  alert('Access denied. Please contact your system administrator.');
  return;
   }
if (window.frames["aifr"].contentDocument.frmIFR.txtAliasID.value !=="")
{	
	
	if (confirm('Do you really want to delete this Alias?')){
		window.frames["aifr"].src ="FacilityAlias.asp?txtFrmAction=DELETE&AliasID="
		+window.frames["aifr"].contentDocument.frmIFR.txtAliasID.value
		+"&hdnUpdateDateTime="+window.frames["aifr"].contentDocument.frmIFR.hdnUpdateDateTime.value
		+"&CircuitID="+window.frames["aifr"].contentDocument.frmIFR.txtCircuitID.value;
		}
 }
 else
 {
  alert('You must select a record to delete!');
 }
}

function body_onLoad(){
	iFrame_display();
}



function btnNew_click(){
 var strFacType;
  //this is a new record remove old circuit id
  if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {alert('Access denied. Please contact your system administrator.'); return;}
	DeleteCookie("ParentCircuitID");
	strFacType = document.fmfacDetail.selfactyp.item(document.fmfacDetail.selfactyp.selectedIndex).value;
	if (isWhitespace(strFacType ))
	{
	  alert("You must enter a facility");
	 }
	 else{
	SetCookie("FacilityType",strFacType);
	self.document.location.href ="FacilityDetail.asp?NewFacility=NEW&CircuitTyp="+strFacType;
	}
}


function fct_onChange(){
	bolSaveRequired = true
}

function fct_facility_number_onChange(){

if (document.fmfacDetail.txtcktname.value ==""){
document.fmfacDetail.txtcktname.value = document.fmfacDetail.txtcktnum.value;
}
document.fmfacDetail.btnSave.disabled = false;

}
function fct_facility_type_onChange(){
   var strFac;
   strFac = document.fmfacDetail.selfactyp.item(document.fmfacDetail.selfactyp.selectedIndex).value;
   //alert(document.fmfacDetail.hdnfacilityType.value);
   if (document.fmfacDetail.txtcktnum.value !="") {
   alert("You are about to change facility type, this will cause fields not related to the new facility to be nulled out!!");

   switch(document.fmfacDetail.hdnfacilityType.value){
   case  'FR':
     document.fmfacDetail.txtfrdlcifrom.value ="";
     document.fmfacDetail.txtfrdlcito.value ="";
     break;
   case  'ADSL':
     document.fmfacDetail.selfacadsltyp.selectedIndex=0;
     document.fmfacDetail.selmonth2.selectedIndex=0;
     document.fmfacDetail.selday2.selectedIndex=0;
     document.fmfacDetail.selyear2.selectedIndex=0;
     document.fmfacDetail.txtadslshelf.value ="";
     document.fmfacDetail.txtadslldb.value ="";
     document.fmfacDetail.txtadslslot.value ="";
     document.fmfacDetail.txtadsltsp.value ="";
     document.fmfacDetail.txtadsldisbl.value ="";
     document.fmfacDetail.txtadslorder.value ="";
     document.fmfacDetail.chkadslcpe.checked = false;
     break;
   case  'FIBRE':
     document.fmfacDetail.txtfibreord.value ="";
     document.fmfacDetail.txtfibrechk.value ="";
     break;
   case  'ATMPVC':
    document.fmfacDetail.selmgtbyems.selectedIndex=0;
    document.fmfacDetail.txtcusservb.value ="";
    document.fmfacDetail.txtcustomerb.value ="";
    document.fmfacDetail.selbilltype.SelectedIndex=0;
    document.fmfacDetail.selusgcalc.SelectedIndex=0;
    document.fmfacDetail.hdnCustomerIdB.value ="";
    document.fmfacDetail.hdnCustomerServIDB.value ="";
    break;
    } //end switch


	fct_onChange();

	}
	if(strFac=='ADSL')
	{document.fmfacDetail.txtformat.value= 'Format:604-261-7789'}
	else if(strFac!='ATMPVC')
	{document.fmfacDetail.txtformat.value='Format:12/DHAT/163789/000/BCTC/000'}

	document.fmfacDetail.CircuitTyp.value = strFac;
	//document.fmfacDetail.submit();

}



function fct_onDelete(){
 if (((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) || (document.fmfacDetail.txtRecordStatusInd.value == "D"))
  {
  alert('Access denied. Please contact your system administrator.');
  return;
   }
	if (confirm('Do you really want to delete this object?')){
		document.location = "FacilityDetail.asp?txtFrmAction=DELETE&CircuitID="+document.fmfacDetail.CircuitID.value+"&UpdateDateTime="+document.fmfacDetail.hdnUpdateDateTime.value;
	}

}


function fct_onUnload(){

}



</script>


<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>


function fmfacDetail_onsubmit() {

 var strMonth,strDay,strYear,strDate;

 if	((((intAccessLevel & intConst_Access_Create) == intConst_Access_Create) && (document.fmfacDetail.CircuitID.value == "")) || ((intAccessLevel & intConst_Access_Update) == intConst_Access_Update) && (document.fmfacDetail.CircuitID.value != ""))
 {
    //this cookie maintains the former operational status
	  SetCookie("OperStat", document.fmfacDetail.hdnoperationalstatus.value);
//check mandatory fields
	if (isWhitespace(document.fmfacDetail.txtcktnum.value)) {
		alert('Please enter a Circuit Number');
		document.fmfacDetail.txtcktnum.focus();
		return(false);
	}
	if (isWhitespace(document.fmfacDetail.txtcktname.value)) {
		alert('Please enter a Circuit Name');
		document.fmfacDetail.txtcktname.focus();
		return(false);
	}
	if (isWhitespace(document.fmfacDetail.selfactyp.item(document.fmfacDetail.selfactyp.selectedIndex).value)) {
		alert('Please enter a Facility Type');
		document.fmfacDetail.selfactyp.focus();
		return(false);
	}
	if (isWhitespace(document.fmfacDetail.selcktprov.item(document.fmfacDetail.selcktprov.selectedIndex).value)) {
		alert('Please enter a Circuit Provider');
		document.fmfacDetail.selcktprov.focus();
		return(false);
	}

//alert(document.fmfacDetail.selfacstat.item(document.fmfacDetail.selfacstat.selectedIndex).value);
// Attempt to allow for PVC to save without both Side A or Side B.
 if (((document.fmfacDetail.selfactyp.item(document.fmfacDetail.selfactyp.selectedIndex).value!='ATMPVC')
		&& (document.fmfacDetail.hdnfacilityType.value!='ATMPVC'))
	|| (document.fmfacDetail.selfacstat.item(document.fmfacDetail.selfacstat.selectedIndex).value=='OPER'))
{
	if (isWhitespace(document.fmfacDetail.hdnCustomerServIDA.value)) {
		alert('Please enter Customer Service A');
		document.fmfacDetail.btnCustomerServiceLookupA.focus();
		return(false);
	}

	if (isWhitespace(document.fmfacDetail.hdnCustomerIdA.value)) {
		alert('Please enter Customer A');
		document.fmfacDetail.btncustomerlookupA.focus();
		return(false);
	}

	if (isWhitespace(document.fmfacDetail.hdnServiceLocIdA.value)) {
		alert('Please enter Service Location A');
		document.fmfacDetail.btnServiceLocationLookupA.focus();
		return(false);
	}


	if (isWhitespace(document.fmfacDetail.hdnServiceLocIdB.value)) {
		alert('Please enter Service Location B');
		document.fmfacDetail.btnServiceLocationLookupB.focus();
		return(false);
	}


	if ((document.fmfacDetail.selfactyp.item(document.fmfacDetail.selfactyp.selectedIndex).value=='ATMPVC')&& (document.fmfacDetail.hdnfacilityType.value=='ATMPVC'))
	{


	if (isWhitespace(document.fmfacDetail.hdnCustomerServIDB.value)) {
		alert('Please enter Customer Service B');
		document.fmfacDetail.btnCustomerServiceBLookupB.focus();
		return(false);
	}


	if (isWhitespace(document.fmfacDetail.hdnCustomerIdB.value)) {
		alert('Please enter Customer B');
		document.fmfacDetail.btncustomerBlookupB.focus();
		return(false);
	}

	if (isWhitespace(document.fmfacDetail.selmgtbyems.item(document.fmfacDetail.selmgtbyems.selectedIndex).value)) {
		alert('Please enter an EMS code');
		document.fmfacDetail.selmgtbyems.focus();
		return(false);
		}

		if (isWhitespace(document.fmfacDetail.selusgcalc.item(document.fmfacDetail.selusgcalc.selectedIndex).value)) {
		alert('Please select a Usage Calc Type code');
		document.fmfacDetail.selusgcalc.focus();
		return(false);
		}

	}
}
	if (isWhitespace(document.fmfacDetail.selfacstat.item(document.fmfacDetail.selfacstat.selectedIndex).value)) {
		alert('Please enter the Facility Status');
		document.fmfacDetail.selfacstat.focus();
		return(false);
	}

	if ((document.fmfacDetail.selfactyp.item(document.fmfacDetail.selfactyp.selectedIndex).value=='ADSL')&& (document.fmfacDetail.hdnfacilityType.value=='ADSL'))
	{
	if (isWhitespace(document.fmfacDetail.chkadslcpe.value)) {
		alert('Please enter the CPE flag');
		document.fmfacDetail.chkadslcpe.focus();
		return(false);
	}

	if (isWhitespace(document.fmfacDetail.selfacadsltyp.item(document.fmfacDetail.selfacadsltyp.selectedIndex).value)) {
		alert('Please enter the ADSL Service Type');
		document.fmfacDetail.selfacadsltyp.focus();
		return(false);
	  }

	}

	if (isWhitespace(document.fmfacDetail.selregion.item(document.fmfacDetail.selregion.selectedIndex).value)) {
		alert('Please enter a region');
		document.fmfacDetail.selregion.focus();
		return(false);
	}


	if ((document.fmfacDetail.selfactyp.item(document.fmfacDetail.selfactyp.selectedIndex).value=='FIBRE')&& (document.fmfacDetail.hdnfacilityType.value=='FIBRE'))
	{

	if (isNaN(document.fmfacDetail.txtfibrelength.value))
	  {
	   alert('Please enter a valid Fibre Length');
	   document.fmfacDetail.txtfibrelength.focus();
	   return(false);
	  }

	  if (isNaN(document.fmfacDetail.txtfibrecost.value))
	  {
	   alert('Please enter a valid Riser Cost');
	   document.fmfacDetail.txtfibrecost.focus();
	   return(false);
	  }

	  if (isNaN(document.fmfacDetail.fibreinstallationcost.value))
	  {
	   alert('Please enter a valid Fibre Installation Cost');
	   document.fmfacDetail.fibreinstallationcost.focus();
	   return(false);
	  }


	}

	//Circuit Start Date

	strMonth = document.fmfacDetail.selmonth.item(document.fmfacDetail.selmonth.selectedIndex).value;
	strDay = document.fmfacDetail.selday.item(document.fmfacDetail.selday.selectedIndex).value;
	strYear = document.fmfacDetail.selyear.item(document.fmfacDetail.selyear.selectedIndex).value;

	if ((strMonth != "") & (strDay !="") & (strYear !=""))
	{
	strDate = strMonth + "/" + strDay + "/" + strYear;
	document.fmfacDetail.hdnCircuitStartDt.value = strDate;
	}
	else
	  if ((strMonth != "")||(strDay != "" || strYear != ""  )) {
	      alert('Please enter a complete facility start date');
	      document.fmfacDetail.selmonth.focus();
	      return(false);
		  }
	  else
		 {
	      strDate = "";
	      document.fmfacDetail.hdnCircuitStartDt.value = strDate;
	     }

	if ((document.fmfacDetail.selfactyp.item(document.fmfacDetail.selfactyp.selectedIndex).value=='ADSL')&& (document.fmfacDetail.hdnfacilityType.value=='ADSL')){
	//adsl due date

	strMonth = document.fmfacDetail.selmonth2.item(document.fmfacDetail.selmonth2.selectedIndex).value;
	strDay = document.fmfacDetail.selday2.item(document.fmfacDetail.selday2.selectedIndex).value;
	strYear = document.fmfacDetail.selyear2.item(document.fmfacDetail.selyear2.selectedIndex).value;

	if ((strMonth != "") & (strDay != "") & (strYear != ""))
	   {
	   strDate = strMonth + "/" + strDay + "/" + strYear;
	   document.fmfacDetail.hdnAdslDueDt.value = strDate;
	   //disable Save button
	   }
	else
	  if ((strMonth != "")||(strDay != "" || strYear != ""  )) {
	      alert('Please enter a complete ASDL due date');
	      document.fmfacDetail.selmonth2.focus();
	      return(false);
		  }
	  else
	      {
	       strDate = "";
	       document.fmfacDetail.hdnAdslDueDt.value = strDate;
	      }
	}
	bolSaveRequired = false;

	//submit the form
	document.fmfacDetail.txtFrmAction.value = "SAVE";

	return(true);
	} //end if intAccessLevel >= intConst_Access_Create
	else {
	alert('Access denied. Please contact your system administrator.');
	 return(false);
	 }

}

function btnSave_onclick()
{
   var bolretval;
   bolretval= fmfacDetail_onsubmit();
    if(bolretval)
     document.fmfacDetail.submit();
}

function fct_lookupCustomerService(CustService){
 var strFacility = document.fmfacDetail.selfactyp.item(document.fmfacDetail.selfactyp.selectedIndex).value;

 switch(CustService){
   case  'A':
     SetCookie("ServiceEnd", CustService);

	if (document.fmfacDetail.txtcusserva.value != "")
	{
	 SetCookie("CustomerService", document.fmfacDetail.txtcusserva.value);
	 break;
	 }
	case  'B':
	 SetCookie("ServiceEnd", CustService);
	if( strFacility == 'ATMPVC'){
	  if (document.fmfacDetail.txtcusservb.value != "") {
	  SetCookie("CustomerService", document.fmfacDetail.txtcusservb.value);
	  break;
	  }
	 }
   }

	SetCookie("WinName", 'Popup');
	if( strFacility == 'ATMPVC'){
		window.open('SearchFrame.asp?fraSrc=CustServPVC', 'Popup', 'top=50, left=100,  WIDTH=800, HEIGHT=600' ) ;
		}
	else {
		window.open('SearchFrame.asp?fraSrc=CustServ', 'Popup', 'top=50, left=100,  WIDTH=800, HEIGHT=600' ) ;
	}
	//enable Save button - may not need it but onChange event suck (is not fired when use lookup)
	document.fmfacDetail.btnSave.disabled = false;
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
 var strCustomerIdA,strCustomerIdB,strServiceLocAId,strServiceLocBId;
 var strCustomerSvcIdA,strCustomerSvcIdB,strCustomerServiceA,strCustomerServiceB,strFacTyp;


    strFacTyp = document.fmfacDetail.selfactyp.item(document.fmfacDetail.selfactyp.selectedIndex).value;

	strCustomerIdA = document.fmfacDetail.hdnCustomerIdA.value;
	strCustomerIdB = document.fmfacDetail.hdnCustomerIdB.value;
	strServiceLocAId = document.fmfacDetail.hdnServiceLocA.value;
	strServiceLocBId = document.fmfacDetail.hdnServiceLocB.value;
	strCustomerSvcIdA = document.fmfacDetail.hdnCustomerServIDA.value;
	strCustomerSvcIdB = document.fmfacDetail.hdnCustomerServIDB.value;
	strCustomerServiceA = document.fmfacDetail.txtcusserva.value;
	if  (strFacTyp == 'ATMPVC')
	{strCustomerServiceB = document.fmfacDetail.txtcusservb.value;}



	strPageName = document.fmfacDetail.selNavigate.item(document.fmfacDetail.selNavigate.selectedIndex).value ;


	// from Customer Detail Page, user will always navigate to lists i.e. search pages.

	switch (strPageName) {

	  case 'CustA':
	    if (strCustomerIdA !="")
	    {
	    document.fmfacDetail.selNavigate.selectedIndex=0;
		self.location.href = "CustDetail.asp?CustomerId=" + strCustomerIdA;
		}
		else{
		  alert("Unable to Navigate to Customer Detail as Customer does not exist");
		  }
		break;
	case 'CustB':
	  if (strCustomerIdB !="")
	    {
	    document.fmfacDetail.selNavigate.selectedIndex=0;
		self.location.href = "CustDetail.asp?CustomerId=" + strCustomerIdB;
		}
		else{
		  alert("Unable to Navigate to Customer Detail as Customer does not exist");
		  }
		break;
	 case 'ServLocA':
	   if (strServiceLocAId !="")
	    {
	    document.fmfacDetail.selNavigate.selectedIndex=0;
		self.location.href = "ServLocDetail.asp?ServLocID=" + strServiceLocAId ;
		}
		else{
		  alert("Unable to Navigate to Location Detail as Location does not exist");
		  }
		break;
	 case 'ServLocB':
	  if (strServiceLocBId !="")
	      {
	      document.fmfacDetail.selNavigate.selectedIndex=0;
		   self.location.href = "ServLocDetail.asp?ServLocID=" + strServiceLocBId ;
		  }
		else{
		  alert("Unable to Navigate to Location Detail as Location does not exist");
		  }
		  break;
	case 'CustServA':
	   if (strCustomerSvcIdA !="")
	      {
	      document.fmfacDetail.selNavigate.selectedIndex=0;
		  self.location.href = "CustServDetail.asp?CustServID=" + strCustomerSvcIdA ;
		  }
		else{
		  alert("Unable to Navigate to Service Detail as Service does not exist");
		  }
		break;
	case 'CustServB':
	   if (strCustomerSvcIdB !="")
	      {
	      document.fmfacDetail.selNavigate.selectedIndex=0;
		  self.location.href = "CustServDetail.asp?CustServID=" + strCustomerSvcIdB ;
		  }
		else{
		  alert("Unable to Navigate to Service Detail as Service does not exist");
		  }
		break;
	case 'Correlation':
	 {
	  document.fmfacDetail.selNavigate.selectedIndex=0;
	  SetCookie("ObjectName", document.fmfacDetail.txtcktname.value);
	  SetCookie("Type", "Facility");
	  self.location.href = "searchFrame.asp?fraSrc=Correlation";
	  break;
	  }
	  case 'OrderHistoryA':
	  if (strCustomerSvcIdA !="")
	  {
	   document.fmfacDetail.selNavigate.selectedIndex=0;
		//SetCookie("CustomerServiceName", strCustomerServiceA);
		SetCookie("CustomerServiceID", strCustomerSvcIdA);
		self.location.href = 'SearchFrame.asp?fraSrc=OrderHistory';
	  }
	  else
	  {
	    alert("Unable to Navigate to Order History as Service ID does not exist");
	  }
	  break ;
	  case 'OrderHistoryB':
	  if (strCustomerSvcIdB !="")
	  {
	    document.fmfacDetail.selNavigate.selectedIndex=0;
		//SetCookie("CustomerServiceName", strCustomerServiceB);
		SetCookie("CustomerServiceID", strCustomerSvcIdB);
		self.location.href = 'SearchFrame.asp?fraSrc=OrderHistory';
	  }
	  else
	  {
	    alert("Unable to Navigate to Order History as Service ID does not exist");
	  }
	  break ;
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
var strFacility = document.fmfacDetail.selfactyp.item(document.fmfacDetail.selfactyp.selectedIndex).value;
var strOperStat = document.fmfacDetail.selfacstat.item(document.fmfacDetail.selfacstat.selectedIndex).value;
var strWinStatus='<%=strWinMessage%>';
fct_displayStatus(strWinStatus);


//Used to control  the changing of facilities
DeleteCookie("CustomerServID") ;
DeleteCookie("CustName") ;
DeleteCookie("ServiceLocName") ;
DeleteCookie("CustID") ;
DeleteCookie("ServiceLocID") ;
DeleteCookie("CustomerServA") ;
DeleteCookie("Address") ;

//Maintains original facility type when changing facilities
document.fmfacDetail.hdnfacilityType.value = strFacility;
document.fmfacDetail.hdnoperationalstatus.value = strOperStat;

 iFrame_display();
 return true;
}

function btnCalendar_onclick(intDateFieldNo) {
	var NewWin;
	    SetCookie("Field",intDateFieldNo);
		NewWin=window.open("calendar.asp","NewWin","toolbar=no,status=no,width=260,height=225,menubar=no resize=no");
		//NewWin.creator=self;
	NewWin.focus();
}

function fct_lookupCustomer(End){
   var strFacility = document.fmfacDetail.selfactyp.item(document.fmfacDetail.selfactyp.selectedIndex).value;
   switch(End){
   case  'A':
     SetCookie("ServiceEnd", End);
	if (document.fmfacDetail.txtcustomera.value != "")
	{
	 SetCookie("CustomerName", document.fmfacDetail.txtcustomera.value);
	 break;
	 }
	case  'B':
	  SetCookie("ServiceEnd", End);
	 if( strFacility == 'ATMPVC'){
	  if (document.fmfacDetail.txtcustomerb.value != "") {
	  SetCookie("CustomerName", document.fmfacDetail.txtcustomerb.value);
	  break;
	 }
	 }
  }
	SetCookie("WinName", 'Popup');
	window.open('SearchFrame.asp?fraSrc=Cust', 'Popup', 'top=50, left=100, height=600, width=800' ) ;
	//enable Save button - may not need it but onChange event suck (is not fired when use lookup)
	document.fmfacDetail.btnSave.disabled = false;
}

function btnServiceLocationLookup_onclick(End) {
//******************************************************************************************
//
//
//
//******************************************************************************************


  switch(End){

   case  'A':
    SetCookie("ServiceEnd", End);
	if (document.fmfacDetail.txtsrvloca.value != "")
	{
	 SetCookie("ServLocName", document.fmfacDetail.txtsrvloca.value);
	 }
	 else
	 {
	   SetCookie("ServLocName", "new") ;
	 }
	 break;
	case  'B':
	  SetCookie("ServiceEnd", End);
	  if (document.fmfacDetail.txtsrvlocb.value != "") {
	  SetCookie("ServLocName", document.fmfacDetail.txtsrvlocb.value);
	  }
	  else
	  {
	   SetCookie("ServLocName", "new") ;
	  }
	  break;
    }

	SetCookie("WinName", "Popup") ;
	window.open('SearchFrame.asp?fraSrc=ServLoc','Popup','top=50, left=100, height=600, width=800') ;
 }


 function fct_onClone(){
 if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {alert('Access denied. Please contact your system administrator.'); return;}

 document.location ="FacilityDetail.asp?NewFacility=CLONED&CircuitTyp="+strCircuitType+'&CircuitID='+intCircuitID;
 alert("Record Cloned. Please make changes then save!");
}

function body_onBeforeUnload(){
    document.fmfacDetail.btnSave.focus();
	if (bolSaveRequired) {
		if	(((intAccessLevel & intConst_Access_Create == intConst_Access_Create) && (document.fmfacDetail.CircuitID.value == "")) || ((intAccessLevel & intConst_Access_Update == intConst_Access_Update) && (document.fmfacDetail.CircuitID.value != ""))) {
			event.returnValue = "There is unsaved data on the screen. To save changes, click CANCEL below then click SAVE on the main form.";
		}
	}
}


function btnReferences_onclick() {
var strOwner = 'CRP' ;
var strTableName = 'CIRCUIT' ;
var strRecordID = document.fmfacDetail.CircuitID.value ;
var URL ;
  if (document.fmfacDetail.CircuitID.value ==""){
     alert('No references. This is a new record.');
     return false;
     }
  else
	{
	URL ='Dependency.asp?Owner=' + strOwner + '&TableName=' + strTableName + '&RecordID='+ strRecordID   ;
	window.open(URL, 'Popup', 'top=100, left=100, WIDTH=500, HEIGHT=300'  ) ;
	return true;
	}
}


function fct_onReset(){
	if(confirm('All changes will be lost. Do you really want to reset this page?')){
		bolSaveRequired = false;
		<%if not bolclone then%>
			document.location = "FacilityDetail.asp?CircuitID=<%=StrCircuitID %>&CircuitTyp=<%=StrCircuitTyp %>";
		<%else%>
			document.location = "FacilityDetail.asp?CircuitID=<%=strTmpCircuitID %>&CircuitTyp=<%=StrCircuitTyp %>&NewFacility=CLONED";
		<%end if%>
	}
}

function email_setup()
{
var strSubject;
var strBody;
var strMonth,strDay,strYear,strDate,URL

strMonth = document.fmfacDetail.selmonth2.item(document.fmfacDetail.selmonth2.selectedIndex).value;
strDay = document.fmfacDetail.selday2.item(document.fmfacDetail.selday2.selectedIndex).value;
strYear = document.fmfacDetail.selyear2.item(document.fmfacDetail.selyear2.selectedIndex).value;

if ((strMonth != "") & (strDay != "") & (strYear != ""))
	{

	strDate = strMonth + "/" + strDay + "/" + strYear;
    }
else
   {
   strDate = "";
   }

strSubject = "SMA - ADSL Circuit Status Change -" + document.fmfacDetail.txtcktnum.value+"/"+document.fmfacDetail.txtcktname.value;
strBody = "The following is a list of ADSL Circuit(s) in SMA that have been made OPERATIONAL\n\n";
strBody = strBody+"Circuit ID:  " + document.fmfacDetail.txtcktnum.value +"\n";
strBody = strBody+"Circuit Name:  " +document.fmfacDetail.txtcktname.value+"\n";
strBody = strBody+"Circuit Type:  " + document.fmfacDetail.selfactyp.item(document.fmfacDetail.selfactyp.selectedIndex).value+"\n";
strBody = strBody+"Former Status: " + GetCookie ("OperStat")+"\n";
strBody = strBody+"New Status: " + document.fmfacDetail.selfacstat.item(document.fmfacDetail.selfacstat.selectedIndex).value+"\n";
strBody = strBody+"ADSL Due Date: " + strDate+"\n";
strBody = strBody+"ADSL Train Speed: " + document.fmfacDetail.txtadsltsp.value+"\n";
//alert(strBody);

URL ="ADSLEmailSend.asp?CircuitID="+ document.fmfacDetail.CircuitID.value+"&CustServID="+document.fmfacDetail.hdnCustomerServIDA.value+"&body="+escape(strBody)+"&subject="+escape(strSubject);
EmailWin=window.open(URL ,"EmailWin","toolbar=no,status=no,width=800,height=600,menubar=no resize=no");
EmailWin.focus();
}

</SCRIPT>

</HEAD>
<BODY LANGUAGE=javascript onUnload="fct_onUnload();" onBeforeUnload="body_onBeforeUnload();" onload="return window_onload();">
<FORM NAME=fmfacDetail METHOD=POST ACTION="FacilityDetail.asp"  onsubmit="return fmfacDetail_onsubmit(this)">
<table width="100%" border=0 COLS=4>
<%if (StrCircuitTyp <> "ATMPVC") THEN %>
<thead>
	<tr>
	<td colspan=3 align=left>Facility Detail</td>
	<td><SELECT ALIGN=RIGHT valign=top id=selNavigate name=selNavigate LANGUAGE=javascript onchange="return selNavigate_onchange()" <%if StrCircuitID = 0 then Response.Write "disabled" end if%>>
		<OPTION value="DEFAULT">Quickly Goto ...</OPTION>
		<OPTION value="CustA" >Customer A</OPTION>
		<OPTION value="ServLocA" >Service Location A</OPTION>
		<OPTION value="ServLocB" >Service Location B</OPTION>
		<OPTION value="CustServA" >Customer Service A</OPTION>
		<OPTION value="Correlation" >Correlation</OPTION>
		<OPTION value="OrderHistoryA" >Order History</OPTION>

	</SELECT> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	</td>
	</tr>
<%else%>
<thead>
	<tr>
	<td colspan=3 align=left>PVC Detail</td>
	<td><SELECT ALIGN=RIGHT valign=top id=selNavigate name=selNavigate LANGUAGE=javascript onchange="return selNavigate_onchange()" <% if StrCircuitID = 0 then Response.Write "disabled" end if%>>
		<OPTION value="DEFAULT">Quickly Goto ...</OPTION>
		<OPTION value="CustA" >CustomerA</OPTION>
		<OPTION value="CustB" >CustomerB</OPTION>
		<OPTION value="ServLocA" >Service Location A</OPTION>
		<OPTION value="ServLocB" >Service Location B</OPTION>
		<OPTION value="CustServA" >Customer Service A</OPTION>
		<OPTION value="CustServB" >Customer Service B</OPTION>
		<OPTION value="Correlation" >Correlation</OPTION>
		<OPTION value="OrderHistoryA" >Order History A</OPTION>
		<OPTION value="OrderHistoryB" >Order History B</OPTION>
		</SELECT>
	</td>
	</tr>
<%end if %>
	<tr><td width=100% align=left colSpan=4>General Information</td></tr>
</thead>
	<tbody>
<TR>
    <%if (StrCircuitTyp <> "ATMPVC") THEN %>
	 <TD ALIGN=RIGHT NOWRAP WIDTH=25%>Facility Type<font color=red>*</font></TD>
	 <%else %>
	 <TD ALIGN=RIGHT NOWRAP WIDTH=25%>Type<font color=red>*</font></TD>
	 <%end if %>

	 <TD COLSPAN=3>
	 <SELECT id=selfactyp name=selfactyp style="HEIGHT: 20px; WIDTH: 120px" onchange ="fct_facility_type_onChange();">
		<OPTION></OPTION>
		<%Do while Not objRsFacTyp.EOF
			Response.write "<OPTION "
		if StrCircuitID <> 0 then
			if objRsFacTyp("CIRCUIT_TYPE_CODE") = objRs("CIRCUIT_TYPE_CODE") then
			  Response.Write " selected "
			end if
		else
		 if objRsFacTyp("CIRCUIT_TYPE_CODE") = Request.Cookies("FacilityType") then
		  Response.Write " selected "
		 end if
		end if
			Response.write "VALUE ="& """"&objRsFacTyp("CIRCUIT_TYPE_CODE")&"""" & ">" & objRsFacTyp("CIRCUIT_TYPE_CODE") & "</OPTION>" & vbCrLf
			objRsFacTyp.MoveNext
		 Loop
		%>
		</SELECT>
		<%
		dim Format
		select case ucase(StrCircuitTyp)
		case "ADSL"
		  Format = "Format : 604-261-7789"
		case "RFLINK"
		  Format = "Format : 604-261-7789"
		case "EVDO"
		  Format = "Format : 604-261-7789"
		case "MOBILE"
		  Format = "Format : 604-261-7789"
		case "SATELL"
		  Format = "Format : e.g. BCLC8B2T"
		case "CMI4"
		  Format = "Format : 604-261-7789"
		case "BUS4"
		  Format = "Format : 604-261-7789"
		case "CON15"
		   Format = "Format : 604-261-7789"
		case "PRO25"
		   Format = "Format : 604-261-7789"
		case "E-ADSL"
		   Format = "Format : 604-261-7789"
		case "ATMPVC"
		  Format = ""
		 case ""
		  Format = "Format : 604-261-7789"
		case else
		  Format = "Format : 12/DHAT/163789/000/BCTC/000"
		end select

		%>
		<INPUT READONLY name=txtformat style="border:0;width:250px;background:none" value="<%=Format%>">
	</TD>
</TR>
<TR>
   <%if (StrCircuitTyp <> "ATMPVC") THEN %>
	<TD ALIGN=RIGHT NOWRAP width=25%>Facility #<font color=red>*</font></TD>
	<%else%>
	<TD ALIGN=RIGHT NOWRAP width=25%>PVC #<font color=red>*</font></TD>
	<%end if %>

	<TD COLSPAN=3><INPUT id=txtcktnum name=txtcktnum size=50 maxlength=50 value= <%if StrCircuitID <> 0 then  Response.Write """"&routineHtmlString(objRS("CIRCUIT_NUMBER"))&"""" else Response.Write """""" end if%> onchange ="fct_facility_number_onChange();"></TD>
</TR>
<TR>
	<%if (StrCircuitTyp <> "ATMPVC") THEN %>
		<TD ALIGN=RIGHT NOWRAP width=25% >Facility Name<font color=red>*</font></TD>
	<%else%>
		<TD ALIGN=RIGHT NOWRAP width=25% >PVC Name<font color=red>*</font></TD>
	<%end if %>

	<TD ALIGN=LEFT COLSPAN=3><INPUT id=txtcktname name=txtcktname size=65 maxlength=65 value= <%if StrCircuitID <> 0  then  Response.Write """"&routineHtmlString(objRS("CIRCUIT_NAME"))&"""" else Response.Write """""" end if%> onchange ="fct_onChange();"></TD>
</TR>
 <%
  'IF objRs("CIRCUIT_TYPE_CODE") = "FIBRE" THEN
   if (StrCircuitTyp<> "ATMPVC") AND (StrCircuitTyp<> "ADSL") AND (StrCircuitTyp<> "DCS")THEN

  %>
<TR>
	<TD ALIGN=RIGHT NOWRAP width=25%>USSO #</TD>
	<TD width=25%><INPUT id=txtfibreord name=txtfibreord style="HEIGHT: 21px; WIDTH: 200px" value= <%if StrCircuitID <> 0  then  Response.Write """"&routineHtmlString(objRS("FIBRE_ORDER_NUMBER"))&"""" else Response.Write """""" end if%> onchange ="fct_onChange();"></TD>
	<TD ALIGN=RIGHT NOWRAP width=25%>NSRC Checklist #</TD>
	<TD width=25%><INPUT id=txtfibrechk name=txtfibrechk style="HEIGHT: 21px; WIDTH: 200px" value= <%if StrCircuitID <> 0 then  Response.Write """"&routineHtmlString(objRS("FIBRE_CHECK_NUMBER"))&"""" else Response.Write """""" end if%> onchange ="fct_onChange();"></TD>
</TR>
<% END IF
  IF (StrCircuitTyp= "FIBRE") THEN
%>
<TR>
	<TD ALIGN=RIGHT NOWRAP width=25% >Fibre Length</TD>
	<TD width=25%><INPUT name=txtfibrelength  style="HEIGHT: 21px; WIDTH: 200px" value= <%if StrCircuitID <> 0 then  Response.Write """"&objRS("FIBRE_LENGTH_M")&"""" else Response.Write """""" end if%> onchange ="fct_onChange();">(M)</TD>
	<TD ALIGN=RIGHT NOWRAP width=25% >Riser Cost</TD>
	<TD width=25%><INPUT name=txtfibrecost  style="HEIGHT: 21px; WIDTH: 200px" value= <%if StrCircuitID <> 0 then  Response.Write """"&objRS("FIBRE_RISER_COST")&"""" else Response.Write """""" end if%> onchange ="fct_onChange();"></TD>
</TR>

<TR>
	<TD ALIGN=RIGHT NOWRAP width=25% >Fibre Installation Cost</TD>
	<TD width=25%><INPUT  name=fibreinstallationcost style="HEIGHT: 22px; WIDTH: 115px" value= <%if StrCircuitID <> 0 then  Response.Write """"&objRS("FIBRE_INSTALLATION_COST")&"""" else Response.Write """""" end if%> onchange ="fct_onChange();"></TD>
	<td width=25%>&nbsp;</td>
	<td width=25%>&nbsp;</td>
</TR>
<%END IF %>

<%
'END IF
'IF objRs("CIRCUIT_TYPE_CODE") = "FR" THEN
if (StrCircuitTyp = "FR") THEN
%>
<TR>
	<TD ALIGN=RIGHT NOWRAP width=25%>F.R DLCI From</TD>
	<TD width=25%><INPUT id=txtfrdlcifrom name=txtfrdlcifrom style="HEIGHT: 21px; WIDTH: 200px" value= <%if StrCircuitID <> 0 then  Response.Write """"&routineHtmlString(objRS("FRAME_RELAY_DLCI_FROM"))&"""" else Response.Write """""" end if%> onchange ="fct_onChange();"></TD>
	<TD ALIGN=RIGHT NOWRAP width=25%>F.R DLCI To</TD>
	<TD width=25%><INPUT id=txtfrdlcito name=txtfrdlcito style="HEIGHT: 21px; WIDTH: 200px" value= <%if StrCircuitID <> 0 then  Response.Write """"&routineHtmlString(objRS("FRAME_RELAY_DLCI_TO"))&"""" else Response.Write """""" end if%> onchange ="fct_onChange();"></TD>
</TR>
<%
END IF
'IF objRs("CIRCUIT_TYPE_CODE") = "ADSL" THEN
if (StrCircuitTyp = "ADSL") THEN
%>
<TR>
	<TD ALIGN=RIGHT NOWRAP width=25%>ADSL Shelf #</TD>
	<TD width=25%><INPUT id=txtadslshelf name=txtadslshelf style="HEIGHT: 21px; WIDTH: 111px" value= <%if StrCircuitID <> 0 then  Response.Write """"&routineHtmlString(objRS("ADSL_SHELF_NUMBER"))&"""" else Response.Write """""" end if%> onchange ="fct_onChange();"></TD>
	<TD ALIGN=RIGHT NOWRAP width=25%>ADSL Loop Loss db</TD>
	<TD><INPUT id=txtadslldb name=txtadslldb style="HEIGHT: 21px; WIDTH: 101px" value= <%if StrCircuitID <> 0 then  Response.Write """"&routineHtmlString(objRS("ADSL_LOOP_LOSS_DECIBEL"))&"""" else Response.Write """""" end if%> onchange ="fct_onChange();"></TD>
</TR>
<TR>
	<TD ALIGN=RIGHT NOWRAP width=25%>ADSL Service Type<font color=red>*</font></TD>
	<TD width=25%><SELECT id=selfacadsltyp name=selfacadsltyp style="HEIGHT: 20px; WIDTH: 120px" onchange ="fct_onChange();">
		<OPTION></OPTION>
		<%
		Do while Not objRsAdslTyp.EOF
			Response.write "<OPTION "
		if StrCircuitID <> 0 then
			if objRsAdslTyp("ADSL_TYPE_CODE") = objRs("ADSL_TYPE_CODE") then
				Response.Write " selected "
			end if
		end if
			Response.write "VALUE ="& """"&objRsAdslTyp("ADSL_TYPE_CODE")&"""" & ">" & routineHtmlString(objRsAdslTyp("ADSL_TYPE_DESC")) & "</OPTION>" &vbCrLf
			objRsAdslTyp.MoveNext
		Loop
		%>
		</SELECT>
	</TD>
	<TD ALIGN=RIGHT NOWRAP>ADSL Slot #</TD>
	<TD><INPUT id=txtadslslot name=txtadslslot style="HEIGHT: 22px; WIDTH: 115px" value= <%if StrCircuitID <> 0 then  Response.Write """"&routineHtmlString(objRS("ADSL_SLOT_NUMBER"))&"""" else Response.Write """""" end if%> onchange ="fct_onChange();"></TD>
</TR>
<TR>
	<TD ALIGN=RIGHT NOWRAP width=25%>ADSL Trained Speed</TD>
	<TD width=25%><INPUT id=txtadsltsp name=txtadsltsp style="HEIGHT: 22px; WIDTH: 115px" value= <%if StrCircuitID <> 0 then  Response.Write """"&routineHtmlString(objRS("ADSL_TRAINED_SPEED"))&"""" else Response.Write """""" end if%> onchange ="fct_onChange();">
		<INPUT id=btnemail name=btnemail type=button value=Email LANGUAGE=javascript onclick="email_setup();">
	</TD>
	<TD ALIGN=RIGHT NOWRAP>ADSL Dist. Block #</TD>
	<TD width=25%><INPUT id=txtadsldisbl name=txtadsldisbl style="HEIGHT: 22px; WIDTH: 115px" value= <%if StrCircuitID <> 0 then  Response.Write """"&routineHtmlString(objRS("ADSL_DISTRIBUTION_BLOCK_NUMBER"))&"""" else Response.Write """""" end if%> onchange ="fct_onChange();"></TD>
</TR>
<TR>
	<TD ALIGN=RIGHT NOWRAP width=25%>ADSL Due Date</TD>
	<TD width=25% COLSPAN=1><SELECT name=selmonth2 style="HEIGHT: 20px; WIDTH: 70px" onchange ="fct_onChange();">
		<OPTION></OPTION>
		<%
		dim j
		for j = 1 to 12
			Response.Write "<option "
		if StrCircuitID <> 0 then
			if j = month(objRS("ADSL_DUE_DATE")) then
				Response.Write " selected "
			end if
		 end if
			if j < 10 then
				j="0"&j
			end if
			Response.write " VALUE ="& j & ">" &ucase(monthName(j,true)) & "</OPTION>" &vbCrLf
		next
		%>
		</SELECT>

		<SELECT  name=selday2 style="HEIGHT: 20px; WIDTH: 60px" onchange ="fct_onChange();">
		<OPTION></OPTION>
		<%
		for j = 1 to 31
			Response.Write "<option "
		if StrCircuitID <> 0 then
			if j = day(objRS("ADSL_DUE_DATE")) then
				Response.Write " selected "
			end if
		end if
			if j < 10 then
				j="0"&j
			end if
			Response.write " VALUE ="& j & ">" &j & "</OPTION>"
		next
		%>
		</SELECT>

		<SELECT  name=selyear2 style="HEIGHT: 20px; WIDTH: 60px" onchange ="fct_onChange();">
		<OPTION></OPTION>
		<%
		dim m,baseYear1
		baseYear1 = 1994
		for m = 0 to 30
			Response.Write "<option "
		if StrCircuitID <> 0 then
			if (baseYear1+m) = year(objRS("ADSL_DUE_DATE")) then
				Response.Write " selected "
			end if
		end if
			Response.write " VALUE ="& baseYear1+m & ">" &baseYear1+m & "</OPTION>"
		next
		%>
		</SELECT>
		<INPUT type="button" value="..." id=btnCalendar name=btnCalendar LANGUAGE=javascript onclick="btnCalendar_onclick(2);fct_onChange();">
	</TD>
	<TD ALIGN=RIGHT NOWRAP width=25%>ADSL Order #</TD>
	<TD width=25%><INPUT id=txtadslorder name=txtadslorder style="HEIGHT: 22px; WIDTH: 115px" value= <%if StrCircuitID <> 0 then  Response.Write """"&objRS("ADSL_ORDER_NUMBER")&"""" else Response.Write """""" end if%> onchange ="fct_onChange();"></TD>
</TR>
<TR>
	<%
	Response.Write "<TD ALIGN=RIGHT NOWRAP >ADSL CPE Ownership<font color=red>*</font></TD><TD><INPUT TYPE=CHECKBOX NAME=""chkadslcpe""  "
  if StrCircuitID <> 0 then
	if objRs("ADSL_CPE_OWNERSHIP_FLAG") = "Y" then
	   Response.Write " CHECKED "
	'else
	 'Response.Write " VALUE=""N"" "
	end if
  'else
   ' Response.Write " VALUE=""N"" "
  end if
	Response.write " onclick =""fct_onChange();"" > Tac Owned</TD> "
END IF
	%>
	<TD align=RIGHT nowrap width=25%>Region<font color=red>*</font></TD>
	<TD align=left width=25%><SELECT name=selregion onchange ="fct_onChange();">
		<OPTION></OPTION>
		<%Do while Not objRsRegion.EOF
			Response.write "<OPTION "
		if StrCircuitID <> 0 then
			if objRsRegion("NOC_REGION_LCODE") = objRs("NOC_REGION_LCODE") then
				Response.Write " selected "
			end if
		end if
			Response.write " VALUE ="& """"&objRsRegion("NOC_REGION_LCODE")&"""" & ">" & objRsRegion("NOC_REGION_DESC") & "</OPTION>"
			objRsRegion.MoveNext
		 Loop
		%>
		</SELECT>
	</TD>
</TR>
<TR>
	<%if (StrCircuitTyp <> "ATMPVC") THEN %>
		<TD align=RIGHT nowrap width=25%>Facility Start Date</TD>
	<%ELSE%>
		<TD align=RIGHT nowrap width=25%>PVC Install Date</TD>
	<%END IF %>

	<TD width=25%>
		<SELECT name=selmonth style="HEIGHT: 20px; WIDTH: 70px" onchange ="fct_onChange();">
		<OPTION></OPTION>
		<%
		dim k

		for k = 1 to 12
		  Response.Write "<option "
	   if StrCircuitID <> 0 then
		 if k = month(objRS("CIRCUIT_START_DATE")) then
		   Response.Write " selected "
		 end if
	   end if
		 if k < 10 then
		 k="0"&k
		 end if
		 Response.write " VALUE ="& k & ">" &ucase(monthName(k,true)) & "</OPTION>"
		 next
		%>
		</SELECT>

		<SELECT  name=selday style="HEIGHT: 20px; WIDTH: 60px" onchange ="fct_onChange();">
			<OPTION></OPTION>
			<%
			for k = 1 to 31
				Response.Write "<option "
			 if StrCircuitID <> 0 then
				if k = day(objRS("CIRCUIT_START_DATE")) then
					Response.Write " selected "
				end if
			 end if
				if k < 10 then
					k="0"&k
				end if
				Response.write " VALUE ="& k & ">" &k & "</OPTION>" & vbCrLf
			next
			%>
		</SELECT>
		<SELECT  name=selyear style="HEIGHT: 20px; WIDTH: 60px" onchange ="fct_onChange();">
		<OPTION></OPTION>
		<%
		dim i,baseYear
		baseYear = 1994
		for i = 0 to 30
		  Response.Write "<option "
		if StrCircuitID <> 0 then
		 if (baseYear+i) = year(objRS("CIRCUIT_START_DATE")) then
		   Response.Write " selected "
		 end if
		end if
		 Response.write " VALUE ="& baseYear+i & ">" &baseYear+i & "</OPTION>"
		 next
		%>
		</SELECT>
		<INPUT type="button" value="..." id=btnCalendar name=btnCalendar LANGUAGE=javascript onclick="btnCalendar_onclick(1);fct_onChange();">
	</TD>
	<TD ALIGN=RIGHT NOWRAP width=25%>Operational Status<font color=red>*</font></TD>
	<TD width=25%>
		<SELECT id=selfacstat name=selfacstat style="HEIGHT: 20px; WIDTH: 120px" onchange ="fct_onChange();">
			<OPTION></OPTION>
			<%Do while Not objRsFacStat.EOF
			 Response.write "<OPTION"
			 if StrCircuitID <> 0 then
			    if objRsFacStat("CIRCUIT_STATUS_CODE") = objRs("CIRCUIT_STATUS_CODE") then
					   Response.Write " selected "
				end if
			 end if
			  Response.write " VALUE ="& """"&objRsFacStat("CIRCUIT_STATUS_CODE")&"""" & ">" & objRsFacStat("CIRCUIT_STATUS_CODE") & "</OPTION>"
			   objRsFacStat.MoveNext
			 Loop
			%>
		</SELECT>
	</TD>
</TR>
<TR>
	<%if (StrCircuitTyp <> "ATMPVC") THEN %>
		<TD ALIGN=RIGHT NOWRAP width=25%>Facility Provider<font color=red>*</font></TD>
	<%ELSE%>
		<TD ALIGN=RIGHT NOWRAP width=25%>PVC Provider<font color=red>*</font></TD>
	<%END IF%>

	<TD width=25%>
		<SELECT id=selcktprov name=selcktprov style="HEIGHT: 20px; WIDTH: 240px" onchange ="fct_onChange();">
		<%Do while Not objRsCktProv.EOF
		 Response.write "<OPTION "
		 if StrCircuitID <> 0 then
		    if objRsCktProv("CIRCUIT_PROVIDER_CODE") = objRs("CIRCUIT_PROVIDER_CODE") then
				   Response.Write " selected "
			end if
		 end if
		  Response.Write " VALUE ="& """"&routineHtmlString(objRsCktProv("CIRCUIT_PROVIDER_CODE"))&"""" & ">" & routineHtmlString(objRsCktProv("CIRCUIT_PROVIDER_NAME"))& routineHtmlString(objRsCktProv("IS_ON_NET")) & "</OPTION>"
		   objRsCktProv.MoveNext
		 Loop
		%>
		</SELECT>
	</TD>
	<%
	'IF objRs("CIRCUIT_TYPE_CODE") = "ATMPVC" THEN
	if (StrCircuitTyp = "ATMPVC") THEN
	%>
	<TD ALIGN=RIGHT NOWRAP width=25%>Managed By Ems<font color=red>*</font></TD><TD width=25%>
		<SELECT id=selmgtbyems name=selmgtbyems style="HEIGHT: 20px; WIDTH: 120px" onchange ="fct_onChange();">
			<OPTION></OPTION>
			<%Do while Not objRsManEms.EOF
			 Response.write "<OPTION "
			 if StrCircuitID <> 0 then
			   if objRsManEms("ELEMENT_MANAGEMENT_SYSTEM_CODE") = objRs("MANAGED_BY_EMS_CODE") then
					   Response.Write " selected "
				end if
			 end if
			 Response.write "VALUE ="&""""& routineHtmlString(objRsManEms("ELEMENT_MANAGEMENT_SYSTEM_CODE")) &""""& ">" & routineHtmlString(objRsManEms("ELEMENT_MANAGEMENT_SYSTEM_CODE")) & "</OPTION>"
			   objRsManEms.MoveNext
			 Loop
			%>
		</SELECT>
	</TD>
	<%
	END IF
	%>
</TR>

<%if (StrCircuitTyp = "ATMPVC") THEN %>
<TR>
	<!--<TD ALIGN=RIGHT NOWRAP width=25%>Billing Type:</TD>
	<TD width=25%>
		<SELECT id=selbilltype name=selbilltype style="HEIGHT: 20px; WIDTH: 120px" onchange ="fct_onChange();">
			<OPTION></OPTION>
			<%Do while Not objRsBilTyp.EOF
				Response.write "<OPTION "
			if StrCircuitID <> 0 then
				if objRsBilTyp("BILLING_TYPE_CODE") = objRs("BILLING_TYPE_CODE") then
					   Response.Write " selected "
				end if
			end if
				Response.write "VALUE ="& objRsBilTyp("BILLING_TYPE_CODE") & ">" & objRsBilTyp("BILLING_TYPE_CODE") & "</OPTION>" & vbCrLf
				objRsBilTyp.MoveNext
			Loop
			%>
		</SELECT></TD> -->
	<TD ALIGN=RIGHT NOWRAP width=25%>Usage Calc Type<font color=red>*</font></TD>
	<TD width=25%>
		<SELECT id=selusgcalc name=selusgcalc style="HEIGHT: 20px; WIDTH: 120px" onchange ="fct_onChange();">
			<OPTION></OPTION>
			<%Do while Not objUsgCalc.EOF
			Response.write "<OPTION "
			if StrCircuitID <> 0 then
			  if objUsgCalc("USAGE_CALCULATION_TYPE_CODE") = objRs("USAGE_CALCULATION_TYPE_CODE") then
			   	   Response.Write " selected "
			   end if
			end if
			 Response.write "VALUE ="& """"&objUsgCalc("USAGE_CALCULATION_TYPE_CODE")&"""" & ">" & routineHtmlString(objUsgCalc("USAGE_CALCULATION_TYPE_CODE")) & "</OPTION>"
			 objUsgCalc.MoveNext
			Loop
			%>
		</SELECT>
	</TD>
</TR>
<%END IF %>



<TR>
	<TD ALIGN=RIGHT NOWRAP VALIGN=TOP width=25%>Comments</TD>
	<TD ALIGN=LEFT NOWRAP COLSPAN=3 VALIGN="TOP" width=25%><TEXTAREA id=txtacomments name=txtacomments ROWS=3 style="WIDTH: 100%" onchange ="fct_onChange();"><%if StrCircuitID <> 0 then  Response.Write objRS("COMMENTS") else Response.Write null end if%></TEXTAREA></TD>
</TR>

</tbody>

</table>

<table width="100%" border=0>
<thead>
	<tr><td width="100%" colSpan=4>Side A</td></tr>
</thead>
<tbody>
	<TR>
	<TD align=RIGHT nowrap width=25%>Customer Service A<font color=red>*</font></TD>
	<TD colspan=3 width=75%>
	<INPUT id=txtcusserva name=txtcusserva DISABLED  style="HEIGHT: 21px; WIDTH: 500px" value= <%if StrCircuitID <> 0 then  Response.Write """"&routineHtmlString(objRS("CUSTOMER_SERVICE_DESC"))&"""" else Response.Write """"& routineHtmlString(Request.Cookies("CustomerServA"))&""""  end if%> onchange ="fct_onChange();">
	<INPUT  name=btnCustomerServiceLookupA type=button onClick="fct_lookupCustomerService('A');fct_onChange();" value=...>
	<!--<INPUT  name=btnCustomerServiceLookupClear type=button onClick="document.fmfacDetail.txtcusserva.value ='';document.fmfacDetail.hdnCustomerServIDA.value ='';" value="X"> -->
	</TD>
	</TR>
	<TR>
	<TD ALIGN=RIGHT NOWRAP width=25%>Customer A<font color=red>*</font></TD>
	<TD colspan=3 width=75%>
	<INPUT id=txtcustomera name=txtcustomera DISABLED  style="HEIGHT: 21px; WIDTH: 500px" value= <%if StrCircuitID <> 0 then  Response.Write """"&routineHtmlString(objRS("CUSTOMER_NAME"))&"""" else Response.Write """"& routineHtmlString(Request.Cookies("CustName"))&"""" end if%> onchange ="fct_onChange();">
	<INPUT  name=btncustomerlookupA type=button value=... LANGUAGE=javascript onclick="fct_lookupCustomer('A');fct_onChange();" >
	<!--<INPUT  name=btncustomerlookupClear type=button value="X" LANGUAGE=javascript onclick="document.fmfacDetail.txtcustomera.value ='';document.fmfacDetail.hdnCustomerIdA.value ='';"> -->
	</TD>
	</TR>
	<TR>
	<TD ALIGN=RIGHT NOWRAP width=25%>Service Location A<font color=red>*</font></TD>
	<TD colspan=3 width=75%><INPUT id=txtsrvloca name=txtsrvloca  DISABLED style="HEIGHT: 21px; WIDTH: 500px" value= <%if StrCircuitID <> 0 then  Response.Write """"&routineHtmlString(objRS("LOCATION_A"))&"""" else Response.Write """"& routineHtmlString(Request.Cookies("ServiceLocName"))&"""" end if%> onchange ="fct_onChange();">
	<INPUT id=btnServiceLocationLookupA name=btnServiceLocationLookupA type=button  value=... LANGUAGE=javascript onclick="btnServiceLocationLookup_onclick('A');fct_onChange();">
	<!--<INPUT  name=btnServiceLocationLookupClear type=button value="X" LANGUAGE=javascript onclick="document.fmfacDetail.txtsrvloca.value ='';document.fmfacDetail.hdnServiceLocIdA.value ='';document.fmfacDetail.txtaaddressa.value ='';"> -->
	</TD>
	</TR>
	<TR>
	<TD ALIGN=RIGHT VALIGN=TOP NOWRAP ></TD>
	<TD ALIGN=LEFT NOWRAP COLSPAN=3><TEXTAREA id=txtaaddressa name=txtaaddressa style="HEIGHT: 90px; WIDTH: 400px" DISABLED COLS=1 ROWS=5 ><%if StrCircuitID <> 0 then  Response.Write routineHtmlString(objRS("ADDRESS_A")) else Response.Write  routineHtmlString(Request.Cookies("Address")) end if%> </TEXTAREA></TD></TR>
	</TR>
	<%
	  'IF objRs("CIRCUIT_TYPE_CODE") = "FIBRE" THEN
	  if (StrCircuitTyp = "FIBRE") THEN
	%>
	<TR>
	 <TD ALIGN=RIGHT NOWRAP width=25% >Floor</TD>
	 <TD width=25% COLSPAN=3><INPUT name=txtfloora  style="HEIGHT: 22px; WIDTH: 115px" value= <%if StrCircuitID <> 0 then  Response.Write """"&routineHtmlString(objRS("FIBRE_FLOOR_A"))&"""" else Response.Write """""" end if%> onchange ="fct_onChange();">
	 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Bay</font>
	 <INPUT name=txtbaya  style="HEIGHT: 22px; WIDTH: 115px" value= <%if StrCircuitID <> 0 then  Response.Write """"&routineHtmlString(objRS("FIBRE_BAY_A"))&"""" else Response.Write """""" end if%> onchange ="fct_onChange();"></TD>
	 </TR>
	 <TR>
	 <TD ALIGN=RIGHT NOWRAP width=25% >Shelf/Module</TD>
	 <TD width=25% colspan=3><INPUT name=txtshellmodulea  style="HEIGHT: 22px; WIDTH: 115px" value= <%if StrCircuitID <> 0 then  Response.Write """"&routineHtmlString(objRS("FIBRE_SHELF_MODULE_A"))&"""" else Response.Write """""" end if%> onchange ="fct_onChange();">
	 Position:
	 <INPUT name=txtpositiona style="HEIGHT: 22px; WIDTH: 115px" value= <%if StrCircuitID <> 0 then  Response.Write """"&routineHtmlString(objRS("FIBRE_POSITION_A"))&"""" else Response.Write """""" end if%> onchange ="fct_onChange();"></TD>
	 </TR>
	 <TR>
	 <TD ALIGN=RIGHT NOWRAP width=25%>Cable #</TD>
	 <TD width=25% colspan=3><INPUT name=txtcablenumbera  style="HEIGHT: 22px; WIDTH: 115px" value= <%if StrCircuitID <> 0 then  Response.Write """"&routineHtmlString(objRS("FIBRE_CABLE_A"))&"""" else Response.Write """""" end if%> onchange ="fct_onChange();">
	 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Pairs:<INPUT name=txtpairsa style="HEIGHT: 22px; WIDTH: 115px" value= <%if StrCircuitID <> 0 then  Response.Write """"&routineHtmlString(objRS("FIBRE_PAIRS_A"))&"""" else Response.Write """""" end if%> onchange ="fct_onChange();"></TD>
	 </TR>
	 <%
	   end if
	  %>
</tbody>

</table>
<table width="100%" border=0>
<thead>
	<tr><td align=left colspan=4 width="100%">Side B</td></tr>
</thead>
<tbody>
	<%if (StrCircuitTyp = "ATMPVC") THEN %>
	<TR>
	<TD align=RIGHT nowrap width=25%>Customer Service B<font color=red>*</font></TD>
	<TD colspan=3 width=25%><INPUT id=txtcusservb name=txtcusservb DISABLED  style="HEIGHT: 21px; WIDTH: 500px" value= <%if StrCircuitID <> 0 then  Response.Write """"&routineHtmlString(objRS("SERVICE_B"))&"""" else Response.Write """""" end if%> onchange ="fct_onChange();">
	<INPUT name=btnCustomerServiceBLookupB type=button onClick="fct_lookupCustomerService('B');fct_onChange();" value=...>
	<!--<INPUT name=btnCustomerServiceBLookupClear type=button onClick="document.fmfacDetail.txtcusservb.value ='';document.fmfacDetail.hdnCustomerServIDB.value ='';document.fmfacDetail.hdnCustomerServIDB.value ='';" value="X">	-->
	</TD>

	</TR>
	<TR>
	<TD ALIGN=RIGHT NOWRAP width=25%>Customer B<font color=red>*</font></TD>
	<TD colspan=3 width=75%><INPUT id=txtcustomerb name=txtcustomerb DISABLED style="HEIGHT: 21px; WIDTH: 500px" value= <%if StrCircuitID <> 0 then  Response.Write """"&routineHtmlString(objRS("CUSTOMER_B"))&"""" else Response.Write """""" end if%> onchange ="fct_onChange();">
	<INPUT  name=btncustomerBlookupB type=button value=... LANGUAGE=javascript onclick="fct_lookupCustomer('B');fct_onChange();">
	<!--<INPUT  name=btncustomerBlookupClear type=button value="X" LANGUAGE=javascript onclick="document.fmfacDetail.txtcustomerb.value ='';document.fmfacDetail.hdnCustomerIdB.value ='';"> -->
	</TD>

	</TR>
	<%END IF%>
	<TR>
		<TD ALIGN=RIGHT NOWRAP width=25%>Service Location B<font color=red>*</font></TD>
		<TD colspan=3 width=75%><INPUT id=txtsrvlocb name=txtsrvlocb DISABLED  style="HEIGHT: 21px; WIDTH: 500px" value= <%if StrCircuitID <> 0 then  Response.Write """"&routineHtmlString(objRS("LOCATION_B"))&"""" else Response.Write """""" end if%> onchange ="fct_onChange();">
		<INPUT id=btnServiceLocationBLookupB name=btnServiceLocationLookupB type=button value=... LANGUAGE=javascript onclick="btnServiceLocationLookup_onclick('B');fct_onChange();">
		<!--<INPUT name=btnServiceLocationBLookupClear type=button value="X" LANGUAGE=javascript onclick="document.fmfacDetail.txtsrvlocb.value ='';document.fmfacDetail.hdnServiceLocIdB.value ='';document.fmfacDetail.txtaaddressb.value ='';">-->
		</TD>

	</TR>
	<TR>
	<TD ALIGN=RIGHT VALIGN=TOP NOWRAP ></TD>
	<TD ALIGN=LEFT NOWRAP COLSPAN=3><TEXTAREA id=txtaaddressb name=txtaaddressb style="HEIGHT: 90px; WIDTH: 400px" DISABLED COLS=1 ROWS=5 ><%if StrCircuitID <> 0 then  Response.Write objRS("ADDRESS_B") else Response.Write null end if%> </TEXTAREA></TD></TR>
	</TR>
	<%
	  'IF objRs("CIRCUIT_TYPE_CODE") = "FIBRE" THEN
	  if (StrCircuitTyp = "FIBRE") THEN
	  %>
	<TR>
	 <TD ALIGN=RIGHT NOWRAP width=25% >Floor</TD>
	 <TD width=25% colspan=3><INPUT name=txtfloorb  style="HEIGHT: 22px; WIDTH: 115px" value= <%if StrCircuitID <> 0 then  Response.Write """"&routineHtmlString(objRS("FIBRE_FLOOR_B"))&"""" else Response.Write """""" end if%> onchange ="fct_onChange();">
	 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Bay:<INPUT name=txtbayb  style="HEIGHT: 22px; WIDTH: 115px" value= <%if StrCircuitID <> 0 then  Response.Write """"&routineHtmlString(objRS("FIBRE_BAY_B"))&"""" else Response.Write """""" end if%> onchange ="fct_onChange();"></TD>
	 </TR>
	 <TR>
	 <TD ALIGN=RIGHT NOWRAP width=25%>Shelf/Module</TD>
	 <TD width=25% colspan=3><INPUT name=txtshellmoduleb  style="HEIGHT: 22px; WIDTH: 115px" value= <%if StrCircuitID <> 0 then  Response.Write """"&routineHtmlString(objRS("FIBRE_SHELF_MODULE_B"))&"""" else Response.Write """""" end if%> onchange ="fct_onChange();">
	 Position:<INPUT name=txtpositionb style="HEIGHT: 22px; WIDTH: 115px" value= <%if StrCircuitID <> 0 then  Response.Write """"&routineHtmlString(objRS("FIBRE_POSITION_B"))&"""" else Response.Write """""" end if%> onchange ="fct_onChange();"></TD>
	 </TR>
	 <TR>
	 <TD ALIGN=RIGHT NOWRAP width=15%>Cable #</TD>
	 <TD width=25% colspan=3><INPUT name=txtcablenumberb  style="HEIGHT: 22px; WIDTH: 115px" value= <%if StrCircuitID <> 0 then  Response.Write """"&routineHtmlString(objRS("FIBRE_CABLE_B"))&"""" else Response.Write """""" end if%> onchange ="fct_onChange();">
	 &nbsp;&nbsp;&nbsp;&nbsp;Pairs:<INPUT name=txtpairsb style="HEIGHT: 22px; WIDTH: 115px" value= <%if StrCircuitID <> 0 then  Response.Write """"&routineHtmlString(objRS("FIBRE_PAIRS_B"))&"""" else Response.Write """""" end if%> onchange ="fct_onChange();"></TD>
	 </TR>
	 <%
	   end if
	  %>
	<tr>
	<td>
	<INPUT id=hdnCustomerServIDA name=hdnCustomerServIDA TYPE=HIDDEN style="HEIGHT: 21px; WIDTH: 500px" value=<%if StrCircuitID <> 0 then  Response.Write objRS("CUSTOMER_SERVICE_ID_A") else Response.Write Request.Cookies("CustomerServID") end if%> >
	<INPUT id=hdnCustomerServIDB name=hdnCustomerServIDB TYPE=HIDDEN style="HEIGHT: 21px; WIDTH: 500px" value=<%if StrCircuitID <> 0 then  Response.Write objRS("CUSTOMER_SERVICE_ID_B") else Response.Write """""" end if%> >
	<INPUT id=hdnServiceLocA name=hdnServiceLocIdA TYPE=HIDDEN style="HEIGHT: 21px; WIDTH: 500px" value=<%if StrCircuitID <> 0 then  Response.Write objRS("SERVICE_LOCATION_ID_A") else Response.Write Request.Cookies("ServiceLocID") end if%>  >
	<INPUT id=hdnServiceLocB name=hdnServiceLocIdB TYPE=HIDDEN style="HEIGHT: 21px; WIDTH: 500px" value=<%if StrCircuitID <> 0 then  Response.Write objRS("SERVICE_LOCATION_ID_B") else Response.Write """""" end if%> >
	<INPUT id=hdnCustomerIdA name=hdnCustomerIdA TYPE=HIDDEN style="HEIGHT: 21px; WIDTH: 500px" value=<%if StrCircuitID <> 0 then  Response.Write objRS("BILLING_CUSTOMER_ID_A") else Response.Write Request.Cookies("CustID") end if%> >
	<INPUT id=hdnCustomerIdB name=hdnCustomerIdB TYPE=HIDDEN style="HEIGHT: 21px; WIDTH: 500px" value=<%if StrCircuitID <> 0 then  Response.Write objRS("BILLING_CUSTOMER_ID_B") else Response.Write """""" end if%>  >
	<INPUT id=CircuitID name=CircuitID TYPE=HIDDEN style="HEIGHT: 21px; WIDTH: 500px" value=<%if StrCircuitID <> 0 then  Response.Write StrCircuitID else Response.Write """""" end if%>  >
	<INPUT id=CircuitTyp name=CircuitTyp TYPE=HIDDEN style="HEIGHT: 21px; WIDTH: 500px" value=<%if StrCircuitID <> 0 then  Response.Write objRS("CIRCUIT_TYPE_CODE") else Response.Write StrCircuitTyp end if%> >
	<INPUT type="hidden" name=hdnUpdateDateTime value=<%if StrCircuitID <> 0 then  Response.Write """"&objRS("UPDATE_DATE_TIME")&"""" else Response.Write """""" end if%>>
    <INPUT type="hidden" name=txtFrmAction value="">
    <INPUT type="hidden" name=hdnNameAlias value="">
    <INPUT type="hidden" name=hdnCircuitStartDt value="">
    <INPUT type="hidden" name=hdnAdslDueDt value="">
    <INPUT type="hidden" name=hdnfacilityType value="">
    <INPUT type="hidden" name=hdnoperationalstatus value="">

    </TD>
	</tr>

</tbody>
</table>

<table width="100%" border=0>
<thead>
	<tr><td colSpan=4 align=left>Alias</td></tr>
</thead>
<tbody>
	<TR>
		<td width="25%" rowSpan="5" colspan=2 valign="top"  align=left>
			<iframe id=aifr width=100% height=100 src="" scrolling=yes marginheight=1 marginwidth=1></iframe>
			<br>
			<input type="button" style= "width: 2cm" value="Delete"  <%if StrCircuitID =0 or StrCircuitID="" then  Response.Write "DISABLED" end if%> name="btn_iFrameDelete" onClick="btn_iFrmDelete();fct_onChange();">&nbsp;&nbsp;
			<input type="button" style= "width: 2cm" value="Refresh" <%if StrCircuitID =0 or StrCircuitID="" then Response.Write "DISABLED" end if%>   name="btn_iFrameRefresh"    onClick="iFrame_display();">&nbsp;&nbsp;
			<input type="button" style= "width: 2cm" value="New"  <%if StrCircuitID =0 or StrCircuitID="" then Response.Write "DISABLED" end if%>   name="btn_iFrameAdd"    onClick="btn_iFrmAdd();fct_onChange();">&nbsp;&nbsp;
			<input type="button" style= "width: 2cm" value="Update"  <%if StrCircuitID =0 or StrCircuitID="" then Response.Write "DISABLED" end if%>   name="btn_iFrameupdate"    onClick="btn_iFrmUpdate();fct_onChange();">
	</td>
	<td width=15%></td>
	</TR>
</tbody>
</table>

<TABLE>
	<tfoot>
	  <TR><TD align=right colspan=5>
	       <input name=btnReferences type=button style= "width: 2.2cm" value=References LANGUAGE=javascript onclick="return btnReferences_onclick()">&nbsp;&nbsp;
			<INPUT name=btnDelete type=button style= "width: 2cm" value=Delete LANGUAGE=javascript onclick="return fct_onDelete();">&nbsp;&nbsp;
			<INPUT name=btnReset type=reset style= "width: 2cm" value=Reset onClick="fct_onReset();">&nbsp;&nbsp;
			<INPUT name=btnAddNew type=button style= "width: 2cm"value="New" LANGUAGE=javascript onclick="return btnNew_click()">&nbsp;&nbsp;
			<INPUT name=btnClone type=button style="width: 2cm" value=Clone onclick="fct_onClone();">&nbsp;&nbsp;
			<INPUT name=btnSave type=button style= "width: 2cm" value=Save onclick="btnSave_onclick();">&nbsp;&nbsp;
	  </TD></TR>
	 </tfoot>
</table>

<FIELDSET >
   <%if bolClone then StrCircuitID = 0%>
	<LEGEND ALIGN=RIGHT><B>Audit Information</B></LEGEND>
	<Div SIZE=8pt ALIGN=RIGHT>
		Record Status Indicator:
		<INPUT align = left name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 18px"disabled value=<%if StrCircuitID <> 0 then  Response.Write """"&objRS("RECORD_STATUS_IND")&"""" else Response.Write """""" end if%>  >&nbsp;&nbsp;&nbsp;
		Create Date:&nbsp;&nbsp;
		<INPUT align = center name=txtcrdate type=text style="HEIGHT: 20px; WIDTH: 140px"disabled value=<%if StrCircuitID <> 0 then  Response.Write """"&objRS("CREATE_DATE_TIME")&"""" else Response.Write """""" end if%> >&nbsp;
		&nbsp;
		Created By:
		<INPUT align = right name=txtcrby type=text style="HEIGHT: 20px; WIDTH: 100px"disabled value=<%if StrCircuitID <> 0 then  Response.Write """"&routineHtmlString(objRS("CREATE_REAL_USERID"))&"""" else Response.Write """""" end if%> ><BR>
		Update Date:
		<INPUT align= center name=txtupdate type=text style="HEIGHT: 20px; WIDTH: 140px"disabled value=<%if StrCircuitID <> 0 then  Response.Write """"&objRS("UPDATE_DATE_TIME_CONV")&"""" else Response.Write """""" end if%>  >&nbsp;&nbsp;
		Updated By:
		<INPUT align=right name=txtupby type=text style="HEIGHT: 20px; WIDTH: 100px"disabled value=<%if StrCircuitID <> 0 then  Response.Write """"&routineHtmlString(objRS("UPDATE_REAL_USERID"))&"""" else Response.Write """""" end if%>  >
	</DIV>
</FIELDSET>
</FORM>
<%

   'Move to the next row
   'if StrCircuitID <> 0 then
   'if not objRS.EOF THEN
    'objRS.MoveNext
   'end if
 'Loop

 'Clean up our ADO objects
    objRS.close
    set objRS = Nothing

    objRsFacTyp.close
    set objRsFacTyp = Nothing

    objRsFacStat.close
    set objRsFacStat = Nothing

    objRsAdslTyp.close
    set objRsAdslTyp = Nothing

    objRsRegion.close
    set objRsRegion = Nothing

    objRsManEms.close
    set objRsManEms = Nothing

    objRsCktProv.close
    set objRsCktProv = Nothing

    IF StrCircuitTyp = "ATMPVC" THEN
    objRsBilTyp.close
    set objRsBilTyp = Nothing

    objUsgCalc.close
    set objUsgCalc = Nothing
    END IF

    objConn.close
    set ObjConn = Nothing

%>

</BODY>
</HTML>





















