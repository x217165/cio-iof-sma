<%@  language="VBScript" %>
<% Option Explicit %>
<% on error resume next %>
<!--#include file="smaConstants.inc"-->
<!--#include file="databaseconnect.asp"-->
<!--#include file="smaProcs.inc"-->
<!--
*************************************************************************************
* File Name:	ManObjPortDetail.asp
*
* Purpose:		To display Managed Object Port Name and LAN IP and allow user to make changes.
*
* Created by:	Dan S. Ty	03/13/2002f
*
* Edited by:
**************************************************************************************
		 Date		Author		Changes/enhancements made
		 2003/10/15	DTy			Add field required for IP Mediation:
								Customer Service ID & Name & Billable Port.
								Remove LAN IP as a mandatory field.
								Validate LAN IP if entered.
		 2006/11/27	Anthony Cheung		Add field required for Government of Ontario:
								Reportable ?
								Port Function
		 2007/07/27     Anthony Cheung          Add free form test field required for Government of Ontario:
								Managed Objects Unqiue ID (MSUID) or Port Identification
		 2010/07/22	Anthony Cheung		Increase Port Name size to 50 from 25.
**************************************************************************************
-->

<%
    
        Function in_array(element, arr)
    dim i
    For i=0 To Ubound(arr) 
        If Trim(arr(i)) = Trim(element) Then 
            in_array = True
            Exit Function
        Else 
            in_array = False
        End If  
    Next 
End Function
         
'check user's rights
dim intAccessLevel
    
intAccessLevel = CInt(CheckLogon(strConst_ManagedObjects))
if intAccessLevel and intConst_Access_ReadOnly <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to managed object. Please contact your system administrator"
end if

     dim MgmtSystems,MGMT_Systems
    MgmtSystems = Request("selMGMT_SYSTEMS_NAME")
    MGMT_Systems = Request("selMGMT_SYSTEMS_NAME")

    if MGMT_Systems = Empty then
    MGMT_Systems = "null"
    end if
    
dim sql, strWinMessage, rsPort, bolclone ,strCustId , SQLscript , rsMgmtSystem ,sqlMgmtSystemName
    sqlMgmtSystemName ="select mgmt_system_id,mgmt_system_name from CRP.LCODE_MGMT_SYSTEMS"
    set rsMgmtSystem=server.CreateObject("ADODB.Recordset")
     rsMgmtSystem.CursorLocation = adUseClient
     rsMgmtSystem.Open sqlMgmtSystemName, objConn
dim strAction
strAction = Request("action")			'get the action code from caller
if strAction = "" then
	Response.write "No action requested"
	Response.End						'no action requested
end if
    
    
    strCustId = Request("CustId")
If strAction = "clone" then
   bolClone = true
else
   bolClone = false
end if
    dim cmdinsertAlias
dim strMasterID , strNetworkElementPortSequence
strMasterID = Request("masterID")		'get master id
dim strPortID
strPortID = Request("PortID")			'get port id
dim strRealUserID
strRealUserID = Session("username")
if err then
	'unexpected error
	DisplayError "BACK", "", 0, "UNEXPECTED ERROR", "Close port window to return to managed objects form."
end if
strLastUpdate = Request("hdnLastUpdate")

' Setup Port Type drop-down list
dim rsPortType , rsSequence
dim strSQL, strSIteNameSQL,strOrganisationSQL,strCTRINNameSQL,strCTROutNameSQL,strVTRInNameSQL,strVTROutNameSQL,strETRInNameSQL,strETROutNameSQL , strCIStatusSQL

    strNetworkElementPortSequence = "select Update_date_time,NETWORK_ELEMENT_PORT_ID from  CRP.NETWORK_ELEMENT_PORT where NETWORK_ELEMENT_ID =" & strMasterID & " order by NETWORK_ELEMENT_PORT_ID"

strSQL = "select network_element_port_type_name" & _
		 " from crp.network_element_port_type" & _
		 " where record_status_ind = 'A'" & _
		 " order by network_element_port_type_name"

set rsPortType = Server.CreateObject("ADODB.Recordset")
rsPortType.CursorLocation = adUseClient
rsPortType.Open strSQL, objConn 
     
    set rsSequence  =  Server.CreateObject("ADODB.Recordset")
    rsSequence.CursorLocation = adUseClient
    rsSequence.Open strNetworkElementPortSequence, objConn
    
'set rsPortType.ActiveConnection = nothing

' Setup Port Function drop-down list
dim rsPortFunction , rsSiteNameFunction , rsOrganisationFunction ,rsCTRINFunction,rsCTROutFunction,rsVTRInFunction,rsVTROutFunction,rsETRInFunction,rsETROutFunction ,rsCIStatusFunction

strSQL = "SELECT ne_port_function_name, ne_port_function_lcode" & _
		 " FROM crp.lcode_ne_port_function" & _
		 " WHERE record_status_ind = 'A'" & _
		 " ORDER BY ne_port_function_lcode"


    strCIStatusSQL = "select CI_STATUS_id,CI_STATUS_name from CRP.LCODE_CI_STATUS "

    strSIteNameSQL = "select site_id,site_name from CRP.SITE_NAME_CODE"

    strOrganisationSQL="select ORGANIZATION_ID,ORGANIZATION_NAME   from CRP.CUSTOMER_ORGANIZATION where customer_id = "& strCustId

    strCTRINNameSQL = "select * from crp.lcode_ctr_in order by TO_NUMBER(ctr_in_value)"

    strCTROutNameSQL = "select ctr_out_id,ctr_out_name from CRP.LCODE_CTR_out order by TO_NUMBER(ctr_out_value)  "

    strVTROutNameSQL = "select vtr_out_id,vtr_out_name from CRP.LCODE_VTR_out order by TO_NUMBER(vtr_out_value)"

    strVTRInNameSQL = "select vtr_in_id,vtr_in_name from CRP.LCODE_VTR_IN order by TO_NUMBER(vtr_in_value)"

     strETROutNameSQL = "select Etr_out_id,Etr_out_name from CRP.LCODE_ETR_out order by TO_NUMBER(etr_out_value)"

    strETRInNameSQL = "select Etr_in_id,Etr_in_name from CRP.LCODE_ETR_IN order by TO_NUMBER(etr_in_value)"

    set rsCIStatusFunction = Server.CreateObject("ADODB.Recordset")

    rsCIStatusFunction.CursorLocation = adUseClient
      rsCIStatusFunction.Open strCIStatusSQL, objConn


    set rsETRInFunction = Server.CreateObject("ADODB.Recordset")

    rsETRInFunction.CursorLocation = adUseClient
      rsETRInFunction.Open strETRInNameSQL, objConn
    
    set rsETROutFunction = Server.CreateObject("ADODB.Recordset")

    rsETROutFunction.CursorLocation = adUseClient
      rsETROutFunction.Open strETROutNameSQL, objConn

    
    set rsSiteNameFunction = Server.CreateObject("ADODB.Recordset")

    rsSiteNameFunction.CursorLocation = adUseClient
      rsSiteNameFunction.Open strSIteNameSQL, objConn
    
    set rsCTRINFunction = Server.CreateObject("ADODB.Recordset")

    rsCTRINFunction.CursorLocation = adUseClient
      rsCTRINFunction.Open strCTRINNameSQL, objConn

    set rsCTROutFunction = Server.CreateObject("ADODB.Recordset")

    rsCTROutFunction.CursorLocation = adUseClient
      rsCTROutFunction.Open strCTROutNameSQL, objConn

    '--start--
    set rsVTROutFunction = Server.CreateObject("ADODB.Recordset")

    rsVTROutFunction.CursorLocation = adUseClient
      rsVTROutFunction.Open strVTROutNameSQL, objConn

    set rsVTRInFunction = Server.CreateObject("ADODB.Recordset")

    rsVTRInFunction.CursorLocation = adUseClient
      rsVTRInFunction.Open strVTRInNameSQL, objConn

    '--complete--
       set rsOrganisationFunction = Server.CreateObject("ADODB.Recordset")

    rsOrganisationFunction.CursorLocation = adUseClient
rsOrganisationFunction.Open strOrganisationSQL, objConn

set rsPortFunction = Server.CreateObject("ADODB.Recordset")
rsPortFunction.CursorLocation = adUseClient
rsPortFunction.Open strSQL, objConn

'set rsPortFunction.ActiveConnection = nothing

'save changes?
     
if strAction = "save" then 
	dim strPortName, strPortIP, strCSID, strBillable, strLastUpdate, strreportable, strMSUID 
	strMasterID     = Request("masterID")
	strPortName     = Request("txtPortName")
    if strPortName ="" then
    strPortName = " "
    end if
	strPortIP       = Request("txtPortIP")
	strCSID         = Request("lngCSID")
	strBillable     = Request("chkBillable")
	strreportable   = Request("chkReportable")
	strMSUID        = Request("MSUID")
    

	if lcase(strBillable) = "on" then
		strBillable = "Y"
	else
		strBillable = "N"
	end if

	if lcase(strreportable) = "on" then
		strreportable = "Y"
	else
		strreportable = "N"
	end if

	'call stored proc to save the record
	if (strPortID <> "") and (intAccessLevel and intConst_Access_Update = intConst_Access_Update) then
		'create command object for update stored proc
		dim cmdUpdateObj
		set cmdUpdateObj = server.CreateObject("ADODB.Command")
		set cmdUpdateObj.ActiveConnection = objConn
		cmdUpdateObj.CommandType = adCmdText
		'cmdUpdateObj.CommandText = "sma_sp_userid.spk_sma_mo_inter.sp_mo_port_update"
    dim updateETRIN
    if Request("selETR_IN_ID") <> "" then
    updateETRIN =  cint( Request("selETR_IN_ID"))
    else updateETRIN = "null"
    end if
    dim updateETROUT
    if Request("selETR_OUT_ID") <> "" then
    updateETROUT =  cint( Request("selETR_OUT_ID"))
    else updateETROUT = "null"
    end if
        dim updateCTRIN
    if Request("selCTR_IN_ID") <> "" then
     updateCTRIN =  cint( Request("selCTR_IN_ID"))
    else updateCTRIN = "null"
    end if
    dim updateCTROUT
    if Request("selCTR_OUT_ID") <> "" then
    updateCTROUT =  cint( Request("selCTR_OUT_ID"))
    else updateCTROUT = "null"
    end if
    dim updateVTRIN
    if Request("selVTR_IN_ID") <> "" then
     updateVTRIN =  cint( Request("selVTR_IN_ID"))
    else updateVTRIN = "null"
    end if
    dim updateVTROUT
    if Request("selVTR_OUT_ID") <> "" then
     updateVTROUT =  cint( Request("selVTR_OUT_ID"))
    else updateVTROUT = "null"
    end if
    dim updateCISTATUS
    if Request("selCI_Status_ID")  <> "" then
     updateCISTATUS =cint(Request("selCI_Status_ID"))
    else updateCISTATUS ="null"
    end if
     dim updateORGANIZATION
    if Request("selORGANIZATION_NAME")  <> "" then
    updateORGANIZATION =cint(Request("selORGANIZATION_NAME"))
    else updateORGANIZATION ="null"
    end if
     dim updateSITENAME
    if Request("selSITE_NAME") <> "" then
    updateSITENAME =cint(Request("selSITE_NAME"))
    else updateSITENAME ="null"
    end if



dim custSerId 

    if strCSID <> "" then
      custSerId = Clng(strCSID)
    else custSerId = "null"
    end if
    cmdUpdateObj.CommandText = "UPDATE crp.network_element_port " &_
            "  SET network_element_id        =" & CLng(strMasterID) & "," &_
                " network_element_port_name =  '" & strPortName & "',"&_
                "  network_element_port_ip   = '" & strPortIP & "', " &_
               "  customer_service_id       = " &  custSerId & "," &_
               "  billable_port             = '" &  strBillable & "', " &_
              "   reportable                ='"& strreportable & "', " &_
              "   ne_port_function_lcode    =  " & CInt(Request("selPortFunction"))  & "," &_
              "   msuid                     = '" & strMSUID &  "' ," &_
              "   CTR_IN_ID                = " & updateCTRIN  & ", " &_ 
              "   CTR_OUT_ID    =  " & updateCTROUT &" ," &_
              "   VN_NAME       = '" &  Request("txtVN_NAME") & "'," &_
              "   VTR_IN_ID     =  " & updateVTRIN &  "," &_
              "   VTR_OUT_ID     = "& updateVTROUT & "," &_
              "   QOS_NAME          = '" & Request("txtQOS_NAME") &  "', " &_               
              "   ETR_IN_ID      =  " & updateETRIN & "," &_
             "   ETR_OUT_ID     =  " & updateETROUT & "," &_
             "   CI_STATUS_ID    =  " & updateCISTATUS & "," &_
             "   ORGANIZATION_ID  =  " & updateORGANIZATION & "," &_
             "   SITE_ID          =  " & updateSITENAME & "," &_
               "   update_real_userid   ='" & strRealUserID & "'," &_ 
              "   PORT_NAME_ALIAS   ='" & Request("txtPortNameAlias") & "'" &_ 
             " WHERE network_element_port_id =" &  CLng(strPortID) & ""   
           

     set cmdinsertAlias = server.CreateObject("ADODB.Command")
		set cmdinsertAlias.ActiveConnection = objConn

         SQLscript = "delete CRP.NETWORK_PORT_NAME_ALIAS where NETWORK_ELEMENT_PORT_ID =" & strPortID 
    cmdinsertAlias.CommandText = SQLscript
    cmdinsertAlias.Execute


    if Request("selAlias") <> "" then
  SQLscript = "insert into CRP.NETWORK_PORT_NAME_ALIAS(NETWORK_ELEMENT_PORT_ID ,NETWORK_PORT_NAME_ALIAS ,CREATE_REAL_USERID,CREATE_DB_USERID, UPDATE_REAL_USERID) values (" & strPortID & ",'" & Replace( Request("selAlias"), "_", " ")   & "','"& strRealUserID & "','"& strRealUserID & "','"& strRealUserID & "')"
    cmdinsertAlias.CommandText = SQLscript
    cmdinsertAlias.Execute
    end if
    
		on error resume next
		cmdUpdateObj.Execute
		if objConn.Errors.Count <> 0 then
			DisplayError "BACK", "x", objConn.Errors(0).NativeError, "CANNOT UPDATE OBJECT", objConn.Errors(0).Description
			objConn.Errors.Clear
		end if

     dim cmdinsertMGMT,MgMtSystemsSQLscript
    set cmdinsertMGMT = server.CreateObject("ADODB.Command")
		set cmdinsertMGMT.ActiveConnection = objConn
    cmdinsertMGMT.CommandType = adCmdText
    
    MgMtSystemsSQLscript = "delete  CRP.NETWORK_ELEMENT_MGMT_SYS where NETWORK_ELEMENT_Port_ID ="& CLng(strPortID)
     cmdinsertMGMT.CommandText = MgMtSystemsSQLscript
     cmdinsertMGMT.Execute

      if IsEmpty(MGMT_Systems) <> true and MGMT_Systems <> "null" then
	    dim arr , i
      arr = split(MGMT_Systems,",",10)
             For i=0 To Ubound(arr) 
    set cmdinsertMGMT = server.CreateObject("ADODB.Command")
		set cmdinsertMGMT.ActiveConnection = objConn
         MgMtSystemsSQLscript = "insert into CRP.NETWORK_ELEMENT_MGMT_SYS(NETWORK_ELEMENT_Port_ID,MGMT_SYSTEM_ID,CREATE_REAL_USERID,CREATE_DB_USERID) values (" & strPortID & "," & arr(i) & ",'"& strRealUserID & "','"& strRealUserID & "')"
    cmdinsertMGMT.CommandText = MgMtSystemsSQLscript
    cmdinsertMGMT.Execute
    next
    end if

		strWinMessage = "Record saved successfully. You can now see the changes you made."
	elseif (strPortID = "") and (intAccessLevel and intConst_Access_Create = intConst_Access_Create) then
		'create command object for insert stored proc
		dim cmdInsertObj
		set cmdInsertObj = server.CreateObject("ADODB.Command")
		set cmdInsertObj.ActiveConnection = objConn
		cmdInsertObj.CommandType = adCmdStoredProc
		cmdInsertObj.CommandText = "sma_sp_userid.spk_sma_mo_inter.sp_mo_port_insert"
		'create params
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_user_id",       adVarChar, adParamInput, 30, strRealUserID )
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_port_id",       adNumeric, adParamOutput,20,0)
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_ne_id",         adNumeric, adParamInput,   , CLng(strMasterID))
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_port_name",     adVarChar, adParamInput, 50, strPortName)
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_port_ip",       adVarChar, adParamInput, 50, strPortIP)
		if strCSID <> "" then
     	   cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_CSID",       adNumeric, adParamInput,   , CLng(strCSID))
		else
     	   cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_CSID",       adNumeric, adParamInput,   , null)
		end if
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_billable_port", adVarChar, adParamInput,  1, strBillable)
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_reportable",   adVarChar,     adParamInput,  1, strreportable)
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_ne_port_function", adVarChar, adParamInput, 20, Request("selPortFunction"))	        'NE port function
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_msuid",         adVarChar,     adParamInput, 50, strMSUID) 'New free form MSUID
         If Request("selCTR_IN_ID") <> "" then
        cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_CTR_IN_ID", adNumeric, adParamInput, 20, cint( Request("selCTR_IN_ID")))
    else
     cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_CTR_IN_ID", adNumeric, adParamInput, 20, null)
    end if
    If Request("selCTR_OUT_ID") <> "" then
      cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_CTR_OUT_ID",   adNumeric,     adParamInput,  20, cint( Request("selCTR_OUT_ID")) )
    else
    cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_CTR_OUT_ID",   adNumeric,     adParamInput,  20, null )
    end if
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_VN_NAME", adVarChar, adParamInput, 256, Request("txtVN_NAME"))
    if Request("selVTR_Out_ID") <> "" then	        
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_VTR_OUT_ID",  adNumeric,     adParamInput, 20, cint( Request("selVTR_Out_ID")))
    else
    cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_VTR_OUT_ID",         adNumeric,     adParamInput, 20, null) 
    end if
     cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_QOS_NAME",   adVarChar,     adParamInput,  256, Request("txtQOS_NAME"))
	If Request("selETR_IN_ID") <> "" then	
    cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_ETR_IN_ID", adNumeric, adParamInput, 20,cint( Request("selETR_IN_ID")))	
    else 
    cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_ETR_IN_ID", adNumeric, adParamInput, 20, null)
    end if
       if  Request("selETR_OUT_ID") <> "" then   
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_ETR_OUT_ID",         adNumeric,     adParamInput, 20,cint( Request("selETR_OUT_ID")))
    else
    cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_ETR_OUT_ID",         adNumeric,     adParamInput, 20, null) 
    end if 
    if Request("selCI_Status_ID") <> "" then
        cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_CI_STATUS_ID", adNumeric, adParamInput,  20, cint(Request("selCI_Status_ID")))
    else 
     cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_CI_STATUS_ID", adNumeric, adParamInput,  20, null)
    end if
    if Request("selORGANIZATION_NAME") <> "" then
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_ORGANIZATION_ID",   adNumeric,     adParamInput,  20,cint( Request("selORGANIZATION_NAME")))
    else 
    cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_ORGANIZATION_ID",   adNumeric,     adParamInput,  20, null)
    end if
      ' cmdcmdInsertObjrameters.Append cmdInsertObj.CreateParameter("p_ORGANIZATION_CODE",   adVarChar,     adParamInput,  1, Request("txtVN_NAME"))
    if Request("selSITE_NAME") <> "" then
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_SITE_ID", adNumeric, adParamInput, 20, cint( Request("selSITE_NAME"))) 
    else
    cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_SITE_ID", adNumeric, adParamInput, 20, null)
    end if
    if Request("selVTR_In_ID") <> "" then
           cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_VTR_IN_ID", adNumeric, adParamInput,  20,cint( Request("selVTR_In_ID")))
    else
    cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_VTR_IN_ID", adNumeric, adParamInput,  20, null)
    end if	     
    
    cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_port_name_alias",     adVarChar, adParamInput, 50,  Request("txtPortNameAlias"))   
		
		'call the update stored proc
		on error resume next
		cmdInsertObj.Execute
		if objConn.Errors.Count <> 0 then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT INSERT OBJECT", objConn.Errors(0).Description
			objConn.Errors.Clear
		end if
		strPortID = cmdInsertObj.Parameters("p_port_id").Value	'set return parameter
		if strPortID = "" then
			DisplayError "BACK", "", 2100, "CANNOT DISPLAY NEW PORT INFORMATION.", "Most probably the new Port Information has been saved successfully even if there was an error retrieving the new id. Close the Port Information window to return to the managed objects screen."
			objConn.Errors.Clear
		end if

    if Request("selAlias") <> "" then
    
    set cmdinsertAlias = server.CreateObject("ADODB.Command")
		set cmdinsertAlias.ActiveConnection = objConn
         SQLscript = "insert into CRP.NETWORK_PORT_NAME_ALIAS(NETWORK_ELEMENT_PORT_ID ,NETWORK_PORT_NAME_ALIAS ,CREATE_REAL_USERID,CREATE_DB_USERID, UPDATE_REAL_USERID) values (" & strPortID & ",'" & Replace( Request("selAlias"), "_", " ")  & "','"& strRealUserID & "','"& strRealUserID & "','"& strRealUserID & "')"
    cmdinsertAlias.CommandText = SQLscript
    cmdinsertAlias.Execute
    end if

     if IsEmpty(MGMT_Systems) <> true and MGMT_Systems <> "null" then
     
		set cmdinsertMGMT = server.CreateObject("ADODB.Command")
		set cmdinsertMGMT.ActiveConnection = objConn
		cmdinsertMGMT.CommandType = adCmdText
   ' dim arr , i
    arr = split(MGMT_Systems,",",10)
     For i=0 To Ubound(arr) 
    dim cmdTxt
        cmdTxt   = "insert into CRP.NETWORK_ELEMENT_MGMT_SYS(NETWORK_ELEMENT_Port_ID,MGMT_SYSTEM_ID,CREATE_REAL_USERID,CREATE_DB_USERID) values (" & strPortID & "," & arr(i) & ",'"& strRealUserID & "','" & strRealUserID & "')"
    cmdinsertMGMT.CommandText=cmdTxt 
    cmdinsertMGMT.Execute
    next
    end if
		strWinMessage = "Record saved successfully. You can now see the changes you made."
	else
		DisplayError "BACK", "", err.Number, "ACCESS DENIED. PLEASE CONTACT YOUR SYSTEM ADMINISTRATOR.", err.Description
	end if
	if err then
		DisplayError "BACK", "", err.Number, "CANNOT UPDATE PORT INFORMATION - TRY AGAIN", err.Description
	end if

end if

    if strAction = "undelete" then
     stop
	'call stor proc to delete current Port Information
	if intAccessLevel and intConst_Access_Delete = intConst_Access_Delete then
		'create command object for update stored proc
		dim cmdDeleteObj
   
		set cmdDeleteObj = server.CreateObject("ADODB.Command")
		set cmdDeleteObj.ActiveConnection = objConn
		cmdDeleteObj.CommandType = adCmdText
		'cmdDeleteObj.CommandText = "sma_sp_userid.spk_sma_mo_inter.sp_mo_port_delete"
         cmdDeleteObj.CommandText = "update CRP.NETWORK_ELEMENT_PORT set RECORD_STATUS_IND ='A',CI_STATUS_ID = 3 where NETWORK_ELEMENT_PORT_ID = " & CLng(strPortID)
		'create params
		'cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_port_id", adNumeric , adParamInput,, CLng(strPortID))
		'cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , CDate(strLastUpdate))
		'call the update stored proc
		if objConn.Errors.Count <> 0 then
	'		DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT - PARAMETER ERROR", objConn.Errors(0).Description
			objConn.Errors.Clear
		end if
		cmdDeleteObj.Execute
		if objConn.Errors.Count <> 0 then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT un DELETE OBJECT", objConn.Errors(0).Description
			objConn.Errors.Clear
		end if
		strWinMessage = "Record activated successfully. "
	else
		DisplayError "BACK", "", err.Number, "ACCESS DENIED. PLEASE CONTACT YOUR SYSTEM ADMINISTRATOR.", err.Description
	end if
	if err then
		DisplayError "BACK", "", err.Number, "CANNOT un DELETE PORT INFORMATION", err.Description
	end if
	'called from main form?
	if Request("back") = "true" then
		Response.Redirect "ManObjPort.asp?ne_id="+strMasterID
	end if
	'ready to enter a new Port Information?
	'strPortID=""
	'strAction="new"
end if
'delete Port Information?
if strAction = "delete" then
     stop
	'call stor proc to delete current Port Information
	if intAccessLevel and intConst_Access_Delete = intConst_Access_Delete then
		'create command object for update stored proc
		dim cmdUnDeleteObj
   
		set cmdUnDeleteObj = server.CreateObject("ADODB.Command")
		set cmdUnDeleteObj.ActiveConnection = objConn
		cmdUnDeleteObj.CommandType = adCmdText
		'cmdDeleteObj.CommandText = "sma_sp_userid.spk_sma_mo_inter.sp_mo_port_delete"
         cmdUnDeleteObj.CommandText = "update CRP.NETWORK_ELEMENT_PORT set RECORD_STATUS_IND ='D',CI_STATUS_ID = 6 where NETWORK_ELEMENT_PORT_ID = " & CLng(strPortID)
		'create params
		'cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_port_id", adNumeric , adParamInput,, CLng(strPortID))
		'cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , CDate(strLastUpdate))
		'call the update stored proc
		if objConn.Errors.Count <> 0 then
	'		DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT - PARAMETER ERROR", objConn.Errors(0).Description
			objConn.Errors.Clear
		end if
		cmdUnDeleteObj.Execute
		if objConn.Errors.Count <> 0 then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT", objConn.Errors(0).Description
			objConn.Errors.Clear
		end if
		strWinMessage = "Record deleted successfully. You can now create a new Port Information."
	else
		DisplayError "BACK", "", err.Number, "ACCESS DENIED. PLEASE CONTACT YOUR SYSTEM ADMINISTRATOR.", err.Description
	end if
	if err then
		DisplayError "BACK", "", err.Number, "CANNOT DELETE PORT INFORMATION", err.Description
	end if
	'called from main form?
	if Request("back") = "true" then
		Response.Redirect "ManObjPort.asp?ne_id="+strMasterID
	end if
	'ready to enter a new Port Information?
	'strPortID=""
	'strAction="new"
end if

'display the Port Information info
if strAction <> "new" then
	sql =	"SELECT "&_
				"NE.NETWORK_ELEMENT_PORT_ID, "&_
				"NE.NETWORK_ELEMENT_ID, "&_
				"NE.NETWORK_ELEMENT_PORT_NAME, "&_
				"NE.NETWORK_ELEMENT_PORT_IP, "&_
				"NE.CUSTOMER_SERVICE_ID, "&_
				"NE.BILLABLE_PORT, "&_
				"NE.CREATE_DATE_TIME, "&_
				"sma_sp_userid.spk_sma_library.sf_get_full_username(NE.CREATE_REAL_USERID) as create_real_userid, "&_
				"NE.UPDATE_DATE_TIME," &_
				"sma_sp_userid.spk_sma_library.sf_get_full_username(NE.UPDATE_REAL_USERID) as update_real_userid, "&_
				"NE.RECORD_STATUS_IND, "&_
				"CS.CUSTOMER_SERVICE_DESC, " &_
				"NE.REPORTABLE, " &_
				"NE.NE_PORT_FUNCTION_LCODE, " &_
				"NE.MSUID ," &_
                "NE.CTR_IN_ID," &_
                 "NE.CTR_OUT_ID,"&_
                 "NE.VN_NAME," &_
          "NE.VTR_OUT_ID," &_
          "NE.QOS_NAME," &_
          "NE.ETR_IN_ID," &_
          "NE.ETR_OUT_ID," &_
          "NE.CI_STATUS_ID," &_
          "CO.ORGANIZATION_NAME," &_
          "NE.ORGANIZATION_ID," &_
          "CO.ORGANIZATION_CODE," &_
          "NE.SITE_ID," &_
          "SNC.SITE_CODE," &_
          "NE.VTR_IN_ID , " &_
          "NPNA.NETWORK_PORT_NAME_ALIAS , NE.PORT_NAME_ALIAS" &_
			
    " FROM CRP.NETWORK_ELEMENT_PORT NE, CRP.CUSTOMER_SERVICE CS  , CRP.CUSTOMER_ORGANIZATION CO, CRP.NETWORK_PORT_NAME_ALIAS NPNA, CRP.SITE_NAME_CODE SNC "&_
			"WHERE NE.CUSTOMER_SERVICE_ID = CS.CUSTOMER_SERVICE_ID(+) and NE.ORGANIZATION_ID = CO.ORGANIZATION_ID(+) AND NE.NETWORK_ELEMENT_PORT_ID = NPNA.NETWORK_ELEMENT_PORT_ID(+) and NPNA.RECORD_STATUS_IND (+)='A'  and NE.SITE_ID = SNC.SITE_ID(+)  " &_
			    "AND NE.NETWORK_ELEMENT_PORT_ID = " & strPortID

	set rsPort=server.CreateObject("ADODB.Recordset")
	rsPort.CursorLocation = adUseClient
	rsPort.Open sql, objConn
	if err then
		DisplayError "BACK", "", err.Number, "DATABASE ERROR", err.Description
	end if
	if rsPort.EOF then
		DisplayError "BACK", "", err.Number, "CANNOT FIND PORT INFORMATION", err.Description
	end if
	set rsPort.ActiveConnection = nothing
    strLastUpdate = rsPort.Fields("UPDATE_DATE_TIME").value

end if
     
    dim rsAlias 
	sql = "SELECT NETWORK_ELEMENT_NAME_ALIAS, NETWORK_ELEMENT_NAME_ALIAS_ID from CRP.NETWORK_ELEMENT_NAME_ALIAS WHERE NETWORK_ELEMENT_ID = " & strMasterID
	set rsAlias=server.CreateObject("ADODB.Recordset")
	rsAlias.CursorLocation = adUseClient
	rsAlias.Open sql, objConn


    dim previId, NextId , prevUpdatetime , nextupdatetime
    
                  while not rsSequence.EOF
					if cint(rsSequence("NETWORK_ELEMENT_PORT_ID")) < cint(rsPort("NETWORK_ELEMENT_PORT_ID"))  then
					   previId = rsSequence("NETWORK_ELEMENT_PORT_ID")
    prevUpdatetime = rsSequence("Update_date_time")
					else
                         if (cint(rsSequence("NETWORK_ELEMENT_PORT_ID")) > cint(rsPort("NETWORK_ELEMENT_PORT_ID"))) and IsEmpty(NextId) = true then
					     NextId = rsSequence("NETWORK_ELEMENT_PORT_ID")
    nextupdatetime = rsSequence("Update_date_time")
                         end if
					end if
					rsSequence.MoveNext

				  wend
				  rsSequence.Close
				  set rsSequence = nothing

   

%>
<html>
<head>
    <meta http-equiv="Pragma" content="no-cache">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
    <link rel="stylesheet" type="text/css" href="stylesheets/styles.css">
    <title>Port Information Detail</title>

    <script type="text/javascript" src="GeneralJavaFunctions.js"></script>
    <script type="text/javascript" src="AccessLevels.js"></script>
    <script type="text/javascript">
        var bolSaveRequired = false;
        intAccessLevel = "<%=intAccessLevel%>";
        var intConst_MessageDisplay = "<%=intConst_MessageDisplay%>";

        function fct_onChange() {
            bolSaveRequired = true;
        }
        function btnPopulate_onclick() {
            var strSuggestPortName = document.frmPort.selPortType.value + document.frmPort.txtPortNumber.value;
            fct_onChange();
            document.frmPort.txtPortName.value = strSuggestPortName;
        }
        function btnNew_click() {
            if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) { alert('Access denied. Please contact your system administrator.'); return; }
            var strMasterID = "<%=strMasterID%>";
            var custId = "<%=strCustId %>";
            document.location.href = "ManObjPortDetail.asp?action=new&masterID=" + strMasterID + "&CustId=" + custId;;
        }

        function fct_onDelete() {
            if (document.frmPort.PortID.value != '') {
                var strMasterID = "<%=strMasterID%>";
                var strPortID = "<%=strPortID%>";
                var strLastUpdate = "<%=strLastUpdate%>";
                var custId = "<%=strCustId %>";

                if ((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) { alert('Access denied. Please contact your system administrator.'); return; }


                if (confirm('Do you really want to delete this object?')) {

                    if (document.frmPort.txtRecordStatusInd.value == "D") {
                        document.location.href = "ManObjPortDetail.asp?action=undelete&PortID=" + strPortID + "&masterID=" + strMasterID + "&hdnLastUpdate=" + strLastUpdate + "&CustId=" + custId;
                    }
                    else {
                        document.location.href = "ManObjPortDetail.asp?action=delete&PortID=" + strPortID + "&masterID=" + strMasterID + "&hdnLastUpdate=" + strLastUpdate + "&CustId=" + custId;
                    }

                }
            } else { fct_displayStatus('There is no need to delete an empty Port Information.'); }
        }

        function btnClose_onclick() {
            window.close();
        }

        function frmPort_onsubmit() {
            if ((((intAccessLevel & intConst_Access_Create) == intConst_Access_Create) && (document.frmPort.PortID.value == "")) || (((intAccessLevel & intConst_Access_Update) == intConst_Access_Update) && (document.frmPort.PortID.value != ""))) {


                //		var strMSUIDComments = document.frmPort.MSUID.value ;
                //		if (strMSUIDComments.length > 50 ) {
                //			alert('The Port Identification (MSUID) can be at most 50 characters.\n\nYou entered ' + strMSUIDComments.length + ' character(s).');
                //			document.frmPort.MSUID.focus();
                //			bolSaveRequired = false;
                //			return(false);
                //		}
                //If Billable port is yes, CSID must be provided
                if ((document.frmPort.chkBillable.checked) &&
                    (document.frmPort.txtCSName.value == '' || (document.frmPort.lngCSID.value == ''))) {
                    alert("If port is billable, you must provide a Customer Service Name or ID. Please re-enter.");
                    document.frmPort.txtCSName.focus();
                    return (false);
                }

                if ((document.frmPort.chkReportable.checked ||
                    document.frmPort.chkBillable.checked) && (document.frmPort.txtPortName.value == " " || document.frmPort.txtPortName.value == "")) {
                    alert("If port is reportable or billable , you must provide a Port Name.");
                    if (document.frmPort.txtPortName)
                        document.frmPort.txtPortName.focus();
                    return (false);
                }


                if ((document.frmPort.chkReportable.checked) &&
                   (document.frmPort.txtCSName.value == '' || (document.frmPort.lngCSID.value == ''))) {
                    alert("If port is reportable, you must provide a Customer Service Name or ID. Please re-enter.");
                    document.frmPort.txtCSName.focus();
                    return (false);
                }

                //Validate LAN IP address.
                var strPortIP = document.frmPort.txtPortIP.value;
                if (strPortIP != "") {

                    //Check if LAN IP address has the correct format.
                    //var re = /[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}/;
                    var re = /^((\d{1,3}\.){3}\d{1,3})|(([0-9,A-F]{0,4}:){2,7}[0-9,A-F]{0,4})$/i;
                    var rv = strPortIP.search(re);
                    if (rv != 0) {
                        //alert('The standard format for LAN IP is:\n\tnnn.nnn.nnn.nnn\nwhere n is a digit. Please re-enter.');
                        alert('The standard format for LAN IP is either \n\tnnn.nnn.nnn.nnn\nwhere n is a digit, or \n\tx:x:x:x:x:x:x:x\n where the x is the hexadecimal values, Please re-enter.');
                        document.frmPort.txtPortIP.focus();
                        return (false);
                    }

                    //check if IP address has 4 segments.
                    re = /^((\d{1,3}\.){3}\d{1,3})$/;
                    rv = strPortIP.search(re);
                    if (rv == 0) {
                        var IPOctet = strPortIP.split(".", 5);
                        if (IPOctet.length != 4) {
                            alert('The LAN IP should have 4 octets.\n\tPlease re-enter.');
                            document.frmPort.txtPortIP.focus();
                            return (false);
                        }

                        //Check the value range of each segment.
                        for (var i = 0; i < IPOctet.length; i++) {
                            if (IPOctet[i] < 0 || IPOctet[i] > 255) {
                                alert('Invalid value in LAN IP Octet #' + (i + 1) + '.\nPlease re-enter.');
                                document.frmPort.txtPortIP.focus();
                                return (false);
                            }
                        }
                    }
                }

                //Check if Port Name is filled
                //  if (document.frmPort.txtPortName.value != "") {
                document.frmPort.action.value = "save";
                bolSaveRequired = false;
                document.frmPort.submit();
                return (true);
                //}
                //else {
                //    alert("You cannot save an empty Port Information record.  Please re-enter.");
                //    return (false);
                //}

            }
            else { alert('Access denied. Please contact your system administrator.'); return (false); }
        }

        function fct_clearStatus() {
            window.status = "";
        }

        function fct_displayStatus(strMessage) {
            window.status = strMessage;
            setTimeout('fct_clearStatus()', intConst_MessageDisplay);
        }

        function body_onLoad(strWinStatus) {
            var strWinStatus = '<%=strWinMessage%>';
            fct_displayStatus(strWinStatus);
        }

        function body_onBeforeUnload() {
            if (bolSaveRequired) {
                event.returnValue = "There is unsaved data on the screen. To save changes, click CANCEL below then click SAVE on the main form.";
            }
        }

        function btnReset_onclick() {
            var strMasterID = '<%=strMasterID%>';
            var strPortID = '<%=strPortID%>';
            var action = '<%=strAction%>';
            var custId = "<%=strCustId %>";
            if (confirm('All changes will be lost. Do you really want to reset the page?')) {
                bolSaveRequired = false;
                document.location = "ManObjPortDetail.asp?action=" + action + "&PortID=" + strPortID + "&masterID=" + strMasterID + "&CustId=" + custId;;
            }
        }

        function body_onUnload() {
            
            opener.document.frmMODetails.btn_iFrame2Refresh.click();
            // window.opener.location.reload();
        }


        function fct_CSLookup(CSKey) {
            var strCSName = document.frmPort.txtCSName.value;
            var strCSID = document.frmPort.lngCSID.value;


            switch (CSKey) {
                case 'CSID':
                    SetCookie("CustomerService", "");
                    SetCookie("CustomerServiceID", strCSID);
                    break;

                case 'CSName':
                    SetCookie("CustomerService", strCSName);
                    SetCookie("CustomerServiceID", "");
                    break;
            }

            SetCookie("WinName", 'Popup');
            SetCookie("ServiceEnd", 'D');
            window.open('SearchFrame.asp?fraSrc=CustServ', 'Popup', 'top=150, left=150,  WIDTH=1000, HEIGHT=700');

            document.frmPort.btnSave.disabled = false;
        }

        function gotoPort(portId, dateTime) {
            
            var url = window.location.href;

            if (url.indexOf('?') > -1) {
                url = url.split('?')[0];
            }
            if (dateTime == 'prev') {
                dateTime = document.frmPort.hdnPrevLastUpdate.value
            }
            else {
                dateTime = document.frmPort.hdnNextLastUpdate.value
            }
            url += "?action=update&PortID=" + portId + "&masterID=" + document.frmPort.MasterID.value + "&hdnLastUpdate=" + dateTime + "&CustId=" + document.frmPort.CustId.value;
            window.location.href = url;
            //   document.frmPort.PortID = portId;
            // document.frmPort.submit();
        }

    </script>
</head>

<body onload="body_onLoad();" onbeforeunload="body_onBeforeUnload();" onunload="body_onUnload();">
    <form name="frmPort" language="javascript" onsubmit="return frmPort_onsubmit()">
        <input type="hidden" name="action" value="">
        <input type="hidden" name="hdnLastUpdate" value="<%Response.Write rsPort.Fields("UPDATE_DATE_TIME").value%>">
        <input type="hidden" name="PortID" value="<%if strPortID <> "" and not bolclone then Response.Write rsPort("NETWORK_ELEMENT_PORT_ID")%>">
        <input type="hidden" name="MasterID" value="<%=strMasterID%>">
        <input type="hidden" name="CustId" value="<%=strCustId %>" />
        <input type="hidden" name="hdnPrevLastUpdate" value="<%Response.Write prevUpdatetime%>">
        <input type="hidden" name="hdnNextLastUpdate" value="<%Response.Write nextupdatetime%>">
        <table style="display: <% if strAction <> "new"   then Response.Write "block" else Response.Write "none"  %>">
            <tr>
                <td colspan="4"></td>
                <td>
                    <p align="right">
                        <img src="images/back_002.gif" id="_imgPortPrev" <% if IsEmpty(previId) then Response.Write "disabled"  %> alt="Go Back" onclick="gotoPort(<%= previId %> ,  'prev');" width="31" height="31">&nbsp;
		                <img src="images/forward_002.gif" id="_imgPortNext" <% if IsEmpty(NextId) then Response.Write "disabled"  %> alt="Go Forward" onclick="gotoPort( <%= NextId %> ,  'next');" width="31" height="31">&nbsp;
                    </p>

                </td>
            </tr>
        </table>
        <table border="0" width="100%">
            <thead>
                <tr>
                    <td colspan="5">Port Information Detail</td>
                </tr>
            </thead>
            <%         %>
            <tbody>
                <tr>
                    <td align="RIGHT" nowrap>Port Type/Port No.</td>
                    <td align="LEFT">
                        <select name="selPortType" onchange="fct_onChange();">
                            <%
                                 
                                Response.Write "<OPTION selected value=></option> vbCrLf"
				  dim strPortNumber
                                
				  while not rsPortType.EOF
					if ((rsPort(2) = Empty) or (IsNull(rsPort(2).Value) = true) or (mid(rsPort(2), 1, len(rsPortType(0))) <> rsPortType(0))) then
					   Response.write "<OPTION  value=" & rsPortType(0) & ">" & rsPortType(0) & " </option> vbCrLf"
					else
					   Response.write "<OPTION  selected value=" & rsPortType(0) & ">" & rsPortType(0) & "</option> vbCrLf"
					   strPortNumber = mid(rsPort(2), len(rsPortType(0))+1)
					end if
					rsPortType.MoveNext

				  wend
				  rsPortType.Close
				  set rsPortType = nothing
                            %>
                        </select>
                        <input size="30" maxlength="40" name="txtPortNumber" value="<%if strPortNumber <> "" then Response.Write strPortNumber%>" onchange="fct_onChange();">
                        <input id="btnPopulate" name="btnPopulate" style="height: 22px; width: 65px" type="button" value="Populate" language="javascript" onclick="return btnPopulate_onclick()">
                    </td>
                    <td></td>
                    <td align="RIGHT" nowrap>IP</td>
                    <td>
                        <input size="50" maxlength="50" name="txtPortIP" value="<%if strPortID <> "" then Response.write rsPort("NETWORK_ELEMENT_PORT_IP")%>" onchange="fct_onChange();"></td>
                    <td></td>
                </tr>
                <tr>
                    <td align="right">Port Name<font color="red">*</font></td>
                    <td colspan="2">
                        <input readonly style="color: silver" size="55" maxlength="50" name="txtPortName" value="<%if strPortID <> "" then Response.write rsPort("NETWORK_ELEMENT_PORT_NAME")%>" onchange="fct_onChange();"></td>
                    <td align="right" nowrap>Billable Port?</td>
                    <td align="left">
                        <input id="chkBillable" name="chkBillable" type="checkbox" <%if (rsPort(5)= "N" or rsPort(5) = Empty) then Response.Write "unchecked" else Response.Write "checked"%> onclick="fct_onChange();"></td>
                    <td></td>
                </tr>

                <tr>
                    <td align="right" nowrap>Customer Service Name</td>
                    <td colspan="2">
                        <input id="txtCSName" name="txtCSName" style="height: 21px; width: 300px" value="<%if rsPort(4) <> 0 then  Response.Write rsPort(11)%>" onchange="fct_onChange();">
                        <input name="btnCSLookup" type="button" onclick="fct_CSLookup('CSName'); fct_onChange();" value="..."></td>
                    <td align="right" nowrap>Reporting Required?</td>
                    <td align="left">
                        <input id="chkReportable" name="chkReportable" type="checkbox" <%if (rsPort(12)= "N" or rsPort(12) = Empty) then Response.Write "unchecked" else Response.Write "checked"%> onclick="fct_onChange();"></td>
                    <td></td>
                </tr>

                <%         %>
                <!--  New changes-->
                <tr>
                    <td align="RIGHT" nowrap>CTR IN</td>
                    <td>
                        <!--<input size="50" maxlength="50" name="txtCTR_IN" value="<%if strCTR_IN <> "" then Response.write rsPort("CTR_IN_ID")%>" onchange="fct_onChange();"></td>-->

                        <select id="setCTR_IN_ID" name="selCTR_IN_ID" onchange="fct_onChange();">
                            <option selected value=""></option>
                            <%   
                                
			        while not rsCTRINFunction.EOF
					Response.Write "<OPTION"
					if strPortID <> ""  then 
                                 if IsNull( rsPort("CTR_IN_ID").Value) <>  true and IsEmpty(rsPort("CTR_IN_ID").Value) <> true then
                                  if CLng(rsPort("CTR_IN_ID").Value) = CLng(rsCTRINFunction(0)) then Response.write " selected" else Response.write " " end if
                                end if
                                end if
					   Response.write " value=" & rsCTRINFunction(0) & ">" & routineHtmlString(rsCTRINFunction(1)) & "</option>" &vbCrLf
					rsCTRINFunction.MoveNext
				wend
				
                                rsCTRINFunction.Close
                            %>
                        </select>
                    <td></td>

                    <td align="RIGHT" nowrap>CTR OUT</td>
                    <td>
                        <!--<input size="50" maxlength="50" name="txtCTR_OUT" value="<%if strCTR_OUT <> "" then Response.write rsPort("CTR_OUT_ID")%>" onchange="fct_onChange();"></td>-->

                        <select id="setCTR_OUT_ID" name="selCTR_OUT_ID" onchange="fct_onChange();">
                            <option value=""></option>
                            <%   
			        while not rsCTROutFunction.EOF
					Response.Write "<OPTION"

                                if strPortID <> ""  then 
                                 if IsNull( rsPort("CTR_OUT_ID").Value) <>  true and IsEmpty(rsPort("CTR_OUT_ID").Value) <> true then
                                  if CLng(rsPort("CTR_OUT_ID").Value) = CLng(rsCTROutFunction(0)) then Response.write " selected" else Response.write " " end if
                                end if
                                end if

					   Response.write " value=" & rsCTROutFunction(0).Value & ">" & routineHtmlString(rsCTROutFunction(1).Value) & "</option>" &vbCrLf
					rsCTROutFunction.MoveNext
				wend
				
                                rsCTROutFunction.Close
                            %>
                        </select>
                    <td></td>
                </tr>
                <tr>
                    <td align="RIGHT" nowrap>VN NAME</td>
                    <td>
                        <input size="50" maxlength="50" name="txtVN_NAME" value="<%if strVN_NAME <> "" then Response.write rsPort("VN_NAME")%>" onchange="fct_onChange();"></td>
                    <td></td>


                    <td align="RIGHT" nowrap>QOS NAME</td>
                    <td>
                        <input size="50" maxlength="50" name="txtQOS_NAME" value="<%if strQOS_NAME <> "" then Response.write rsPort("QOS_NAME")%>" onchange="fct_onChange();"></td>
                    <td></td>
                </tr>
                <tr>
                    <td align="RIGHT" nowrap>VTR IN</td>
                    <td>
                        <!--<input size="50" maxlength="50" name="txtVTR_IN" value="<%if strVTR_IN <> "" then Response.write rsPort("VTR_IN_ID")%>" onchange="fct_onChange();">-->

                        <select id="setVTR_In_ID" name="selVTR_In_ID" onchange="fct_onChange();">
                            <option value=""></option>
                            <%   
			        while not rsVTRInFunction.EOF
					Response.Write "<OPTION"
					
                                 if strPortID <> ""  then 
                                 if IsNull( rsPort("VTR_In_ID").Value) <>  true and IsEmpty(rsPort("VTR_In_ID").Value) <> true then
                                  if CLng(rsPort("VTR_In_ID").Value) = CLng(rsVTRInFunction(0)) then Response.write " selected" else Response.write " " end if
                                end if
                                end if
					   Response.write " value=" & rsVTRInFunction(0) & ">" & routineHtmlString(rsVTRInFunction(1)) & "</option>" &vbCrLf
					rsVTRInFunction.MoveNext
				wend
				
                                rsVTRInFunction.Close
                            %>
                        </select>
                    </td>
                    <td></td>

                    <td align="RIGHT" nowrap>VTR OUT</td>
                    <td>
                        <!--<input size="50" maxlength="50" name="txtVTR_OUT" value="<%if strVTR_OUT <> "" then Response.write rsPort("VTR_OUT_ID")%>" onchange="fct_onChange();">-->

                        <select id="setVTR_Out_ID" name="selVTR_Out_ID" onchange="fct_onChange();">
                            <option value=""></option>

                            <%   
                                                 
			        while not rsVTROutFunction.EOF 
					Response.Write "<OPTION"
                                if strPortID <> ""  then 
                                 if IsNull( rsPort("VTR_OUT_ID").Value) <>  true and IsEmpty(rsPort("VTR_OUT_ID").Value) <> true then
                                  if CLng(rsPort("VTR_OUT_ID").Value) = CLng(rsVTROutFunction(0)) then Response.write " selected" else Response.write " " end if
                                end if
                                end if
					   Response.write " value=" & rsVTROutFunction(0) & ">" & routineHtmlString(rsVTROutFunction(1)) & "</option>" &vbCrLf
					rsVTROutFunction.MoveNext
				wend
				
                                rsVTROutFunction.Close
                            %>
                        </select>
                    </td>
                    <td></td>
                </tr>

                <tr>
                    <td align="RIGHT" nowrap>ETR IN</td>
                    <td>
                        <!-- <input size="50" maxlength="50" name="txtETR_IN" value="<%if strETR_IN <> "" then Response.write rsPort("ETR_IN_ID")%>" onchange="fct_onChange();">-->

                        <select id="setETR_In_ID" name="selETR_In_ID" onchange="fct_onChange();">
                            <option value=""></option>
                            <%   
			        while not rsETRInFunction.EOF
					Response.Write "<OPTION"
					
                                 if strPortID <> ""  then 
                                 if IsNull( rsPort("ETR_In_ID").Value) <>  true and IsEmpty(rsPort("ETR_In_ID").Value) <> true then
                                  if CLng(rsPort("ETR_In_ID").Value) = CLng(rsETRInFunction(0)) then Response.write " selected" else Response.write " " end if
                                end if
                                end if
					   Response.write " value=" & rsETRInFunction(0) & ">" & routineHtmlString(rsETRInFunction(1)) & "</option>" &vbCrLf
					rsETRInFunction.MoveNext
				wend
				
                                rsETRInFunction.Close
                            %>
                        </select>
                    </td>
                    <td></td>

                    <td align="RIGHT" nowrap>ETR OUT</td>
                    <td>
                        <!--<input size="50" maxlength="50" name="txtETR_OUT" value="<%if strETR_OUT <> "" then Response.write rsPort("ETR_OUT_ID")%>" onchange="fct_onChange();">-->
                        <select id="setETR_OUT_ID" name="selETR_OUT_ID" onchange="fct_onChange();">
                            <option value=""></option>
                            <%   
			        while not rsETROutFunction.EOF
					Response.Write "<OPTION"
					
                                 if strPortID <> ""  then 
                                 if IsNull( rsPort("ETR_OUT_ID").Value) <>  true and IsEmpty(rsPort("ETR_OUT_ID").Value) <> true then
                                  if CLng(rsPort("ETR_OUT_ID").Value) = CLng(rsETROutFunction(0)) then Response.write " selected" else Response.write " " end if
                                end if
                                end if
					   Response.write " value=" & rsETROutFunction(0) & ">" & routineHtmlString(rsETROutFunction(1)) & "</option>" &vbCrLf
					rsETROutFunction.MoveNext
				wend
				
                                rsETROutFunction.Close
                            %>
                        </select>

                    </td>
                </tr>
                <tr>
                    <td align="RIGHT" nowrap>CI STATUS</td>
                    <td>
                        <!-- <input size="50" maxlength="50" name="txtCI_STATUS" value="<%if strCI_STATUS <> "" then Response.write rsPort("CI_STATUS_ID")%>" onchange="fct_onChange();">-->
                        <select id="selCI_Status_ID" name="selCI_Status_ID" onchange="fct_onChange();">
                            <option value=""></option>
                            <%   
			        while not rsCIStatusFunction.EOF
					Response.Write "<OPTION"
					
                                 if strPortID <> ""  then 
                                 if IsNull( rsPort("CI_STATUS_ID").Value) <>  true and IsEmpty(rsPort("CI_STATUS_ID").Value) <> true then
                                  if CLng(rsPort("CI_STATUS_ID").Value) = CLng(rsCIStatusFunction(0)) then Response.write " selected" else Response.write " " end if
                                end if
                                end if
					   Response.write " value=" & rsCIStatusFunction(0) & ">" & routineHtmlString(rsCIStatusFunction(1)) & "</option>" &vbCrLf
					rsCIStatusFunction.MoveNext
				wend
				
                                rsCIStatusFunction.Close
                            %>
                        </select>
                    </td>
                    <td></td>
                    <td align="right" nowrap>NAME ALIAS</td>
                    <td valign="top">
                        <select id="selAlias" name="selAlias" onchange="fct_onChange();">
                            <option value=""></option>
                            <%   
                                 dim temp 
			        while not rsAlias.EOF
					Response.Write "<OPTION"
					
                                 if strMasterID <> "" and IsEmpty(rsPort) <> true  then 
                                 if  IsEmpty(rsPort("NETWORK_PORT_NAME_ALIAS")) <> true and   IsNull( rsAlias("NETWORK_ELEMENT_NAME_ALIAS_ID").Value) <>  true and IsEmpty(rsAlias("NETWORK_ELEMENT_NAME_ALIAS_ID").Value) <> true then
                                  
                                if rsPort("NETWORK_PORT_NAME_ALIAS").Value = rsAlias("NETWORK_ELEMENT_NAME_ALIAS").Value then Response.write " selected" else Response.write " " end if
                            
                                 end if
                                end if
                                
                                temp = Replace(rsAlias("NETWORK_ELEMENT_NAME_ALIAS").Value, " ", "_")  
					   Response.write " value=" & routineHtmlString(temp)& ">" & routineHtmlString(rsAlias("NETWORK_ELEMENT_NAME_ALIAS").Value) & "</option>" &vbCrLf
					rsAlias.MoveNext
				wend
				
                                rsAlias.Close
                            %>
                        </select>
                    </td>
                </tr>

                <tr>
                    <%        %>
                    <!--<td align="RIGHT" nowrap>ORGANIZATION_NAME</td>
        <td>
            <input size="50" maxlength="50" name="txtORGANIZATION_NAME" value="<%if strORGANIZATION_NAME <> "" then Response.write rsPort("ORGANIZATION_NAME")%>" onchange="fct_onChange();"></td>
        <td></td>-->
                    <td align="RIGHT" nowrap>ORGANIZATION NAME</td>
                    <td valign="top" colspan="2">
                        <select id="setORGANIZATION_NAME" name="selORGANIZATION_NAME" onchange="fct_onChange();">
                            <option value=""></option>
                            <%    dim strORGANIZATION_NAME
				 
			        while not rsOrganisationFunction.EOF
					Response.Write "<OPTION"
					
                                 if strPortID <> ""  then 
                                 if rsPort("ORGANIZATION_ID") <> empty then
                                if IsNull( rsPort("ORGANIZATION_ID").Value) <>  true and IsEmpty(rsPort("ORGANIZATION_ID").Value) <> true then
                                  if CLng(rsPort("ORGANIZATION_ID").Value) = CLng(rsOrganisationFunction(0)) then Response.write " selected" else Response.write " " end if
                                end if
                                 end if
                                end if
					   Response.write " value=" & rsOrganisationFunction(0) & ">" & routineHtmlString(rsOrganisationFunction(1)) & "</option>" &vbCrLf
					rsOrganisationFunction.MoveNext
				wend
				
                                rsOrganisationFunction.Close
                            %>
                        </select>

                        <!-- <td align="RIGHT" nowrap>ORGANIZATION_CODE</td>
                    <td>
                        <input size="50" maxlength="50" name="txtORGANIZATION_CODE" value="<%if strORGANIZATION_CODE <> "" then Response.write rsPort("ORGANIZATION_CODE")%>" onchange="fct_onChange();">

                    </td>-->
                        <!--</tr>
                <tr>
                    <td align="RIGHT" nowrap>SITE_NAME</td>
                    <td>
                        <input size="50" maxlength="50" name="txtSITE_NAME" value="<%if strSITE_NAME <> "" then Response.write rsPort("SITE_NAME")%>" onchange="fct_onChange();">

                    </td>-->


                        <td align="RIGHT" nowrap>SITE NAME</td>
                    <td valign="top" colspan="2">
                        <select id="setSite_NAME" name="selSITE_NAME" onchange="fct_onChange();">
                            <option value=""></option>
                            <%    dim strSITE_NAME
				 
			        while not rsSiteNameFunction.EOF
					Response.Write "<OPTION"
					
                                 if strPortID <> ""  then 
                                 if IsNull( rsPort("SITE_ID").Value) <>  true and IsEmpty(rsPort("SITE_ID").Value) <> true then
                                  if CLng(rsPort("SITE_ID").Value) = CLng(rsSiteNameFunction(0)) then Response.write " selected" else Response.write " " end if
                                end if
                                end if
					   Response.write " value=" & rsSiteNameFunction(0) & ">" & routineHtmlString(rsSiteNameFunction(1)) & "</option>" &vbCrLf
					rsSiteNameFunction.MoveNext
				wend
				rsSiteNameFunction.Close
                            %>
                        </select>
                    </td>
                </tr>

                <tr>
                    <td align="right" nowrap>Customer Service ID</td>
                    <td>
                        <input id="lngCSID" name="lngCSID" style="height: 21px; width: 75px" value="<%if rsPort(4) <> 0 then  Response.Write rsPort(4)%>" onchange="fct_onChange();">
                        <input name="btnCSIDLookup" type="button" onclick="fct_CSLookup('CSID'); fct_onChange();" value="...">
                    </td>
                    <td></td>
                    <td align="RIGHT" nowrap>Port Function</td>
                    <td valign="top">
                        <select id="selPortFunction" name="selPortFunction" onchange="fct_onChange();">
                            <%  dim strPortFunction
				
			        while not rsPortFunction.EOF
					Response.Write "<OPTION"
					if strPortID <> "" then if CLng(rsPort(13)) = CLng(rsPortFunction(1)) then Response.write " selected"
					   Response.write " value=" & rsPortFunction(1) & ">" & routineHtmlString(rsPortFunction(0)) & "</option>" &vbCrLf
					rsPortFunction.MoveNext
				wend
				rsPortFunction.Close
                            %>
                        </select>
                    </td>
                    <td></td>
                    <td></td>
                    <td></td>
                </tr>

                <tr>
                    <td align="right" nowrap>Port Identification</td>
                    <td>
                        <input name="MSUID" size="50" maxlength="50" value="<% if rsPort(0) <> 0 then Response.write  rsPort(14) end if%>" onchange="fct_onChange();"></td>
                    <td></td>
                    <td align="RIGHT" width="20%" nowrap>MGMT SYSTEMS NAME</td>
                    <td width="80%">
                        <!-- <input size='40' maxlength='30' name='txtMGMT_SYSTEMS_NAME' value='<%   rsNE("mgmtSystemName")  %>'>-->
                        <select id="selMGMT_SYSTEMS_NAME" name="selMGMT_SYSTEMS_NAME" multiple="multiple" onchange="fct_onChange();" style="width: 172px">
                            <%		
                                		
                                		
                              dim cmdMGMTSelect , mgmtSelect , comaSeperated 	
                              
                              		
if  strPortID <> empty then		
                                if CLng(strPortID) <> 0 then		
     set cmdMGMTSelect = server.CreateObject("ADODB.Recordset")		
		set cmdMGMTSelect.ActiveConnection = objConn		
      mgmtSelect = "select MGMT_SYSTEM_ID from CRP.NETWORK_ELEMENT_MGMT_SYS where NETWORK_ELEMENT_Port_ID = "  & CLng(strPortID)		
      cmdMGMTSelect.Open mgmtSelect, objConn		
    while not cmdMGMTSelect.EOF 		
    comaSeperated = comaSeperated &  cmdMGMTSelect("MGMT_SYSTEM_ID").Value & ","		
            cmdMGMTSelect.MoveNext		
			wend		
                                end if		
                                 end if	
                                
                                	
                                 if comaSeperated <> Empty then		
                                comaSeperated = mid(comaSeperated,1,len(comaSeperated)-1)		
                                end if		
				while not rsMgmtSystem.EOF		
				Response.Write "<OPTION"		
				if strNE_ID <> "" then		
                                if comaSeperated <> Empty then		
                                if   in_array(rsMgmtSystem("mgmt_system_id").Value,split(comaSeperated,",",10)) <> Empty  and in_array(rsMgmtSystem("mgmt_system_id").Value,split(comaSeperated,",",10)) <> False then		
                                 Response.write " selected"		
                                end if		
                                end if		
                                end if		
					Response.Write " VALUE="& rsMgmtSystem("mgmt_system_id").Value &">" & routineHtmlString(rsMgmtSystem("mgmt_system_name").Value) & "</OPTION>" &vbCrLf		
				rsMgmtSystem.MoveNext		
			wend		
			rsMgmtSystem.Close		
                            %>
                        </select>
                    </td>

                </tr>
                 <tr>
                    <td align="right">Port Name Alias</td>
                    <td colspan="2">
                        <input  size="55" maxlength="50" name="txtPortNameAlias" value="<%if strPortID <> "" then Response.write rsPort("PORT_NAME_ALIAS")%>" "></td>
                    
                </tr>

                <tr>
                    <td align="right" colspan="5">
                        <input type="button" name="btnClose" value="Close" style="width: 2cm" onclick="return btnClose_onclick();">&nbsp;&nbsp;
	  	<input type="button" name="btnDelete" value='<% if rsPort("RECORD_STATUS_IND").Value = "A" then  Response.write "Delete" else Response.write "UnDelete" end if  %>' style="width: 2cm" <%if bolclone then Response.write " disabled " end if%> onclick="return fct_onDelete();">&nbsp;&nbsp;
	  	<input type="button" name="btnReset" value="Reset" style="width: 2cm" onclick="return btnReset_onclick();">&nbsp;&nbsp;
	  	<input type="button" name="btnNew" value="New" style="width: 2cm" <%if bolclone then Response.write " disabled " end if%> onclick="return btnNew_click();">&nbsp;&nbsp;
	  	<input type="button" name="btnSave" value="Save" style="width: 2cm" onclick="return frmPort_onsubmit();">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    </td>
                </tr>
            </tbody>


        </table>

        <fieldset>
            <legend align="right"><b>Audit Information</b></legend>
            <div size="8pt" align="RIGHT">
                Record Status Indicator
	<input align="left" name="txtRecordStatusInd" type="text" style="height: 20px; width: 18px" disabled value="<%if strPortID <> "" and not bolclone then Response.write rsPort("RECORD_STATUS_IND")%>">&nbsp;&nbsp;&nbsp;
	Create Date&nbsp;<input align="center" name="txtCreateDateTime" type="text" style="height: 20px; width: 150px" disabled value="<%if strPortID <> "" and not bolclone then Response.write rsPort("CREATE_DATE_TIME")%>">&nbsp;
	Created By&nbsp;
                <input align="right" name="txtCreateRealUser" type="text" style="height: 20px; width: 200px" disabled value="<%if strPortID <> "" and not bolclone then Response.write rsPort("CREATE_REAL_USERID")%>"><br>
                Update Date&nbsp;<input align="center" name="txtUpdateDateTime" type="text" style="height: 20px; width: 150px" disabled value="<%if strPortID <> "" and not bolclone then Response.write rsPort("UPDATE_DATE_TIME")%>">
                Updated By&nbsp;
                <input align="right" name="txtUpdateRealUser" type="text" style="height: 20px; width: 200px" disabled value="<%if strPortID <> "" and not bolclone then Response.write rsPort("UPDATE_REAL_USERID")%>">
            </div>
        </fieldset>
        <%
'Clean up our ADO objects
if strPortID <> "" then
	rsPort.close
	set rsPort = Nothing
	objConn.close
	set objConn = Nothing
end if
        %>
    </form>
</body>
</html>
