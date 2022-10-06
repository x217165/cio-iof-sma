<%@ Language=VBScript %>
<%   OPTION EXPLICIT
on error resume next %>
<!-- #include file="smaConstants.inc" -->
<!-- #include file="smaProcs.inc" -->
<!-- #include file="databaseconnect.asp" -->
<!--
*************************************************************************************************
* Page name:	STypeInstList.asp																*	
* Purpose:		To display Service Instances for the Service Type inside a frame  							*														*
*																								*
* Created by:					Date															*
* Linda Chen					09/12/2008														*
* ====================================															*
* Modifications By				Date				Modifcations								*
*
*************************************************************************************************
-->
<%

Dim objRs, strSQL, detailsRS, strserv_inst_det_count, strErrCode, strErrMsg
Dim strServiceTypeID, strXRefID, strUsageID, strSerInstAttrID, strSIAseqId
Dim intAccessLevel

strServiceTypeID = Request("hdnServiceTypeID")
strXRefID = Request("txtXRefID")
strUsageID = Request("hdnUsageID")

strSIAseqId = Request("hdnSIASeqID")




intAccessLevel = CInt(CheckLogon(strConst_ServiceCatalogue))
' The following 3 lines temp commented for my test LC	
if intAccessLevel and intConst_Access_ReadOnly <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access Service Instance. Please contact your system administrator"
end if

select case Request("txtFrmAction")

	
	case "DELETE"  
	' The following 3 lines temp commented for my test LC	
	'if ((intAccessLevel and intConst_Access_Delete) <> intConst_Access_Delete) then
	 ' DisplayError "BACK", "", 0, "DELETE DENIED", "You don't have access to delete Service Instance Attribute. Please contact your system administrator"
	'end if


	dim cmdObjSP_CHK_ORD_INSTNC_ATT
	set cmdObjSP_CHK_ORD_INSTNC_ATT = server.CreateObject("ADODB.Command")
	set cmdObjSP_CHK_ORD_INSTNC_ATT.ActiveConnection = objConn
	cmdObjSP_CHK_ORD_INSTNC_ATT.CommandType = adCmdStoredProc
	cmdObjSP_CHK_ORD_INSTNC_ATT.CommandText = "jagora.SP_CHK_ORD_INSTNC_ATT" 
	

	cmdObjSP_CHK_ORD_INSTNC_ATT.Parameters.Append cmdObjSP_CHK_ORD_INSTNC_ATT.CreateParameter("p_SRVC_INSTNC_ATT_XREF_ID",adNumeric, adParamInput,9,strXRefID) 	
	cmdObjSP_CHK_ORD_INSTNC_ATT.Parameters.Append cmdObjSP_CHK_ORD_INSTNC_ATT.CreateParameter("p_serv_inst_det_count",adNumeric,adParamOutput,9)
	cmdObjSP_CHK_ORD_INSTNC_ATT.Parameters.Append cmdObjSP_CHK_ORD_INSTNC_ATT.CreateParameter("p_err_code", adVarChar,adParamOutput,9)
	cmdObjSP_CHK_ORD_INSTNC_ATT.Parameters.Append cmdObjSP_CHK_ORD_INSTNC_ATT.CreateParameter("p_err_message",adVarChar,adParamOutput,200)					


			'****************************
			'check parameter values		
  			'****************************
  			
  			'dim objparm
  			'for each objparm in cmdObjSP_CHK_ORD_INSTNC_ATT.Parameters
  			'	  Response.Write "<b>" & objparm.name & "</b>"
  			'	  Response.Write " has size:  " & objparm.Size & " "
  			'	  Response.Write " and value:  " & objparm.value & " "
  			'	  Response.Write " and datatype:  " & objparm.Type & "<br> "
  			'next
  									 
  			'response.write (cmdObjSP_CHK_ORD_INSTNC_ATT.CommandText)
  			'response.write (cmdObjSP_CHK_ORD_INSTNC_ATT.Parameters("p_serv_inst_det_count"))
  			
	cmdObjSP_CHK_ORD_INSTNC_ATT.Execute
	strserv_inst_det_count = cmdObjSP_CHK_ORD_INSTNC_ATT.Parameters("p_serv_inst_det_count")
	strErrCode = cmdObjSP_CHK_ORD_INSTNC_ATT.Parameters("p_err_code")
	strErrMsg  = cmdObjSP_CHK_ORD_INSTNC_ATT.Parameters("p_err_message")

	if objConn.Errors.Count <> 0 then
		DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT QUERY BACK RECORD STATUS DUE TO SRT2 DEPENDENCY", objConn.Errors(0).Description
		objConn.Errors.Clear
	end if




	if strserv_inst_det_count = 0 then
	    if strXRefID <> "" then
	    	' replace the following update with calling to SP
		    'strSQL = "UPDATE SO.SRVC_INSTNC_ATT_VAL_USAGE" &_
			'	" SET RECORD_STATUS_IND  = 'D'"  &_
			'	" WHERE SRVC_INSTNC_ATT_VAL_USAGE_ID = " & strUsageID
		    'set objRs = objConn.Execute(strSQL)
		    'if err then
			'   DisplayError "BACK", "", err.Number, "CANNOT DELETE RECORD", err.Description
		    'end if
			
			
			
			
			dim cmdUpdateObj
			set cmdUpdateObj = server.CreateObject("ADODB.Command")
			set cmdUpdateObj.ActiveConnection = objConn
			cmdUpdateObj.CommandType = adCmdStoredProc
			cmdUpdateObj.CommandText = "SMA_SP_USERID.Sp_Srvinst_Val_Xrefusg_Delete"			
 
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_service_type_id",adNumeric , adParamInput,, clng(strServiceTypeID))
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_srvc_Inst_usage_id",adNumeric , adParamInput,,  clng(strUsageID))
	        cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_srvc_instnc_att_xref_id", adNumeric, adParamInput,,  clng(strXRefID))
			'****************************
			'check parameter values		
  			'****************************
  		
  			'dim objparm
  			'for each objparm in cmdUpdateObj.Parameters
  			'	  Response.Write "<b>" & objparm.name & "</b>"
  			'	  Response.Write " has size:  " & objparm.Size & " "
  			'	  Response.Write " and value:  " & objparm.value & " "
  			'	  Response.Write " and datatype:  " & objparm.Type & "<br> "
  			'next
           
		 'response.end
			On Error Resume Next
			cmdUpdateObj.Execute
			If objConn.Errors.Count <> 0 Then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT Delete RECORD", objConn.Errors(0).Description
				objConn.Errors.Clear
			End If

			 response.redirect "StypeInstList.asp?hdnServiceTypeID=" + strServiceTypeID


	
	
 	    end if
	else 
         if strserv_inst_det_count > 0 then 
	     	DisplayError "BACK", "", "999", "CANNOT DELETE RECORD DUE TO SRT2 DEPENDENCY", "there is SRT2 data dependency on this record and therefore cannot be deleted"
	     end if	
	end if

	Set cmdObjSP_CHK_ORD_INSTNC_ATT = Nothing

 end select 			
			
		 
if isnumeric(strServiceTypeID)  then
	
	'Response.Write "Service Type :" & strServiceTypeID & "<P>"	
	
	StrSql =" SELECT a.srvc_instnc_att_name, " &_
					"a.srvc_instnc_att_id, " &_
				    "b.SRVC_instnc_att_val, " &_
				    "b.SRVC_instnc_ATT_VAL_ID, " &_
				    "c.srvc_instnc_att_xref_id, " &_
				    "d.srvc_instnc_att_val_usage_id, " &_
				    "decode (C.DISPLAY_ORDER,0, 9999,C.DISPLAY_ORDER) as seq_ord " &_
			" FROM so.srvc_instnc_att a," &_
				   "so.srvc_instnc_att_val b, " &_
				   "so.srvc_instnc_att_xref c, " &_
				   "so.srvc_instnc_att_val_usage d " &_
			" WHERE a.srvc_instnc_att_id = c.srvc_instnc_att_id  AND " &_
				   "d.srvc_instnc_att_xref_id = c.srvc_instnc_att_xref_id " &_
			" AND  b.SRVC_instnc_ATT_VAL_ID = d.SRVC_INSTNC_ATT_VALUE_ID " &_
			" and  d.RECORD_STATUS_IND ='A' " &_
			" AND  c.service_type_id = " & strServiceTypeID &_
			" order by seq_ord asc, b.srvc_instnc_att_val" 

'response.write StrSql
'response.end
	set objRs = objConn.Execute(strSQL)
	if err then
		DisplayError "BACK", "", err.Number, "ERROR IN SELECTING SERVICE TYPE ATTRIBUTES", err.Description
	end if
	
end if
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK href="stylesheets/styles.css" rel="stylesheet" type="text/css">
<STYLE>
.regularItem {
	cursor: hand;
}
.whiteItem {
	cursor: hand;
	background-color: white; }
	
.Highlight {
	cursor: hand; 
	background-color: #00974f;
	color: white;
}
</STYLE>

<script type="text/javascript">

var oldHighlightedElement;
var oldHighlightedElementClassName;

function cell_onClick(intXRefID, intUsgeID, intatID, intatvID, intServiceType, intSIASeqID){

	document.frmIFR.txtXRefID.value = intXRefID;
	document.frmIFR.hdnUsageID.value = intUsgeID; 
	document.frmIFR.txtInstID.value = intatID;
	document.frmIFR.txtInstvID.value = intatvID;
	document.frmIFR.hdnServiceTypeID.value = intServiceType;
	
	document.frmIFR.hdnSIASeqID.value = intSIASeqID;

	//highlight current record

	if (oldHighlightedElement != null) {
		oldHighlightedElement.className = oldHighlightedElementClassName
	}
	oldHighlightedElement = window.event.srcElement.parentElement;
	oldHighlightedElementClassName = oldHighlightedElement.className;
	oldHighlightedElement.className = "Highlight";

}

</script>

</HEAD>
<BODY>
<form name="frmIFR" action="STypeInstList.asp" method="POST">
		<input type=hidden name=hdnServiceTypeID value="">
		<input type=hidden name=txtXRefID value="">
		<input type=hidden name=txtInstID value="">
		<input type=hidden name=txtInstvID value="">
		<input type=hidden name=hdnUsageID value="">
		
		<input type=hidden name=hdnSIASeqID value="">

		<input type=hidden name=UpdateDateTime value="">


<TABLE border=1 cellspacing=0 cellpadding=2>
	<thead>
		<th class="style1">Service Instance Attribute</th>
		
		<th>Value</th>
		<th>Display Order</th>


	</thead>
	<tbody>
		
<%	if isnumeric(strServiceTypeID)  then

		dim k
		k = 0
		while not objRs.EOF
			if Int(k/2) = k/2 then
				Response.Write "<tr class=""regularItem"">"
			else
				Response.Write "<tr class=""whiteItem"">"
			end if
			k = k+1 %>
		<td nowrap onClick="cell_onClick(<%=objRs(4)%>, <%=objRs(5)%>, <%=objRs(1)%>, <%=objRs(3)%>, <%=strServiceTypeID%>, <%=objRs(6)%>);"><%=objRs(0)%>&nbsp;</td>
		<td nowrap onClick="cell_onClick(<%=objRs(4)%>, <%=objRs(5)%>, <%=objRs(1)%>, <%=objRs(3)%>, <%=strServiceTypeID%>, <%=objRs(6)%>);"><%=objRs(2)%>&nbsp;</td>
				<td nowrap onClick="cell_onClick(<%=objRs(4)%>, <%=objRs(5)%>, <%=objRs(1)%>, <%=objRs(3)%>, <%=strServiceTypeID%>, <%=objRs(6)%>);">
		<% if objRs(6) = "9999" then response.write "" else response.write objRs(6) end if%>&nbsp;</td>


			</tr>
			<% objRs.MoveNext
		wend
		
		objRs.Close
		set objRs = Nothing
		
 end if %>
</tbody>
</table>
</FORM>
</BODY>
</HTML>


