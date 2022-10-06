<%@ Language=VBScript %>
<%  
OPTION EXPLICIT
on error resume next
%>
<!-- #include file="smaConstants.inc" -->
<!-- #include file="databaseconnect.asp" -->
<!-- #include file="smaProcs.inc" -->

<%

'Get asset Id?

dim strAliasID,StrAssetID,StrSql,objRsAddCost,dblAssetVal,intAccessLevel
dblAssetVal=0

strAliasID = Request("AliasID")
StrAssetID = Request("AssetID")
intAccessLevel = CInt(CheckLogon(strConst_AssetAdditionalCosts))

Dim strWhereClause,objRS,strNewFacility,strWinMessage,objRsAssetType,strUpdDate



 select case Request("txtFrmAction")
  case "DELETE" 
  if ((intAccessLevel and intConst_Access_Delete) <> intConst_Access_Delete) then
     DisplayError "BACK", "", 0, "DELETE DENIED", "You don't have access to delete Asset Alias. Please contact your system administrator"
  end if	
    if (Request("AliasID") <>"") then
	    
			dim cmdDeleteObj
			set cmdDeleteObj = server.CreateObject("ADODB.Command")
			set cmdDeleteObj.ActiveConnection = objConn
			cmdDeleteObj.CommandType = adCmdStoredProc
			cmdDeleteObj.CommandText = "sma_sp_userid.spk_sma_asset_inter.sp_aacost_delete"

			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_asset_alias_id", adNumeric, adParamInput,,Clng(Request("AliasID")))	
			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , CDate(Request("hdnUpdateDateTime")))
			
			
			'dim objparm
  			'for each objparm in cmdDeleteObj.Parameters
  			 ' Response.Write "<b>" & objparm.name & "</b>"
  			 'Response.Write " has size:  " & objparm.Size & " "
  			  'Response.Write " and value:  " & objparm.value & " "
  			'Response.Write " and datatype:  " & objparm.Type & "<br> "
  		    'next
  		  
			cmdDeleteObj.Execute
			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE ASSET ADDITIONAL COST", objConn.Errors(0).Description
				objConn.Errors.Clear
			end if
		
			'strWinMessage = "Record deleted successfully."
	 end if	
		'else
	   'DisplayError "BACK", "", 0, "DELETE DENIED", "You don't have access to delete Asset Additional Cost. Please contact your system administrator"
'end if  			
		
	 	
 end select 			
			
	

if strAssetID = "" then
	Response.End
end if

StrSql = "select asset_additional_cost_id,asset_id,to_char(dollar_date,'mon-dd-yyyy') dollar_date_conv,dollar_amount,asset_cost_type_code,UPDATE_DATE_TIME,dollar_date from crp.asset_additional_cost" &_
            "  where asset_id= "& StrAssetID & " ORDER BY DOLLAR_DATE"
            
            
  set objRsAddCost = objConn.Execute(StrSql)

 


if err then
	DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET 32132", err.Description
end if
'release connection


%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<LINK href="stylesheets/styles.css" rel="stylesheet" type="text/css">
<STYLE>
.regularItem {
	cursor: hand;
}
.whiteItem {
	cursor: hand;
	background-color: white;
}
.Highlight {
	cursor: hand; 
	background-color: #00974f;
	color: white;
}
</STYLE>



<script type="text/javascript">

var oldHighlightedElement;
var oldHighlightedElementClassName;

function cell_onClick(dtUpdate,intAliasID,intAssetID){
	document.frmIFR.txtAliasID.value = intAliasID;
	document.frmIFR.txtAssetID.value = intAssetID;
	document.frmIFR.hdnUpdateDateTime.value = dtUpdate; 
	//highlight current record
	if (oldHighlightedElement != null) {oldHighlightedElement.className = oldHighlightedElementClassName}
	oldHighlightedElement = window.event.srcElement.parentElement;
	oldHighlightedElementClassName = oldHighlightedElement.className;
	oldHighlightedElement.className = "Highlight";
}


</script>

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
var dblCost =0.00;
//Derive the asset total cost
if ((document.frmIFR.hdnAssetCost.value !="") && (parent.document.fmAssetDetail.txtpurprice.value==""))
{
 dblCost =parseFloat(document.frmIFR.hdnAssetCost.value);
 dblCost = Math.round(dblCost*100)/100;
}

if ((document.frmIFR.hdnAssetCost.value =="") && (parent.document.fmAssetDetail.txtpurprice.value!=""))
{
 dblCost =parseFloat(parent.document.fmAssetDetail.txtpurprice.value);
 dblCost = Math.round(dblCost*100)/100;
}

if ((document.frmIFR.hdnAssetCost.value !="") && (parent.document.fmAssetDetail.txtpurprice.value!=""))
{
 var dblCost = (parseFloat(document.frmIFR.hdnAssetCost.value)+parseFloat(parent.document.fmAssetDetail.txtpurprice.value));
  dblCost = Math.round(dblCost*100)/100;
}


parent.document.fmAssetDetail.txtassetval.value = dblCost;
 
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">
<form name="frmIFR" action="AssetAlias.asp" method="POST">
<input type="hidden" name="txtAliasID" value="">
<input type="hidden" name="txtAssetID" value="">
<input type="hidden" name="hdnUpdateDateTime" value="">

<TABLE border=1 cellspacing=0 cellpadding=2 width="100%">
	<thead>
		<th nowrap title="Dollar Date">Date</th>
		<th nowrap title="Dollar">Dollar</th>
		<th nowrap title="Cost Type">Type</th>
	</thead>
	<tbody>
		<%
		dim k
		k = 0
		
		while not objRsAddCost.EOF
		
		   if not isnull(objRsAddCost("dollar_amount")) then
			  dblAssetVal = cdbl(dblAssetVal) + cdbl(objRsAddCost("dollar_amount"))
		   end if
			if Int(k/2) = k/2 then
				Response.Write "<tr class=""regularItem"">"
			else
				Response.Write "<tr class=""whiteItem"">"
			end if
			k = k+1
		%>
			<td nowrap onClick="cell_onClick('<%=objRsAddCost("UPDATE_DATE_TIME")%>',<%=objRsAddCost("asset_additional_cost_id")%>, <%=objRsAddCost("ASSET_ID")%>);"><%=objRsAddCost("dollar_date_conv")%>&nbsp;</td>
			<td nowrap onClick="cell_onClick('<%=objRsAddCost("UPDATE_DATE_TIME")%>',<%=objRsAddCost("asset_additional_cost_id")%>, <%=objRsAddCost("ASSET_ID")%>);"><%=FormatNumber(objRsAddCost("dollar_amount"),-1,-2,-2,0)%>&nbsp;</td>
			<td nowrap onClick="cell_onClick('<%=objRsAddCost("UPDATE_DATE_TIME")%>',<%=objRsAddCost("asset_additional_cost_id")%>, <%=objRsAddCost("ASSET_ID")%>);"><%=objRsAddCost("asset_cost_type_code")%>&nbsp;</td>
			
		</tr>
		<%
		objRsAddCost.MoveNext
		wend
		objRsAddCost.Close
		set objRsAddCost = Nothing
		%>
	</tbody>
</table>
  <input type="hidden" name="hdnAssetCost" value="<%=dblAssetVal%>">
</FORM>
</BODY>
</HTML>


