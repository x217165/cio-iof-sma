<%@ LANGUAGE=VBSCRIPT %>
<% option explicit
'on error resume next %>
<!-- #include file=databaseconnect.asp -->
<!-- #include file=smaConstants.inc -->
<!-- #include file=smaProcs.inc -->
<!--
*************************************************************************************
* File Name:	AssetCriteria.asp
*
* Purpose:	
*
* In Param:		
*
* Out Param:
*
* Created By:	
* Edited by:    Adam Haydey Mar 2, 2001
*               CR 1550 Added TAC Asset Code (Barcode) search fields.
*        29-Jul-15   PSmith  Set Cookies in validation so the back key works
*        05-Oct-15   PSmith  Only sumbit() for pop-up windows
*        03-Feb-16   PSmith  Don't pre-populate search criteria
**************************************************************************************
-->
<%
'check user's rights
dim intAccessLevel,objRs,StrSql,objRsAssetType, strWinName, strCustomerName, strMake, strModel, strAssetCode

intAccessLevel = CInt(CheckLogon(strConst_Asset))

if ((intAccessLevel and intConst_Access_ReadOnly) <> intConst_Access_ReadOnly) then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to the Asset SCREEN. Please contact your system administrator"
end if

  StrSql = "SELECT NOC_REGION_LCODE,NOC_REGION_DESC FROM CRP.LCODE_NOC_REGION WHERE RECORD_STATUS_IND = 'A' ORDER BY  NOC_REGION_LCODE"
      
     'Create Recordset object  
   set objRS = objConn.Execute(StrSql)
   
   StrSql = "SELECT asset_type_id, ASSET_TYPE_DESC FROM CRP.ASSET_TYPE WHERE RECORD_STATUS_IND = 'A' ORDER BY  ASSET_TYPE_DESC"
   'Create Recordset object  
   set objRsAssetType = objConn.Execute(StrSql)
   
    if err then
	  DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
    end if
    
    strWinName = Request.Cookies("WinName")
    strCustomerName = Request.Cookies("CustomerName")
    strMake = Request.Cookies("Make")
    strModel = Request.Cookies("Model")
    
%>

<HTML>
<HEAD>
<LINK href="stylesheets/styles.css" rel="stylesheet" type="text/css">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
<script type="text/javascript" SRC="AccessLevels.js"></script>
<script type="text/javascript">
var intAccessLevel=<%=intAccessLevel%>;
//set section title
setPageTitle("SMA - Asset");

function fct_onLoad() {
	
	var strWinName = document.frmAssetSearch.hdnWinName.value ;
	
 	if (strWinName ==  "Popup"){
 		DeleteCookie("AssetID");
 		DeleteCookie("AssetName");
 		DeleteCookie("WinName");
 	}
 	
 	DeleteCookie("CustomerName");
 	DeleteCookie("Make");
 	DeleteCookie("Model");
 	
 	if (strWinName == "Popup" && ((document.frmAssetSearch.txtassetid.value != "")||(document.frmAssetSearch.txtcustomerName.value != "")||(document.frmAssetSearch.txtassetmake.value != "" )||(document.frmAssetSearch.txtassetmodel.value != "" )) ){
		SetCookie("AssetID",document.frmAssetSearch.txtassetid.value);
		SetCookie("CustomerName",document.frmAssetSearch.txtcustomerName.value);
		SetCookie("Make",document.frmAssetSearch.txtassetmake.value);
		SetCookie("Model",document.frmAssetSearch.txtassetmodel.value);
    thinking(parent.fraResult);
 		document.frmAssetSearch.submit();
 	}
}

function btnSearch_click() {
  // Start thinking
  thinking(parent.fraResult);
}


function btnClear_click() {
	with(document.frmAssetSearch){
		txtcustomerName.value = "";
		txttacname.value = "" ;
		selregion.selectedIndex  = 0;
		selassettype.selectedIndex = 0 ;
		txtassetmake.value = "" ;
		txtassetid.value = "" ;
		txtassetmodel.value = "" ;
		txtserial.value = "" ;
		txtcllicode.value = "" ;
		txtassetcode.value = "";	
	} 
}

function btnNew_click(){
	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {alert('Access denied. Please contact your system administrator.'); return;}
	parent.document.location.href ="AssetDetail.asp?NewFacility=NEW";
}

</script>

</HEAD>
<BODY onLoad="fct_onLoad();">
<form name="frmAssetSearch" action="assetlist.asp" method="post" target="fraResult" onSubmit="btnSearch_click()">
	<!-- hidden variables -->
	<INPUT name="hdnWinName" type=hidden value="<%=strWinName%>">

<table border="0" width="100%">
<thead>
	<tr><td colspan=4>Asset Search</tr></td>
</thead>
<tbody>
  <tr>
    <td width=15% nowrap align=right>Customer Name</td>
    <td width=20% align=left><INPUT type="text" name=txtcustomerName tabindex=1 style="HEIGHT: 20px; WIDTH: 220px" value="<%=strCustomerName%>"></td>
    <TD align=right nowrap width=15%>Region</TD>
	<TD align=left width=50%><SELECT id=selregion name=selregion tabindex=7 style="HEIGHT: 20px; WIDTH: 120px">
		<OPTION></OPTION>
			<%Do while Not objRS.EOF 
				Response.write "<OPTION VALUE ="& objRS("NOC_REGION_LCODE") & ">" & objRS("NOC_REGION_DESC") & "</OPTION>"
				objRS.MoveNext   
			Loop
 
			objRS.close
			set objRS = Nothing
			%></SELECT></TD>
    </TR> 
  <tr>
	<td width=15% nowrap align=right>Asset ID</td>
    <td width=20% align=left><INPUT type=text name=txtassetid tabindex=2 value="<%=Request.Cookies("AssetID")%>"></td>
	<td width=15% nowrap align=right>Tac Name</td>
    <td width=50% align=left><INPUT type="text" name=txttacname tabindex=8 value=""></td>
  </tr>  
  <tr>
	<td width=15% nowrap align=right>Asset Type</td>
    <TD width=20% align=left><SELECT id=selassettype name=selassettype tabindex=3 style="HEIGHT: 20px; WIDTH: 220px">
			<OPTION></OPTION>
			 <%Do while Not objRsAssetType.EOF 
				 Response.write "<OPTION VALUE ="""& objRsAssetType("ASSET_TYPE_ID") & """>" & routineHtmlString(objRsAssetType("ASSET_TYPE_DESC")) & "</OPTION>"
				objRsAssetType.MoveNext   
				Loop
 
				objRsAssetType.close
				set objRsAssetType = Nothing
			%></SELECT>
	</TD>
	<td width=15% nowrap align=right>Serial #</td>
    <td width=50%><INPUT type="text" name=txtserial tabindex=9 value=""></td>
    
  </tr>
  <tr>
	<td width=15% nowrap align=right>Asset Make</td>
    <td width=20% align=left><INPUT type="text" name=txtassetmake tabindex=4 value="<%=strMake%>"></td>
	<td width=15% nowrap align=right >CLLI Code</td>
    <td width=50% align=left><INPUT type="text" name=txtcllicode tabindex=10 value=""></td>
  </tr>
  <tr>
	<td width=15% nowrap align=right>Asset Model</td>
    <td width=20%><INPUT type="text" name=txtassetmodel tabindex=5 value="<%=strModel%>"></td> 
    <TD ALIGN=right width=15%>Active Only</TD>
	<TD ALIGN=LEFT><INPUT TYPE=CHECKBOX NAME="chkactive" tabindex=11 VALUE=YES CHECKED></TD>
  </tr>
  <tr>   
	<td width=15% nowrap align=right>TAC Asset Code</td>
    <td width=20% align=left><INPUT type="text" name=txtassetcode tabindex=6 value="<%=strAssetCode%>"></td> 
    <td align=right colspan=2>
    <% if strWinName <> "Popup" then %>
		<INPUT id=btnAdd name=btnAdd  tabindex=14 type=button style="width: 2cm" value=New style="HEIGHT: 24px; WIDTH: 65px" LANGUAGE=javascript onclick="return btnNew_click();">&nbsp;&nbsp;
	<% end if %>
		<INPUT id=btnClear name=btnClear tabindex=12 type=button style="width: 2cm" value=Clear style="HEIGHT: 24px; WIDTH: 62px" onClick="return btnClear_click();">&nbsp;&nbsp;
		<INPUT id=btnSearch name=btnSearch tabindex=13 type=submit style="width: 2cm" value=Search style="HEIGHT: 24px; WIDTH: 62px">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	</td>
  </tr>   
</tbody>		
</table>
</form>
</BODY>
</HTML>
