
<%@ LANGUAGE=VBSCRIPT %>
<% option explicit%>
<%on error resume next %>

<!-- #include file=smaConstants.inc -->
<!-- #include file=smaProcs.inc -->
<!-- #include file=databaseconnect.asp -->
<!--
********************************************************************************************
* Page name:	AssetClassCriteria.asp
*
* Purpose:		To dynamically set the criteria to search for an asset make.
*				Results are displayed via AssetClassList.asp
*	
* In Param:		This page reads following cookies
*				AssetClassTypeID - the id of the asset class type
*				AssetClassDesc - the id of the asset class type
*				WinName
*
* Created by:	Shawn Myers Oct. 17, 2000
*        29-Jul-15   PSmith  Set Cookies in validation so the back key works
*        05-Oct-15   PSmith  Only sumbit() for pop-up windows  
*        03-Feb-16   PSmith  Don't pre-populate search criteria
********************************************************************************************
-->

<%
const COOKIE_AC_DESC   = "AssetClassDesc"
const COOKIE_AC_TYPEID = "AssetClassTypeID"

'check user's rights
dim intAccessLevel


intAccessLevel = CInt(CheckLogon(strConst_AssetTypeClassification))
if (intAccessLevel and intConst_Access_ReadOnly) <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to asset class. Please contact your system administrator."
end if

dim objRsTypeID, strSQL, objCmd				

'Load the asset class type drop down box
strSQL = " SELECT asset_class_type_id, " &_
		 "        asset_class_type_desc, " &_
		 "		  record_status_code " &_
		 "FROM   crp.asset_class_type"
		 
set objCmd = Server.CreateObject("ADODB.Command")
objCmd.ActiveConnection = objconn
objCmd.CommandText = strSQL  
objCmd.CommandType = adCmdText
	  
set objRsTypeID = Server.CreateObject("ADODB.Recordset")
objRsTypeID.CursorLocation = adUseClient
objRsTypeID.Open strSQL, objConn
		
if err then
	DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
end if

if objRsTypeID.EOF then 
	DisplayError "BACK", "", 999, "CANNOT FIND REQUESTED OBJECT", "EOF condition occurred in objRsTypeID recordset."
end if

%>


<HTML>
<HEAD>
<LINK href="stylesheets/styles.css" rel="stylesheet" type="text/css">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
<script type="text/javascript" SRC="AccessLevels.js"></script>
<script type="text/javascript">


var intAccessLevel = <%=intAccessLevel%>;

//set section title
setPageTitle("SMA - Asset Class");

//OK
function fct_onLoad() {

 	if (document.frmAssetClassSearch.hdnWinName.value ==  "Popup")
 	
 	{
 		
 		DeleteCookie("<%=COOKIE_AC_DESC%>");
 		DeleteCookie("<%=COOKIE_AC_TYPEID%>");
 		DeleteCookie("WinName");
 		
 	}	

	if (document.frmAssetClassSearch.hdnWinName.value == "Popup" && ((document.frmAssetClassSearch.txtAssetClassDesc.value !=  "" ) ||(document.frmAssetClassSearch.selAssetClassType.selectedIndex != 0)))
	{
 		{
    	 SetCookie("<%=COOKIE_AC_DESC%>",document.frmAssetClassSearch.txtAssetClassDesc.value);
    	 SetCookie("<%=COOKIE_AC_TYPEID%>",document.frmAssetClassSearch.selAssetClassType.selectedIndex);
      thinking(parent.fraResult);
 			document.frmAssetClassSearch.submit();
 		}
 	}
 		

}

//OK
function fct_clear() 
{
	document.frmAssetClassSearch.txtAssetClassDesc.value = "";
	document.frmAssetClassSearch.selAssetClassType.selectedIndex = 0;
}

//OK
function btnNew_onclick() 
{

	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create)
	{
		alert('Access denied. Please contact your system administrator.');
		return false;
	}
	
	parent.document.location.href ="AssetClassDetail.asp?hdnAssetClassID=0" ;
	
}


//OK


function validate(theForm){

	var bolConfirm ;

	if (isWhitespace(theForm.txtAssetClassDesc.value) &&
					(theForm.selAssetClassType.selectedIndex == 0 ))
	{
	   bolConfirm = window.confirm("No Search Criteria have been entered. This search may take a long time..Continue?");
	    if (!bolConfirm){
	     return false;
	    }
	  }

     thinking(parent.fraResult);

	   return true;
}
</script>



</HEAD>

<BODY onLoad="fct_onLoad();">

<form name="frmAssetClassSearch" action="AssetClassList.asp" method="post" target="fraResult" onsubmit="return validate(this);">

<INPUT name="hdnWinName"  type="hidden" value="<%=Request.Cookies("WinName")%>">

<table border="0" width="100%">

<tbody>
	
  <thead>
	<tr>
	<td colspan=4>Asset Class Search</td>
	</tr>
  </thead>
  
  
  <TR>
   <TD align=right width=15%>Asset Class Type</TD>
        
        <TD align=left width=25%><SELECT name=selAssetClassType  style="width: 7cm">
			<OPTION></OPTION>
			
			<%
			Do while Not objRsTypeID.EOF 
				Response.write "<OPTION VALUE=" & objRsTypeID(0) 
        If Cint(objRsTypeID(0)) = Cint(Request.Cookies(COOKIE_AC_TYPEID)) - 1 then
        	Response.write " selected "
        End If
        Response.write ">" & objRsTypeID(1) & "</OPTION>"
				objRsTypeID.MoveNext   
			Loop
			%>	
			</SELECT> 
        </TD>
    <td width=15%>&nbsp;</td>
    <td>&nbsp;</td>
  </TR>
  
  <tr>
    <td width=15% align=right>Asset Class</td>
    <td width= 20% align=left><INPUT type="text" name="txtAssetClassDesc" value="<%=Request.Cookies(COOKIE_AC_DESC)%>" style="width: 7cm"></td>
  	<TD width=15% align=right>Active Only</TD>
	<TD align=left width=27px><INPUT id=chkActiveOnly name=chkActiveOnly type=checkbox value="yes" checked></TD>

  </tr>
  
  <TR>  
    <td align=right colspan="4">
		<% if Request.Cookies("WinName") <> "Popup" then %>
			<INPUT id=btnAddNew name=btnAddNew type=button style="HEIGHT: 24px; WIDTH: 62px" value=New LANGUAGE=javascript onclick="return btnNew_onclick()" >&nbsp;&nbsp;
		<% end if %>
		<INPUT name=btnClear type=button style="width: 2cm" value=Clear onClick="fct_clear()">&nbsp;&nbsp;
		<INPUT name=btnSubmit type=submit style="width: 2cm" value=Search>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    </td>
  </TR>
</table>
</form>
</BODY>
</HTML>


