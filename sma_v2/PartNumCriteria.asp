<%@ LANGUAGE=VBSCRIPT %>
<% option explicit
on error resume next %>
<!-- #include file=smaConstants.inc -->
<!-- #include file=smaProcs.inc -->
<!-- #include file=databaseconnect.asp -->
<!--
********************************************************************************************
* Page name:	PartNumCriteria.asp
* Purpose:		To dynamically set the criteria to search for an asset part number.
*				Results are displayed via PartNumList.asp
*
* In Param:		This page reads following cookies
*				PartNumDesc
*				WinName
*
* Created by:	Chris Roe Oct. 04, 2000
*        29-Jul-15   PSmith  Set Cookies in validation so the back key works
*        05-Oct-15   PSmith  Only sumbit() for pop-up windows
*        03-Feb-16   PSmith  Don't pre-populate search criteria
********************************************************************************************
-->

<%
const COOKIE_DESC = "PartNumDesc"
const LIST_PAGE   = "PartNumList.asp"
const DETAIL_PAGE = "PartNumDetail.asp"

'check user's rights
dim intAccessLevel
intAccessLevel = CInt(CheckLogon(strConst_AssetCatalogue))
if (intAccessLevel and intConst_Access_ReadOnly) <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to asset part numbers. Please contact your system administrator."
end if
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<LINK href="stylesheets/styles.css" rel="stylesheet" type="text/css">
<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
<script type="text/javascript" SRC="AccessLevels.js"></script>
<script type="text/javascript">

var intAccessLevel = <%=intAccessLevel%>;

//set section title
setPageTitle("SMA - Part Number");

function fct_onLoad()
{
 		DeleteCookie("<%=COOKIE_DESC%>");
 		DeleteCookie("WinName");

    var strWinName = document.frmSearch.hdnWinName.value;

 		if ( strWinName == "Popup" && (document.frmSearch.txtDesc.value != ""))
 		{
    	 SetCookie("<%=COOKIE_DESC%>",document.frmSearch.txtDesc.value);
      thinking(parent.fraResult);
 			document.frmSearch.submit();
 		}
}

function fct_clear() {

	document.frmSearch.txtDesc.value = "";
}

function btnNew_onclick()
{

	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create)
	{
		alert('Access denied. Please contact your system administrator.');
		return false;
	}

	parent.document.location.href ="<%=DETAIL_PAGE%>?NewRecord=NEW" ;

}
function validate(theForm){

	var bolConfirm ;

	if (isWhitespace(theForm.txtDesc.value))
	{
	   bolConfirm = window.confirm("No Search Criteria have been entered. This search may take a long time..Continue?");
	    if (!bolConfirm){
	     return false;
	    }
	  }

  // Start thinking
  thinking(parent.fraResult);

	   return true;
}
</script>

</HEAD>
<BODY onLoad="fct_onLoad();">
<form name="frmSearch" action="<%=LIST_PAGE%>" method="post" target="fraResult" onsubmit="return validate(this);">

<INPUT name="hdnWinName"  type="hidden" value="<%=Request.Cookies("WinName")%>">

<table border="0" width="100%">
<tbody>
	<thead><tr><td colspan=4>Part Number Search</td></tr></thead>
  <tr>
    <td width=15% align=right>Part Number</td>
    <td width= 20% align=left><INPUT type="text" name="txtDesc" value="<%=Request.Cookies(COOKIE_DESC)%>"></td>
    <td width=15%>&nbsp</td>
    <td>&nbsp</td>
  </tr>
  <tr>
    <td align=right colspan="4">
		<% if Request.Cookies("WinName") <> "Popup" then %>
			<INPUT id=btnAddNew name=btnAddNew type=button style="HEIGHT: 24px; WIDTH: 62px" value=New LANGUAGE=javascript onclick="return btnNew_onclick()" >&nbsp;&nbsp;
		<% end if %>
		<INPUT name=btnClear type=button style="width: 2cm" value=Clear onClick="fct_clear()">&nbsp;&nbsp;
		<INPUT name=btnSubmit type=submit style="width: 2cm" value=Search>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    </td>
  </tr>
<tbody></tbody>
</table>
</form>
</BODY>
</HTML>
