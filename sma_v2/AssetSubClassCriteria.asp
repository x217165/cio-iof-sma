<%@ LANGUAGE=VBSCRIPT   %>
<% option explicit      %>
<% on error resume next %>
<!-- #include file=smaConstants.inc -->
<!-- #include file=smaProcs.inc -->
<!-- #include file=databaseconnect.asp -->
<!--
********************************************************************************************
* Page name:	MakeCriteria.asp
* Purpose:		To dynamically set the criteria to search for an asset make.
*				Results are displayed via MakeList.asp
*	
* In Param:		This page reads following cookies
*				AssetClassType - the id of the asset class type
*				AssetClass - the id of the asset class type
*				AssetSubClassDesc
*				WinName
*
* Created by:	Chris Roe Oct. 04, 2000
*        29-Jul-15   PSmith  Set Cookies in validation so the back key works
*        05-Oct-15   PSmith  Only sumbit() for pop-up windows
*        03-Feb-16   PSmith  Don't pre-populate search criteria
********************************************************************************************
-->

<%
const COOKIE_TYPE  = "AssetClassType"
const COOKIE_CLASS = "AssetClass"
const COOKIE_DESC  = "AssetSubClassDesc"
const LIST_PAGE    = "AssetSubClassList.asp"
const DETAIL_PAGE  = "AssetSubClassDetail.asp"

'check user's rights
dim intAccessLevel
intAccessLevel = CInt(CheckLogon(strConst_AssetTypeClassification))
if (intAccessLevel and intConst_Access_ReadOnly) <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to asset subclasses. Please contact your system administrator."
end if

Dim strSql

Dim objRsType
strSql = " SELECT asset_class_type_id" &_
         " ,      asset_class_type_desc" &_
         " FROM   crp.asset_class_type" &_
         " ORDER BY asset_class_type_desc"
         
set objRSType = objConn.execute(strSql)

Dim objRsClass
strSql = " SELECT asset_class_id" &_
         " ,      asset_class_desc" &_
         " ,      asset_class_type_id" &_
         " FROM   crp.asset_class" &_
         " ORDER BY asset_class_desc"
Set objRSClass = objConn.execute(strSql)

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
setPageTitle("SMA - Asset Subclass");

//load all of the assetclass id's into an array so that we can dynamically populated the list box
var arrAssetClassOptions = new Array();

function fct_onLoad() 
{
  	var strWinName = document.frmSearch.hdnWinName.value ;
	  	
 		DeleteCookie("<%=COOKIE_TYPE%>");
 		DeleteCookie("<%=COOKIE_CLASS%>");
 		DeleteCookie("<%=COOKIE_DESC%>");
 		DeleteCookie("WinName");
 		
 		if (strWinName == "Popup" && ((document.frmSearch.txtDesc.value != "") || (document.frmSearch.selType.selectedIndex != 0) || (document.frmSearch.selClass.selectedIndex != 0)))
 		{
    	 SetCookie("<%=COOKIE_TYPE%>",document.frmSearch.selType.selectedIndex);
    	 SetCookie("<%=COOKIE_CLASS%>",document.frmSearch.selClass.selectedIndex);
    	 SetCookie("<%=COOKIE_DESC%>",document.frmSearch.txtDesc.value);
      thinking(parent.fraResult);
 			document.frmSearch.submit();
 		}
 	
	//load all of the class options into an array for
	for (var i = 1; i < document.frmSearch.selClass.options.length; i++)
	{
		//each array element holds:  assetClassID¿assetTypeId¿assetClassDesc
		arrAssetClassOptions[i - 1] = document.frmSearch.selClass.options[i].value + '¿' + document.frmSearch.selClass.options[i].text
	}
	
	selType_onChange();
}

function selType_onChange()
{
	var newOptions;   //these are the new options that are going to be placed in the drop down
	var lngTypeID = document.frmSearch.selType.options(document.frmSearch.selType.selectedIndex).value;
	//remove all old options (except the very first blank one) from the drop-down
	for (var i = document.frmSearch.selClass.options.length; i > 0; i--)
	{
		//alert('removing options' + document.frmSearch.selClass.options(i).value);
		document.frmSearch.selClass.options.remove(i);
	}
	


	//fill the array with the new options
	for (var j = 1; j < arrAssetClassOptions.length; j++)
	{
		var tmpElement;
		tmpElement = arrAssetClassOptions[j].split('¿');
		if ((tmpElement[1] == lngTypeID) || (lngTypeID == ""))
		{
			var tmpStr = "<OPTION value=\"" + tmpElement[0] + "¿" + tmpElement[1] + "\">" + tmpElement[2] + "</OPTION>";
			var tmpOption = document.createElement(tmpStr);
			document.frmSearch.selClass.options.add(tmpOption);
			tmpOption.innerText = tmpElement[2];
		}
	}

}
function fct_clear() 
{
	document.frmSearch.selType.selectedIndex = 0;
	document.frmSearch.selClass.selectedIndex = 0;
	document.frmSearch.txtDesc.value = "";
	selType_onChange();
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

	if (isWhitespace(theForm.txtDesc.value) && theForm.selClass.selectedIndex == 0 && theForm.selType.selectedIndex == 0)
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
<form name="frmSearch" id="frmSearch" action="<%=LIST_PAGE%>" method="post" target="fraResult" onsubmit="return validate(this);">

<INPUT name="hdnWinName"  type="hidden" value="<%=Request.Cookies("WinName")%>">

<table border="0" width="100%">
<tbody>
  <thead><tr><td colspan=5>Asset Subclass Search</td></tr></thead>
  <tr>
    <td width=20% align=right>Asset Class Type</td>
    <td width=25% align=left>
		<SELECT id="selType" name="selType" style="width: 7cm" onChange="selType_onChange();">
		<OPTION> </OPTION>
		<%
			Do while Not objRSType.EOF 
				Response.write "<OPTION "
				Response.Write " VALUE =""" & routineHTMLString(objRSType("ASSET_CLASS_TYPE_ID")) & """>" & routineHTMLString(objRSType("ASSET_CLASS_TYPE_DESC")) & "</OPTION>"
				objRSType.MoveNext   
			Loop
		%>
		</SELECT>
	</td>
  </tr>
  <tr>
    <td width=20% align=right>Asset Class</td>
    <td width=25% align=left>
		<SELECT id="selClass" name="Selclass" style="width: 7cm">
		<OPTION> </OPTION>
		<%
			Do while Not objRSClass.EOF 
				Response.write "<OPTION "
				Response.Write " VALUE =""" & routineHTMLString(objRSClass("ASSET_CLASS_ID") & strDelimiter & objRSClass("ASSET_CLASS_TYPE_ID")) & """>" & routineHTMLString(objRSClass("ASSET_CLASS_DESC")) & "</OPTION>"
				objRSClass.MoveNext   
			Loop
		%>
		</SELECT>
	</td>
	<td>&nbsp;</td>

  </tr>
  <tr>
    <td width=20% align=right>Asset Subclass</td>
    <td width=25% align=left><INPUT type="text" name="txtDesc" value="<%=Request.Cookies(COOKIE_DESC)%>" style="width: 7cm"></td>
    <td width=20% align=right>&nbsp;</td>
    <td width=25% align=left>Active Only<INPUT type="checkbox" name="chkActiveOnly" checked ></td>
	<td>&nbsp;</td>
  </tr>
  <tr>
    <td align=right colspan="5">
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
