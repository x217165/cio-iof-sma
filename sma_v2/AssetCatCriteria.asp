<%@ LANGUAGE=VBSCRIPT %>
<% option explicit %>

<!-- #include file=smaConstants.inc -->
<!-- #include file=smaProcs.inc -->
<!-- #include file=databaseconnect.asp -->
<!--
********************************************************************************************
        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
       29-Jul-15   PSmith  Set Cookies in validation so the back key works
       05-Oct-15   PSmith  Only sumbit() for pop-up windows
       03-Feb-16   PSmith  Don't pre-populate search criteria
**************************************************************************************************
-->
<%

'check user's rights
dim intAccessLevel

intAccessLevel = CInt(CheckLogon(strConst_AssetCatalogue))

if intAccessLevel < intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to asset catalogue. Please contact your system administrator."
end if




'retrieve the cookie variables

dim strAssetCatID, strAssetCatMake, strAssetCatModel, strAssetCatPartNumber, strWinName
	
strAssetCatID = Request.Cookies("AssetCatID")
strAssetCatMake = Request.Cookies("AssetCatMake")
strAssetCatModel = Request.Cookies("AssetCatModel")
strAssetCatPartNumber = Request.Cookies("AssetCatPartNumber")
strWinName	= Request.Cookies("WinName") 

%>


<HTML>
<HEAD>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<LINK href="stylesheets/styles.css" rel="stylesheet" type="text/css">

<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
<script type="text/javascript" SRC="AccessLevels.js"></script> 
<script type="text/javascript">


var intAccessLevel=<%=intAccessLevel%>;


//set section title
setPageTitle("SMA - Asset Catalogue");

function window_onload() {
	
	/************************************************************************************************
	*  Function:	window_onload																	*
	*																								*
	*  Purpose:		To submit the form automatically when values have been received from a cookie   *				
	*				and have been stored in hidden form controls.									*	
	*																								*			
	*  															
	*  Modified By: Shawn Myers  10/05/2000																								*
	*************************************************************************************************/
	
		var strAssetCatID, strAssetCatMake, strAssetCatModel, strAssetCatPartNumber, strWinName
		
		strAssetCatID = document.frmAssetCatSearch.hdnAssetCatID.value ;
	 	strAssetCatMake = document.frmAssetCatSearch.txtAssetCatalogueMake.value ;
	 	strAssetCatModel = document.frmAssetCatSearch.txtAssetCatalogueModel.value ;
	 	strAssetCatPartNumber = document.frmAssetCatSearch.txtAssetCataloguePartNumber.value ;
	 	strWinName = document.frmAssetCatSearch.hdnWinName.value ;
	 	
	 	//alert ('deleting cookies');
	 	
	 	//delete the associated cookies
	 	DeleteCookie("AssetCatID");
 		DeleteCookie("AssetCatMake");
 		DeleteCookie("AssetCatModel");
 		DeleteCookie("AssetCatPartNumber");
 		DeleteCookie("WinName");
 		
 		if ((strWinName ==  "Popup" )&&((strAssetCatID != "")||(strAssetCatMake != "")||(strAssetCatModel != "")||(strAssetCatPartNumber !=  "" ))){

			SetCookie("AssetCatID",document.frmAssetCatSearch.hdnAssetCatID.value);
			SetCookie("AssetCatMake",document.frmAssetCatSearch.txtAssetCatalogueMake.value);
			SetCookie("AssetCatModel",document.frmAssetCatSearch.txtAssetCatalogueModel.value);
			SetCookie("AssetCatPartNumber",document.frmAssetCatSearch.txtAssetCataloguePartNumber.value);
      thinking(parent.fraResult); 			
 			document.frmAssetCatSearch.submit();  
 		}	
	
		document.frmAssetCatSearch.txtAssetCatalogueMake.focus();
	
	}



function fct_clear() {

	document.frmAssetCatSearch.hdnAssetCatID.value = "";
	document.frmAssetCatSearch.txtAssetCatalogueMake.value = "";
	document.frmAssetCatSearch.txtAssetCatalogueModel.value = "";
	document.frmAssetCatSearch.txtAssetCataloguePartNumber.value = "";
}


function fct_addNew(){
	
	  if ((intAccessLevel & intConst_Access_Create)!= intConst_Access_Create) 
		{
			alert('Access denied. Please contact your system administrator.'); 
			return;
		}

		{
		
		parent.document.location.href ="AssetCatDet.asp?hdntxtAssetCatalogueID=0" ;
		}
	
	}
	
	
function validate(theForm){
	//**********************************************************************************************

	var bolConfirm ;
				
	if (isWhitespace(theForm.txtAssetCatalogueMake.value) && 
		    isWhitespace(theForm.txtAssetCatalogueModel.value) &&
		    isWhitespace(theForm.txtAssetCataloguePartNumber.value)) 
		     	
		 {
		   bolConfirm = window.confirm("No search criteria have been entered. This search may take a long time...Continue?")
		    if (!bolConfirm){
			 // abort search
		     return false;			
		    }
		  }

  thinking(parent.fraResult);

		  // search critiera have been entered so continue search
		  return true ;				
	}	

</script>

</HEAD>

<BODY Language=javascript onLoad="return window_onload()">
<form name="frmAssetCatSearch" action="AssetCatList.asp" method=post target="fraResult" onsubmit="return validate(this);">

<!--gets its value from the "WinName" cookie -->

<INPUT name="hdnWinName" type="hidden" value="<%=strWinName%>">

<!--gets its value from the "AssetCatID" cookie -->
<INPUT name="hdnAssetCatID" type="hidden" value="<%=strAssetCatID%>">

<table border="0" width="100%">

<tbody>
	
	<thead><tr><td colspan=4>Asset Catalogue Search</td></tr></thead>
   
   <tr>
    <td width=15% align=right>Make&nbsp;</td>
    <td width= 20% align=left><INPUT type="text" name="txtAssetCatalogueMake" value="<%=strAssetCatMake%>"></td>
    <td width=15%>&nbsp</td>
    <td>&nbsp</td>
   </tr>
   
   <tr>
    <td width=15% align=right>Model&nbsp;</td>
    <td width=20% align=left><INPUT type="text" name="txtAssetCatalogueModel" value="<%=strAssetCatModel%>"></td>
   </tr>
   
   <tr>
    <td width=15% align=right>Part Number&nbsp;</td>  
    <td width=20%><INPUT type="text" name="txtAssetCataloguePartNumber" value="<%=strAssetCatPartNumber%>"></td>
   </tr>
  
   <tr>
    <td align=right colspan="4">
        
        <% if strWinName <> "Popup" then %>
				<input name=btnNew type=button value=New style="width: 2cm" tabindex=10 style="HEIGHT: 24px; WIDTH: 62px" onClick="fct_addNew();">&nbsp;&nbsp;
		<% end if %>
		
		<INPUT name=btnClear type=button style="width: 2cm" value=Clear onClick="fct_clear()">&nbsp;&nbsp;
		<INPUT name=btnSubmit type=submit style="width: 2cm" value=Search>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    
    </td>
  </tr>
<tbody>
</tbody>
</table>

</form>

</BODY>
</HTML>
