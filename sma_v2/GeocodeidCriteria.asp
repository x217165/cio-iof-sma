<%@ Language=VBScript %>
<% option explicit %>
<% Response.Buffer = true %>

<!--#include file="smaConstants.inc"-->
<!--#include file="smaProcs.inc"-->
<!--#include file = "databaseconnect.asp"-->
<!--
********************************************************************************************
* Page name:	AddressCriteria.asp
* Purpose:		To dynamically set the criteria to search for an address.
*				Results are displayed via AddressList.asp
*
* In Param:		This page reads following cookies
*				CustomerName
*				WinName
*
* Created by:	Sara Sangha	Aug. 14th, 2000
*
********************************************************************************************
        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
       14-Nov-01	  DTy		Turn buffer on.
       24-Feb-02	  DTy		Simplified Search Screen.
***************************************************************************************************
-->
<%

dim intAccessLevel
intAccessLevel = CInt(CheckLogon(strConst_Address))
if (intAccessLevel and intConst_Access_ReadOnly) <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to Address. Please contact your system administrator"
end if

dim objRsProvince, strSQL
dim strGeocodeid, strWinName

strGeocodeid = Request.cookies("Geocodeid")
strWinName	= Request.Cookies("WinName")

strSQL = "select distinct province as province from crp.lcode_geocodeid "

set objRsProvince = objConn.Execute(StrSql)


%>

<HTML>
<HEAD>
	<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<LINK rel="stylesheet" type="text/css" href="stylesheets/styles.css">
	<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
	<script type="text/javascript" SRC="AccessLevels.js"></script>
	<SCRIPT type = "text/javascript">

//*************************************Java Functions*******************************************

var intAccessLevel = "<%=intAccessLevel%>" ;

//set section title
setPageTitle("SMA - CLLI Code");

function validate(theForm){
//***********************************************************************************************
// Function:	validate()																		*
//																								*
// Purpose:		To validate that minimum search critiera has been entered so that full database *
//				search can be avoided.															*
//																								*
// Created By:	Sara Sangha		Aug. 28th, 2000													*
//																								*
// Updated By:																					*
//***********************************************************************************************

var bolConfirm
   if(isWhitespace(theForm.txtGeocodeid.value) &&
    isWhitespace(theForm.txtAddress.value) &&
    isWhitespace(theForm.txtCity.value) &&
    isWhitespace(theForm.txtPostal.value) &&
    (theForm.selProvince.selectedIndex == 0) &&
    isWhitespace(theForm.txtCllicode.value)  &&
    isWhitespace(theForm.txtDescription.value)
 {
   bolConfirm = window.confirm("No search criteria have been entered. This search may take a long time...Continue?");
    if (bolConfirm){
	  // continue search even no creitiera has been entered
      return true;
    }
    else
    {
	 // abort search
     return false;
    }
  }
   // search critiera has been entered so continue search
   return true;
 }

function window_onload() {
/************************************************************************************************
*  Function:	window_onload																	*
*																								*
*  Purpose:		To submit the form automatically when txtCustomerName has a value.				*
*				This will happen when this page is called by a lookup function or by the Quick	*
*    		    Navigation drop-down box, which had already saved some values in the cookies	*
* 				and this form's VBScript has read those values and put in the form controls.	*
*																								*
*  Created By:	Sara Sangha Aug 25, 2000														*
*																								*
*  Updated By:																					*
*************************************************************************************************/


		var strGeocodeid;
	 	var strWinName;
	 	var hdnCustomerID;

	 	strGeocodeid = document.frmAddSearch.txtGeocodeid.value ;
	 	strWinName = document.frmAddSearch.hdnWinName.value ;
	 	hdnCustomerID = GetCookie("CustomerID");

	 	if (strWinName == "Simple"){
	 	   SetCookie("strSimple", "Simple");
	 	   SetCookie("txtGeocodeid", strGeocodeid);
	 	   SetCookie("CustomerID", hdnCustomerID);}
	 	else
	 	{DeleteCookie("strSimple")}

	 	DeleteCookie("Geocodeid");
 		DeleteCookie("WinName") ;
	 	if (strGeocodeid != "") {
 			document.frmAddSearch.submit() ;
 		}
}
function btnClear_onclick() {

	document.frmAddSearch.txtGeocodeid.value = "" ;
	document.frmAddSearch.txtAddress.value = "" ;
	document.frmAddSearch.txtCity.value = ""  ;
	document.frmAddSearch.selProvince.selectedIndex = 0 ;
	document.frmAddSearch.txtPostal.value = "";
	document.frmAddSearch.txtDescription = "";
	document.frmAddSearch.chkActiveOnly.checked=true;
	document.frmAddSearch.txtlCllicode.value = "" ;
}

//********************************* end of java functions******************************************
//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">
<FORM name=frmAddSearch method=post action="GeocodeidList.asp" target="fraResult" onSubmit="return validate(this)">
	<!-- hidden fields -->
	<INPUT name=hdnWinName type=hidden value="<%=strWinName%>">

<TABLE width="96%">
	<thead>
    <% if strWinName = "Simple" then %>
  	   <TR><TD align=left colspan=3>Simplified CLLI Code Search</td></TR>
    <% else %>
  	   <TR><TD align=left colspan=4>Full CLLI Code Search</td></TR>
    <% end if %>
	</thead>
    <TR>
        <TD align=right width="20%">GEOCODE ID</TD>
        <TD align=left width="22%"><INPUT id=txtGeocodeid name=txtGeocodeid tabindex=1 size=30 value="<%=strGeocodeid%>"></TD>

        <% if strWinName <> "Simple" then %>
           <TD align=right width="18%">CLLI Code</TD>
           <TD width="33%">
			<INPUT id=txtCllicode name=txtCllicode size=30 tabindex=4></td>
        <% end if %>

    <TR>
        <TD align=right width="20%">ADDRESS</TD>
        <TD align=left width="22%"><INPUT id=txtAddress name=txtAddress tabindex=2 size=30 ></TD>

        <% if strWinName <> "Simple" then %>
              <TD align=right width="18%">DESCRIPTION:</TD>
              <TD width="33%">
				<INPUT id=txtDescription name=txtDescription size=30 tabindex=4>
		      </TD>
        <% end if %>
		</TR>

    <TR>
        <TD align=right width="20%">CITY </TD>
        <TD width="22%"><INPUT id=txtCity name=txtCity size=30 tabindex=3></td>
        <td width="18%">&nbsp;</td>

        <% if strWinName <> "Simple" then %>
		<td width="33%">&nbsp;</TD>
        <% end if %>

	</TR>

    <% if strWinName <> "Simple" then %>
       <TR>
       <TD align=right>POSTAL CODE</TD>
       <TD><INPUT id=txtPostal name=txtPostal size=30 tabindex=4></TD>
       <TD align=right width="18%">&nbsp;</TD>
       <TD width="33%">&nbsp;</TD>
       </TR>
    <% end if %>

          <TR>
          <TD align=right width="20%">PROVINCE </td>
          <TD align=left width=25%><SELECT name=selProvince tabindex=5>
			<OPTION></OPTION>
			<%  objRsProvince.MoveFirst
				Do while Not objRsProvince.EOF %>
				<OPTION> <% response.write(objRsProvince(0))  %> </OPTION>
				<%objRsProvince.MoveNext
				Loop
			%>
			</SELECT>
        </TD>

          </TD>
          <TD>&nbsp;</TD>


       <TD colSpan=2 align=center>
    <% if strWinName <> "Popup" then %><INPUT id=btnClear name=btnClear type=button style="width: 2cm" value=Clear LANGUAGE=javascript onclick="return btnClear_onclick()" tabindex=10>&nbsp;

      &nbsp;&nbsp;
    <% end if %>&nbsp;  &nbsp;&nbsp;
      <INPUT id=btnSearch name=btnSearch type=submit style="width: 2cm" value=Search tabindex=11> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      </TD>
    </TR>
    <TR>
       <TD colSpan=1 align=right>
       <TD align=right width="22%">Active Only</TD>
       <TD width="18%"><INPUT id=chkActiveOnly name=chkActiveOnly type=checkbox VALUE=yes checked tabindex=8></TD>
    </TR>

</TABLE>
</FORM>
</BODY>
</HTML>
