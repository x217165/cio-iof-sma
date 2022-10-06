<%@ Language=VBScript %>
<% Option Explicit %>
<!--#include file="smaConstants.inc"-->
<!--#include file="smaProcs.inc"-->
<!--#include file="databaseconnect.asp"-->
<!--
********************************************************************************************
* Page name:	ServLocCriteria.asp
* Purpose:		To dynamically set the criteria to search for service locations.
*				Results are displayed via CustList.asp
*
* In Param:		This page reads following cookies
*				Cookie - WinName
*				Cookie - AddressID
*				Cookie - ServLocName
*				Cookie - CustomerName
*				Cookie - IncludeTelus - must be set to "yes" or "no"
*
* Out Param:	None
*
* Created by:	Nancy Mooney Aug. 8th, 2000
*
********************************************************************************************
-->
<%
Dim intAccessLevel
intAccessLevel = CInt(CheckLogon(strConst_ServiceLocation))

dim objRsProvince, strSQL

	strSQL = "select s.province_state_lcode, s.province_state_name " &_
		 "from crp.lcode_province_state s, " &_
				"crp.lcode_country c " &_
		 "where s.record_status_ind = 'A' " &_
		 "and	  s.country_lcode = c.country_lcode " &_
		 "order by s.country_lcode, s.province_state_name "

	set objRsProvince = objConn.Execute(StrSql)

	'retrieve cookie values
dim  strWinName, strCustomerName, strAddressID, strServLocName, strCityName, strProvinceName,strServiceEnd, strIncludeTelus, strStreet

	strWinName = Request.Cookies("WinName")
	strCustomerName = Request.Cookies("CustomerName")
	strAddressID = Request.Cookies("AddressID")
	strProvinceName = Request.Cookies("ProvinceName")
	strCityName = Request.Cookies("CityName")
	strServLocName = Request.Cookies("ServLocName")
	strServiceEnd = Request.Cookies("ServiceEnd")
	strIncludeTelus = LCase(Request.Cookies("IncludeTelus"))
	strStreet = Request.Cookies("Street")

	if strServLocName = "new" then
		strServLocName = ""
	end if

	if strIncludeTelus = "" or strIncludeTelus <> "yes" then
		strIncludeTelus = "no"
	end if
%>

<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<LINK rel="stylesheet" type="text/css" href="stylesheets/styles.css">
<SCRIPT type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
<SCRIPT type="text/javascript" SRC="AccessLevels.js"></script>
<SCRIPT type = "text/javascript">

<!-- hide script from old browsers
//********************************** Java Functions *****************************************
var intAccessLevel = <%=intAccessLevel%>;

//set section title
try
	{window.parent.PageTitle.value = "SMA - Service Location"}
catch(e) //do nothing, don't need to set up title when in Lookup mode
	{}

function window_onLoad() {

	var strWinName, strServLocName, strCustomerName, strAddressID, strCityName, strStreet;

	strWinName = document.frmServLocCriteria.hdnWinName.value;
	strServLocName = document.frmServLocCriteria.txtServiceLocationName.value;
	strCustomerName = document.frmServLocCriteria.txtCustomerName.value;
	strAddressID = document.frmServLocCriteria.hdnAddressID.value;
	strCityName = document.frmServLocCriteria.txtCity.value;
	strStreet = document.frmServLocCriteria.txtStreetName.value;

	DeleteCookie("WinName");
	DeleteCookie("ServLocName");
	DeleteCookie("ProvinceName");
	DeleteCookie("AddressID");
	DeleteCookie("CustomerName");
	DeleteCookie("CityName");
	DeleteCookie("Street");
	DeleteCookie("IncludeTelus");

	if ((strServLocName != "")||(strCustomerName != "")||(strAddressID != "")||(strCityName != "")){
		document.frmServLocCriteria.submit();
	}
}

function validate(theForm){

	var bolConfirm ;

	if (isWhitespace(theForm.txtCustomerName.value)
	    && isWhitespace(theForm.txtCity.value)
	    && isWhitespace(theForm.txtServiceLocationName.value)
	    && isWhitespace(theForm.txtSpecificLocationDesc.value)
	    && isWhitespace(theForm.txtStreetName.value)
	    && (theForm.selProvince.selectedIndex == 0))
	{
	   bolConfirm = window.confirm("No Search Criteria have been entered. This search may take a long time..Continue?");
	    if (bolConfirm){
	      return true;
	    }
	    else
	    {
	     return false;
	    }
	  }
	   return true;
}

function btnAddNew_onclick()
{
	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create)
	{
		alert('Access denied.  Please contact your system administrator.');
		return fale;
	}

	parent.document.location.href ="ServLocDetail.asp?NewServLoc=NEW";
}

function fct_Clear(){
	//clear the hidden variables
	//document.frmServLocCriteria.hdnWinName.value="";
	//document.frmServLocCriteria.hdnAddressID.value="";
	//clear input areas
	document.frmServLocCriteria.txtCustomerName.value="";
	document.frmServLocCriteria.txtStreetName.value="";
	document.frmServLocCriteria.txtServiceLocationName.value="";
	document.frmServLocCriteria.txtCity.value="";
	document.frmServLocCriteria.txtSpecificLocationDesc.value="";
	document.frmServLocCriteria.selProvince.value="";
	document.frmServLocCriteria.chkActiveOnly.checked=true;
}
//-->end hide script
//************************************End of Java Functions **********************************
</SCRIPT>

</HEAD>
<BODY LANGUAGE=javascript onload="window_onLoad();" >
<FORM name="frmServLocCriteria" method="post" action="ServLocList.asp" target="fraResult" onSubmit="return validate(this);">
	<!--hidden variable-->
	<INPUT name="hdnWinName" type="hidden" value="<%=strWinName%>">
	<INPUT name=hdnAddressID type=hidden value ="<%=strAddressID%>">
	<INPUT id=hdnServiceEnd name=hdnServiceEnd type=hidden value="<%=strServiceEnd%>">
<TABLE >
	<thead><tr><td align=left colspan=4>Service Location Search</td></tr></thead>
	<tbody>
    <TR>
		<TD align=right width=20%>Customer</TD>
        <TD align=left width=20%><INPUT name=txtCustomerName tabindex=1 size=35 maxlength=50 value="<%=routineHTMLString(strCustomerName)%>"></TD>
		<TD align=right width =15%>Street</TD>
        <TD align=left ><INPUT name="txtStreetName" tabindex=5 size=25  value="<%=routineHTMLString(strStreet)%>"></TD>
    </TR>
    <TR>
		<TD align=right width=20%>Service Location</TD>
        <TD align= left width=20%><INPUT name="txtServiceLocationName" tabindex=2 size=35 maxlength=50 value="<%=routineHTMLString(strServLocName)%>"></td>
		<TD align=right width = 15%>City</TD>
        <TD align = left ><INPUT id=txtCity name=txtCity tabindex=6 size=25 value="<%=routineHTMLString(strCityName)%>"></td>
    <TR>
		<TD align=right width=20%>Specific Loc Desc</TD>
        <TD align=left width=20%><INPUT id=txtSpecificLocationDesc name=txtSpecificLocationDesc tabindex=3 size=35 maxlength=80></td>
        <TD align=right width=15%>Province</TD>
        <TD align=left >
			<SELECT id=selProvince name=selProvince tabindex=7 style="width: 190px">
				<OPTION value=""></OPTION>
        		<%Do while Not objRsProvince.EOF
        			If objRsProvince(0) = strProvinceName then
        				Response.write "<OPTION selected VALUE=" & routineHTMLString(objRsProvince(0)) &  ">" & routineHTMLString(objRsProvince(0)) & "&nbsp;&nbsp;&nbsp;" & routineHTMLString(objRsProvince(1)) & "</OPTION>"
        			else
						Response.write "<OPTION VALUE=" & routineHTMLString(objRsProvince(0)) &  ">" & routineHTMLString(objRsProvince(0)) & "&nbsp;&nbsp;&nbsp;" & routineHTMLString(objRsProvince(1)) & "</OPTION>"
					end if
					objRsProvince.MoveNext
					Loop
				%>
			</SELECT>
		</td>
    <TR>
		<TD width=20% align="right">Include Telus Locations</TD>
		<TD width=20%><input id=chkIncludeTelus name=chkIncludeTelus tabindex=4 type=checkbox value=YES <% if strIncludeTelus = "yes" then Response.Write "checked" end if%>></TD>
		<TD align=right width=15%>Active Only</TD>
        <TD align=left ><INPUT id=chkActiveOnly name=chkActiveOnly tabindex=8 type=checkbox value=YES checked ></TD>
    </TR>
	<TR>
        <TD COLSPAN=4 ALIGN=RIGHT>
			<% if strWinName <> "Popup" then %>
				<INPUT id=btnAddNew name=btnAddNew type=button style="HEIGHT: 24px; WIDTH: 62px" value=New LANGUAGE=javascript onclick="return btnAddNew_onclick()" tabindex=9 >&nbsp;&nbsp;
			<% end if %>
			<INPUT id=btnClear name=btnClear type=button style="HEIGHT: 24px; WIDTH: 62px" value=Clear onClick="fct_Clear();" tabindex=10> &nbsp;&nbsp;
			<INPUT id=btnSearch name=btnSearch type=submit style="HEIGHT: 24px; WIDTH: 62px" value=Search tabindex=11>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD></TR>


		</TD>
	</TR>
    </tbody>
</TABLE>
</FORM>
</BODY>
</HTML>
