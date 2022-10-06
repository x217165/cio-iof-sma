<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<!--#include file="smaConstants.inc"-->
<!--#include file="databaseconnect.asp"-->
<!-- #include file="kenanconnect.asp" -->
<!--#include file="smaProcs.inc"-->
<!--
*****************************************************************************************
* File Name:	SKenanACriteria.asp
* Author:		Linda chen
* Purpoase:		Start Service Kenan Attribute Search
* Date:			August 2009
******************************************************************************************
-->
<%
Dim strWinName, strBusinessID, strServiceCategoryID, strSTypeID, strSTypeDesc, strServiceKenanID, lIndex, strLANG
Dim lRow, arrLOBList, pfRow, arrSCategoryList, arrSCategoryListEN, arrRevenueList
Dim objRS, strSQL, strWhereClause
Dim intAccessLevel
Dim objRsSTKenanCP, objRsSTKenanPK, strken

	intAccessLevel = CInt(CheckLogon(strConst_ServiceCatalogue))
	If (intAccessLevel And intConst_Access_ReadOnly) <> intConst_Access_ReadOnly Then
		DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to Service Category. Please contact your system administrator"
	End If

	'strWinName = Request.Cookies("WinName")
	'strBusinessID = Request.Cookies("BusinessID")
	'strServiceCategoryID = Request.Cookies("ServiceCategoryID")
	'strServiceKenanID = Request.Cookies("ServiceKenanAID")
	'strSTypeID = Request.Cookies("ServiceType")
	'strSTypeDesc = Request.Cookies("STypeDesc")

'TQ_INOSS
	'strLANG = Request.Cookies("UserInformation")("language_preference")
	'if (Len(strLANG) = 0) then strLANG = "EN"
	strLANG = "EN"

	strSQL =" SELECT LOB_ID, LOB_CODE, LOB_DESC " &_
			" FROM CRP.V_LOB " &_
			" WHERE lob_id NOT IN " &_
		    " (SELECT lob_id " &_
		    " FROM crp.v_lob " &_
		    " WHERE language_preference_lcode = '" & strLANG & "' ) " &_
			" AND LANGUAGE_PREFERENCE_LCODE = 'EN' " &_
			" AND RECORD_STATUS_IND = 'A'" &_
			" UNION SELECT LOB_ID, LOB_CODE, LOB_DESC " &_
			" FROM crp.v_lob " &_
			" WHERE language_preference_lcode = '" & strLANG & "' " &_
			" AND RECORD_STATUS_IND = 'A' " &_
			" ORDER BY LOB_DESC "

	'Create Recordset object
	Set objRS = Server.CreateObject("ADODB.Recordset")
	objRS.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
	arrLOBList = objRS.GetRows

	'Get SERVICE_CATEGORY
	'strLANG = Request.Cookies("UserInformation")("language_preference")
	'if (Len(strLANG) = 0) then strLANG = "EN"

	strSQL = "SELECT SERVICE_CATEGORY_ID, LOB_ID, SERVICE_CATEGORY_DESC, LANGUAGE_PREFERENCE_LCODE " &_
			"FROM CRP.V_SERVICE_CATEGORY "&_
			"WHERE SERVICE_CATEGORY_ID NOT IN" &_
			      	"(SELECT SERVICE_CATEGORY_ID " &_
		      		"FROM CRP.V_SERVICE_CATEGORY " &_
			      	"WHERE language_preference_lcode = '" & strLANG & "' ) " &_
			"AND LANGUAGE_PREFERENCE_LCODE = 'EN'" &_
			"AND RECORD_STATUS_IND = 'A'" &_
			"UNION SELECT SERVICE_CATEGORY_ID, LOB_ID, SERVICE_CATEGORY_DESC, LANGUAGE_PREFERENCE_LCODE " &_
			"FROM CRP.V_SERVICE_CATEGORY " &_
			"WHERE language_preference_lcode = '" & strLANG & "'" &_
			"AND RECORD_STATUS_IND = 'A' " &_
			"ORDER BY SERVICE_CATEGORY_DESC"

	'Create Recordset object
	Set objRS = Server.CreateObject("ADODB.Recordset")
	objRS.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
	arrSCategoryList = objRS.GetRows

	' Get English versions
	strSQL = " SELECT service_category_id, service_category_desc " &_
			 " FROM crp.service_category " &_
			 " WHERE record_status_ind = 'A' "

	Set objRS = Server.CreateObject("ADODB.Recordset")
	objRS.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
	arrSCategoryListEN = objRS.GetRows

'	Response.Write (arrSCategoryList(0,0) & " " & arrSCategoryListEN(0,0))
'	Response.End

	for lRow = 0 to UBound(arrSCategoryList, 2)
		if (strcomp(arrSCategoryList(3,lRow), "EN") <> 0 ) then
			' For non-English entries only, append the English
			for pfRow = 0 to UBound(arrSCategoryListEN, 2)
				if (strcomp(arrSCategoryList(0,lRow), arrSCategoryListEN(0,pfRow)) = 0) then
					if (strcomp(arrSCategoryList(2,lRow), arrSCategoryListEN(1,pfRow)) <> 0) then
						' Append English version to the end of any translated service categories
						arrSCategoryList(2,lRow) = arrSCategoryList(2,lRow) & " (" & arrSCategoryListEN(1,pfRow) & ")"
						pfRow = UBound(arrSCategoryListEN, 2) + 1 ' Done: Skip out of the loop early
					end if
				end if
			next
		end if
	next

	'Get the Service Instance Name
'	StrSql = "SELECT SRVC_INSTNC_ATT_ID, SRVC_INSTNC_ATT_DESC, SRVC_INSTNC_ATT_NAME "&_
'			"FROM SO.SRVC_INSTNC_ATT "&_
'			"WHERE RECORD_STATUS_IND = 'A' "&_
'			"ORDER BY SRVC_INSTNC_ATT_DESC"

	'Create Recordset object
'	Set objRS = Server.CreateObject("ADODB.Recordset")
'	objRS.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
'	arrSerKenanList = objRS.GetRows

	'StrSql = "SELECT REVENUE_REGION_LCODE, " &_
	'         "       REVENUE_REGION_DESC " &_
	'		 "FROM SO.LCODE_REVENUE_REGION " &_
	'		 "WHERE RECORD_STATUS_IND = 'A' "

	'Create Recordset object
	'Set objRS = Server.CreateObject("ADODB.Recordset")
	'objRS.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
	'arrRevenueList = objRS.Getrows
 strSQL = "SELECT component_id, COMPONENT_NAME FROM ARBOR.V_PKG_COMPONENTS " &_
    		" order by component_id, COMPONENT_NAME"
 set objRsSTKenanCP = objKenanConn.Execute(strSQL)

 strSQL = "SELECT distinct PACKAGE_ID, PACKAGE_NAME FROM ARBOR.V_PKG_COMPONENTS " &_
    		" order by PACKAGE_ID, PACKAGE_NAME"
 set objRsSTKenanPK = objKenanConn.Execute(strSQL)
%>
<HTML>
<HEAD>
<META name="Generator" content="Microsoft Visual Studio 6.0">
<META http-equiv="Pragma" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<LINK rel="stylesheet" type="text/css" href="stylesheets/styles.css">
<SCRIPT type="text/javascript" language="javascript" src="AccessLevels.js"></SCRIPT>
<SCRIPT type="text/javascript" language="javascript" src="GeneralJavaFunctions.js"></SCRIPT>
<SCRIPT type="text/javascript" language="javascript" id="clientEventHandlersJS">
<!--
var intAccessLevel = <%=intAccessLevel%>;
var strWinName = "<%=strWinName%>";

var arrLOBList = new Array();
var arrServiceCategoryList = new Array();

//set section title
setPageTitle("SMA - Service Type");

function window_onLoad() {
//****************************************************************************************
//
//
//
//
//****************************************************************************************
var intCounter;

	arrLOBList[0] = "";
	arrServiceCategoryList[0] = "";

	for (intCounter = 1; intCounter < document.frmSKenanSearch.selLOB.options.length; intCounter++) {
		
		//giovanni arevalo  before '.options(intCounter);'
		var oOption= document.frmSKenanSearch.selLOB.options[intCounter];
	
		//Each array element holds LOB_ID|LOB_DESC
		arrLOBList[intCounter] = (oOption.value + "|" + oOption.text);
	}

	for (intCounter = 1; intCounter < document.frmSKenanSearch.selSCategory.options.length; intCounter++) {
		console.log('tamanio de select cat: '+document.frmSKenanSearch.selSCategory.options.length);
		//giovanni arevalo  before '.options(intCounter);'
		var oOption = document.frmSKenanSearch.selSCategory.options[intCounter];
		
		console.log('valor de oOption cat: '+oOption);
		console.log('oOption value'+oOption.value);
		console.log('oOption text'+oOption.text);

		//Each array element holds SERVICE_CATEGORY_ID|LOB_ID|SERVICE_CATEGORY_DESC
		arrServiceCategoryList[intCounter] = (oOption.value + "|" + oOption.text);
		console.log("valor del array: "+arrServiceCategoryList[intCounter]);
		var strValue = oOption.value;
		console.log("strValue::"+strValue);
		var arrValue = strValue.split("|");
		oOption.value = arrValue[0];
		console.log("oOption value:"+oOption.value);
	}

	if (document.frmSKenanSearch.selLOB.selectedIndex != 0
			|| document.frmSKenanSearch.selSCategory.selectedIndex != 0
			|| document.frmSKenanSearch.selSTKenanPK.selectedIndex != 0
			|| document.frmSKenanSearch.selSTKenanCP.selectedIndex != 0
			|| document.frmSKenanSearch.hdnSTypeID.value != ""
			|| document.frmSKenanSearch.txtSTypeDescription.value != "") {
		document.frmSKenanSearch.hdnSTypeID.value = "";
		DeleteCookie("BusinessID");
		DeleteCookie("ServiceCategoryID");
		DeleteCookie("ServiceKenanAID");
		DeleteCookie("ServiceType");
		DeleteCookie("STypeDesc");
		DeleteCookie("WinName");
		document.frmSKenanSearch.submit();
	}
}

function fct_onChangeLOB() {
var intCounter = 1;
var strBusinessID;

	if (document.frmSKenanSearch.selLOB.selectedIndex != 0) {
		strBusinessID = document.frmSKenanSearch.selLOB.value;

		//Remove all the OPTION tags from the Service Category
		for (intCounter = document.frmSKenanSearch.selSCategory.length - 1; intCounter > 0; intCounter--) {
			document.frmSKenanSearch.selSCategory.options.remove(intCounter);
		}
		//Add Service Categories that belong to the selected Line of Business
		for (intCounter = 1; intCounter < arrServiceCategoryList.length; intCounter++) {
			var strValue = arrServiceCategoryList[intCounter];
			var arrValue = strValue.split("|");
			if (arrValue[1] == strBusinessID) {
				var strElement = "<OPTION VALUE='" + arrValue[0] + "'>" + arrValue[2] + "</OPTION>";
				var oOption = document.createElement(strElement);
				document.frmSKenanSearch.selSCategory.options.add(oOption);
				oOption.innerText = arrValue[2];	//SERVICE_CATEGORY_DESC
//				oOption.Value = arrValue[1];		//LOB_ID
//				oOption.Value = arrValue[0];		//SERVICE_CATEGORY_ID
			}
		}
	}
	else {
		//Remove all the OPTION tags from the Service Category
		for (intCounter = document.frmSKenanSearch.selSCategory.length - 1; intCounter > 0; intCounter--) {
			document.frmSKenanSearch.selSCategory.options.remove(intCounter);
		}
		//Add all the Service Categories
		for (intCounter = 1; intCounter < arrServiceCategoryList.length; intCounter++) {
			var strValue = arrServiceCategoryList[intCounter];
			var arrValue = strValue.split("|");
			var strElement = "<OPTION VALUE='" + arrValue[0] + "'>" + arrValue[2] + "</OPTION>";
			var oOption = document.createElement(strElement);
			document.frmSKenanSearch.selSCategory.options.add(oOption);
			oOption.innerText = arrValue[2];	//SERVICE_CATEGORY_DESC
//			oOption.Value = arrValue[1];		//LOB_ID
//			oOption.Value = arrValue[0];		//SERVICE_CATEGORY_ID
		}
	}
}

function btnCalendar_onClick(intDateFieldNo) {
var NewWin;

	SetCookie("Field", intDateFieldNo);
	NewWin=window.open("TheCalendar.asp","NewWin","toolbar=no,status=no,width=260,height=225,menubar=no,resize=no");
	//NewWin.creator=self;
	NewWin.focus();
}

function btnNew_onClick() {
//************************************************************************************************
// Function:	btnAddNew_onClick()
//
// Purpose:		To bring up a blank Service Type Detail page so that user can enter a new ST.
//
// Created By:	Gilles Archer Oct 02 2000
//
// Updated By:
//************************************************************************************************
	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
		alert('You do not have permission to CREATE a Service Type.  Please contact your System Administrator.');
		return false;
	}
	parent.document.location.href ="STypeDetail.asp?ServiceTypeID=NEW";
}

//function fct_onChange() {
//	document.frmSKenanSearch.hdnSTypeID.value = "";
//}

function btnClear_onClick() {
		 document.frmSKenanSearch.selSTKenanCP
	with (document.frmSKenanSearch) {
		selLOB.selectedIndex = 0;
		selSCategory.selectedIndex = 0;
		hdnSTypeID.value = "";
		txtSTypeDescription.value = "";
		selSTKenanCP.value="";
		selSTKenanPK.value="";
		chkActiveOnly.checked = true;
	}
}

function fct_setDays(iIndex) {
var intDays = 31;
var strMonth = document.frmSKenanSearch.item("selmonth", iIndex).options[document.frmSKenanSearch.item("selmonth", iIndex).selectedIndex].value;
var strYear = document.frmSKenanSearch.item("selyear", iIndex).options[document.frmSKenanSearch.item("selyear", iIndex).selectedIndex].value;
var intCurrentDay = document.frmSKenanSearch.item("selday", iIndex).options[document.frmSKenanSearch.item("selday", iIndex).selectedIndex].value;
var intCounter = document.frmSKenanSearch.item("selday", iIndex).options.length;

	switch (strMonth) {
		case "02":						//February
			if (strYear % 4 != 0) { intDays = 28; }
			else if (strYear % 400 == 0) { intDays = 29; }
			else if (strYear % 100 == 0) { intDays = 28; }
			else { intDays = 29; }
			break;
		case "04": intDays = 30; break;	//April
		case "06": intDays = 30; break;	//June
		case "09": intDays = 30; break;	//September
		case "11": intDays = 30; break;	//November
		default: intDays = 31; break;	//January, March, May, July, August, October, December
	}
	if (intCounter <= intDays) {
		while (intCounter <= intDays) {
			var oOption = new Option(intCounter, intCounter);
			document.frmSKenanSearch.item("selday", iIndex).options[intCounter++] = oOption;
		}
	}
	else {
		while (intCounter > intDays) {
			document.frmSKenanSearch.item("selday", iIndex).options[intCounter--] = null;
		}
	}
	if (intCurrentDay > intDays) {
		document.frmSKenanSearch.item("selday", iIndex).selectedIndex = intDays;
	}
}
//-->
</SCRIPT>
</HEAD>
<BODY language="javascript" onLoad="window_onLoad();">
<FORM id="frmSKenanSearch" name="frmSKenanSearch" method="post" action="SKenanList.asp" target="fraResult">
	<INPUT type="hidden" id="hdnWinName" name="hdnWinName" value="<%=strWinName%>">
<TABLE cols="4" width=100%>
<THEAD>
	<TH colspan="4" align="left">Kenan Package Component Search</TH>
</THEAD>
<TBODY>
	<TR>
		<TD align="right" width="17%">Line of Business</TD>
		<TD align="left" width="51%"><SELECT id="selLOB" name="selLOB" style="width: 350px">
			<OPTION></OPTION>
			<%For lRow = LBound(arrLOBList, 2) To UBound(arrLOBList, 2)
				If StrComp(strBusinessID, arrLOBList(0, lRow), 0) = 0 Then%>
				<OPTION selected value="<%=arrLOBList(0, lRow)%>"><%=arrLOBList(1, lRow) & " - " & arrLOBList(2, lRow)%></OPTION>
				<%Else%>
				<OPTION value="<%=arrLOBList(0, lRow)%>"><%=arrLOBList(1, lRow) & " - " & arrLOBList(2, lRow)%></OPTION>
			<%	End If
			Next%>
			</SELECT>
		</TD>
	</TR>
	<TR>
		<TD align="right" width="17%">Service Category</TD>
		<TD align="left" width="51%"><SELECT id="selSCategory" name="selSCategory" style="width: 350px">
			<OPTION></OPTION>
			<%For lRow = LBound(arrSCategoryList, 2) To UBound(arrSCategoryList, 2)
				If StrComp(strServiceCategoryID, arrSCategoryList(0, lRow), 0) = 0 Then%>
				<OPTION selected value="<%=arrSCategoryList(0, lRow) & "|" & arrSCategoryList(1, lRow)%>"><%=arrSCategoryList(2, lRow)%></OPTION>
				<%Else%>
				<OPTION value="<%=arrSCategoryList(0, lRow) & "|" & arrSCategoryList(1, lRow)%>"><%=arrSCategoryList(2, lRow)%></OPTION>
			<%	End If
			Next%>
			</SELECT>
		</TD>

	</TR>
	<TR>
		<TD align="right" width="17%">Service Type</TD>
		<TD align="left" width="51%">
			<INPUT type="hidden" id="hdnSTypeID" name="hdnSTypeID" value="<%=strSTypeID%>">
			<INPUT type="text" id="txtSTypeDescription" name="txtSTypeDescription" value="<%if strSTypeDesc<>"" then response.write strSTypeDesc else response.write"" end if%>" maxlength="80" style="width: 350px">
		</TD>
		<TD>&nbsp;</TD>
	</TR>
	<TR>
		<TD align="right" width="17%">Package ID and Name</TD>
		<TD align="left" width="51%">
		<SELECT id=selSTKenanPK name=selSTKenanPK style="HEIGHT: 22; WIDTH: 447" >
		<OPTION ></OPTION>
		<%Do while Not objRsSTKenanPK.EOF
		   strken = objRsSTKenanPK(0)& " | " & objRsSTKenanPK(1) %>
		   <option  value = <% =objRsSTKenanPK(0) %>> <% =strken %> </option>
		<%  objRsSTKenanPK.MoveNext
		Loop		%>
		</SELECT>

		</TD>
	</TR>
	<TR>
		<TD align="right" width="17%">Component ID and Name</TD>
		<TD align="left" width="51%">
		<SELECT id=selSTKenanCP name=selSTKenanCP style="HEIGHT: 22; WIDTH: 447" >
		<OPTION ></OPTION>
		<%Do while Not objRsSTKenanCP.EOF
		   strken = objRsSTKenanCP(0)& " | " & objRsSTKenanCP(1) %>
		   <option  value = <% =objRsSTKenanCP(0) %>
		   > <% =strken %> </option>
		<%  objRsSTKenanCP.MoveNext
		Loop
		%>
		</SELECT>
		</TD>

		<TD align="right" width="6%">Active Only</TD>
		<td width="31%">
		<INPUT id=chkActiveOnly name=chkActiveOnly tabindex=12 type=checkbox value=YES CHECKED style="HEIGHT: 24; WIDTH: 25"></TD>
	</tr>
</TBODY>
<TFOOT>
	<TR>
		<TD colspan="4" align="right">
			<INPUT id="btnClear"  name="btnClear"  type="button" value="Clear"  style="width: 2cm" language="javascript" onClick="btnClear_onClick()">&nbsp;
			<INPUT id="btnSearch" name="btnSearch" type="submit" value="Search" style="width: 2cm" language="javascript">&nbsp;
		</TD>
	</TR>
</TFOOT>
</TABLE>
</FORM>
</BODY>
<%
	'Clean ADO Objects
	objRS.Close
	objRsSTKenanPK.close
	objRsSTKenanCP.close

	Set objRS = Nothing
	set objRsSTKenanCP = Nothing
	set objRsSTKenanPK = Nothing

	objConn.Close
	Set objConn = Nothing

	objKenanConn.Close
	set objKenanConn = Nothing
%>
</HTML>
