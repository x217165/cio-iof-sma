<%@ Language=VBScript %>
<% Option Explicit %>
<% Response.Buffer = true %>
<!--#include file="SmaConstants.inc"-->
<!--#include file="databaseconnect.asp"-->
<!--#include file="smaProcs.inc"-->
<!--
*************************************************************************************
* File Name:	FacilityCriteria.asp
*
* Purpose:	
*
* In Param:		This page reads following cookies
*				CustomerServiveA
*
* Out Param:
*
* Created By:	Ian Harriot
* Edited by:    Adam Haydey Jan 25, 2001
*               Added Customer Service City, Customer Service Address, and Past Facility Start Date search fields.
*				Also Added ADSL Due Date and Facility Start Date to the results.
**************************************************************************************

        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
       01-22-02	     DTy		Add Facility Provider and Is On Net search fields.
								Update Facility Provider drop down list.
								Re-sequence input.
       28-Feb-02	 DTy		Add 'Name/Alias' to 'Customer Service' field name.
       29-Jul-15   PSmith  Set Cookies in validation so the back key works
       05-Oct-15   PSmith  Only sumbit() for pop-up windows       
       03-Feb-16   PSmith  Don't pre-populate search criteria
**************************************************************************************
-->
<%
Dim objRs,Recordcnt,strbgcolor,StrSql,objRsFacTyp,objRsFacStat,StrCkt,objRsAdslTyp,strWinName, lIndex
DIM objRsFacPrvdr
dim strCustomerServiceA, strCustomerA, strServLocName, strServLocCity, strServLocAdd


dim intAccessLevel
strCustomerServiceA = Request.Cookies("CustomerServiceA") 
strCustomerA = Request.Cookies("CustomerA")
StrCkt = Request.QueryString("CktType")
strWinName	= Request.Cookies("WinName") 
strServLocName = Request.Cookies("ServLocName")

IF StrCkt = "PVC" THEN
intAccessLevel = CInt(CheckLogon(strConst_PVC))
ELSE
 intAccessLevel = CInt(CheckLogon(strConst_Facilities))
END IF

if intAccessLevel and intConst_Access_ReadOnly <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to PVC/Facilities. Please contact your system administrator"
end if



StrSql = "SELECT ADSL_TYPE_CODE,ADSL_TYPE_DESC FROM CRP.ADSL_TYPE WHERE RECORD_STATUS_IND = 'A' ORDER BY ADSL_TYPE_CODE"
      
     'Create Recordset object  
 set objRsAdslTyp = objConn.Execute(StrSql)
 
StrSql = "SELECT NOC_REGION_LCODE,NOC_REGION_DESC FROM CRP.LCODE_NOC_REGION WHERE RECORD_STATUS_IND = 'A' ORDER BY  NOC_REGION_LCODE"
      
'Create Recordset object  
set objRS = objConn.Execute(StrSql)
StrSql = "SELECT CIRCUIT_STATUS_CODE FROM CRP.CIRCUIT_STATUS WHERE RECORD_STATUS_IND = 'A' ORDER BY CIRCUIT_STATUS_CODE"
set objRsFacStat = objConn.Execute(StrSql)
   
IF StrCkt = "PVC" THEN
	StrSql = "SELECT CIRCUIT_TYPE_CODE FROM CRP.CIRCUIT_TYPE WHERE CIRCUIT_TYPE_CODE LIKE '%PVC%' AND RECORD_STATUS_IND = 'A' ORDER BY CIRCUIT_TYPE_CODE"
ELSE
	StrSql = "SELECT CIRCUIT_TYPE_CODE FROM CRP.CIRCUIT_TYPE WHERE CIRCUIT_TYPE_CODE NOT LIKE '%PVC%' AND RECORD_STATUS_IND = 'A' ORDER BY CIRCUIT_TYPE_CODE"
END IF
   
set objRsFacTyp = objConn.Execute(StrSql)

StrSql = "SELECT CIRCUIT_PROVIDER_CODE, CIRCUIT_PROVIDER_NAME,  DECODE(IS_ON_NET, 'Y', ' (ON NET)','') AS ""IS_ON_NET"" FROM CRP.CIRCUIT_PROVIDER WHERE RECORD_STATUS_IND='A' ORDER BY CIRCUIT_PROVIDER_NAME"
set objRsFacPrvdr = objConn.Execute(StrSql)

%>
 


<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
<script type="text/javascript" SRC="AccessLevels.js"></script>
<LINK REL="stylesheet" TYPE="text/css" HREF="stylesheets/styles.css">
<script type="text/javascript">
var intAccessLevel=<%=intAccessLevel%>;

//set the heading
		if ('<%=StrCkt%>' == 'PVC') {
			setPageTitle("SMA - PVC");
		} else {
			setPageTitle("SMA - Facility");
		}


function confirm_search(theForm)
{
var bolConfirm;

//alert(theForm.selfactyp.value);
if (theForm.selfactyp.value != 'ATMPVC'){
	if ((isWhitespace(theForm.txtfacname.value) &&isWhitespace(theForm.selfactyp.value)&& isWhitespace(theForm.selfacadsltyp.value) &&
	    isWhitespace(theForm.selFacPrvdr.value) &&isWhitespace(theForm.selOnOffNet.value)&&
		isWhitespace(theForm.selfacstat.value) && isWhitespace(theForm.selregion.value) && isWhitespace(theForm.txtcuserva.value) &&
		isWhitespace(theForm.txtcusta.value) && isWhitespace(theForm.txtservloca.value) && isWhitespace(theForm.txtservcity.value) && 
		isWhitespace(theForm.txtservadd.value) && (theForm.chkutadsl.checked == false) && (theForm.chkPastFacStart.checked == false)))
	{
  
    bolConfirm = window.confirm("No Search Criteria have been entered. This search may take a long time..Continue?");
    if (!bolConfirm){
     return false;
    }
  }
} //end if <> ATMPVC
   
if (theForm.selfactyp.value == 'ATMPVC'){
 if ((isWhitespace(theForm.txtfacname.value)  &&isWhitespace(theForm.selfactyp.value)&& 
    isWhitespace(theForm.selfacstat.value) && isWhitespace(theForm.selregion.value) && isWhitespace(theForm.txtcuserva.value) &&
    isWhitespace(theForm.txtcusta.value) && isWhitespace(theForm.txtservloca.value) && (theForm.chkoutstpvc.checked == false)))
{
  
  bolConfirm = window.confirm("No Search Criteria have been entered. This search may take a long time..Continue?");
    if (!bolConfirm){
     return false;
    }
  }
 } //end if = ATMPVC
   
  // Start thinking
  thinking(parent.fraResult);
  return true;

}

function window_onload() {
//***************************************************************************************
// Function:	window_onload
//
// Purpose:		To sumbit the form automatically if there Customer Service A has a value.
//
// Created By:	Sara Sangha Sept. 1st, 200
//
// Updated By:	Nancy Mooney 09/03/2000
//***************************************************************************************
	var strWinName;
	strWinName = document.fmfacSearch.hdnWinName.value ; 
	if (strWinName !=  "" ){
 		DeleteCookie("WinName") ;
	}

	var strCustomerServiceA = document.fmfacSearch.txtcuserva.value ;
	var strCustomerA = document.fmfacSearch.txtcusta.value ;
	var strServLocName = document.fmfacSearch.txtservloca.value ; 
	
	DeleteCookie("CustomerServiceA");
	DeleteCookie("CustomerA");
	DeleteCookie("ServLocName");
	
	if ( strWinName == "Popup" && ((strCustomerServiceA != "")||(strCustomerA != "" || strServLocName != ""  ))) {
		SetCookie("CustomerServiceA",document.fmfacSearch.txtcuserva.value);
		SetCookie("CustomerA",document.fmfacSearch.txtcusta.value);
		SetCookie("ServLocName",document.fmfacSearch.txtservloca.value);
    thinking(parent.fraResult);
		document.fmfacSearch.submit();
	}  
}

function btnNew_click(){
 var strFacType;
	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {alert('Access denied. Please contact your system administrator.'); return;}
	strFacType = document.fmfacSearch.selfactyp.item(document.fmfacSearch.selfactyp.selectedIndex).value;
	if (isWhitespace(strFacType ))
	{
	  alert("You must enter a facility type.");
	 }
	 else{ 
	SetCookie("FacilityType",strFacType); 
	parent.document.location.href ="FacilityDetail.asp?NewFacility=NEW&CircuitTyp="+strFacType;
	}
}

function btnClear_click(){
var bolPVC='<%=StrCkt%>';
   
	with(document.fmfacSearch){ 
		txtfacname.value = "" ;
		selregion.selectedIndex = 0 ;
		selfactyp.selectedIndex = 0 ;
		selFacPrvdr.selectedIndex = 0 ;
		txtcusta.value = "";
		selfacstat.selectedIndex = 0 ;
		txtcuserva.value = "" ;

		if (bolPVC == 'PVC')
			{ chkoutstpvc.checked=false; 
			  
			  item("selday", 0).selectedIndex = 0 ;
			  item("selmonth", 0).selectedIndex = 0 ;
			  item("selyear", 0).selectedIndex = 0 ;
			  
			  item("selday", 1).selectedIndex = 0 ;
			  item("selmonth", 1).selectedIndex = 0 ;
			  item("selyear", 1).selectedIndex = 0 ;
			
			}
		else
			{ 
			txtservcity.value = "";
			txtservadd.value = "";
			chkutadsl.checked=false; 
			chkPastFacStart.checked=false; }
		
		if (bolPVC == 'OTHER')
		    selfacadsltyp.selectedIndex = 0 ;
		  
		txtservloca.value = "";
		chkactive.checked = true; 
		
	}
}

function fct_setDays(iIndex) {
var intDays = 31;
var strMonth = document.fmfacSearch.item("selmonth", iIndex).options[document.fmfacSearch.item("selmonth", iIndex).selectedIndex].value;
var strYear = document.fmfacSearch.item("selyear", iIndex).options[document.fmfacSearch.item("selyear", iIndex).selectedIndex].value;
var intCurrentDay = document.fmfacSearch.item("selday", iIndex).options[document.fmfacSearch.item("selday", iIndex).selectedIndex].value;	
var intCounter = document.fmfacSearch.item("selday", iIndex).options.length;
	
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
			document.fmfacSearch.item("selday", iIndex).options[intCounter++] = oOption;
		}
	}
	else {
		while (intCounter > intDays) {
			document.fmfacSearch.item("selday", iIndex).options[intCounter--] = null;
		}
	}
	if (intCurrentDay > intDays) {
		document.fmfacSearch.item("selday", iIndex).selectedIndex = intDays;
	}
}

function btnCalendar_onClick(intDateFieldNo) {
var NewWin;

	SetCookie("Field", intDateFieldNo);
	NewWin=window.open("TheCalendar.asp","NewWin","toolbar=no,status=no,width=260,height=225,menubar=no,resize=no");
	//NewWin.creator=self;
	NewWin.focus();
}

function btnSearch_onClick()
{
//********************************************************************************************
// Function:	btnSearch_onClick()
//
// Purpose:		
//				To validate PVC Create To and From date fields before submitting the page to FacilityList.asp
//				The FacilityList.asp requires that either both date values for its BETWEEN clause
//				or no date value at all.
//
// Created By:	Sara Sangha	
//
// Updated By
//
//********************************************************************************************
var bolPVC='<%=StrCkt%>';

if (bolPVC == "PVC") {

 
  var strFromDay = document.fmfacSearch.item("selday", 0).value ;
  var strFromMonth = document.fmfacSearch.item("selmonth",0).value ;
  var strFromYear = document.fmfacSearch.item("selyear", 0).value ;	
 
  var strToDay = document.fmfacSearch.item("selday", 1).value ;
  var strToMonth = document.fmfacSearch.item("selmonth", 1).value ;
  var strToYear = document.fmfacSearch.item("selyear", 1).value ;


  var strToDate = "" ;
  var strFromDate = "" ;

	// Validate Create From Date
	if ( strFromDay != "" || strFromMonth != "" || strFromYear != "" ) {
		// one of the date field has a value
		// now make sure all date fields (day, month and year has values) 

		if ( strFromDay == "" || strFromMonth == "" || strFromYear == "")
		{
			alert('Invalid Date. Please enter a valid date.');
			document.fmfacSearch.item("selmonth",0).focus()  ;
			return(false);
		}	
		else
		{
			strFromDate =  strFromMonth + "/" + strFromDay + "/" + strFromYear ;
			
		}	
	}

	// validate Create To Date
	if ( strToDay != "" || strToMonth != "" || strToYear != "" ) {
		// one of the date field(day, month or year) has a value
		// Now, make sure all date fields (day, month and year has values)
			
			if ( strToDay == ""	|| strToMonth == "" || strToYear == "") 
			{
				alert("Invalid Date. Please enter a valid date.");
				document.fmfacSearch.item("selmonth", 1).focus()  ;
				return(false);
			}
		else
			{
				strToDate = strToMonth + "/" + strToDay +  "/" + strToYear  ;
				
			}
	}	


		
	if ((strToDate != "") && (strFromDate != "")) {	 	
			// both date fields have valid dates
			document.fmfacSearch.hdnToDate.value  = strToDate ;
			document.fmfacSearch.hdnFromDate.value  = strFromDate ;  
			return(true) ;
		}
		
	else { 
		
		if ( (strToDate == "") && (strFromDate == "") ){
			// user has not selected any date ;
			// therefore, reset the value to nothing.
			
			document.fmfacSearch.hdnToDate.value = ""
			document.fmfacSearch.hdnFromDate.value = ""
			
			return(true);
		}
			
		else {
			// only one date field has a value 
			alert('Please enter both To and From dates');
			return(false);	
		}
	}
  }	
} // end of btnSearch_onClick()

// End of script hiding -->
</SCRIPT>
</HEAD>

<BODY  LANGUAGE=javascript onload="return window_onload()">
<FORM NAME=fmfacSearch METHOD=POST ACTION="FacilityList.asp" TARGET="fraResult" onSubmit="return confirm_search(this)">
	<!-- hidden variables -->
	<input type="hidden" name="hdnWinName" value="<%=strWinName%>">
	<input type="hidden" name="hdnToDate" value="" >
	<input type="hidden" name="hdnFromDate" value="" >

<TABLE cols=4 WIDTH="100%" border=0>
<thead>
	<%
	if StrCkt = "PVC" then
		response.write "<tr><td align=left colspan=4>PVC Search</td></tr>"
	else
		Response.Write "<tr><td align=left colspan=4>Facility Search</tr></tr>"
	end if
	%>
</thead>
<tbody>
<TR>
	<%
	if StrCkt = "PVC" then
		Response.Write "<TD width=15% align=right nowrap >Path Name</TD>"
	else
		Response.Write "<TD width=15% align=right nowrap >Name/Number/Alias</TD>"
	end if
	%>
	<TD align=left width=20%><INPUT id=txtfacname name=txtfacname tabindex=1 style="HEIGHT: 23px; WIDTH: 200px"></TD>
	<TD align=right width=15% nowrap>Region</TD>
	<TD align=left >
		<SELECT id=selregion name=selregion tabindex=7 style="HEIGHT: 20px; WIDTH: 120px">
			<OPTION></OPTION>
			<%Do while Not objRS.EOF 
				Response.write "<OPTION VALUE ="& objRS("NOC_REGION_LCODE") & ">" & objRS("NOC_REGION_DESC") & "</OPTION>"
				objRS.MoveNext   
			Loop
			%>
		</SELECT>
	</TD>
</TR>
<TR>
	<%
	if StrCkt = "PVC" then
		Response.Write "<TD align=RIGHT width=15% nowrap>PVC Type</TD>"
	else
		Response.Write "<TD align=RIGHT width=15% nowrap>Facility Type</TD>"
	end if
	%>
	<TD ALIGN=LEFT width=20%>
		<SELECT id=selfactyp  name=selfactyp tabindex=2 style="HEIGHT: 20px; WIDTH: 120px">
			<%
			if StrCkt <> "PVC" then
				Response.Write "<OPTION></OPTION>"&vbCrLf
			end if
			Do while Not objRsFacTyp.EOF 
				Response.write "<OPTION VALUE ="& objRsFacTyp("CIRCUIT_TYPE_CODE") & ">" & objRsFacTyp("CIRCUIT_TYPE_CODE") & "</OPTION>"
				objRsFacTyp.MoveNext   
			Loop
			%>
		</SELECT>
	</TD>
	<%IF StrCkt = "PVC" THEN%>
		<TD align=RIGHT nowrap width=15% >PVC Status</TD>
	<%ELSE%>
		<TD align=RIGHT nowrap width=15% >Facility Status</TD>
	<%END IF%>
	<TD ALIGN=LEFT>
		<SELECT id=selfacstat name=selfacstat tabindex=8 style="HEIGHT: 20px; WIDTH: 120px">
			<OPTION></OPTION>
			<%Do while Not objRsFacStat.EOF 
				Response.write "<OPTION VALUE ="& objRsFacStat("CIRCUIT_STATUS_CODE") & ">" & objRsFacStat("CIRCUIT_STATUS_CODE") & "</OPTION>"
				objRsFacStat.MoveNext   
			Loop
			%>
		</SELECT>
	</TD>	
</TR>

<TR>
	<%
	if StrCkt <> "PVC" then
		
		Response.Write "<TD align=RIGHT width=15% nowrap>Facility Providers</TD>"
	    
	    Response.Write "<TD ALIGN=LEFT>"
	    Response.Write "<SELECT id=selFacPrvdr  name=selFacPrvdr tabindex=3 style=HEIGHT: 20px; WIDTH: 240px>"
		Response.Write "<OPTION></OPTION>"&vbCrLf
		Do while Not objRsFacPrvdr.EOF 
			Response.write "<OPTION VALUE =" & objRsFacPrvdr("CIRCUIT_PROVIDER_CODE") & ">" & objRsFacPrvdr("CIRCUIT_PROVIDER_NAME") & "  " & objRsFacPrvdr("IS_ON_NET") & "</OPTION>"
			objRsFacPrvdr.MoveNext   
		Loop
		Response.Write "</SELECT>"
		Response.Write "</TD>"
		Response.Write "<TD align=RIGHT width=15% nowrap>ON/OFF Net</TD>"
	    Response.Write "<TD ALIGN=LEFT>"
	    Response.Write "<SELECT id=selOnOffNet  name=selOnOffNet tabindex=9 style=HEIGHT: 20px; WIDTH: 120px>"
		Response.Write "<OPTION></OPTION>"&vbCrLf
		Response.Write "<OPTION VALUE=ON >ON NET only</OPTION>"&vbCrLf
		Response.Write "<OPTION VALUE=OFF>OFF NET only</OPTION>"&vbCrLf
		Response.Write "</SELECT>"
		Response.Write "</TD>"
	end if
	%>
</TR>
<TR>
	<TD align=right nowrap width=15% >Customer</TD>
	<TD align=left width=20%><INPUT id=txtcusta name=txtcusta tabindex=4 style="HEIGHT: 23px; WIDTH: 200px" value="<%=strCustomerA%>"></TD>
	 <%if StrCkt <> "PVC" then %>
		<TD ALIGN=RIGHT NOWRAP width=15% >ADSL Service Type</TD>
		<TD ><SELECT id=selfacadsltyp name=selfacadsltyp tabindex=10 style="HEIGHT: 20px; WIDTH: 120px" >
			<OPTION></OPTION>
			<%
			Do while Not objRsAdslTyp.EOF 
				Response.write "<OPTION VALUE ="& objRsAdslTyp("ADSL_TYPE_CODE") & ">" & objRsAdslTyp("ADSL_TYPE_DESC") & "</OPTION>" 
				objRsAdslTyp.MoveNext  
			Loop
			%>
			</SELECT>
		</TD>
	  <% ELSE %>
		 <TD ALIGN=RIGHT NOWRAP width=15% >Create Date ( From</TD>	
		 <TD><SELECT id="selmonth" name="selmonth" onChange="fct_setDays(0);">
		 <OPTION></OPTION>
		 <%For lIndex = 1 To 12
			If lIndex < 10 Then	lIndex = "0" & lIndex
			Response.Write "<OPTION value='" & lIndex & "'>" & monthName(lIndex, False) & "</OPTION>"  
		 Next%>
		</SELECT>
		<SELECT id="selday" name="selday" onChange="fct_setDays(0);">
		<OPTION></OPTION>
		<%For lIndex = 1 To 31
			If lIndex < 10 Then	lIndex = "0" & lIndex
			Response.Write "<OPTION value='" & lIndex & "'>" & lIndex & "</OPTION>"  
		Next%>
		</SELECT>
		<SELECT id="selyear" name="selyear" onChange="fct_setDays(0);">
		<OPTION></OPTION>
		<%For lIndex = intBaseYear To Year(Now) + 7
			Response.Write "<OPTION value='" & lIndex & "'>" & lIndex & "</OPTION>"  
		Next%>
		</SELECT>
		<INPUT id="btnCalendar" name="btnCalendar" type="button" value="..." language="javascript" onClick="btnCalendar_onClick(0);"></TD>
		
	</TR>
	<% end if %>
</TR>
<TR>
	<TD align=right nowrap width=15% >Customer Service Name/Alias</TD>
	<TD align=left width=20% ><INPUT id=txtcuserva name=txtcuserva tabindex=5 style="HEIGHT: 23px; WIDTH: 200px" value="<%=strCustomerServiceA%>"></TD>
	<%IF StrCkt = "PVC" THEN%>
		<TD ALIGN=RIGHT NOWRAP width=15% > To )</TD>	
		 <TD><SELECT id="selmonth" name="selmonth" onChange="fct_setDays(1);">
		 <OPTION></OPTION>
		 <%For lIndex = 1 To 12
			If lIndex < 10 Then	lIndex = "0" & lIndex
			Response.Write "<OPTION value='" & lIndex & "'>" & monthName(lIndex, False) & "</OPTION>"  
		 Next%>
		</SELECT>
		<SELECT id="selday" name="selday" onChange="fct_setDays(1);">
		<OPTION></OPTION>
		<%For lIndex = 1 To 31
			If lIndex < 10 Then	lIndex = "0" & lIndex
			Response.Write "<OPTION value='" & lIndex & "'>" & lIndex & "</OPTION>"  
		Next%>
		</SELECT>
		<SELECT id="selyear" name="selyear" onChange="fct_setDays(1);">
		<OPTION></OPTION>
		<%For lIndex = intBaseYear To Year(Now) + 7
			Response.Write "<OPTION value='" & lIndex & "'>" & lIndex & "</OPTION>"  
		Next%>
		</SELECT>
		<INPUT id="btnCalendar" name="btnCalendar" type="button" value="..." language="javascript" onClick="btnCalendar_onClick(1);"></TD> 
	<%ELSE %>
		<TD ALIGN=right width=15%>Untrained ADSL Past Due Date</TD>
		<TD ALIGN=LEFT><INPUT TYPE=CHECKBOX NAME=chkutadsl tabindex=9 VALUE=YES></TD>
	<%END IF%>
</TR>
<TR>
	<TD align=right nowrap width=15%>Service Location</TD>
	<TD ALIGN=LEFT width=20%><INPUT id=txtservloca name=txtservloca tabindex=5 style="HEIGHT: 23px; WIDTH: 200px" value="<%=strServLocName%>" ></TD> 
	<%IF StrCkt = "PVC" THEN%>
		<TD ALIGN=right width=15%>Outstanding/ Uncorrelated PVCs</TD><TD ALIGN=LEFT>
		<INPUT TYPE=CHECKBOX NAME=chkoutstpvc tabindex=9 VALUE=YES></TD></TR>
	<% ELSE %>
			<TD align=right nowrap width=15%>Service Location City</TD>
			<TD ALIGN=LEFT><INPUT id=txtservcity name=txtservcity tabindex=9 style="HEIGHT: 23px; WIDTH: 200px" value="<%=strServLocCity%>" ></TD>
		</TR>
		<TR>
			<TD align=right nowrap width=15%>Service Location Address</TD>
			<TD ALIGN=LEFT width=20%><INPUT id=txtservadd name=txtservadd tabindex=5 style="HEIGHT: 23px; WIDTH: 200px" value="<%=strServLocAdd%>" ></TD> 
			<TD align=right nowrap width=15%>Past Facility Start Date </TD>
			<TD ALIGN=LEFT ><INPUT TYPE=CHECKBOX NAME="chkPastFacStart" VALUE=YES tabindex=10></TD>
		</TR>
		
	<% END IF %>
	
<TR>
	<TD ALIGN=right width=15%>Active Only</TD>
	<TD ALIGN=LEFT><INPUT TYPE=CHECKBOX NAME="chkactive" VALUE=YES tabindex=10 CHECKED></TD>
	<TD align=right colspan=2>
		<% if strWinName <> "Popup" then %>
		<INPUT id=btnNew name=btnNew  type=button value=New tabindex=14 style="WIDTH: 2cm" LANGUAGE=javascript onclick="return btnNew_click();">&nbsp;&nbsp;
		<% end if %>
		<INPUT id=btnClear name=btnClear type=button tabindex=12 value=Clear style="WIDTH: 2cm" LANGUAGE=javascript onclick="return btnClear_click();" >&nbsp;&nbsp;
		<INPUT id=btnSearch name=btnSearch type=submit tabindex=13 value=Search style="WIDTH: 2cm" LANGUAGE=javascript onclick="return btnSearch_onClick();">&nbsp;&nbsp;
	</TD>
</TR>
</tbody>
</TABLE>
</FORM>
</BODY>
</HTML>

<%
'Clean up our ADO objects
    objRS.close
    set objRS = Nothing
    
    objRsFacTyp.close
    set objRsFacTyp = Nothing
    
    objRsFacStat.close
    set objRsFacStat = Nothing

    objRsAdslTyp.close
    set objRsAdslTyp = Nothing

    objRsFacPrvdr.close
    set objRsFacPrvdr = Nothing
    
    objConn.close
    set ObjConn = Nothing
 %>
    
