<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
function Calendar1_Click() {
var intMonth, intDay, intYear, strField;

	intMonth = document.frmCalendar.Calendar1.month;
	intDay = document.frmCalendar.Calendar1.day;
	intYear = document.frmCalendar.Calendar1.year;
	 
	intField = <%=Request.Cookies("Field")%>;
	 
	if (intMonth < 10) {
		intMonth = '0' + intMonth;
	}
	if (intDay < 10) {
		intDay = '0' + intDay;
	}
  
switch(intField){
 //case 0:
	//{
	//{parent.opener.document.forms[0].item("selmonth", intField).value = intMonth ;}
	//{parent.opener.document.forms[0].item("selday", intField).value = intDay ;}
	//{parent.opener.document.forms[0].item("selyear", intField).value = intYear ;}
	//break;
	//}
 case 1:
   {
   {parent.opener.document.forms[0].selmonth.value = intMonth ;}
   {parent.opener.document.forms[0].selday.value = intDay ;}
   {parent.opener.document.forms[0].selyear.value = intYear ;}
   break;
   }
 case 2:
   {
   {parent.opener.document.forms[0].selmonth2.value = intMonth ;}
   {parent.opener.document.forms[0].selday2.value = intDay ;}
   {parent.opener.document.forms[0].selyear2.value = intYear ;}
    break;
    }
  case 3:{
   {parent.opener.document.forms[0].selmonth3.value = intMonth ;}
   {parent.opener.document.forms[0].selday3.value = intDay ;}
   {parent.opener.document.forms[0].selyear3.value = intYear ;}
    break;
    }
  case 4:{
   {parent.opener.document.forms[0].selmonth4.value = intMonth ;}
   {parent.opener.document.forms[0].selday4.value = intDay ;}
   {parent.opener.document.forms[0].selyear4.value = intYear ;}
    break;
    }
  case 5:{
   {parent.opener.document.forms[0].selmonth5.value = intMonth ;}
   {parent.opener.document.forms[0].selday5.value = intDay ;}
   {parent.opener.document.forms[0].selyear5.value = intYear ;}
    break;
    }
  case 6:{
   {parent.opener.document.forms[0].selmonth6.value = intMonth ;}
   {parent.opener.document.forms[0].selday6.value = intDay ;}
   {parent.opener.document.forms[0].selyear6.value = intYear ;}
    break;
    } 
  case 7:{
   {parent.opener.document.forms[0].selmonth7.value = intMonth ;}
   {parent.opener.document.forms[0].selday7.value = intDay ;}
   {parent.opener.document.forms[0].selyear7.value = intYear ;}
    break;
    } 
  case 8:{
   {parent.opener.document.forms[0].selmonth8.value = intMonth ;}
   {parent.opener.document.forms[0].selday8.value = intDay ;}
   {parent.opener.document.forms[0].selyear8.value = intYear ;}
    break;
    } 
  } //end switch

	{ window.close(); }
	return true;
}

function window_onload() {
var dtCal = new Date();
	document.frmCalendar.Calendar1.month = dtCal.getMonth()+1;
	document.frmCalendar.Calendar1.day = dtCal.getDate();
	document.frmCalendar.Calendar1.year = dtCal.getYear();
	return true;
}

//-->
</SCRIPT>
<SCRIPT LANGUAGE=javascript FOR=Calendar1 EVENT=Click>
	Calendar1_Click();
</SCRIPT>
</HEAD>
<BODY  LANGUAGE=javascript onload="return window_onload()">
<FORM NAME=frmCalendar>
<!--<INPUT name=hdnWinName type=text value="<%=Request.Cookies("CalWinName")%>"> -->
<OBJECT name=Calendar1 classid="clsid:8E27C92B-1264-101C-8A2F-040224009C02" id=Calendar1 style="HEIGHT: 172px; LEFT: 0px; TOP: 0px; WIDTH: 237px" VIEWASTEXT>
	<PARAM NAME="_Version" VALUE="458752">
	<PARAM NAME="_ExtentX" VALUE="6270">
	<PARAM NAME="_ExtentY" VALUE="4551">
	<PARAM NAME="_StockProps" VALUE="1">
	<PARAM NAME="BackColor" VALUE="-2147483633">
	<PARAM NAME="Year" VALUE="2000">
	<PARAM NAME="Month" VALUE="7">
	<PARAM NAME="Day" VALUE="26">
	<PARAM NAME="DayFontColor" VALUE="0">
	<PARAM NAME="DayLength" VALUE="1">
	<PARAM NAME="FirstDay" VALUE="1">
	<PARAM NAME="GridCellEffect" VALUE="1">
	<PARAM NAME="GridFontColor" VALUE="10485760">
	<PARAM NAME="GridLinesColor" VALUE="-2147483632">
	<PARAM NAME="MonthLength" VALUE="0">
	<PARAM NAME="ShowDateSelectors" VALUE="-1">
	<PARAM NAME="ShowDays" VALUE="-1">
	<PARAM NAME="ShowHorizontalGrid" VALUE="-1">
	<PARAM NAME="ShowTitle" VALUE="-1">
	<PARAM NAME="ShowVerticalGrid" VALUE="-1">
	<PARAM NAME="TitleFontColor" VALUE="10485760">
	<PARAM NAME="ValueIsNull" VALUE="0"></OBJECT>
</FORM>
</BODY>
</HTML>
