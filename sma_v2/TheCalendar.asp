<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE></TITLE>
<SCRIPT type="text/javascript" language="javascript" src="GeneralJavaFunctions.js"></SCRIPT>
<SCRIPT type="text/javascript" language="javascript" id="clientEventHandlersJS">
<!--
function frmCalendar_Click() {
var intMonth, intDay, intYear, intField;

	intMonth = document.frmCalendar.Calendar1.Month;
	intDay = document.frmCalendar.Calendar1.Day;
	intYear = document.frmCalendar.Calendar1.Year;
	 
	intField = <%=Request.Cookies("Field")%>;
	 
	if (intMonth < 10) {
		intMonth = '0' + intMonth;
	}
	if (intDay < 10) {
		intDay = '0' + intDay;
	}
  
	parent.opener.document.forms[0].item("selmonth", intField).value = intMonth;
	parent.opener.document.forms[0].item("selday", intField).value = intDay;
	parent.opener.document.forms[0].item("selyear", intField).value = intYear;
	
	window.close();
	return true;
}

function window_onLoad() {
var dtCal = new Date();

	document.frmCalendar.Calendar1.month = dtCal.getMonth()+1;
	document.frmCalendar.Calendar1.day = dtCal.getDate();
	document.frmCalendar.Calendar1.year = dtCal.getYear();
	 
	return true;
}
//-->
</SCRIPT>
<SCRIPT language="javascript" for="Calendar1" event="Click">
<!--
	frmCalendar_Click()
//-->
</SCRIPT>
</HEAD>
<BODY language="javascript" onLoad="window_onLoad()">
<FORM id="frmCalendar" name="frmCalendar">
<P>
<OBJECT classid="clsid:8E27C92B-1264-101C-8A2F-040224009C02" id=Calendar1 style="HEIGHT: 172px; LEFT: 0px; TOP: 0px; WIDTH: 237px">
	<PARAM NAME="_Version" VALUE="524288">
	<PARAM NAME="_ExtentX" VALUE="6271">
	<PARAM NAME="_ExtentY" VALUE="4551">
	<PARAM NAME="_StockProps" VALUE="1">
	<PARAM NAME="BackColor" VALUE="-2147483633">
	<PARAM NAME="Year" VALUE="2000">
	<PARAM NAME="Month" VALUE="7">
	<PARAM NAME="Day" VALUE="26">
	<PARAM NAME="DayLength" VALUE="1">
	<PARAM NAME="MonthLength" VALUE="0">
	<PARAM NAME="DayFontColor" VALUE="0">
	<PARAM NAME="FirstDay" VALUE="1">
	<PARAM NAME="GridCellEffect" VALUE="1">
	<PARAM NAME="GridFontColor" VALUE="10485760">
	<PARAM NAME="GridLinesColor" VALUE="-2147483632">
	<PARAM NAME="ShowDateSelectors" VALUE="-1">
	<PARAM NAME="ShowDays" VALUE="-1">
	<PARAM NAME="ShowHorizontalGrid" VALUE="-1">
	<PARAM NAME="ShowTitle" VALUE="-1">
	<PARAM NAME="ShowVerticalGrid" VALUE="-1">
	<PARAM NAME="TitleFontColor" VALUE="10485760">
	<PARAM NAME="ValueIsNull" VALUE="0"></OBJECT>
</P>
</FORM>
</BODY>
</HTML>
