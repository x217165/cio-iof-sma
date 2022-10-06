<%@ LANGUAGE=VBSCRIPT %>
<%  
OPTION EXPLICIT
on error resume next
%>
<!-- #include file="smaConstants.inc" -->
<!-- #include file="smaProcs.inc" -->
<LINK href="stylesheets/styles.css" rel="stylesheet" type="text/css">
<HTML>
<TITLE>In-line Frame Page</TITLE>

<script type="text/javascript">
function cell_onClick(intRowNo){
	document.frmIFR.txtCurrentRow.value = intRowNo;
}
</script>

<body>
<form name="frmIFR">
<input type="hidden" name="txtCurrentRow" value="">
<%
dim rowCount, colCount, strInnerValues
dim aValues, k, m, intStartBody

'get parameters via POST or GET
rowCount = Request("rowCount")
colCount = Request("colCount")
strInnerValues = unescape(Request("InnerValues"))

'convert input text to an array
'the number of elements in the array should always be equal to rowCount x colCount
aValues = Split(strInnerValues, strDelimiter)

Response.Write "<table width=100% border=1 cellspacing=0 cellpadding=0>"&vbCrLf
'create table caption
if Request("Caption") <> "" then
	Response.Write "<caption>"&Request("Caption")&"</caption>"
end if
'create table header
intStartBody = 1
if UCase(Request("TblTitle")) <> "" then
	Response.Write "<thead>"&vbCrLf
	dim Item
	for each Item in Request.QueryString("tblTitle")
		Response.Write "<th>" & Item & "</th>"&vbCrLf
	next
	Response.Write "</thead>"&vbCrLf
	intStartBody = 2
end if
'get table body
Response.Write "<tbody>"&vbCrLf
for k = 1 to rowCount
	Response.Write "<tr>"&vbCrLf
	for m = 2 to colCount
		Response.Write "<td>"
		'create the hidden index just in the first cell of each row
		if m = 2 then Response.Write "<input name=x0y"&CStr(k-intStartBody+1)&" type=hidden value=""" & aValues((k-1)*colCount) & """>"&vbCrLf
		if Int(k/2) = k/2 then
			Response.Write "<input name=x"&CStr(m-1)&"y"&CStr(k-intStartBody+1)&" style=""BACKGROUND-COLOR: lightgoldenrodyellow; border=0; width=100%"" type=text value=""" & routineHtmlString(aValues((k-1)*colCount+m-1)) & """ onClick=""cell_onClick("&k&");"">"
		else
			Response.Write "<input name=x"&CStr(m-1)&"y"&CStr(k-intStartBody+1)&" style=""BACKGROUND-COLOR: white; border=0; width=100%"" type=text value=""" & routineHtmlString(aValues((k-1)*colCount+m-1)) & """ onClick=""cell_onClick("&k&");"">"
		end if
		Response.Write "</td>"&vbCrLf
	next 
	Response.Write "</tr>"&vbCrLf
next 
Response.Write "</tbody>"&vbCrLf
Response.write "</table>"&vbCrLf
		
%>
</FORM>
</BODY>
</HTML>