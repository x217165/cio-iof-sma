<%@ Language=VBScript %>
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<FORM name=thisForm METHOD=post>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">


<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--


function button1_onclick() {
var NewWin
NewWin=window.open("calendar2.asp" ,"NewWin","toolbar=no,status=no,width=260,height=225,menubar=no resize=no");
NewWin.creator=self;
NewWin.focus();
}

//-->
</SCRIPT>

</HEAD>
<BODY>&nbsp;<INPUT id=text1 name=text1><INPUT id=button1 name=button1  type=button value=.. LANGUAGE=javascript onclick="return button1_onclick()">
<P>
</P>
<P>&nbsp;</P>
<P>
</P>
<P>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
</P>

</BODY>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</HTML>
