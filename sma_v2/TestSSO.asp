<% @Language="VBScript" %>
<HTML><BODY>
	<p>Dumping request headers</p>
	<%
        if Request.Cookies("EIDSSOpro") = "" then
            Response.AddHeader "Set-Cookie", "EIDSSOpro=" + Mid(Request.ServerVariables("QUERY_STRING"),11) + "; HttpOnly"
        end if

		for each x in Request.ServerVariables
  		response.write("<B>" & x & ":</b> " & Request.ServerVariables(x) & "<p />")
		next
  %>
</BODY></HTML>
