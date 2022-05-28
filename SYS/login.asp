<% Server.ScriptTimeout = 300 %>
<% if session("login") <> "" then %>
	<font face=verdana size=2>
	<b><center>
	<br><a href="logout.asp">Logout</a><br>
    </center>
<% else %>
		<form method=post action=validalogin.asp>
		<input type=hidden name=enviando value=sim>
		<table border=0>
		<tr>
		<td><font face=verdana size=2>Login </font></td>
		<td><input type=text name=login size=15></td></tr>
		<tr>
		<td><font face=verdana size=2>Senha </font></td>
		<td><input type=password name=senha size=15></td></tr>
		<tr>
		<td colspan=2 align=center><input type="submit" value="Enviar"></td></tr>
		</table>
		</form>
		</font>
	<%end if%>