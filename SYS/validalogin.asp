<!--#include file="conectar.asp"-->
<%
sUsername=trim(Request.Form("login"))
sPassword=trim(Request.Form("senha"))

Sql = "SELECT * FROM admin WHERE login='" & sUsername & "' AND senha='" & sPassword & "' "
set rs = conexao.execute(Sql)

if not rs.eof then
Session("logado") = cint(rs("id_cod"))+1135716
Session("login") = trim(rs("login"))
Session("id_cod") = trim(rs("id_cod"))
sUsername=""
sPassword=""
Response.Redirect "default.asp"
else
Response.Redirect "default.asp?erro=log"
end if
%>