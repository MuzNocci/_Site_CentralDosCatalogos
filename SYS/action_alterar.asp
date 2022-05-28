<% Server.ScriptTimeout = 300 %>
<!--#include file="restrito.asp"-->
<!--#include file="conectar.asp"-->
<%
Dim objConn, strQuery, sql_query, RsQuery, campo, sql, ID
Dim titulo, noticia1, noticia2 , ObjRs


titulo = Request.Form("titulo")
noticia = Request.Form("noticia")
ID = Request.Form("ID")



Set objConn =  Server.CreateObject("ADODB.Connection")
objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("../../../ch_database/bd_hermes.mdb")

strQuery = "UPDATE noticias SET titulo = '"&titulo&"', noticia = '"&noticia&"', data = '"&Date()&"', hora = '"&Time()&"' WHERE ID ="&ID&""

Set ObjRs = objConn.Execute(strQuery)

objConn.close
Set objRs = Nothing
Set objConn = Nothing 
if err = 0 Then

	response.redirect "sucesso.asp"
end if
%>