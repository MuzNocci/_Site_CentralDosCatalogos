<% Server.ScriptTimeout = 300 %>
<%
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("../../../ch_database/bd_hermes.mdb")

ano = request.form("ano")
mes = request.form("mes")
datavolta1 = request.form("1datavolta")
datavolta2 = request.form("2datavolta")
datavolta3 = request.form("3datavolta")
datavolta4 = request.form("4datavolta")
datavolta5 = request.form("5datavolta")

dataida1 = request.form("1dataida")
dataida2 = request.form("2dataida")
dataida3 = request.form("3dataida")
dataida4 = request.form("4dataida")
dataida5 = request.form("4dataida")

inserir = "INSERT INTO datas (ano,mes,1dataida,2dataida,3dataida,4dataida,5dataida,1datavolta,2datavolta,3datavolta,4datavolta,5datavolta) VALUES ('"&ano&"','"&mes&"','"&dataida1&"','"&dataida2&"','"&dataida3&"','"&dataida4&"','"&dataida5&"','"&datavolta1&"','"&datavolta2&"','"&datavolta3&"','"&datavolta4&"','"&datavolta5&"')"

Conexao.Execute(inserir)
Conexao.Close
Set Conexao = Nothing

	Response.Redirect "sucesso.asp"
%> 