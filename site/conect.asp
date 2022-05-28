<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%	
	Set conexao = Server.CreateObject("ADODB.Connection")
	conexao.open="Driver={MySQL ODBC 5.1 Driver};server=186.202.13.31;database=centraldoscata;uid=centraldoscata;pwd=art@central"
	valor = Replace(Request.QueryString("q"),"'","")
	sql = "SELECT * FROM quatroestacoes where codigo like '"&valor&"%' ORDER BY codigo ASC limit 1;"
	Set query = conexao.execute(sql)
	While Not query.eof
	response.write query("codigo")&"|"&query("descricao")&"|"&query("pagina")&"|"&query("cor")&vbCrLf
	query.movenext
	Wend
	query.close
	Set query = Nothing
	conexao.close
	Set conexao = Nothing
%>
