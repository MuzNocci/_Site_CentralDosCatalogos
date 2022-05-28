<% Server.ScriptTimeout = 300 %>
<%
Dim objConn, stringSQL, strConnection, array_id, i, sql_id, id
id = Request.QueryString("radio")
'Caso ocorra algum erro os precessos não são interrompidos 
'e é passado para a próxima linha de comando
Set objConn =  Server.CreateObject("ADODB.Connection")
objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("../../../ch_database/bd_hermes.mdb")
'Deletando registro da tabela contato onde esta a id
		if err = 0 and id <> "" then
			array_id = split(id,",")
			For i=0 to ubound(array_id)
				sql_id = sql_id & "noticias.id = " & Trim(array_id(i)) & " OR "
														 'campo texto, entao" & Trim(array_id(i)) & " OR "
														 'caso numerico '" & Trim(array_id(i)) & "' OR "
			Next
			sql_id = left(sql_id,(len(sql_id)-4))
			stringSQL = "DELETE * FROM noticias WHERE "&sql_id&""
			objConn.Execute(stringSQL)
			
			
			
objConn.close
Set objConn = Nothing		  
response.Redirect "sucesso.asp"
END IF
%>