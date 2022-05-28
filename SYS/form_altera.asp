
<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<% Server.ScriptTimeout = 300 %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!--#include file="restrito.asp"-->
<!--#include file="conectar.asp"-->
<%
Dim objConn, stringSQL, strConnection, idi
idi = Request.QueryString("radio")
' Conectando com o banco de dados contato.mdb
Set objConn =  Server.CreateObject("ADODB.Connection")
Set objRS =  Server.CreateObject("ADODB.Connection")
objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("../../../ch_database/bd_hermes.mdb")
'objConn.Open "DBQ=" & Server.MapPath("../../../ch_database/bd_hermes.mdb") & ";Driver={Microsoft Access Driver (*.mdb)}","username","password"
stringSQL = "SELECT * FROM noticias WHERE ID = "&idi
Set objRS = objConn.Execute(stringSQL)
%>
<html>
<head>
<title>Administra&ccedil;&atilde;o</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../hermes.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
body {
	margin:0px; background-image: url('../imagens/fundo.jpg');
}
-->
</style>
<script language="JavaScript">
	function validaForm(){
		//validar nome
		d = document.cadastro;
		if (d.titulo.value == ""){
			alert("Preencha o Título da Notícia!");
			d.titulo.focus();
			return false;
		}
		//validar telefone
		if (d.noticia1.value == ""){
			alert("Preencha pelo menos o 1º Parágrafo!");
			d.noticia1.focus();
			return false;
		}
				
}
</script>
</head>
<body>
<!--#include file="topbar.asp"-->
<table width="770" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="300" align="center" background="../imagens/sombraesquerda.gif">
      <form action="action_alterar.asp" method="post" name="cadastro" id="cadastro"  onSubmit="return validaForm()">
        <table width="770" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="50" colspan="2" align="left" width="740"><img src="imagens/alteracaodenticiastit.gif" width="222" height="22"></td>
          </tr>
          <tr>
            <td colspan="2" align="left" class="textoform" width="740">T&iacute;tulo da Not&iacute;cias:</td>
          </tr>
          <tr>
            <td colspan="2" align="left" width="740"><input name="titulo" type="text" class="form" value="<%=objRS("titulo")%>" size="60" maxlength="60"></td>
          </tr>
          <tr>
            <td height="30" colspan="2" align="left" valign="bottom" class="textoform" width="740">Corpo da Not&iacute;cia: </td>
          </tr>
          <tr>
            <td colspan="2" align="left" width="740">
            <textarea name="noticia" cols="90" rows="6" class="formmsg" id="noticia" style="width: 770"><%=objRS("noticia")%></textarea></td>
          </tr>
          <tr>
            <td height="60" align="left" width="668"><div align="right">
              <input name="Submit2" type="reset" class="form" value="Limpar">
              &nbsp; </div></td>
            <td align="left" width="72">&nbsp;
              <input name="Submit" type="submit" class="form" value="Enviar">
              <input name="id" type="hidden" id="id5" value="<%=objRS("id")%>"></td>
          </tr>
        </table>
    </form>    </td>
  </tr>
</table>
</body>
</html>