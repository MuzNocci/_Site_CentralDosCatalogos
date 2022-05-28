<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<% Server.ScriptTimeout = 300 %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!--#include file="restrito.asp"-->
<!--#include file="conectar.asp"-->
<%

Dim objConn, objRs, strQuery
Dim strConnection

'Conectando com o banco de dados contato.mdb
Set objConn =  Server.CreateObject("ADODB.Connection")
objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("../../../ch_database/bd_hermes.mdb")
strQuery = "SELECT * FROM novidades order by novidades.id desc"
Set ObjRs = objConn.Execute(strQuery)
%>
<html>
<head>
<title>Administra&ccedil;&atilde;o</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../HERMES.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
body {
	margin:0px; background-image: url('../imagens/fundo.jpg');
	
}
-->
</style>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}
//-->
</script>
</head>
<body onLoad="MM_preloadImages('../imagens/home_on.gif','imagens/coluna_on.gif','imagens/noticias_on.gif','imagens/suporte_on.gif')">
<!--#include file="topbar.asp"-->
<table width="770" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="300" align="center" valign="top" background="../imagens/sombraesquerda.gif">
        <table width="770" border="0" cellspacing="0" cellpadding="0">
          <form method="get" action="action_excluirnovid.asp">
          <tr>
            <td height="50" align="left"><img src="imagens/exclusaodecolunatit.gif" width="236" height="23"></td>
          </tr>
          <tr>
            <td><table border="0" width="770" cellpadding="2" align="center">
              <tr bgcolor="#000099">
                <td width="454" bgcolor="#CC0000" class="titulocaixa"><div align="left">&nbsp;T&iacute;tulo:</div></td>
                <td width="54" height="2" align="center" bgcolor="#CC0000"><input name="Submit" type="submit" class="form" value="Excluir"></td>
              </tr>
              <%While Not objRS.EOF %>
              <tr bgcolor="#99CCFF">
                <td bgcolor="#FFD9D9" class="admintextos"><%Response.write objRS("titulo")%></td>
                <td width="54" height="2" align="center" bgcolor="#FFD9D9"><input type="radio" name="radio" value="<%=objRS(0)%>"></td>
              </tr>
              <%
  'Move para o pr&oacute;ximo registro
  objRS.MoveNext
  Wend
  'Fechando as conex&otilde;es
  objRs.close
  objConn.close
  Set objRs = Nothing
  Set objConn = Nothing
  %>
              <tr bgcolor="#000099">
                <td height="10" align="center" bgcolor="#CC0000"></td>
                <td height="10" align="center" bgcolor="#CC0000"></td>
              </tr>
            </table></td>
          </tr>
          </form>
        </table>
        </td>
  </tr>
</table>
</body>
</html>