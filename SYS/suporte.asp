<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<% Server.ScriptTimeout = 300 %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!--#include file="restrito.asp"-->
<!--#include file="conectar.asp"-->
<%
dim useraction

useraction = request("action")
select case useraction
	case "login"
		nome = request.form("nome")
		telefone = request.form("telefone")
		duvida = request.Form("duvida")
		email = request.form("email")
		
	htmlemail = htmlemail & "<img src=""http://www.drveit.com/novosite/imagens/tituloemail.gif"" width=""610"" height=""110"">"
	htmlemail = htmlemail & "<BR><BR><font color=""#003399"" face=""Trebuchet MS, Verdana, Arial"" size=""2""><strong>Nome:</strong></font><br>"
	htmlemail = htmlemail & "<font color=""#ff0000"" face=""Trebuchet MS, Verdana, Arial"" size=""2""><strong>" & nome & "</strong></font><br>"
	htmlemail = htmlemail & "<font color=""#003399"" face=""Trebuchet MS, Verdana, Arial"" size=""2""><strong>Telefone:</strong></font><br>"	
	htmlemail = htmlemail & "<font color=""#ff0000"" face=""Trebuchet MS, Verdana, Arial"" size=""2""><strong>" & telefone & "</strong></font><br>"
	htmlemail = htmlemail & "<font color=""#003399"" face=""Trebuchet MS, Verdana, Arial"" size=""2""><strong>E-mail:</strong></font><br>"	
	htmlemail = htmlemail & "<font color=""#ff0000"" face=""Trebuchet MS, Verdana, Arial"" size=""2""><strong>" & email & "</strong></font><br>"
	htmlemail = htmlemail & "<font color=""#003399"" face=""Trebuchet MS, Verdana, Arial"" size=""2""><strong>Dúvida ou Problema:</strong></font><br>"	
	htmlemail = htmlemail & "<font color=""#ff0000"" face=""Trebuchet MS, Verdana, Arial"" size=""2""><strong>" & duvida & "</strong></font><br>"


Set objCDOSYSMail = Server.CreateObject("CDO.Message") 
Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration") 

objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "sharedrelay.dominal.com" 
objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 
objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
objCDOSYSCon.Fields.update 

Set objCDOSYSMail.Configuration = objCDOSYSCon 

objCDOSYSMail.From = "drveit@drveit.com" 
objCDOSYSMail.To = "miranda_rafael@yahoo.com.br" 
objCDOSYSMail.Subject = "Problemas no Site" 

objCDOSYSMail.HTMLBody = htmlemail 
objCDOSYSMail.Send 


Set objCDOSYSMail = Nothing 
Set objCDOSYSCon = Nothing 

			response.redirect "sucesso.asp"
end select
%>
<html>
<head>
<title>Administra&ccedil;&atilde;o - Dr. Veit Odontologia &amp; Sa&uacute;de</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../stilos.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
body {
	background-image: url(../imagens/fundo.jpg);
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style>
<script language="JavaScript">
	function validaForm(){
		//validar nome
		d = document.cadastro;
		if (d.nome.value == ""){
			alert("O campo " + d.nome.name + " deve ser preenchido!");
			d.nome.focus();
			return false;
		}
		//validar telefone
		if (d.telefone.value == ""){
			alert("O campo " + d.telefone.name + " deve ser preenchido!");
			d.telefone.focus();
			return false;
		}
				//validar email
		if (d.email.value == ""){
			alert("O campo " + d.email.name + " deve ser preenchido!");
			d.email.focus();
			return false;
		}
		//validar email(verificao de endereco eletronico)
		parte1 = d.email.value.indexOf("@");
		parte2 = d.email.value.indexOf(".");
		parte3 = d.email.value.length;
		if (!(parte1 >= 3 && parte2 >= 6 && parte3 >= 9)) {
			alert("O campo " + d.email.name + " deve ser conter um endereco eletronico!");
			d.email.focus();
			return false;
		}
}
</script>
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
    <td width="7" rowspan="2" background="../imagens/sombraesquerda.gif">&nbsp;</td>
    <td width="758" valign="top" bgcolor="#FFFFFF"><form action="suporte.asp?action=login" method="post" name="cadastro" id="cadastro"  onSubmit="return validaForm()">
      <table width="650" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td height="50" valign="middle"><img src="imagens/suporte.gif" width="87" height="22"></td>
        </tr>
        <tr>
          <td class="titulosadmin">&nbsp;</td>
          </tr>
        <tr>
          <td class="textoform">Nome:</td>
          </tr>
        <tr>
          <td><input name="nome" type="text" class="form" id="nome" size="50"></td>
          </tr>
        <tr>
          <td class="textoform">Telefone:</td>
          </tr>
        <tr>
          <td><input name="telefone" type="text" class="form" id="telefone" size="30"></td>
          </tr>
        <tr>
          <td class="textoform">E-mail:</td>
        </tr>
        <tr>
          <td class="textoform"><input name="email" type="text" class="form" id="email" size="40"></td>
        </tr>
        <tr>
          <td class="textoform">Sua d&uacute;vida/problema:</td>
          </tr>
        <tr>
          <td><textarea name="duvida" cols="80" rows="6" class="formtextos" id="duvida"></textarea></td>
          </tr>
        <tr>
          <td>&nbsp;</td>
          </tr>
        <tr>
          <td><input name="Limpar" type="reset" class="form" id="Limpar" value="Limpar">            <input name="enviar" type="submit" class="form" id="enviar" value="  Enviar  "></td>
          </tr>
        <tr>
          <td>&nbsp;</td>
          </tr>
      </table>
    </form>      </td>
    <td width="5" rowspan="2" background="../imagens/sombradireita.gif">&nbsp;</td>
  </tr>
  <tr>
    <td height="55" valign="bottom" bgcolor="#FFFFFF"><!--#include file="rodape.asp"--></td>
  </tr>
  <tr>
    <td height="8" colspan="3"><img src="../imagens/fundopagina.gif" width="768" height="8"></td>
  </tr>
</table>
</body>
</html>