<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<% Server.ScriptTimeout = 300 %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!--#include file="restrito.asp"-->
<!--#include file="conectar.asp"-->
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
.style1 {
	font-size: 9px;
	font-family: "Trebuchet MS", Verdana, Arial;
	color: #990000;
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
}
</script>

      <script language="VBScript">
SUB enviar_ONCLICK()
  cadastro.enviar.Value = "Aguarde enviando..."
END SUB
</script>
</head>
<body>
<!--#include file="topbar.asp"-->
<table width="770" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="300" background="../imagens/sombraesquerda.gif">
      <form action="action_inserir_data.asp" method="post" name="cadastro" id="cadastro"  onSubmit="return validaForm()">
        <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
          <tr>
            <td align="center" valign="top"><table width="770" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td height="50" colspan="4" align="left"><img src="imagens/inserirprofissional.gif" width="186" height="23"></td>
              </tr>
              <tr>
                <td width="173" align="left" class="textoform">M&ecirc;s de Cadastramento:</td>
                <td colspan="3" align="left" class="textoform">Ano de Cadastramento: </td>
              </tr>
              <tr>
                <td align="left"><select name="mes" class="form" id="mes">
                  <option value="Janeiro">Janeiro</option>
                  <option value="Fevereiro">Fevereiro</option>
                  <option value="Mar&ccedil;o">Mar&ccedil;o</option>
                  <option value="Abril">Abril</option>
                  <option value="Maio">Maio</option>
                  <option value="Junho">Junho</option>
                  <option value="Julho">Julho</option>
                  <option value="Agosto">Agosto</option>
                  <option value="Setembro">Setembro</option>
                  <option value="Outubro">Outubro</option>
                  <option value="Novembro">Novembro</option>
                  <option value="Dezembro">Dezembro</option>
                </select></td>
                <td colspan="3" align="left"><input name="ano" type="text" class="form" id="ano" size="6" maxlength="4" value="<%=year(Now)%>"></td>
              </tr>
              <tr valign="bottom">
                <td height="80" colspan="2" align="left" valign="middle" class="textoform">CAMPOS DOS GOYTACAZES<br></td>
                <td height="80" colspan="2" align="left" valign="middle" class="textoform">OUTRAS REGI&Oacute;ES</td>
                </tr>
              <tr valign="bottom">
                <td align="left" class="textoform">1&ordm; Data (IDA):</td>
                <td width="204" align="left" class="textoform">1&ordm; Data (VOLTA):</td>
                <td width="136" align="left" class="textoform">Data (IDA): </td>
                <td width="137" align="left" class="textoform">Data (VOLTA):</td>
              </tr>
              <tr valign="bottom">
                <td align="left" class="textoform"><input name="1dataida" type="text" class="form" id="1dataida" size="13" maxlength="13"></td>
                <td align="left" class="textoform"><input name="1datavolta" type="text" class="form" id="1datavolta" size="13" maxlength="13"></td>
                <td align="left" class="textoform"><input name="5dataida" type="text" class="form" id="5dataida" size="13" maxlength="13"></td>
                <td align="left" class="textoform"><input name="5datavolta" type="text" class="form" id="5datavolta" size="13" maxlength="13"></td>
              </tr>
              <tr valign="bottom">
                <td align="left" class="textoform">2&ordm; Data (IDA): </td>
                <td align="left" class="textoform">2&ordm; Data (VOLTA):</td>
                <td colspan="2" align="left" class="textoform">&nbsp;</td>
              </tr>
              <tr>
                <td align="left"><span class="textoform">
                  <input name="2dataida" type="text" class="form" id="2dataida" size="13" maxlength="13">
                  </span></td>
                <td align="left"><span class="textoform">
                  <input name="2datavolta" type="text" class="form" id="2datavolta" size="13" maxlength="13">
                  </span></td>
                <td colspan="2" align="left">&nbsp;</td>
              </tr>
              <tr valign="bottom">
                <td align="left" class="textoform">3&ordm; Data (IDA): </td>
                <td align="left" class="textoform">3&ordm; Data (VOLTA):</td>
                <td colspan="2" align="left" class="textoform">&nbsp;</td>
              </tr>
              <tr>
                <td align="left" valign="top"><span class="textoform">
                  <input name="3dataida" type="text" class="form" id="3dataida" size="13" maxlength="13">
                  </span></td>
                <td align="left" valign="top"><span class="textoform">
                  <input name="3datavolta" type="text" class="form" id="3datavolta" size="13" maxlength="13">
                  </span></td>
                <td colspan="2" align="left" valign="top">&nbsp;</td>
              </tr>
              <tr valign="bottom">
                <td align="left" class="textoform">4&ordm; Data (IDA): </td>
                <td align="left" class="textoform">4&ordm; Data (VOLTA):</td>
                <td colspan="2" align="left" class="textoform">&nbsp;</td>
              </tr>
              <tr>
                <td align="left" valign="top"><span class="textoform">
                  <input name="4dataida" type="text" class="form" id="4dataida" size="13" maxlength="13">
                  </span></td>
                <td align="left" valign="top"><span class="textoform">
                  <input name="4datavolta" type="text" class="form" id="4datavolta" size="13" maxlength="13">
                  </span></td>
                <td colspan="2" align="left" valign="top">&nbsp;</td>
              </tr>
              <tr>
                <td height="80" colspan="4" align="left"><span class="textoform">(coloque apenas os dias do m&ecirc;s, o m&ecirc;s e o ano ser&aacute;o cadastrados automaticamente, se houver mais de uma data coloque xx-xx. Ex. 01-05/00/2009)</span><br></td>
                </tr>
              <tr>
                <td height="50" align="left"><div align="right">
                  <input name="Submit2" type="reset" class="form" value="Limpar">
                  &nbsp; </div></td>
                <td colspan="3" align="left">&nbsp;
                  <input name="enviar" type="submit" class="form" id="enviar" value="Enviar"></td>
              </tr>
            </table></td>
          </tr>
        </table>
    </form>    </td>
  </tr>
</table>
</body>
</html>