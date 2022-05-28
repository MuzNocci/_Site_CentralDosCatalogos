<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!--#include file="restrito.asp"-->
<!--#include file="conectar.asp"-->
<html>
<head>
<title>Administra&ccedil;&atilde;o</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="..site/estilo/site.css" rel="stylesheet" type="text/css">
</head>
<body>
<!--#include file="topbar.asp"-->
<table width="934" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="300" background="../imagens/sombraesquerda.gif">
        <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
          <tr>
            <td align="center" valign="top"> 
              <table width="770" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td height="50" align="left">
					<font face="Arial" color="#CC0000">
					<span style="text-transform: uppercase; font-weight:700">Altere as datas dos pedidos</span></font></td>
                </tr>
                <%if request.querystring("save") = "ok" then%>
                <tr>
                  <td height="50" align="center" bgcolor="#006600" bordercolor="#003300" style="border-style: solid; border-width: 1px">
					<b><font size="2" face="Arial" color="#FFFFFF">Suas 
					alterações foram salvas com sucesso!</font></b></td>
                </tr>
                	<tr>
		            <td height="10">
		            <span style="font-size: 1pt">&nbsp;</span></td>
	            </tr>
	            <%end if%>
                <tr>
                  <td>
<table border="0" width="100%" cellspacing="0" cellpadding="2" id="table1">
	<tr>
		<td colspan="3" height="10" bgcolor="#F5F5F5">
		<span style="font-size: 1pt">&nbsp;</span></td>
	</tr>
	<form method="POST" action="up.datas.asp">
	<input type="hidden" name="revista" id="revista" value="C1">
	<tr>
		<td colspan="3" height="50" bgcolor="#F5F5F5">
		<img border="0" src="../site/imagem/C1.jpg" width="934" height="60"></td>
	</tr>
	<tr>
		<td colspan="3" height="10" bgcolor="#F5F5F5">
		<span style="font-size: 1pt">&nbsp;</span></td>
	</tr>
	<tr>
		<td width="33%" bgcolor="#F5F5F5" valign="top" align="center">
<table border="0" width="280" cellspacing="0" cellpadding="5" id="table1" style="border-width: 0px">
	<tr>
		<td width="81" height="35" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="2" face="Arial">Mês</font></b></td>
		<td colspan="3" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font face="Arial" color="#CC0000"><%
		mes = month(Now)
		select case mes
		case 1
		mmes = "Janeiro"
		case 2
		mmes = "Fevereiro"
		case 3
		mmes = "Março"
		case 4
		mmes = "Abril"
		case 5
		mmes = "Maio"
		case 6
		mmes = "Junho"
		case 7
		mmes = "Julho"
		case 8
		mmes = "Agosto"
		case 9
		mmes = "Setembro"
		case 10
		mmes = "Outubro"
		case 11
		mmes = "Novembro"
		case 12
		mmes = "Dezembro"
		end select
		response.write mmes		
		%>
		</font></b>
		</td>
	</tr>
	<tr>
		<td width="81" height="20" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b>
		<font size="1" face="Arial">Data</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Código</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vai</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vem</font></b></td>
	</tr>
<%
for u = 1 to 5
Set c1 = conexao.execute("SELECT * FROM datas WHERE iddata = 'C1"&u&"' and mes = "&mes&" order by iddata DESC")
if not c1.eof then
%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C1<%=u%>0cod" name="C1<%=u%>0cod" size="1" tabindex="2" value="<%=c1("codcat")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C1<%=u%>0vai" id="C1<%=u%>0vai" size="7" tabindex="3"  value="<%=c1("datavai")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C1<%=u%>0vem" id="C1<%=u%>0vem" size="7" tabindex="3"  value="<%=c1("datavem")%>"></td>
	</tr>
<%
else%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C1<%=u%>0cod" name="C1<%=u%>0cod" size="1" tabindex="2" value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C1<%=u%>0vai" id="C1<%=u%>0vai" size="7" tabindex="3"  value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C1<%=u%>0vem" id="C1<%=u%>0vem" size="7" tabindex="3"  value=""></td>
	</tr>
<%
end if
c1.close
Set c1 = Nothing
next
%>
</table></td>
		<td width="33%" bgcolor="#F5F5F5" valign="top" align="center">
<table border="0" width="280" cellspacing="0" cellpadding="5" id="table1" style="border-width: 0px">
	<tr>
		<td width="81" height="35" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="2" face="Arial">Mês</font></b></td>
		<td colspan="3" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font face="Arial" color="#CC0000"><%
		mes = month(Now)+1
		if mes > 12 then
		mes = mes-12
		end if
		select case mes
		case 1
		mmes = "Janeiro"
		case 2
		mmes = "Fevereiro"
		case 3
		mmes = "Março"
		case 4
		mmes = "Abril"
		case 5
		mmes = "Maio"
		case 6
		mmes = "Junho"
		case 7
		mmes = "Julho"
		case 8
		mmes = "Agosto"
		case 9
		mmes = "Setembro"
		case 10
		mmes = "Outubro"
		case 11
		mmes = "Novembro"
		case 12
		mmes = "Dezembro"
		end select
		response.write mmes		
		%>
		</font></b>
		</td>
	</tr>
	<tr>
		<td width="81" height="20" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b>
		<font size="1" face="Arial">Data</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Código</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vai</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vem</font></b></td>
	</tr>
<%
for u = 1 to 5
Set c1 = conexao.execute("SELECT * FROM datas WHERE iddata = 'C1"&u&"' and mes = "&mes&" order by iddata DESC")
if not c1.eof then
%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C1<%=u%>1cod" name="C1<%=u%>1cod" size="1" tabindex="2" value="<%=c1("codcat")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C1<%=u%>1vai" id="C1<%=u%>1vai" size="7" tabindex="3"  value="<%=c1("datavai")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C1<%=u%>1vem" id="C1<%=u%>1vem" size="7" tabindex="3"  value="<%=c1("datavem")%>"></td>
	</tr>
<%
else%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C1<%=u%>1cod" name="C1<%=u%>1cod" size="1" tabindex="2" value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C1<%=u%>1vai" id="C1<%=u%>1vai" size="7" tabindex="3"  value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C1<%=u%>1vem" id="C1<%=u%>1vem" size="7" tabindex="3"  value=""></td>
	</tr>
<%
end if
c1.close
Set c1 = Nothing
next
%>
</table>
		</td>
		<td width="33%" bgcolor="#F5F5F5" valign="top" align="center">
<table border="0" width="280" cellspacing="0" cellpadding="5" id="table1" style="border-width: 0px">
	<tr>
		<td width="81" height="35" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="2" face="Arial">Mês</font></b></td>
		<td colspan="3" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font face="Arial" color="#CC0000"><%
		mes = month(Now)+2
		if mes > 12 then
		mes = mes-12
		end if
		select case mes
		case 1
		mmes = "Janeiro"
		case 2
		mmes = "Fevereiro"
		case 3
		mmes = "Março"
		case 4
		mmes = "Abril"
		case 5
		mmes = "Maio"
		case 6
		mmes = "Junho"
		case 7
		mmes = "Julho"
		case 8
		mmes = "Agosto"
		case 9
		mmes = "Setembro"
		case 10
		mmes = "Outubro"
		case 11
		mmes = "Novembro"
		case 12
		mmes = "Dezembro"
		end select
		response.write mmes		
		%>
		</font></b>
		</td>
	</tr>
	<tr>
		<td width="81" height="20" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b>
		<font size="1" face="Arial">Data</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Código</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vai</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vem</font></b></td>
	</tr>
<%
for u = 1 to 5
Set c1 = conexao.execute("SELECT * FROM datas WHERE iddata = 'C1"&u&"' and mes = "&mes&" order by iddata DESC")
if not c1.eof then
%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C1<%=u%>2cod" name="C1<%=u%>2cod" size="1" tabindex="2" value="<%=c1("codcat")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C1<%=u%>2vai" id="C1<%=u%>2vai" size="7" tabindex="3"  value="<%=c1("datavai")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C1<%=u%>2vem" id="C1<%=u%>2vem" size="7" tabindex="3"  value="<%=c1("datavem")%>"></td>
	</tr>
<%
else%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C1<%=u%>2cod" name="C1<%=u%>2cod" size="1" tabindex="2" value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C1<%=u%>2vai" id="C1<%=u%>2vai" size="7" tabindex="3"  value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C1<%=u%>2vem" id="C1<%=u%>2vem" size="7" tabindex="3"  value=""></td>
	</tr>
<%
end if
c1.close
Set c1 = Nothing
next
%>
</table>		</td>
	</tr>
		<tr>
		<td colspan="3" height="10" bgcolor="#F5F5F5">
		<span style="font-size: 1pt">&nbsp;</span></td>
	</tr>
		<tr>
		<td colspan="3" height="60" bgcolor="#F5F5F5" align="right">
		<input type="image" src="../site/imagem/salvar.png">&nbsp;&nbsp;&nbsp; </td>
	</tr>
	</form>
</table>
<table border="0" width="100%" cellspacing="0" cellpadding="2" id="table1">
	<tr>
		<td colspan="3" height="10">
		</td>
	</tr>
	<tr>
		<td colspan="3" height="10" bgcolor="#F5F5F5">
		<span style="font-size: 1pt">&nbsp;</span></td>
	</tr>
	<form method="POST" action="up.datas.asp">
	<input type="hidden" name="revista" id="revista" value="C2">
	<tr>
		<td colspan="3" height="50" bgcolor="#F5F5F5">
		<img border="0" src="../site/imagem/C2.jpg" width="934" height="60"></td>
	</tr>
	<tr>
		<td colspan="3" height="10" bgcolor="#F5F5F5">
		<span style="font-size: 1pt">&nbsp;</span></td>
	</tr>
	<tr>
		<td width="33%" bgcolor="#F5F5F5" valign="top" align="center">
<table border="0" width="280" cellspacing="0" cellpadding="5" id="table1" style="border-width: 0px">
	<tr>
		<td width="81" height="35" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="2" face="Arial">Mês</font></b></td>
		<td colspan="3" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font face="Arial" color="#CC0000"><%
		mes = month(Now)
		select case mes
		case 1
		mmes = "Janeiro"
		case 2
		mmes = "Fevereiro"
		case 3
		mmes = "Março"
		case 4
		mmes = "Abril"
		case 5
		mmes = "Maio"
		case 6
		mmes = "Junho"
		case 7
		mmes = "Julho"
		case 8
		mmes = "Agosto"
		case 9
		mmes = "Setembro"
		case 10
		mmes = "Outubro"
		case 11
		mmes = "Novembro"
		case 12
		mmes = "Dezembro"
		end select
		response.write mmes		
		%>
		</font></b>
		</td>
	</tr>
	<tr>
		<td width="81" height="20" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b>
		<font size="1" face="Arial">Data</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Código</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vai</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vem</font></b></td>
	</tr>
<%
for u = 1 to 5
Set c1 = conexao.execute("SELECT * FROM datas WHERE iddata = 'C2"&u&"' and mes = "&mes&" order by iddata DESC")
if not c1.eof then
%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C2<%=u%>0cod" name="C2<%=u%>0cod" size="1" tabindex="2" value="<%=c1("codcat")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C2<%=u%>0vai" id="C2<%=u%>0vai" size="7" tabindex="3"  value="<%=c1("datavai")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C2<%=u%>0vem" id="C2<%=u%>0vem" size="7" tabindex="3"  value="<%=c1("datavem")%>"></td>
	</tr>
<%
else%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C2<%=u%>0cod" name="C2<%=u%>0cod" size="1" tabindex="2" value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C2<%=u%>0vai" id="C2<%=u%>0vai" size="7" tabindex="3"  value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C2<%=u%>0vem" id="C2<%=u%>0vem" size="7" tabindex="3"  value=""></td>
	</tr>
<%
end if
c1.close
Set c1 = Nothing
next
%>
</table></td>
		<td width="33%" bgcolor="#F5F5F5" valign="top" align="center">
<table border="0" width="280" cellspacing="0" cellpadding="5" id="table1" style="border-width: 0px">
	<tr>
		<td width="81" height="35" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="2" face="Arial">Mês</font></b></td>
		<td colspan="3" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font face="Arial" color="#CC0000"><%
		mes = month(Now)+1
		if mes > 12 then
		mes = mes-12
		end if
		select case mes
		case 1
		mmes = "Janeiro"
		case 2
		mmes = "Fevereiro"
		case 3
		mmes = "Março"
		case 4
		mmes = "Abril"
		case 5
		mmes = "Maio"
		case 6
		mmes = "Junho"
		case 7
		mmes = "Julho"
		case 8
		mmes = "Agosto"
		case 9
		mmes = "Setembro"
		case 10
		mmes = "Outubro"
		case 11
		mmes = "Novembro"
		case 12
		mmes = "Dezembro"
		end select
		response.write mmes		
		%>
		</font></b>
		</td>
	</tr>
	<tr>
		<td width="81" height="20" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b>
		<font size="1" face="Arial">Data</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Código</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vai</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vem</font></b></td>
	</tr>
<%
for u = 1 to 5
Set c1 = conexao.execute("SELECT * FROM datas WHERE iddata = 'C2"&u&"' and mes = "&mes&" order by iddata DESC")
if not c1.eof then
%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C2<%=u%>1cod" name="C2<%=u%>1cod" size="1" tabindex="2" value="<%=c1("codcat")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C2<%=u%>1vai" id="C2<%=u%>1vai" size="7" tabindex="3"  value="<%=c1("datavai")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C2<%=u%>1vem" id="C2<%=u%>1vem" size="7" tabindex="3"  value="<%=c1("datavem")%>"></td>
	</tr>
<%
else%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C2<%=u%>1cod" name="C2<%=u%>1cod" size="1" tabindex="2" value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C2<%=u%>1vai" id="C2<%=u%>1vai" size="7" tabindex="3"  value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C2<%=u%>1vem" id="C2<%=u%>1vem" size="7" tabindex="3"  value=""></td>
	</tr>
<%
end if
c1.close
Set c1 = Nothing
next
%>
</table>
		</td>
		<td width="33%" bgcolor="#F5F5F5" valign="top" align="center">
<table border="0" width="280" cellspacing="0" cellpadding="5" id="table1" style="border-width: 0px">
	<tr>
		<td width="81" height="35" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="2" face="Arial">Mês</font></b></td>
		<td colspan="3" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font face="Arial" color="#CC0000"><%
		mes = month(Now)+2
		if mes > 12 then
		mes = mes-12
		end if
		select case mes
		case 1
		mmes = "Janeiro"
		case 2
		mmes = "Fevereiro"
		case 3
		mmes = "Março"
		case 4
		mmes = "Abril"
		case 5
		mmes = "Maio"
		case 6
		mmes = "Junho"
		case 7
		mmes = "Julho"
		case 8
		mmes = "Agosto"
		case 9
		mmes = "Setembro"
		case 10
		mmes = "Outubro"
		case 11
		mmes = "Novembro"
		case 12
		mmes = "Dezembro"
		end select
		response.write mmes		
		%>
		</font></b>
		</td>
	</tr>
	<tr>
		<td width="81" height="20" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b>
		<font size="1" face="Arial">Data</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Código</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vai</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vem</font></b></td>
	</tr>
<%
for u = 1 to 5
Set c1 = conexao.execute("SELECT * FROM datas WHERE iddata = 'C2"&u&"' and mes = "&mes&" order by iddata DESC")
if not c1.eof then
%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C2<%=u%>2cod" name="C2<%=u%>2cod" size="1" tabindex="2" value="<%=c1("codcat")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C2<%=u%>2vai" id="C2<%=u%>2vai" size="7" tabindex="3"  value="<%=c1("datavai")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C2<%=u%>2vem" id="C2<%=u%>2vem" size="7" tabindex="3"  value="<%=c1("datavem")%>"></td>
	</tr>
<%
else%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C2<%=u%>2cod" name="C2<%=u%>2cod" size="1" tabindex="2" value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C2<%=u%>2vai" id="C2<%=u%>2vai" size="7" tabindex="3"  value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C2<%=u%>2vem" id="C2<%=u%>2vem" size="7" tabindex="3"  value=""></td>
	</tr>
<%
end if
c1.close
Set c1 = Nothing
next
%>
</table>		</td>
	</tr>
		<tr>
		<td colspan="3" height="10" bgcolor="#F5F5F5">
		<span style="font-size: 1pt">&nbsp;</span></td>
	</tr>
		<tr>
		<td colspan="3" height="60" bgcolor="#F5F5F5" align="right">
		<input type="image" src="../site/imagem/salvar.png">&nbsp;&nbsp;&nbsp; </td>
	</tr>
	</form>
</table>
<!--
<table border="0" width="100%" cellspacing="0" cellpadding="2" id="table1">
	<tr>
		<td colspan="3" height="10">
		</td>
	</tr>
	<tr>
		<td colspan="3" height="10" bgcolor="#F5F5F5">
		<span style="font-size: 1pt">&nbsp;</span></td>
	</tr>
	<form method="POST" action="up.datas.asp">
	<input type="hidden" name="revista" id="revista" value="C3">
	<tr>
		<td colspan="3" height="50" bgcolor="#F5F5F5">
		<img border="0" src="../site/imagem/C3.jpg" width="934" height="60"></td>
	</tr>
	<tr>
		<td colspan="3" height="10" bgcolor="#F5F5F5">
		<span style="font-size: 1pt">&nbsp;</span></td>
	</tr>
	<tr>
		<td width="33%" bgcolor="#F5F5F5" valign="top" align="center">
<table border="0" width="280" cellspacing="0" cellpadding="5" id="table1" style="border-width: 0px">
	<tr>
		<td width="81" height="35" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="2" face="Arial">Mês</font></b></td>
		<td colspan="3" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font face="Arial" color="#CC0000"><%
		mes = month(Now)
		select case mes
		case 1
		mmes = "Janeiro"
		case 2
		mmes = "Fevereiro"
		case 3
		mmes = "Março"
		case 4
		mmes = "Abril"
		case 5
		mmes = "Maio"
		case 6
		mmes = "Junho"
		case 7
		mmes = "Julho"
		case 8
		mmes = "Agosto"
		case 9
		mmes = "Setembro"
		case 10
		mmes = "Outubro"
		case 11
		mmes = "Novembro"
		case 12
		mmes = "Dezembro"
		end select
		response.write mmes		
		%>
		</font></b>
		</td>
	</tr>
	<tr>
		<td width="81" height="20" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b>
		<font size="1" face="Arial">Data</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Código</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vai</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vem</font></b></td>
	</tr>
<%
for u = 1 to 5
Set c1 = conexao.execute("SELECT * FROM datas WHERE iddata = 'C3"&u&"' and mes = "&mes&" order by iddata DESC")
if not c1.eof then
%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C3<%=u%>0cod" name="C3<%=u%>0cod" size="1" tabindex="2" value="<%=c1("codcat")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C3<%=u%>0vai" id="C3<%=u%>0vai" size="7" tabindex="3"  value="<%=c1("datavai")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C3<%=u%>0vem" id="C3<%=u%>0vem" size="7" tabindex="3"  value="<%=c1("datavem")%>"></td>
	</tr>
<%
else%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C3<%=u%>0cod" name="C3<%=u%>0cod" size="1" tabindex="2" value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C3<%=u%>0vai" id="C3<%=u%>0vai" size="7" tabindex="3"  value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C3<%=u%>0vem" id="C3<%=u%>0vem" size="7" tabindex="3"  value=""></td>
	</tr>
<%
end if
c1.close
Set c1 = Nothing
next
%>
</table></td>
		<td width="33%" bgcolor="#F5F5F5" valign="top" align="center">
<table border="0" width="280" cellspacing="0" cellpadding="5" id="table1" style="border-width: 0px">
	<tr>
		<td width="81" height="35" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="2" face="Arial">Mês</font></b></td>
		<td colspan="3" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font face="Arial" color="#CC0000"><%
		mes = month(Now)+1
		select case mes
		case 1
		mmes = "Janeiro"
		case 2
		mmes = "Fevereiro"
		case 3
		mmes = "Março"
		case 4
		mmes = "Abril"
		case 5
		mmes = "Maio"
		case 6
		mmes = "Junho"
		case 7
		mmes = "Julho"
		case 8
		mmes = "Agosto"
		case 9
		mmes = "Setembro"
		case 10
		mmes = "Outubro"
		case 11
		mmes = "Novembro"
		case 12
		mmes = "Dezembro"
		end select
		response.write mmes		
		%>
		</font></b>
		</td>
	</tr>
	<tr>
		<td width="81" height="20" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b>
		<font size="1" face="Arial">Data</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Código</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vai</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vem</font></b></td>
	</tr>
<%
for u = 1 to 5
Set c1 = conexao.execute("SELECT * FROM datas WHERE iddata = 'C3"&u&"' and mes = "&mes&" order by iddata DESC")
if not c1.eof then
%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C3<%=u%>1cod" name="C3<%=u%>1cod" size="1" tabindex="2" value="<%=c1("codcat")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C3<%=u%>1vai" id="C3<%=u%>1vai" size="7" tabindex="3"  value="<%=c1("datavai")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C3<%=u%>1vem" id="C3<%=u%>1vem" size="7" tabindex="3"  value="<%=c1("datavem")%>"></td>
	</tr>
<%
else%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C3<%=u%>1cod" name="C3<%=u%>1cod" size="1" tabindex="2" value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C3<%=u%>1vai" id="C3<%=u%>1vai" size="7" tabindex="3"  value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C3<%=u%>1vem" id="C3<%=u%>1vem" size="7" tabindex="3"  value=""></td>
	</tr>
<%
end if
c1.close
Set c1 = Nothing
next
%>
</table>
		</td>
		<td width="33%" bgcolor="#F5F5F5" valign="top" align="center">
<table border="0" width="280" cellspacing="0" cellpadding="5" id="table1" style="border-width: 0px">
	<tr>
		<td width="81" height="35" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="2" face="Arial">Mês</font></b></td>
		<td colspan="3" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font face="Arial" color="#CC0000"><%
		mes = month(Now)+2
		select case mes
		case 1
		mmes = "Janeiro"
		case 2
		mmes = "Fevereiro"
		case 3
		mmes = "Março"
		case 4
		mmes = "Abril"
		case 5
		mmes = "Maio"
		case 6
		mmes = "Junho"
		case 7
		mmes = "Julho"
		case 8
		mmes = "Agosto"
		case 9
		mmes = "Setembro"
		case 10
		mmes = "Outubro"
		case 11
		mmes = "Novembro"
		case 12
		mmes = "Dezembro"
		end select
		response.write mmes		
		%>
		</font></b>
		</td>
	</tr>
	<tr>
		<td width="81" height="20" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b>
		<font size="1" face="Arial">Data</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Código</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vai</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vem</font></b></td>
	</tr>
<%
for u = 1 to 5
Set c1 = conexao.execute("SELECT * FROM datas WHERE iddata = 'C3"&u&"' and mes = "&mes&" order by iddata DESC")
if not c1.eof then
%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C3<%=u%>2cod" name="C3<%=u%>2cod" size="1" tabindex="2" value="<%=c1("codcat")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C3<%=u%>2vai" id="C3<%=u%>2vai" size="7" tabindex="3"  value="<%=c1("datavai")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C3<%=u%>2vem" id="C3<%=u%>2vem" size="7" tabindex="3"  value="<%=c1("datavem")%>"></td>
	</tr>
<%
else%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C3<%=u%>2cod" name="C3<%=u%>2cod" size="1" tabindex="2" value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C3<%=u%>2vai" id="C3<%=u%>2vai" size="7" tabindex="3"  value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C3<%=u%>2vem" id="C3<%=u%>2vem" size="7" tabindex="3"  value=""></td>
	</tr>
<%
end if
c1.close
Set c1 = Nothing
next
%>
</table>		</td>
	</tr>
		<tr>
		<td colspan="3" height="10" bgcolor="#F5F5F5">
		<span style="font-size: 1pt">&nbsp;</span></td>
	</tr>
		<tr>
		<td colspan="3" height="60" bgcolor="#F5F5F5" align="right">
		<input type="image" src="../site/imagem/salvar.png">&nbsp;&nbsp;&nbsp; </td>
	</tr>
	</form>
</table>
-->
<table border="0" width="100%" cellspacing="0" cellpadding="2" id="table1">
	<tr>
		<td colspan="3" height="10">
		</td>
	</tr>
	<tr>
		<td colspan="3" height="10" bgcolor="#F5F5F5">
		<span style="font-size: 1pt">&nbsp;</span></td>
	</tr>
	<form method="POST" action="up.datas.asp">
	<input type="hidden" name="revista" id="revista" value="C4">
	<tr>
		<td colspan="3" height="50" bgcolor="#F5F5F5">
		<img border="0" src="../site/imagem/C4.jpg" width="934" height="60"></td>
	</tr>
	<tr>
		<td colspan="3" height="10" bgcolor="#F5F5F5">
		<span style="font-size: 1pt">&nbsp;</span></td>
	</tr>
	<tr>
		<td width="33%" bgcolor="#F5F5F5" valign="top" align="center">
<table border="0" width="280" cellspacing="0" cellpadding="5" id="table1" style="border-width: 0px">
	<tr>
		<td width="81" height="35" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="2" face="Arial">Mês</font></b></td>
		<td colspan="3" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font face="Arial" color="#CC0000"><%
		mes = month(Now)
		select case mes
		case 1
		mmes = "Janeiro"
		case 2
		mmes = "Fevereiro"
		case 3
		mmes = "Março"
		case 4
		mmes = "Abril"
		case 5
		mmes = "Maio"
		case 6
		mmes = "Junho"
		case 7
		mmes = "Julho"
		case 8
		mmes = "Agosto"
		case 9
		mmes = "Setembro"
		case 10
		mmes = "Outubro"
		case 11
		mmes = "Novembro"
		case 12
		mmes = "Dezembro"
		end select
		response.write mmes		
		%>
		</font></b>
		</td>
	</tr>
	<tr>
		<td width="81" height="20" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b>
		<font size="1" face="Arial">Data</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Código</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vai</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vem</font></b></td>
	</tr>
<%
for u = 1 to 5
Set c1 = conexao.execute("SELECT * FROM datas WHERE iddata = 'C4"&u&"' and mes = "&mes&" order by iddata DESC")
if not c1.eof then
%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C4<%=u%>0cod" name="C4<%=u%>0cod" size="1" tabindex="2" value="<%=c1("codcat")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C4<%=u%>0vai" id="C4<%=u%>0vai" size="7" tabindex="3"  value="<%=c1("datavai")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C4<%=u%>0vem" id="C4<%=u%>0vem" size="7" tabindex="3"  value="<%=c1("datavem")%>"></td>
	</tr>
<%
else%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C4<%=u%>0cod" name="C4<%=u%>0cod" size="1" tabindex="2" value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C4<%=u%>0vai" id="C4<%=u%>0vai" size="7" tabindex="3"  value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C4<%=u%>0vem" id="C4<%=u%>0vem" size="7" tabindex="3"  value=""></td>
	</tr>
<%
end if
c1.close
Set c1 = Nothing
next
%>
</table></td>
		<td width="33%" bgcolor="#F5F5F5" valign="top" align="center">
<table border="0" width="280" cellspacing="0" cellpadding="5" id="table1" style="border-width: 0px">
	<tr>
		<td width="81" height="35" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="2" face="Arial">Mês</font></b></td>
		<td colspan="3" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font face="Arial" color="#CC0000"><%
		mes = month(Now)+1
		if mes > 12 then
		mes = mes-12
		end if
		select case mes
		case 1
		mmes = "Janeiro"
		case 2
		mmes = "Fevereiro"
		case 3
		mmes = "Março"
		case 4
		mmes = "Abril"
		case 5
		mmes = "Maio"
		case 6
		mmes = "Junho"
		case 7
		mmes = "Julho"
		case 8
		mmes = "Agosto"
		case 9
		mmes = "Setembro"
		case 10
		mmes = "Outubro"
		case 11
		mmes = "Novembro"
		case 12
		mmes = "Dezembro"
		end select
		response.write mmes		
		%>
		</font></b>
		</td>
	</tr>
	<tr>
		<td width="81" height="20" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b>
		<font size="1" face="Arial">Data</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Código</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vai</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vem</font></b></td>
	</tr>
<%
for u = 1 to 5
Set c1 = conexao.execute("SELECT * FROM datas WHERE iddata = 'C4"&u&"' and mes = "&mes&" order by iddata DESC")
if not c1.eof then
%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C4<%=u%>1cod" name="C4<%=u%>1cod" size="1" tabindex="2" value="<%=c1("codcat")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C4<%=u%>1vai" id="C4<%=u%>1vai" size="7" tabindex="3"  value="<%=c1("datavai")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C4<%=u%>1vem" id="C4<%=u%>1vem" size="7" tabindex="3"  value="<%=c1("datavem")%>"></td>
	</tr>
<%
else%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C4<%=u%>1cod" name="C4<%=u%>1cod" size="1" tabindex="2" value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C4<%=u%>1vai" id="C4<%=u%>1vai" size="7" tabindex="3"  value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C4<%=u%>1vem" id="C4<%=u%>1vem" size="7" tabindex="3"  value=""></td>
	</tr>
<%
end if
c1.close
Set c1 = Nothing
next
%>
</table>
		</td>
		<td width="33%" bgcolor="#F5F5F5" valign="top" align="center">
<table border="0" width="280" cellspacing="0" cellpadding="5" id="table1" style="border-width: 0px">
	<tr>
		<td width="81" height="35" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="2" face="Arial">Mês</font></b></td>
		<td colspan="3" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font face="Arial" color="#CC0000"><%
		mes = month(Now)+2
		if mes > 12 then
		mes = mes-12
		end if
		select case mes
		case 1
		mmes = "Janeiro"
		case 2
		mmes = "Fevereiro"
		case 3
		mmes = "Março"
		case 4
		mmes = "Abril"
		case 5
		mmes = "Maio"
		case 6
		mmes = "Junho"
		case 7
		mmes = "Julho"
		case 8
		mmes = "Agosto"
		case 9
		mmes = "Setembro"
		case 10
		mmes = "Outubro"
		case 11
		mmes = "Novembro"
		case 12
		mmes = "Dezembro"
		end select
		response.write mmes		
		%>
		</font></b>
		</td>
	</tr>
	<tr>
		<td width="81" height="20" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b>
		<font size="1" face="Arial">Data</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Código</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vai</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vem</font></b></td>
	</tr>
<%
for u = 1 to 5
Set c1 = conexao.execute("SELECT * FROM datas WHERE iddata = 'C4"&u&"' and mes = "&mes&" order by iddata DESC")
if not c1.eof then
%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C4<%=u%>2cod" name="C4<%=u%>2cod" size="1" tabindex="2" value="<%=c1("codcat")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C4<%=u%>2vai" id="C4<%=u%>2vai" size="7" tabindex="3"  value="<%=c1("datavai")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C4<%=u%>2vem" id="C4<%=u%>2vem" size="7" tabindex="3"  value="<%=c1("datavem")%>"></td>
	</tr>
<%
else%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C4<%=u%>2cod" name="C4<%=u%>2cod" size="1" tabindex="2" value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C4<%=u%>2vai" id="C4<%=u%>2vai" size="7" tabindex="3"  value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C4<%=u%>2vem" id="C4<%=u%>2vem" size="7" tabindex="3"  value=""></td>
	</tr>
<%
end if
c1.close
Set c1 = Nothing
next
%>
</table>		</td>
	</tr>
		<tr>
		<td colspan="3" height="10" bgcolor="#F5F5F5">
		<span style="font-size: 1pt">&nbsp;</span></td>
	</tr>
		<tr>
		<td colspan="3" height="60" bgcolor="#F5F5F5" align="right">
		<input type="image" src="../site/imagem/salvar.png">&nbsp;&nbsp;&nbsp; </td>
	</tr>
	</form>
</table>
<table border="0" width="100%" cellspacing="0" cellpadding="2" id="table1">
	<tr>
		<td colspan="3" height="10">
		</td>
	</tr>
	<tr>
		<td colspan="3" height="10" bgcolor="#F5F5F5">
		<span style="font-size: 1pt">&nbsp;</span></td>
	</tr>
	<form method="POST" action="up.datas.asp">
	<input type="hidden" name="revista" id="revista" value="C5">
	<tr>
		<td colspan="3" height="50" bgcolor="#F5F5F5">
		<img border="0" src="../site/imagem/C5.jpg" width="934" height="60"></td>
	</tr>
	<tr>
		<td colspan="3" height="10" bgcolor="#F5F5F5">
		<span style="font-size: 1pt">&nbsp;</span></td>
	</tr>
	<tr>
		<td width="33%" bgcolor="#F5F5F5" valign="top" align="center">
<table border="0" width="280" cellspacing="0" cellpadding="5" id="table1" style="border-width: 0px">
	<tr>
		<td width="81" height="35" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="2" face="Arial">Mês</font></b></td>
		<td colspan="3" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font face="Arial" color="#CC0000"><%
		mes = month(Now)
		select case mes
		case 1
		mmes = "Janeiro"
		case 2
		mmes = "Fevereiro"
		case 3
		mmes = "Março"
		case 4
		mmes = "Abril"
		case 5
		mmes = "Maio"
		case 6
		mmes = "Junho"
		case 7
		mmes = "Julho"
		case 8
		mmes = "Agosto"
		case 9
		mmes = "Setembro"
		case 10
		mmes = "Outubro"
		case 11
		mmes = "Novembro"
		case 12
		mmes = "Dezembro"
		end select
		response.write mmes		
		%>
		</font></b>
		</td>
	</tr>
	<tr>
		<td width="81" height="20" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b>
		<font size="1" face="Arial">Data</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Código</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vai</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vem</font></b></td>
	</tr>
<%
for u = 1 to 5
Set c1 = conexao.execute("SELECT * FROM datas WHERE iddata = 'C5"&u&"' and mes = "&mes&" order by iddata DESC")
if not c1.eof then
%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C5<%=u%>0cod" name="C5<%=u%>0cod" size="1" tabindex="2" value="<%=c1("codcat")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C5<%=u%>0vai" id="C5<%=u%>0vai" size="7" tabindex="3"  value="<%=c1("datavai")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C5<%=u%>0vem" id="C5<%=u%>0vem" size="7" tabindex="3"  value="<%=c1("datavem")%>"></td>
	</tr>
<%
else%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C5<%=u%>0cod" name="C5<%=u%>0cod" size="1" tabindex="2" value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C5<%=u%>0vai" id="C5<%=u%>0vai" size="7" tabindex="3"  value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C5<%=u%>0vem" id="C5<%=u%>0vem" size="7" tabindex="3"  value=""></td>
	</tr>
<%
end if
c1.close
Set c1 = Nothing
next
%>
</table></td>
		<td width="33%" bgcolor="#F5F5F5" valign="top" align="center">
<table border="0" width="280" cellspacing="0" cellpadding="5" id="table1" style="border-width: 0px">
	<tr>
		<td width="81" height="35" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="2" face="Arial">Mês</font></b></td>
		<td colspan="3" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font face="Arial" color="#CC0000"><%
		mes = month(Now)+1
		if mes > 12 then
		mes = mes-12
		end if
		select case mes
		case 1
		mmes = "Janeiro"
		case 2
		mmes = "Fevereiro"
		case 3
		mmes = "Março"
		case 4
		mmes = "Abril"
		case 5
		mmes = "Maio"
		case 6
		mmes = "Junho"
		case 7
		mmes = "Julho"
		case 8
		mmes = "Agosto"
		case 9
		mmes = "Setembro"
		case 10
		mmes = "Outubro"
		case 11
		mmes = "Novembro"
		case 12
		mmes = "Dezembro"
		end select
		response.write mmes		
		%>
		</font></b>
		</td>
	</tr>
	<tr>
		<td width="81" height="20" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b>
		<font size="1" face="Arial">Data</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Código</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vai</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vem</font></b></td>
	</tr>
<%
for u = 1 to 5
Set c1 = conexao.execute("SELECT * FROM datas WHERE iddata = 'C5"&u&"' and mes = "&mes&" order by iddata DESC")
if not c1.eof then
%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C5<%=u%>1cod" name="C5<%=u%>1cod" size="1" tabindex="2" value="<%=c1("codcat")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C5<%=u%>1vai" id="C5<%=u%>1vai" size="7" tabindex="3"  value="<%=c1("datavai")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C5<%=u%>1vem" id="C5<%=u%>1vem" size="7" tabindex="3"  value="<%=c1("datavem")%>"></td>
	</tr>
<%
else%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C5<%=u%>1cod" name="C5<%=u%>1cod" size="1" tabindex="2" value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C5<%=u%>1vai" id="C5<%=u%>1vai" size="7" tabindex="3"  value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C5<%=u%>1vem" id="C5<%=u%>1vem" size="7" tabindex="3"  value=""></td>
	</tr>
<%
end if
c1.close
Set c1 = Nothing
next
%>
</table>
		</td>
		<td width="33%" bgcolor="#F5F5F5" valign="top" align="center">
<table border="0" width="280" cellspacing="0" cellpadding="5" id="table1" style="border-width: 0px">
	<tr>
		<td width="81" height="35" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="2" face="Arial">Mês</font></b></td>
		<td colspan="3" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font face="Arial" color="#CC0000"><%
		mes = month(Now)+2
		if mes > 12 then
		mes = mes-12
		end if
		select case mes
		case 1
		mmes = "Janeiro"
		case 2
		mmes = "Fevereiro"
		case 3
		mmes = "Março"
		case 4
		mmes = "Abril"
		case 5
		mmes = "Maio"
		case 6
		mmes = "Junho"
		case 7
		mmes = "Julho"
		case 8
		mmes = "Agosto"
		case 9
		mmes = "Setembro"
		case 10
		mmes = "Outubro"
		case 11
		mmes = "Novembro"
		case 12
		mmes = "Dezembro"
		end select
		response.write mmes		
		%>
		</font></b>
		</td>
	</tr>
	<tr>
		<td width="81" height="20" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b>
		<font size="1" face="Arial">Data</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Código</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vai</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vem</font></b></td>
	</tr>
<%
for u = 1 to 5
Set c1 = conexao.execute("SELECT * FROM datas WHERE iddata = 'C5"&u&"' and mes = "&mes&" order by iddata DESC")
if not c1.eof then
%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C5<%=u%>2cod" name="C5<%=u%>2cod" size="1" tabindex="2" value="<%=c1("codcat")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C5<%=u%>2vai" id="C5<%=u%>2vai" size="7" tabindex="3"  value="<%=c1("datavai")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C5<%=u%>2vem" id="C5<%=u%>2vem" size="7" tabindex="3"  value="<%=c1("datavem")%>"></td>
	</tr>
<%
else%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C5<%=u%>2cod" name="C5<%=u%>2cod" size="1" tabindex="2" value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C5<%=u%>2vai" id="C5<%=u%>2vai" size="7" tabindex="3"  value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C5<%=u%>2vem" id="C5<%=u%>2vem" size="7" tabindex="3"  value=""></td>
	</tr>
<%
end if
c1.close
Set c1 = Nothing
next
%>
</table>		</td>
	</tr>
		<tr>
		<td colspan="3" height="10" bgcolor="#F5F5F5">
		<span style="font-size: 1pt">&nbsp;</span></td>
	</tr>
		<tr>
		<td colspan="3" height="60" bgcolor="#F5F5F5" align="right">
		<input type="image" src="../site/imagem/salvar.png">&nbsp;&nbsp;&nbsp; </td>
	</tr>
	</form>
</table>
<table border="0" width="100%" cellspacing="0" cellpadding="2" id="table1">
	<tr>
		<td colspan="3" height="10">
		</td>
	</tr>
	<tr>
		<td colspan="3" height="10" bgcolor="#F5F5F5">
		<span style="font-size: 1pt">&nbsp;</span></td>
	</tr>
	<form method="POST" action="up.datas.asp">
	<input type="hidden" name="revista" id="revista" value="C6">
	<tr>
		<td colspan="3" height="50" bgcolor="#F5F5F5">
		<img border="0" src="../site/imagem/C6.jpg" width="934" height="60"></td>
	</tr>
	<tr>
		<td colspan="3" height="10" bgcolor="#F5F5F5">
		<span style="font-size: 1pt">&nbsp;</span></td>
	</tr>
	<tr>
		<td width="33%" bgcolor="#F5F5F5" valign="top" align="center">
<table border="0" width="280" cellspacing="0" cellpadding="5" id="table1" style="border-width: 0px">
	<tr>
		<td width="81" height="35" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="2" face="Arial">Mês</font></b></td>
		<td colspan="3" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font face="Arial" color="#CC0000"><%
		mes = month(Now)
		select case mes
		case 1
		mmes = "Janeiro"
		case 2
		mmes = "Fevereiro"
		case 3
		mmes = "Março"
		case 4
		mmes = "Abril"
		case 5
		mmes = "Maio"
		case 6
		mmes = "Junho"
		case 7
		mmes = "Julho"
		case 8
		mmes = "Agosto"
		case 9
		mmes = "Setembro"
		case 10
		mmes = "Outubro"
		case 11
		mmes = "Novembro"
		case 12
		mmes = "Dezembro"
		end select
		response.write mmes		
		%>
		</font></b>
		</td>
	</tr>
	<tr>
		<td width="81" height="20" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b>
		<font size="1" face="Arial">Data</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Código</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vai</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vem</font></b></td>
	</tr>
<%
for u = 1 to 5
Set c1 = conexao.execute("SELECT * FROM datas WHERE iddata = 'C6"&u&"' and mes = "&mes&" order by iddata DESC")
if not c1.eof then
%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C6<%=u%>0cod" name="C6<%=u%>0cod" size="1" tabindex="2" value="<%=c1("codcat")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C6<%=u%>0vai" id="C6<%=u%>0vai" size="7" tabindex="3"  value="<%=c1("datavai")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C6<%=u%>0vem" id="C6<%=u%>0vem" size="7" tabindex="3"  value="<%=c1("datavem")%>"></td>
	</tr>
<%
else%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C6<%=u%>0cod" name="C6<%=u%>0cod" size="1" tabindex="2" value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C6<%=u%>0vai" id="C6<%=u%>0vai" size="7" tabindex="3"  value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C6<%=u%>0vem" id="C6<%=u%>0vem" size="7" tabindex="3"  value=""></td>
	</tr>
<%
end if
c1.close
Set c1 = Nothing
next
%>
</table></td>
		<td width="33%" bgcolor="#F5F5F5" valign="top" align="center">
<table border="0" width="280" cellspacing="0" cellpadding="5" id="table1" style="border-width: 0px">
	<tr>
		<td width="81" height="35" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="2" face="Arial">Mês</font></b></td>
		<td colspan="3" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font face="Arial" color="#CC0000"><%
		mes = month(Now)+1
		if mes > 12 then
		mes = mes-12
		end if
		select case mes
		case 1
		mmes = "Janeiro"
		case 2
		mmes = "Fevereiro"
		case 3
		mmes = "Março"
		case 4
		mmes = "Abril"
		case 5
		mmes = "Maio"
		case 6
		mmes = "Junho"
		case 7
		mmes = "Julho"
		case 8
		mmes = "Agosto"
		case 9
		mmes = "Setembro"
		case 10
		mmes = "Outubro"
		case 11
		mmes = "Novembro"
		case 12
		mmes = "Dezembro"
		end select
		response.write mmes		
		%>
		</font></b>
		</td>
	</tr>
	<tr>
		<td width="81" height="20" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b>
		<font size="1" face="Arial">Data</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Código</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vai</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vem</font></b></td>
	</tr>
<%
for u = 1 to 5
Set c1 = conexao.execute("SELECT * FROM datas WHERE iddata = 'C6"&u&"' and mes = "&mes&" order by iddata DESC")
if not c1.eof then
%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C6<%=u%>1cod" name="C6<%=u%>1cod" size="1" tabindex="2" value="<%=c1("codcat")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C6<%=u%>1vai" id="C6<%=u%>1vai" size="7" tabindex="3"  value="<%=c1("datavai")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C6<%=u%>1vem" id="C6<%=u%>1vem" size="7" tabindex="3"  value="<%=c1("datavem")%>"></td>
	</tr>
<%
else%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C6<%=u%>1cod" name="C6<%=u%>1cod" size="1" tabindex="2" value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C6<%=u%>1vai" id="C6<%=u%>1vai" size="7" tabindex="3"  value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C6<%=u%>1vem" id="C6<%=u%>1vem" size="7" tabindex="3"  value=""></td>
	</tr>
<%
end if
c1.close
Set c1 = Nothing
next
%>
</table>
		</td>
		<td width="33%" bgcolor="#F5F5F5" valign="top" align="center">
<table border="0" width="280" cellspacing="0" cellpadding="5" id="table1" style="border-width: 0px">
	<tr>
		<td width="81" height="35" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="2" face="Arial">Mês</font></b></td>
		<td colspan="3" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font face="Arial" color="#CC0000"><%
		mes = month(Now)+2
		if mes > 12 then
		mes = mes-12
		end if
		select case mes
		case 1
		mmes = "Janeiro"
		case 2
		mmes = "Fevereiro"
		case 3
		mmes = "Março"
		case 4
		mmes = "Abril"
		case 5
		mmes = "Maio"
		case 6
		mmes = "Junho"
		case 7
		mmes = "Julho"
		case 8
		mmes = "Agosto"
		case 9
		mmes = "Setembro"
		case 10
		mmes = "Outubro"
		case 11
		mmes = "Novembro"
		case 12
		mmes = "Dezembro"
		end select
		response.write mmes		
		%>
		</font></b>
		</td>
	</tr>
	<tr>
		<td width="81" height="20" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b>
		<font size="1" face="Arial">Data</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Código</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vai</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vem</font></b></td>
	</tr>
<%
for u = 1 to 5
Set c1 = conexao.execute("SELECT * FROM datas WHERE iddata = 'C6"&u&"' and mes = "&mes&" order by iddata DESC")
if not c1.eof then
%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C6<%=u%>2cod" name="C6<%=u%>2cod" size="1" tabindex="2" value="<%=c1("codcat")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C6<%=u%>2vai" id="C6<%=u%>2vai" size="7" tabindex="3"  value="<%=c1("datavai")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C6<%=u%>2vem" id="C6<%=u%>2vem" size="7" tabindex="3"  value="<%=c1("datavem")%>"></td>
	</tr>
<%
else%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C6<%=u%>2cod" name="C6<%=u%>2cod" size="1" tabindex="2" value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C6<%=u%>2vai" id="C6<%=u%>2vai" size="7" tabindex="3"  value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C6<%=u%>2vem" id="C6<%=u%>2vem" size="7" tabindex="3"  value=""></td>
	</tr>
<%
end if
c1.close
Set c1 = Nothing
next
%>
</table>		</td>
	</tr>
		<tr>
		<td colspan="3" height="10" bgcolor="#F5F5F5">
		<span style="font-size: 1pt">&nbsp;</span></td>
	</tr>
		<tr>
		<td colspan="3" height="60" bgcolor="#F5F5F5" align="right">
		<input type="image" src="../site/imagem/salvar.png">&nbsp;&nbsp;&nbsp; </td>
	</tr>
	</form>
</table>
<table border="0" width="100%" cellspacing="0" cellpadding="2" id="table1">
	<tr>
		<td colspan="3" height="10">
		</td>
	</tr>
	<tr>
		<td colspan="3" height="10" bgcolor="#F5F5F5">
		<span style="font-size: 1pt">&nbsp;</span></td>
	</tr>
	<form method="POST" action="up.datas.asp">
	<input type="hidden" name="revista" id="revista" value="C7">
	<tr>
		<td colspan="3" height="50" bgcolor="#F5F5F5">
		<img border="0" src="../site/imagem/C7.jpg" width="934" height="60"></td>
	</tr>
	<tr>
		<td colspan="3" height="10" bgcolor="#F5F5F5">
		<span style="font-size: 1pt">&nbsp;</span></td>
	</tr>
	<tr>
		<td width="33%" bgcolor="#F5F5F5" valign="top" align="center">
<table border="0" width="280" cellspacing="0" cellpadding="5" id="table1" style="border-width: 0px">
	<tr>
		<td width="81" height="35" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="2" face="Arial">Mês</font></b></td>
		<td colspan="3" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font face="Arial" color="#CC0000"><%
		mes = month(Now)
		select case mes
		case 1
		mmes = "Janeiro"
		case 2
		mmes = "Fevereiro"
		case 3
		mmes = "Março"
		case 4
		mmes = "Abril"
		case 5
		mmes = "Maio"
		case 6
		mmes = "Junho"
		case 7
		mmes = "Julho"
		case 8
		mmes = "Agosto"
		case 9
		mmes = "Setembro"
		case 10
		mmes = "Outubro"
		case 11
		mmes = "Novembro"
		case 12
		mmes = "Dezembro"
		end select
		response.write mmes		
		%>
		</font></b>
		</td>
	</tr>
	<tr>
		<td width="81" height="20" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b>
		<font size="1" face="Arial">Data</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Código</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vai</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vem</font></b></td>
	</tr>
<%
for u = 1 to 5
Set c1 = conexao.execute("SELECT * FROM datas WHERE iddata = 'C7"&u&"' and mes = "&mes&" order by iddata DESC")
if not c1.eof then
%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C7<%=u%>0cod" name="C7<%=u%>0cod" size="1" tabindex="2" value="<%=c1("codcat")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C7<%=u%>0vai" id="C7<%=u%>0vai" size="7" tabindex="3"  value="<%=c1("datavai")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C7<%=u%>0vem" id="C7<%=u%>0vem" size="7" tabindex="3"  value="<%=c1("datavem")%>"></td>
	</tr>
<%
else%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C7<%=u%>0cod" name="C7<%=u%>0cod" size="1" tabindex="2" value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C7<%=u%>0vai" id="C7<%=u%>0vai" size="7" tabindex="3"  value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C7<%=u%>0vem" id="C7<%=u%>0vem" size="7" tabindex="3"  value=""></td>
	</tr>
<%
end if
c1.close
Set c1 = Nothing
next
%>
</table></td>
		<td width="33%" bgcolor="#F5F5F5" valign="top" align="center">
<table border="0" width="280" cellspacing="0" cellpadding="5" id="table1" style="border-width: 0px">
	<tr>
		<td width="81" height="35" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="2" face="Arial">Mês</font></b></td>
		<td colspan="3" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font face="Arial" color="#CC0000"><%
		mes = month(Now)+1
		if mes > 12 then
		mes = mes-12
		end if
		select case mes
		case 1
		mmes = "Janeiro"
		case 2
		mmes = "Fevereiro"
		case 3
		mmes = "Março"
		case 4
		mmes = "Abril"
		case 5
		mmes = "Maio"
		case 6
		mmes = "Junho"
		case 7
		mmes = "Julho"
		case 8
		mmes = "Agosto"
		case 9
		mmes = "Setembro"
		case 10
		mmes = "Outubro"
		case 11
		mmes = "Novembro"
		case 12
		mmes = "Dezembro"
		end select
		response.write mmes		
		%>
		</font></b>
		</td>
	</tr>
	<tr>
		<td width="81" height="20" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b>
		<font size="1" face="Arial">Data</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Código</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vai</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vem</font></b></td>
	</tr>
<%
for u = 1 to 5
Set c1 = conexao.execute("SELECT * FROM datas WHERE iddata = 'C7"&u&"' and mes = "&mes&" order by iddata DESC")
if not c1.eof then
%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C7<%=u%>1cod" name="C7<%=u%>1cod" size="1" tabindex="2" value="<%=c1("codcat")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C7<%=u%>1vai" id="C7<%=u%>1vai" size="7" tabindex="3"  value="<%=c1("datavai")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C7<%=u%>1vem" id="C7<%=u%>1vem" size="7" tabindex="3"  value="<%=c1("datavem")%>"></td>
	</tr>
<%
else%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C7<%=u%>1cod" name="C7<%=u%>1cod" size="1" tabindex="2" value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C7<%=u%>1vai" id="C7<%=u%>1vai" size="7" tabindex="3"  value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C7<%=u%>1vem" id="C7<%=u%>1vem" size="7" tabindex="3"  value=""></td>
	</tr>
<%
end if
c1.close
Set c1 = Nothing
next
%>
</table>
		</td>
		<td width="33%" bgcolor="#F5F5F5" valign="top" align="center">
<table border="0" width="280" cellspacing="0" cellpadding="5" id="table1" style="border-width: 0px">
	<tr>
		<td width="81" height="35" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="2" face="Arial">Mês</font></b></td>
		<td colspan="3" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font face="Arial" color="#CC0000"><%
		mes = month(Now)+2
		if mes > 12 then
		mes = mes-12
		end if
		select case mes
		case 1
		mmes = "Janeiro"
		case 2
		mmes = "Fevereiro"
		case 3
		mmes = "Março"
		case 4
		mmes = "Abril"
		case 5
		mmes = "Maio"
		case 6
		mmes = "Junho"
		case 7
		mmes = "Julho"
		case 8
		mmes = "Agosto"
		case 9
		mmes = "Setembro"
		case 10
		mmes = "Outubro"
		case 11
		mmes = "Novembro"
		case 12
		mmes = "Dezembro"
		end select
		response.write mmes		
		%>
		</font></b>
		</td>
	</tr>
	<tr>
		<td width="81" height="20" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b>
		<font size="1" face="Arial">Data</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Código</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vai</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vem</font></b></td>
	</tr>
<%
for u = 1 to 5
Set c1 = conexao.execute("SELECT * FROM datas WHERE iddata = 'C7"&u&"' and mes = "&mes&" order by iddata DESC")
if not c1.eof then
%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C7<%=u%>2cod" name="C7<%=u%>2cod" size="1" tabindex="2" value="<%=c1("codcat")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C7<%=u%>2vai" id="C7<%=u%>2vai" size="7" tabindex="3"  value="<%=c1("datavai")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C7<%=u%>2vem" id="C7<%=u%>2vem" size="7" tabindex="3"  value="<%=c1("datavem")%>"></td>
	</tr>
<%
else%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C7<%=u%>2cod" name="C7<%=u%>2cod" size="1" tabindex="2" value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C7<%=u%>2vai" id="C7<%=u%>2vai" size="7" tabindex="3"  value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C7<%=u%>2vem" id="C7<%=u%>2vem" size="7" tabindex="3"  value=""></td>
	</tr>
<%
end if
c1.close
Set c1 = Nothing
next
%>
</table>		</td>
	</tr>
	<tr>
		<td colspan="3" height="10" bgcolor="#F5F5F5">
		<span style="font-size: 1pt">&nbsp;</span></td>
	</tr>
		<tr>
		<td colspan="3" height="60" bgcolor="#F5F5F5" align="right">
		<input type="image" src="../site/imagem/salvar.png">&nbsp;&nbsp;&nbsp; </td>
	</tr>
	</form>
</table>
<table border="0" width="100%" cellspacing="0" cellpadding="2" id="table1">
	<tr>
		<td colspan="3" height="10">
		</td>
	</tr>
	<tr>
		<td colspan="3" height="10" bgcolor="#F5F5F5">
		<span style="font-size: 1pt">&nbsp;</span></td>
	</tr>
	<form method="POST" action="up.datas.asp">
	<input type="hidden" name="revista" id="revista" value="C9">
	<tr>
		<td colspan="3" height="50" bgcolor="#F5F5F5">
		<img border="0" src="../site/imagem/C9.jpg" width="934" height="60"></td>
	</tr>
	<tr>
		<td colspan="3" height="10" bgcolor="#F5F5F5">
		<span style="font-size: 1pt">&nbsp;</span></td>
	</tr>
	<tr>
		<td width="33%" bgcolor="#F5F5F5" valign="top" align="center">
<table border="0" width="280" cellspacing="0" cellpadding="5" id="table1" style="border-width: 0px">
	<tr>
		<td width="81" height="35" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="2" face="Arial">Mês</font></b></td>
		<td colspan="3" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font face="Arial" color="#CC0000"><%
		mes = month(Now)
		select case mes
		case 1
		mmes = "Janeiro"
		case 2
		mmes = "Fevereiro"
		case 3
		mmes = "Março"
		case 4
		mmes = "Abril"
		case 5
		mmes = "Maio"
		case 6
		mmes = "Junho"
		case 7
		mmes = "Julho"
		case 8
		mmes = "Agosto"
		case 9
		mmes = "Setembro"
		case 10
		mmes = "Outubro"
		case 11
		mmes = "Novembro"
		case 12
		mmes = "Dezembro"
		end select
		response.write mmes		
		%>
		</font></b>
		</td>
	</tr>
	<tr>
		<td width="81" height="20" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b>
		<font size="1" face="Arial">Data</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Código</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vai</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vem</font></b></td>
	</tr>
<%
for u = 1 to 5
Set c1 = conexao.execute("SELECT * FROM datas WHERE iddata = 'C9"&u&"' and mes = "&mes&" order by iddata DESC")
if not c1.eof then
%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C9<%=u%>0cod" name="C9<%=u%>0cod" size="1" tabindex="2" value="<%=c1("codcat")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C9<%=u%>0vai" id="C9<%=u%>0vai" size="7" tabindex="3"  value="<%=c1("datavai")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C9<%=u%>0vem" id="C9<%=u%>0vem" size="7" tabindex="3"  value="<%=c1("datavem")%>"></td>
	</tr>
<%
else%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C9<%=u%>0cod" name="C9<%=u%>0cod" size="1" tabindex="2" value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C9<%=u%>0vai" id="C9<%=u%>0vai" size="7" tabindex="3"  value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C9<%=u%>0vem" id="C9<%=u%>0vem" size="7" tabindex="3"  value=""></td>
	</tr>
<%
end if
c1.close
Set c1 = Nothing
next
%>
</table></td>
		<td width="33%" bgcolor="#F5F5F5" valign="top" align="center">
<table border="0" width="280" cellspacing="0" cellpadding="5" id="table1" style="border-width: 0px">
	<tr>
		<td width="81" height="35" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="2" face="Arial">Mês</font></b></td>
		<td colspan="3" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font face="Arial" color="#CC0000"><%
		mes = month(Now)+1
		if mes > 12 then
		mes = mes-12
		end if
		select case mes
		case 1
		mmes = "Janeiro"
		case 2
		mmes = "Fevereiro"
		case 3
		mmes = "Março"
		case 4
		mmes = "Abril"
		case 5
		mmes = "Maio"
		case 6
		mmes = "Junho"
		case 7
		mmes = "Julho"
		case 8
		mmes = "Agosto"
		case 9
		mmes = "Setembro"
		case 10
		mmes = "Outubro"
		case 11
		mmes = "Novembro"
		case 12
		mmes = "Dezembro"
		end select
		response.write mmes		
		%>
		</font></b>
		</td>
	</tr>
	<tr>
		<td width="81" height="20" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b>
		<font size="1" face="Arial">Data</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Código</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vai</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vem</font></b></td>
	</tr>
<%
for u = 1 to 5
Set c1 = conexao.execute("SELECT * FROM datas WHERE iddata = 'C9"&u&"' and mes = "&mes&" order by iddata DESC")
if not c1.eof then
%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C9<%=u%>1cod" name="C9<%=u%>1cod" size="1" tabindex="2" value="<%=c1("codcat")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C9<%=u%>1vai" id="C9<%=u%>1vai" size="7" tabindex="3"  value="<%=c1("datavai")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C9<%=u%>1vem" id="C9<%=u%>1vem" size="7" tabindex="3"  value="<%=c1("datavem")%>"></td>
	</tr>
<%
else%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C9<%=u%>1cod" name="C9<%=u%>1cod" size="1" tabindex="2" value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C9<%=u%>1vai" id="C9<%=u%>1vai" size="7" tabindex="3"  value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C9<%=u%>1vem" id="C9<%=u%>1vem" size="7" tabindex="3"  value=""></td>
	</tr>
<%
end if
c1.close
Set c1 = Nothing
next
%>
</table>
		</td>
		<td width="33%" bgcolor="#F5F5F5" valign="top" align="center">
<table border="0" width="280" cellspacing="0" cellpadding="5" id="table1" style="border-width: 0px">
	<tr>
		<td width="81" height="35" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="2" face="Arial">Mês</font></b></td>
		<td colspan="3" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font face="Arial" color="#CC0000"><%
		mes = month(Now)+2
		if mes > 12 then
		mes = mes-12
		end if
		select case mes
		case 1
		mmes = "Janeiro"
		case 2
		mmes = "Fevereiro"
		case 3
		mmes = "Março"
		case 4
		mmes = "Abril"
		case 5
		mmes = "Maio"
		case 6
		mmes = "Junho"
		case 7
		mmes = "Julho"
		case 8
		mmes = "Agosto"
		case 9
		mmes = "Setembro"
		case 10
		mmes = "Outubro"
		case 11
		mmes = "Novembro"
		case 12
		mmes = "Dezembro"
		end select
		response.write mmes		
		%>
		</font></b>
		</td>
	</tr>
	<tr>
		<td width="81" height="20" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b>
		<font size="1" face="Arial">Data</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Código</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vai</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vem</font></b></td>
	</tr>
<%
for u = 1 to 5
Set c1 = conexao.execute("SELECT * FROM datas WHERE iddata = 'C9"&u&"' and mes = "&mes&" order by iddata DESC")
if not c1.eof then
%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C9<%=u%>2cod" name="C9<%=u%>2cod" size="1" tabindex="2" value="<%=c1("codcat")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C9<%=u%>2vai" id="C9<%=u%>2vai" size="7" tabindex="3"  value="<%=c1("datavai")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C9<%=u%>2vem" id="C9<%=u%>2vem" size="7" tabindex="3"  value="<%=c1("datavem")%>"></td>
	</tr>
<%
else%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C9<%=u%>2cod" name="C9<%=u%>2cod" size="1" tabindex="2" value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C9<%=u%>2vai" id="C9<%=u%>2vai" size="7" tabindex="3"  value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C9<%=u%>2vem" id="C9<%=u%>2vem" size="7" tabindex="3"  value=""></td>
	</tr>
<%
end if
c1.close
Set c1 = Nothing
next
%>
</table>
		</td>

	</tr>
	<tr>
		<td colspan="3" height="10" bgcolor="#F5F5F5">
		<span style="font-size: 1pt">&nbsp;</span></td>
	</tr>
		<tr>
		<td colspan="3" height="60" bgcolor="#F5F5F5" align="right">
		<input type="image" src="../site/imagem/salvar.png">&nbsp;&nbsp;&nbsp; </td>
	</tr>
	</form>
</table>
<table border="0" width="100%" cellspacing="0" cellpadding="2" id="table1">
	<tr>
		<td colspan="3" height="10">
		</td>
	</tr>
	<tr>
		<td colspan="3" height="10" bgcolor="#F5F5F5">
		<span style="font-size: 1pt">&nbsp;</span></td>
	</tr>
	<form method="POST" action="up.datas.asp">
	<input type="hidden" name="revista" id="revista" value="C10">
	<tr>
		<td colspan="3" height="50" bgcolor="#F5F5F5">
		<img border="0" src="../site/imagem/C10.jpg" width="934" height="60"></td>
	</tr>
	<tr>
		<td colspan="3" height="10" bgcolor="#F5F5F5">
		<span style="font-size: 1pt">&nbsp;</span></td>
	</tr>
	<tr>
		<td width="33%" bgcolor="#F5F5F5" valign="top" align="center">
<table border="0" width="280" cellspacing="0" cellpadding="5" id="table1" style="border-width: 0px">
	<tr>
		<td width="81" height="35" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="2" face="Arial">Mês</font></b></td>
		<td colspan="3" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font face="Arial" color="#CC0000"><%
		mes = month(Now)
		select case mes
		case 1
		mmes = "Janeiro"
		case 2
		mmes = "Fevereiro"
		case 3
		mmes = "Março"
		case 4
		mmes = "Abril"
		case 5
		mmes = "Maio"
		case 6
		mmes = "Junho"
		case 7
		mmes = "Julho"
		case 8
		mmes = "Agosto"
		case 9
		mmes = "Setembro"
		case 10
		mmes = "Outubro"
		case 11
		mmes = "Novembro"
		case 12
		mmes = "Dezembro"
		end select
		response.write mmes		
		%>
		</font></b>
		</td>
	</tr>
	<tr>
		<td width="81" height="20" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b>
		<font size="1" face="Arial">Data</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Código</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vai</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vem</font></b></td>
	</tr>
<%
for u = 1 to 5
Set c1 = conexao.execute("SELECT * FROM datas WHERE iddata = 'C10"&u&"' and mes = "&mes&" order by iddata DESC")
if not c1.eof then
%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C10<%=u%>0cod" name="C10<%=u%>0cod" size="1" tabindex="2" value="<%=c1("codcat")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C10<%=u%>0vai" id="C10<%=u%>0vai" size="7" tabindex="3"  value="<%=c1("datavai")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C10<%=u%>0vem" id="C10<%=u%>0vem" size="7" tabindex="3"  value="<%=c1("datavem")%>"></td>
	</tr>
<%
else%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C10<%=u%>0cod" name="C10<%=u%>0cod" size="1" tabindex="2" value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C10<%=u%>0vai" id="C10<%=u%>0vai" size="7" tabindex="3"  value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C10<%=u%>0vem" id="C10<%=u%>0vem" size="7" tabindex="3"  value=""></td>
	</tr>
<%
end if
c1.close
Set c1 = Nothing
next
%>
</table></td>
		<td width="33%" bgcolor="#F5F5F5" valign="top" align="center">
<table border="0" width="280" cellspacing="0" cellpadding="5" id="table1" style="border-width: 0px">
	<tr>
		<td width="81" height="35" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="2" face="Arial">Mês</font></b></td>
		<td colspan="3" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font face="Arial" color="#CC0000"><%
		mes = month(Now)+1
		if mes > 12 then
		mes = mes-12
		end if
		select case mes
		case 1
		mmes = "Janeiro"
		case 2
		mmes = "Fevereiro"
		case 3
		mmes = "Março"
		case 4
		mmes = "Abril"
		case 5
		mmes = "Maio"
		case 6
		mmes = "Junho"
		case 7
		mmes = "Julho"
		case 8
		mmes = "Agosto"
		case 9
		mmes = "Setembro"
		case 10
		mmes = "Outubro"
		case 11
		mmes = "Novembro"
		case 12
		mmes = "Dezembro"
		end select
		response.write mmes		
		%>
		</font></b>
		</td>
	</tr>
	<tr>
		<td width="81" height="20" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b>
		<font size="1" face="Arial">Data</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Código</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vai</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vem</font></b></td>
	</tr>
<%
for u = 1 to 5
Set c1 = conexao.execute("SELECT * FROM datas WHERE iddata = 'C10"&u&"' and mes = "&mes&" order by iddata DESC")
if not c1.eof then
%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C10<%=u%>1cod" name="C10<%=u%>1cod" size="1" tabindex="2" value="<%=c1("codcat")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C10<%=u%>1vai" id="C10<%=u%>1vai" size="7" tabindex="3"  value="<%=c1("datavai")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C10<%=u%>1vem" id="C10<%=u%>1vem" size="7" tabindex="3"  value="<%=c1("datavem")%>"></td>
	</tr>
<%
else%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C10<%=u%>1cod" name="C10<%=u%>1cod" size="1" tabindex="2" value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C10<%=u%>1vai" id="C10<%=u%>1vai" size="7" tabindex="3"  value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C10<%=u%>1vem" id="C10<%=u%>1vem" size="7" tabindex="3"  value=""></td>
	</tr>
<%
end if
c1.close
Set c1 = Nothing
next
%>
</table>
		</td>
		<td width="33%" bgcolor="#F5F5F5" valign="top" align="center">
<table border="0" width="280" cellspacing="0" cellpadding="5" id="table1" style="border-width: 0px">
	<tr>
		<td width="81" height="35" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="2" face="Arial">Mês</font></b></td>
		<td colspan="3" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font face="Arial" color="#CC0000"><%
		mes = month(Now)+2
		if mes > 12 then
		mes = mes-12
		end if
		select case mes
		case 1
		mmes = "Janeiro"
		case 2
		mmes = "Fevereiro"
		case 3
		mmes = "Março"
		case 4
		mmes = "Abril"
		case 5
		mmes = "Maio"
		case 6
		mmes = "Junho"
		case 7
		mmes = "Julho"
		case 8
		mmes = "Agosto"
		case 9
		mmes = "Setembro"
		case 10
		mmes = "Outubro"
		case 11
		mmes = "Novembro"
		case 12
		mmes = "Dezembro"
		end select
		response.write mmes		
		%>
		</font></b>
		</td>
	</tr>
	<tr>
		<td width="81" height="20" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b>
		<font size="1" face="Arial">Data</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Código</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vai</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20">
		<b>
		<font size="1" face="Arial">Data vem</font></b></td>
	</tr>
<%
for u = 1 to 5
Set c1 = conexao.execute("SELECT * FROM datas WHERE iddata = 'C10"&u&"' and mes = "&mes&" order by iddata DESC")
if not c1.eof then
%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C10<%=u%>2cod" name="C10<%=u%>2cod" size="1" tabindex="2" value="<%=c1("codcat")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C10<%=u%>2vai" id="C10<%=u%>2vai" size="7" tabindex="3"  value="<%=c1("datavai")%>"></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C10<%=u%>2vem" id="C10<%=u%>2vem" size="7" tabindex="3"  value="<%=c1("datavem")%>"></td>
	</tr>
<%
else%>
	<tr>
		<td width="81" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="1" face="Arial">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" id="C10<%=u%>2cod" name="C10<%=u%>2cod" size="1" tabindex="2" value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C10<%=u%>2vai" id="C10<%=u%>2vai" size="7" tabindex="3"  value=""></td>
		<td style="border-style:none; border-width:medium; " align="right" bgcolor="#F0F0F0" height="30">
		<input type="text" name="C10<%=u%>2vem" id="C10<%=u%>2vem" size="7" tabindex="3"  value=""></td>
	</tr>
<%
end if
c1.close
Set c1 = Nothing
next
%>
</table>
		</td>

	</tr>
	
		<tr>
		<td colspan="3" height="10" bgcolor="#F5F5F5">
		<span style="font-size: 1pt">&nbsp;</span></td>
	</tr>
		<tr>
		<td colspan="3" height="60" bgcolor="#F5F5F5" align="right">
		<input type="image" src="../site/imagem/salvar.png">&nbsp;&nbsp;&nbsp; </td>
	</tr>
	</form>
</table>

				  </td>
                </tr>
            </table>                
            </td>
          </tr>
        </table>
    </td>
  </tr>
</table>
</body>
</html>