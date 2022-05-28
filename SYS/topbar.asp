
<head>
<style type="text/css">
<!--
body {
	margin:0px; background-image: url('../imagens/fundo.jpg');
	
}
-->
</style>
<script src="SpryAssets/SpryMenuBar.js" type="text/javascript"></script>
<script language="javascript">
navHover = function() {
	var lis = document.getElementById("navmenu").getElementsByTagName("LI");
	for (var i=0; i<lis.length; i++) {
		lis[i].onmouseover=function() {
			this.className+=" iehover";
		}
		lis[i].onmouseout=function() {
			this.className=this.className.replace(new RegExp(" iehover\\b"), "");
		}
	}
}
if (window.attachEvent) window.attachEvent("onload", navHover);
</script>
<link href="menu.css" rel="stylesheet" type="text/css" />
<link href="SpryAssets/SpryMenuBarHorizontal.css" rel="stylesheet" type="text/css" />
</head>

<div align="center">

<table width="934" height="110" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="164" height="100">
	<img border="0" src="../site/imagem/logocentral.jpg" width="346" height="50"></td>
    <td align="right"><br>
	<img src="imagens/administracao.gif" width="145" height="23" align="middle" /><p>
	<span style="text-transform: uppercase; font-weight: 700">
	<a href="logout.asp" style="text-decoration: none;"><font face="Arial" color="#CC0000" size="1">
	Sair</font></a><font face="Arial" color="#CC0000" size="1">&nbsp;&nbsp; </font></span></td>
  </tr>
  <tr>
    <td width="758" height="20" colspan="2" bgcolor="#CC0000">
<ul id="MenuBar" class="MenuBarHorizontal">
      <li><b><font face="Arial" size="2"><a href="default.asp">Home</a></font></b></li>
      <li><b><font face="Arial" size="2"><a class="MenuBarItemSubmenu" href="#">Data de Pedidos</a>
        </font></b>
        <ul>
        <li><b><font face="Arial" size="2"><a href="update_datas.asp">Editar Datas</a></font></b></li>
        </ul>
      </li>
      </ul></td>
  </tr>
</table>

</div>

<script type="text/javascript">
<!--
var MenuBar1 = new Spry.Widget.MenuBar("MenuBar", {imgDown:"../../SpryAssets/SpryMenuBarDownHover.gif", imgRight:"../../SpryAssets/SpryMenuBarRightHover.gif"});
//-->
</script>