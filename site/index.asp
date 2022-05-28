<!-- #include file="conection.asp" --><%
'Applications
Application("pag") = request.querystring("h")
if Application("pag") = "" then
Application("pag") = request.form("h")
if Application("pag") = "" then
Application("pag") = "inicial"
end if
end if
Application("erro")       = request.querystring("error")
Application("erro1")      = request.querystring("error1")
Application("erro2")      = request.querystring("error2")
Application("erro3")      = request.querystring("error3")
Application("erro4")      = request.querystring("error4")
Application("erro5")      = request.querystring("error5")
Application("erro6")      = request.querystring("error6")
Application("erro7")      = request.querystring("error7")
Application("erro8")      = request.querystring("error8")
Application("erro9")      = request.querystring("error9")
Application("erro10")     = request.querystring("error10")
Application("erro11")     = request.querystring("error11")
Application("erro13")     = request.querystring("error13")
Application("catalogo")   = request.querystring("revista")
Application("revista")    = request.form("catalogo")
Application("nome")       = request.querystring("nome")
Application("cpf")        = request.querystring("cpf")
Application("mail")       = request.querystring("email")
Application("ddd")        = request.querystring("ddd")
Application("fone")       = request.querystring("telefone")
Application("telefone")   = request.querystring("telefone")
Application("assu")       = request.querystring("assunto")
Application("mens")       = request.querystring("mensagem")
Application("status")     = request.querystring("status")
Application("endereco")   = request.querystring("endereco")
Application("cidade")     = request.querystring("cidade")
Application("estado")     = request.querystring("estado")
Application("qualcat")    = request.querystring("qualcat")

'Actions
Application("action") = request.form("action")
if Application("action") <> "" then
Application("nome")       = request.form("nome")
Application("mail")       = request.form("email")
Application("ddd")        = request.form("ddd")
Application("fone")       = request.form("telefone")
Application("assu")       = request.form("assunto")
Application("mens")       = request.form("mensagem")
Application("catalogo")   = request.form("revista")
Application("qualcat")    = request.form("qualcat")
Application("endereco")   = request.form("endereco")
Application("cidade")     = request.form("cidade")
Application("estado")     = request.form("estado")
end if
Select case Application("action")
case "gomail"
Set oMail = Server.CreateObject("Persits.MailSender")
oMail.Host = "smtp.centraldoscatalogos.com.br"
oMail.Port = 587
oMail.UserName = "atendimento@centraldoscatalogos.com.br"
oMail.PassWord = "central123"
oMail.From = "atendimento@centraldoscatalogos.com.br"
oMail.FromName = Application("nome")
oMail.AddAddress "atendimento@centraldoscatalogos.com.br", "Atendimento Central dos Catálogos"
oMail.AddReplyTo Application("mail"), Application("nome")
oMail.Subject = "Pedido On-Line ("&Application("catalogo")&")"
oMail.Body = Session("MSN")
oMail.IsHTML = True
oMail.Send
If Err <> 0 Then
	response.redirect("?h=central.pp.ok&status=erro")
Else
    response.redirect("?h=central.pp.ok&status=ok")
End If
case "crevendedora"
Application("erroconfirm") = 0
if Application("nome") = "" or len(Application("nome")) < 4 then
if Application("nome") = "" then
session("erro_nomec") = "Digite seu nome!<br>"
Application("erroconfirm") = Application("erroconfirm")+1
else
session("erro_nomec") = "Nome inválido!<br>"
Application("erroconfirm") = Application("erroconfirm")+1
end if
else
session("erro_nomec") = ""
end if
if Application("ddd") = "" or len(Application("ddd")) <> 2 or isnumeric(Application("ddd")) = false then
if Application("ddd") = "" then
session("erro_dddc") = "Digite o DDD!<br>"
Application("erroconfirm") = Application("erroconfirm")+1
else
session("erro_dddc") = "DDD inválido!<br>"
Application("erroconfirm") = Application("erroconfirm")+1
end if
else
session("erro_dddc") = ""
end if
if Application("fone") = "" or len(Application("fone")) <> 8 or isnumeric(Application("fone")) = false then
if Application("fone") = "" then
session("erro_telc") = "Digite o número do seu telefone!<br>"
Application("erroconfirm") = Application("erroconfirm")+1
else
session("erro_telc") = "Telefone inválido!<br>"
Application("erroconfirm") = Application("erroconfirm")+1
end if
else
session("erro_telc") = ""
end if
if Application("mail") = "" or instr(Application("mail"),"@") < 1 or instr(Application("mail"),".com") < 1 and instr(Application("mail"),".net") < 1 and instr(Application("mail"),".org") < 1 and instr(Application("mail"),".info") < 1 and instr(Application("mail"),".gov") < 1 then
if Application("mail") = "" then
session("erro_emailc") = "Digite o e-mail!<br>"
Application("erroconfirm") = Application("erroconfirm")+1
else
session("erro_emailc") = "E-mail inválido!<br>"
Application("erroconfirm") = Application("erroconfirm")+1
end if
else
session("erro_emailc") = ""
end if
if Application("endereco") = "" then
session("erro_enderecoc") = "Digite a endereço!<br>"
Application("erroconfirm") = Application("erroconfirm")+1
end if
if Application("cidade") = "" then
session("erro_cidadec") = "Digite a cidade!<br>"
Application("erroconfirm") = Application("erroconfirm")+1
end if
if Application("estado") = "" then
session("erro_estadoc") = "Digite a estado!<br>"
Application("erroconfirm") = Application("erroconfirm")+1
end if
if Application("erroconfirm") > 0 then
response.redirect("?h=central.cr1&nome="&Application("nome")&"&ddd="&Application("ddd")&"&telefone="&Application("fone")&"&email="&Application("mail")&"&endereco="&Application("endereco")&"&cidade="&Application("cidade")&"&estado="&Application("estado")&"")
end if
Set oMail = Server.CreateObject("Persits.MailSender")
oMail.Host = "smtp.centraldoscatalogos.com.br"
oMail.Port = 587
oMail.UserName = "atendimento@centraldoscatalogos.com.br"
oMail.PassWord = "central123"
oMail.From = "atendimento@centraldoscatalogos.com.br"
oMail.FromName = Application("nome")
oMail.AddAddress "atendimento@centraldoscatalogos.com.br", "Atendimento Central dos Catálogos"
oMail.AddReplyTo Application("mail"), Application("nome")
oMail.Subject = "Solicitação de cadastro"
oMail.Body = Application("mens")
oMail.IsHTML = False
oMail.Send
If Err <> 0 Then
	response.redirect("?h=central.fc&status=erro")
Else
    response.redirect("?h=central.fc&status=ok")
End If
case "contato"
Application("erroconfirm") = 0
if Application("nome") = "" or len(Application("nome")) < 4 then
if Application("nome") = "" then
session("erro_nomec") = "Digite seu nome!<br>"
Application("erroconfirm") = Application("erroconfirm")+1
else
session("erro_nomec") = "Nome inválido!<br>"
Application("erroconfirm") = Application("erroconfirm")+1
end if
else
session("erro_nomec") = ""
end if
if Application("ddd") = "" or len(Application("ddd")) <> 2 or isnumeric(Application("ddd")) = false then
if Application("ddd") = "" then
session("erro_dddc") = "Digite o DDD!<br>"
Application("erroconfirm") = Application("erroconfirm")+1
else
session("erro_dddc") = "DDD inválido!<br>"
Application("erroconfirm") = Application("erroconfirm")+1
end if
else
session("erro_dddc") = ""
end if
if Application("fone") = "" or len(Application("fone")) <> 8 or isnumeric(Application("fone")) = false then
if Application("fone") = "" then
session("erro_telc") = "Digite o número do seu telefone!<br>"
Application("erroconfirm") = Application("erroconfirm")+1
else
session("erro_telc") = "Telefone inválido!<br>"
Application("erroconfirm") = Application("erroconfirm")+1
end if
else
session("erro_telc") = ""
end if
if Application("mail") = "" or instr(Application("mail"),"@") < 1 or instr(Application("mail"),".com") < 1 and instr(Application("mail"),".net") < 1 and instr(Application("mail"),".org") < 1 and instr(Application("mail"),".info") < 1 and instr(Application("mail"),".gov") < 1 then
if Application("mail") = "" then
session("erro_emailc") = "Digite o e-mail!<br>"
Application("erroconfirm") = Application("erroconfirm")+1
else
session("erro_emailc") = "E-mail inválido!<br>"
Application("erroconfirm") = Application("erroconfirm")+1
end if
else
session("erro_emailc") = ""
end if
if Application("assu") = "" then
session("erro_assuntoc") = "Digite o assunto!<br>"
Application("erroconfirm") = Application("erroconfirm")+1
end if
if Application("mens") = "" or len(Application("mens")) < 2 then
if Application("mens") = "" then
session("erro_mensagemc") = "Digite a mensagem!<br>"
Application("erroconfirm") = Application("erroconfirm")+1
else
session("erro_mensagemc") = "A mensagem deve conter no mínimo 2 caractéres!<br>"
Application("erroconfirm") = Application("erroconfirm")+1
end if
else
session("erro_mensagemc") = ""
end if
if Application("erroconfirm") > 0 then
response.redirect("?h=central.fc&nome="&Application("nome")&"&ddd="&Application("ddd")&"&telefone="&Application("fone")&"&email="&Application("mail")&"&assunto="&Application("assu")&"&mensagem="&Application("mens")&"")
end if
Set oMail = Server.CreateObject("Persits.MailSender")
oMail.Host = "smtp.centraldoscatalogos.com.br"
oMail.Port = 587
oMail.UserName = "atendimento@centraldoscatalogos.com.br"
oMail.PassWord = "central123"
oMail.From = "atendimento@centraldoscatalogos.com.br"
oMail.FromName = Application("nome")
oMail.AddAddress "atendimento@centraldoscatalogos.com.br", "Atendimento Central dos Catálogos"
oMail.AddReplyTo Application("mail"), Application("nome")
oMail.Subject = Application("assu")
oMail.Body = Application("mens")
oMail.IsHTML = False
oMail.Send
If Err <> 0 Then
	response.redirect("?h=central.fc&status=erro")
Else
    response.redirect("?h=central.fc&status=ok")
End If

End select
%><html>
<head>
<title>Central dos Catálogos</title>
<meta name="title" content="Central dos Catálogos" />
<meta name="url" content="http://www.centraldoscatalogos.com.br/" />
<meta name="description" content="Seja um(a) revendedor(a)!" />
<meta name="revisit-after" content="1" />
<meta name="keywords" content="central dos catálogos, central, catalogo, revista, lingerie, revenda, abelha rainha, luzon, miro star, sexy star, hiroshima, essencial, ivete, l'arc en ciel, quatro estacoes, quatro estações, natubelly, quintess, marcyn, cavie" />
<meta name="autor" content="ARTlizando LTDA" />
<meta name="company" content="ARTlizando LTDA" />
<link rev="made" href="mailto:atendimento@artlizando.com.br" />
<link rel="shortcut icon" href="imagem/centralicon.ico" />
<link rel="icon" href="imagem/centralicon.ico" />
<link href="estilo/site.css" type="text/css" rel="stylesheet">
<script language="JavaScript" type="text/JavaScript" src="http://www.artlizando.com.br/v1/_script/001.js"></script>
<meta content="NO-CACHE" http-equiv="pragma"></meta>
<meta content="no-cache" http-equiv="Cache-control"></meta>
<meta content="must-revalidate" http-equiv="Cache-control"></meta>
<meta content="max-age=0" http-equiv="Cache-control"></meta>
<meta content="Mon, 04 Dec 1999 21:29:02 GMT" http-equiv="Expires"></meta> 
<meta content="document" name="resource-type"></meta>
<meta content="15" name="revisit-after"></meta>
<meta content="ALL" name="robots"></meta>
<meta content="Global" name="distribution"></meta>
<meta content="General" name="rating"></meta>
<meta content="IE=8" http-equiv="X-UA-Compatible"></meta>
<meta http-equiv="X-UA-Compatible" content="IE=8" />
<script src="script/jquery-1.5.2.min.js" language="javascript"></script>
<script type="text/Javascript">
		function abrirRevista(){
			$('#catalogoon').fadeIn('slow');
			$('body').css('overflow', 'hidden');
		}
		function fecharRevista(){
			$('#catalogoon').fadeOut('slow');
			$('body').css('overflow', '');
		}
</script>
<script language="JavaScript" type="text/javascript">  
      function mascaraGenerica(evt, campo, padrao) {  
           var charCode = (evt.which) ? evt.which : evt.keyCode;  
           if (charCode == 8) return true;
           if (charCode != 46 && (charCode < 48 || charCode > 57)) return false;
           campo.maxLength = padrao.length;  
           entrada = campo.value;  
           if (padrao.length > entrada.length && padrao.charAt(entrada.length) != '#') {  
                campo.value = entrada + padrao.charAt(entrada.length);                 
           }  
           return true;  
      }
</script> 
<%if Application("pag") = "central.pp1" then%>
<link href="estilo/jquery.autocomplete.css" type="text/css" rel="stylesheet" />
<script language="javascript">
function MascaraMoeda(objTextBox, SeparadorMilesimo, SeparadorDecimal, e){
    var sep = 0;
    var key = '';
    var i = j = 0;
    var len = len2 = 0;
    var strCheck = '0123456789';
    var aux = aux2 = '';
    var whichCode = (window.Event) ? e.which : e.keyCode;
    if (whichCode == 13) return true;
    key = String.fromCharCode(whichCode); // Valor para o código da Chave
    if (strCheck.indexOf(key) == -1) return false; // Chave inválida
    len = objTextBox.value.length;
    for(i = 0; i < len; i++)
        if ((objTextBox.value.charAt(i) != '0') && (objTextBox.value.charAt(i) != SeparadorDecimal)) break;
    aux = '';
    for(; i < len; i++)
        if (strCheck.indexOf(objTextBox.value.charAt(i))!=-1) aux += objTextBox.value.charAt(i);
    aux += key;
    len = aux.length;
    if (len == 0) objTextBox.value = '';
    if (len == 1) objTextBox.value = '0'+ SeparadorDecimal + '0' + aux;
    if (len == 2) objTextBox.value = '0'+ SeparadorDecimal + aux;
    if (len > 2) {
        aux2 = '';
        for (j = 0, i = len - 3; i >= 0; i--) {
            if (j == 3) {
                aux2 += SeparadorMilesimo;
                j = 0;
            }
            aux2 += aux.charAt(i);
            j++;
        }
        objTextBox.value = '';
        len2 = aux2.length;
        for (i = len2 - 1; i >= 0; i--)
        objTextBox.value += aux2.charAt(i);
        objTextBox.value += SeparadorDecimal + aux.substr(len - 2, len);
    }
    return false;
}
function SomenteNumero(e){
     var tecla=(window.event)?event.keyCode:e.which;
     if((tecla > 47 && tecla < 58)) return true;
     else{
     if (tecla != 8) return false;
     else return true;
     }
}
function SomenteLetras(e) {
                var expressao;
                expressao = /[a-zA-Z çÇãÃõÕáÁéÉíÍóÓúÚâÂêÊôÔ]/;
                if(expressao.test(String.fromCharCode(e.keyCode)))
                {
                        return true;
                }
                else
                {
                        return false;
                }
        }
</script>
<script src="script/jquery.autocomplete.js" language="javascript"></script>
<%for i = 1 to 51
u = "<script type=""text/javascript"" language=""javascript"">"&vbCrLf
u = u&"            function Calc"&i&"(){"&vbCrLf
u = u&"            ValorUm = document.getElementById('quant"&i&"').value;"&vbCrLf
u = u&"            ValorDois = document.getElementById('punit"&i&"').value.replace(',','').replace('.','');"&vbCrLf
u = u&"                document.getElementById('ptotal"&i&"').value = ValorUm*1 * ValorDois*1 / 100;"&vbCrLf
u = u&"                document.getElementById('ptotal"&i&"').value = document.getElementById('ptotal"&i&"').value.replace('.',',');"&vbCrLf
u = u&"				       if (document.getElementById('ptotal"&i&"').value == 0) {"&vbCrLf
u = u&"                         document.getElementById('punit"&i&"').value = '0,00';"&vbCrLf
u = u&"				       		document.getElementById('ptotal"&i&"').value = '0,00';"&vbCrLf
u = u&"			                Totalizar();"&vbCrLf
u = u&"				       }"&vbCrLf
u = u&"				       funccomplete = document.getElementById('ptotal"&i&"').value.split(',');"&vbCrLf
u = u&"					   if (funccomplete[1]==undefined) {"&vbCrLf
u = u&"				            document.getElementById('ptotal"&i&"').value = funccomplete[0]+',00';"&vbCrLf
u = u&"			                Totalizar();"&vbCrLf
u = u&"					   }"&vbCrLf
u = u&"				       if (funccomplete[1].length < 2) {"&vbCrLf
u = u&"				           	document.getElementById('ptotal"&i&"').value = funccomplete[0]+','+funccomplete[1]+'0';"&vbCrLf
u = u&"			                Totalizar();"&vbCrLf
u = u&"				       }"&vbCrLf
u = u&"				       if (funccomplete[0].length > 3) {"&vbCrLf
u = u&"				       		func = funccomplete[0]*1 / 1000;"&vbCrLf
u = u&"				       		funccomplete1 = func.value.split('.');"&vbCrLf
u = u&"							       	if (funccomplete1[1].length == 1) {"&vbCrLf
u = u&"					                      document.getElementById('ptotal"&i&"').value = funccomplete1[0]+'.'+funccomplete1[1]+'00,'+funccomplete[1];"&vbCrLf
u = u&"			                              Totalizar();"&vbCrLf
u = u&"									}"&vbCrLf
u = u&"							       	if (funccomplete1[1].length == 2) {"&vbCrLf
u = u&"					                      document.getElementById('ptotal"&i&"').value = funccomplete1[0]+'.'+funccomplete1[1]+'0,'+funccomplete[1];"&vbCrLf
u = u&"			                              Totalizar();"&vbCrLf
u = u&"									}"&vbCrLf
u = u&"							       	if (funccomplete1[1].length == 3) {"&vbCrLf
u = u&"					                      document.getElementById('ptotal"&i&"').value = funccomplete1[0]+'.'+funccomplete1[1]+','+funccomplete[1];"&vbCrLf
u = u&"			                              Totalizar();"&vbCrLf
u = u&"									}"&vbCrLf 
u = u&"			           Totalizar();"&vbCrLf
u = u&"					   }"&vbCrLf
u = u&"			   Totalizar();"&vbCrLf
u = u&"            }"&vbCrLf
u = u&"</script>"&vbCrLf
response.write u
next
u = "<script type=""text/javascript"" language=""javascript"">"&vbCrLf
u = u&"            function Totalizar(){"&vbCrLf
for i = 1 to 50
u = u&"            ValorUm"&i&" = document.getElementById('quant"&i&"').value;"&vbCrLf
u = u&"            ValorDois"&i&" = document.getElementById('punit"&i&"').value.replace(',','').replace('.','');"&vbCrLf
next
u = u&"            document.getElementById('total').innerHTML = ValorUm1*1 * ValorDois1*1"
for i = 2 to 50
u = u&" + ValorUm"&i&"*1 * ValorDois"&i&"*1"
next
u = u&";"&vbCrLf 
u = u&"            document.getElementById('total').innerHTML = document.getElementById('total').innerHTML / 100"&vbCrLf
u = u&"                document.getElementById('total').innerHTML = document.getElementById('total').innerHTML.replace('.',',');"&vbCrLf
u = u&"				       if (document.getElementById('total').innerHTML == undefined) {"&vbCrLf
u = u&"				       		document.getElementById('total').innerHTML = '0,00';"&vbCrLf
u = u&"				       }"&vbCrLf
u = u&"				       if (document.getElementById('total').innerHTML == 0) {"&vbCrLf
u = u&"				       		document.getElementById('total').innerHTML = '0,00';"&vbCrLf
u = u&"				       }"&vbCrLf
u = u&"				       funccomplete2 = document.getElementById('total').innerHTML.split(',');"&vbCrLf
u = u&"				       if (funccomplete2[0].length < 4) {"&vbCrLf
u = u&"				           	document.getElementById('total').innerHTML = 'R$ '+funccomplete2[0]+','+funccomplete2[1];"&vbCrLf
u = u&"				       }"&vbCrLf
u = u&"					   if (funccomplete2[1]==undefined) {"&vbCrLf
u = u&"				            document.getElementById('total').innerHTML = 'R$ '+funccomplete2[0]+',00';"&vbCrLf
u = u&"					   }"&vbCrLf
u = u&"				       if (funccomplete2[1].length < 2) {"&vbCrLf
u = u&"				           	document.getElementById('total').innerHTML = 'R$ '+funccomplete2[0]+','+funccomplete2[1]+'0';"&vbCrLf
u = u&"				       }"&vbCrLf
u = u&"				       if (funccomplete2[1].length == 2) {"&vbCrLf
u = u&"				           	document.getElementById('total').innerHTML = 'R$ '+funccomplete2[0]+','+funccomplete2[1];"&vbCrLf
u = u&"				       }"&vbCrLf
u = u&"				       if (funccomplete2[0].length > 3) {"&vbCrLf
u = u&"				       		func = funccomplete2[0]*1 / 1000;"&vbCrLf
u = u&"				       		funccomplete3 = func.value.split('.');"&vbCrLf
u = u&"							       	if (funccomplete3[1].length == 1) {"&vbCrLf
u = u&"					                      document.getElementById('total').innerHTML = 'R$ '+funccomplete3[0]+'.'+funccomplete3[1]+'00,'+funccomplete2[1];"&vbCrLf
u = u&"									}"&vbCrLf
u = u&"							       	if (funccomplete3[1].length == 2) {"&vbCrLf
u = u&"					                      document.getElementById('total').innerHTML = 'R$ '+funccomplete3[0]+'.'+funccomplete3[1]+'0,'+funccomplete2[1];"&vbCrLf
u = u&"									}"&vbCrLf
u = u&"							       	if (funccomplete3[1].length == 3) {"&vbCrLf
u = u&"					                      document.getElementById('total').innerHTML = 'R$ '+funccomplete3[0]+'.'+funccomplete3[1]+','+funccomplete2[1];"&vbCrLf
u = u&"									}"&vbCrLf
u = u&"					   }"&vbCrLf
u = u&"            }"&vbCrLf
u = u&"</script>"&vbCrLf
response.write u
end if%>
<script type="text/javascript" language="javascript">
function alta(valor) {
valor.value=valor.value.toUpperCase();
}
</script>
</head>
<body topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" marginwidth="0" marginheight="0">
<div name="catalogoon" id="catalogoon" class="catalogoon">
	<div class="catalogocenter">
    	<div style="position:relative;left:-370px;top:-259px;width:740px;height:50px;" align="right">
		  <img src="./imagem/fechar.png" alt="Fechar" title="Fechar" onClick="fecharRevista()" style="position:relative" >
		</div>
    	<div class="catalogolivro">
<object type="application/x-shockwave-flash" id="flashcontent" name="flashcontent" data="http://static.issuu.com/webembed/viewers/style1/v2/IssuuReader.swf" width="100%" height="100%" style="visibility: visible;">
    <param name="movie" value="http://static.issuu.com/webembed/viewers/style1/v2/IssuuReader.swf">
    <param name="allowfullscreen" value="true">
    <param name="menu" value="false">
    <param name="salign" value="tl">
    <param name="scale" value="noscale">
    <param name="wmode" value="transparent">
    <param name="allowscriptaccess" value="always">
    <param name="flashvars" value="mode=window&amp;pageNumber=1&amp;jsAPIClientDomain=issuu.com&amp;jsAPIInitCallback=jsAPIInitCallback&amp;jsInternalCallback=jsInternalCallback&amp;viewMode=doublePage&amp;documentId=140702202917-1074923b4a1dace6d734fc007d09a186&amp;proSidebarEnabled=false&amp;showHtmlLink=true&amp;shareMenuEnabled=true&amp;backgroundImage=&amp;backgroundStretch=true&amp;shareButtonEnabled=true&amp;searchButtonEnabled=true&amp;clippingEnabled=true&amp;autoFlip=false&amp;linkTarget=_blank&amp;theme=&amp;layout=&amp;logo=&amp;embedId=12469288%2F8480718&amp;name=sexydany&amp;username=mrsistemas&amp;creator=mrsistemas&amp;printButtonEnabled=false&amp;watermarkEnabled=true">
</object>
        </div>
    </div>
</div>
<%If Application("pag") <> "inicial" and Application("pag") <> "central.qs" and Application("pag") <> "central.cr" and Application("pag") <> "central.pp" and Application("pag") <> "central.cr1" and Application("pag") <> "central.dp"  and Application("pag") <> "central.pp1" and Application("pag") <> "central.pp2" and Application("pag") <> "central.pp.ok" and Application("pag") <> "central.js" and Application("pag") <> "central.fc" Then%>
<link href="http://www.artlizando.com.br/v1/_style/style.css" type="text/css" rel="stylesheet">
<table border="0" width="100%" cellspacing="0" cellpadding="0" id="table1" height="100%">
	<tr>
		<td align="center" colspan="3">
		<b style="color: rgb(0, 0, 0); font-family: Times New Roman; font-style: normal; font-variant: normal; letter-spacing: normal; line-height: normal; orphans: 2; text-align: -webkit-center; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-size-adjust: auto; -webkit-text-stroke-width: 0px; font-size: medium">
		<i><font face="Arial" color="#E06022">A página não foi localizada!</font><font face="Arial" size="2" color="#E06022"><br>
		Erro: HTTP 404</font><font face="Arial" size="2" color="#CC0000"><br>
		</font><font face="Arial" size="2" color="#999999"><br>
		</font></i></b>
		<font style="font-style: normal; font-variant: normal; font-weight: normal; letter-spacing: normal; orphans: 2; text-align: -webkit-center; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-size-adjust: auto; -webkit-text-stroke-width: 0px; color: rgb(153, 153, 153); line-height: 11pt; font-size: 8pt; font-family: verdana">
		A página que você está procurando pode ter sido removida,<br>
		o seu nome pode ter sido alterado ou talvez ela não esteja disponível 
		temporariamente.<br>
		<br>
		</font>
		<font style="font-style: normal; font-variant: normal; letter-spacing: normal; orphans: 2; text-align: -webkit-center; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-size-adjust: auto; -webkit-text-stroke-width: 0px; color: rgb(153, 153, 153); line-height: 11pt; font-size: 8pt; font-family: verdana; font-weight: 700">
		Siga um destes procedimentos:</font><font color="#999999" style="font-family: Times New Roman; font-style: normal; font-variant: normal; font-weight: normal; letter-spacing: normal; line-height: normal; orphans: 2; text-align: -webkit-center; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-size-adjust: auto; -webkit-text-stroke-width: 0px; font-size: medium"><br>
		</font>
		<font color="#999999" style="font-style: normal; font-variant: normal; font-weight: normal; letter-spacing: normal; orphans: 2; text-align: -webkit-center; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-size-adjust: auto; -webkit-text-stroke-width: 0px; line-height: 11pt; font-size: 8pt; font-family: verdana">
		- Se você tiver digitado o endereço da página na barra de endereços, 
		certifique-se de que ele tenha sido digitado corretamente.<br>
		- Abra a página inicial<span class="Apple-converted-space">&nbsp;</span>e 
		procure os links para as informações desejadas.<br>
		- Clique no botão<span class="Apple-converted-space">&nbsp;</span></font><font style="color: rgb(0, 0, 0); font-style: normal; font-variant: normal; font-weight: normal; letter-spacing: normal; orphans: 2; text-align: -webkit-center; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-size-adjust: auto; -webkit-text-stroke-width: 0px; line-height: 11pt; font-size: 8pt; font-family: verdana"><a href="javascript:history.back(1)"><font color="#808080">Voltar</font></a></font><font color="#999999" style="font-style: normal; font-variant: normal; font-weight: normal; letter-spacing: normal; orphans: 2; text-align: -webkit-center; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-size-adjust: auto; -webkit-text-stroke-width: 0px; line-height: 11pt; font-size: 8pt; font-family: verdana"><span class="Apple-converted-space">&nbsp;</span>e 
		tente outro link.</font></td>
	</tr>
	<tr>
		<td align="right" height="50" width="10">&nbsp;</td>
		<td align="right" height="50"><a href="http://www.artlizando.com.br/v1/"><img src="http://www.artlizando.com.br/v1/_image/studio_artlizando.png" border="0"></a></td>
		<td align="right" height="50" width="10">&nbsp;</td>
	</tr>
</table>
<%
Else
%>
<div align="center">
	<table border="0" width="100%" cellspacing="0" cellpadding="0" id="table1" height="100%">
		<tr>
			<td align="center" valign="top" height="45">
			<table border="0" width="988" cellspacing="0" cellpadding="0" id="table2">
				<tr>
					<td width="42">
					<img border="0" src="imagem/cantol_barra.jpg" width="42" height="35"></td>
					<td background="imagem/barra_up.jpg" align="right" valign="top">
					<table border="0" width="450" cellspacing="0" cellpadding="0" id="table4" class="menu" height="28">
						<tr>
							<td height="26" align="right"><a href="../site/">Início</a></td>
							<td height="26" align="right"><a href="?h=central.qs">Quem Somos</a></td>
							<td height="26" align="right"><a href="?h=central.dp">Data dos Pedidos</a></td>
							<td height="26" align="right"><a href="?h=central.pp">Passar pedidos</a></td>
							<td height="26" align="right"><a href="?h=central.fc">Fale conosco</a></td>
						</tr>
					</table>
					</td>
					<td width="42" align="right">
					<img border="0" src="imagem/cantor_barra.jpg" width="42" height="35"></td>
				</tr>
			</table>
			<table border="0" width="988" cellspacing="0" cellpadding="0" id="table6">
						<tr>
							<td width="30">&nbsp;
							</td>
							<td width="235">&nbsp;
							</td>
							<td>&nbsp;</td>
						</tr>
						<tr>
							<td width="30">&nbsp;
							</td>
							<td width="235">
							<img border="0" src="imagem/logocentral.jpg"></td>
							<td align="right" style="position:relative;"><img border="0" style="position:absolute;left:160px;top:-5px;width:200px;height:101px;" src="imagem/cerotic.png" onClick="abrirRevista();">
							<a href="?h=central.cr1"><img border="0" src="imagem/seja.jpg" width="139" height="51"></a>&nbsp;&nbsp; &nbsp;&nbsp;<br>
							<a href="http://www.facebook.com/centraldos.catalogos" target="_blanc">
							<img border="0" src="imagem/facebook.png" width="33" height="33"></a>
							<img border="0" src="imagem/twitter.png" width="33" height="33">&nbsp;&nbsp;&nbsp; &nbsp;&nbsp; </td>
						</tr>
						<tr>
							<td width="30">&nbsp;
							</td>
							<td width="235">&nbsp;
							</td>
							<td>&nbsp;</td>
						</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td align="center" valign="top">
<%
Select Case Application("pag")
case "inicial"%>
			<table border="0" width="988" cellspacing="0" cellpadding="0" id="table3" class="conteudo">
				<tr>
					<td height="299" align="center" background="imagem/loading.jpg">
					<object
        				classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"
        				codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=8,0,42,0"
        				id="Movie3"
        				width="985" height="299"
      				>
        			<param name="movie" value="swf/banner.swf">
        			<param name="bgcolor" value="#FFFFFF">
        			<param name="wmode" value="transparent">
        			<param name="quality" value="high">
        			<param name="seamlesstabbing" value="false">
        			<param name="allowscriptaccess" value="samedomain">
        			<embed
          				wmode="transparent"
          				type="application/x-shockwave-flash"
          				pluginspage="http://www.adobe.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash"
          				name="Movie3"
          				width="985" height="299"
          				src="swf/banner.swf"
          				bgcolor="#FF6600"
          				quality="high"
          				seamlesstabbing="false"
          				allowscriptaccess="samedomain"
        				>
          			<noembed>
          			</noembed>
        			</embed>
        			</object>
					</td>
				</tr>
				<tr>
					<td height="380" align="center">
					<img border="0" src="imagem/un.jpg" width="988" height="380"></td>
				</tr>
			</table>
<%
case "central.qs"
%>

<table border="0" width="988" cellspacing="0" cellpadding="0" height="100%" id="table1">
	<tr>
		<td valign="top" align="center">
		<table border="0" width="929" cellspacing="0" cellpadding="0" id="table13" class="conteudo">
			<tr>
				<td>
				<p align="justify"><b><font size="1">&nbsp;</font><font size="4"><br>
				Quem Somos</font><br>
		<br>
		Tecnologia, confiabilidade e respeito, é assim que se descreve nossa 
		empresa! Com experiência em vendas, a Central dos Catálogos dispõe às 
		suas revendedoras um site moderno, prático e seguro, trazendo mais 
		conforto na hora de passar seus pedidos e verificar as promoções.
		A Central dos Catálogos prioriza as necessidades de nossa revendedora 
		trazendo benefícios que tornam suas atividades mais práticas e simples, 
		por esse motivo temos a admiração de quem já trabalha com a gente.<br>
		<br>
		Faça você também parte dessa família e lucre cada vez mais!!!</b></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td height="314" background="imagem/empresa.jpg">&nbsp;</td>
	</tr>
</table>

<%
case "central.dp"
%>
			<table border="0" width="988" cellspacing="0" cellpadding="0" id="table3">
				<tr>
					<td width="30" height="70">&nbsp;</td>
					<td width="479" height="70" valign="bottom">
					<img border="0" src="Imagem/botao_dped.jpg"></td>
					<td width="452" align="right" height="70">&nbsp;
					</td>
					<td width="27" height="70">&nbsp;</td>
				</tr>
				<tr>
					<td align="center" valign="top" colspan="4">
					<table border="0" width="988" cellspacing="0" cellpadding="0" id="table4">
						<tr>
							<td width="27">
							<img border="0" src="Imagem/canto_cinza1.jpg" width="27" height="27"></td>
							<td bgcolor="#F5F5F5">&nbsp;</td>
							<td align="right" width="27">
							<img border="0" src="Imagem/canto_cinza2.jpg" width="27" height="27"></td>
						</tr>
					</table>
					</td>
				</tr>
				<tr>
					<td align="center" valign="top" colspan="4">
<%
Application("cat") = request.querystring("catalogo")
if Application("cat") = "" then
Application("cat") = 1
end if
if Application("cat") = 2 or Application("cat") = 7 then
u = 2
else
u = 1
end if
if Application("cat") = 4 or Application("cat") = 9 then
i = 2
else
i = 1
end if
%>
					<table border="0" width="100%" cellspacing="0" cellpadding="2" id="table1">
	<tr>
		<td colspan="3" height="10" bgcolor="#F5F5F5">
		<span style="font-size: 1pt">&nbsp;</span></td>
	</tr>
	<tr>
		<td colspan="3" height="50" bgcolor="#F5F5F5" align="center">
		<table cellspacing="0" cellpadding="0"><tr><td>
		<%if Application("cat") <= 1 then%>
		<img border="0" src="imagem/setaesqd.jpg" width="23" height="29">
		<%else%>
		<a href="?h=central.dp&catalogo=<%=Application("cat")-i%>"><img border="0" src="imagem/setaesq.jpg" width="23" height="29"></a>
		<%end if%>
		</td>
		<td height="50" bgcolor="#F5F5F5" align="center">
		<img border="0" src="../site/imagem/C<%=Application("cat")%>.jpg" width="934" height="60"></td>
		<td height="50" bgcolor="#F5F5F5" align="center">
		<%if Application("cat") >= 10 then%>
		<img border="0" src="imagem/setadird.jpg" width="23" height="29">
		<%else%>
		<a href="?h=central.dp&catalogo=<%=Application("cat")+u%>"><img border="0" src="imagem/setadir.jpg" width="23" height="29"></a>
		<%end if%>
		</td></tr>
		</table>
		</td>
	</tr>
	<tr>
		<td colspan="3" height="10" bgcolor="#F5F5F5">
		<span style="font-size: 1pt">&nbsp;</span></td>
	</tr>
	<tr>
		<td width="33%" bgcolor="#F5F5F5" valign="top" align="center">
<table border="0" width="300" cellspacing="0" cellpadding="5" id="table1" style="border-width: 0px">
	<tr>
		<td width="40" height="35" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="2" face="Arial" color="#000000">Mês</font></b></td>
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
		<td width="40" height="20" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b>
		<font face="Arial" color="#000000" style="font-size: 9pt">Datas</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20" width="40">
		<b>
		<font face="Arial" color="#000000" style="font-size: 9pt">Código</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20" width="110">
		<b>
		<font face="Arial" color="#000000" style="font-size: 9pt">Data vai</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20" width="110">
		<b>
		<font face="Arial" color="#000000" style="font-size: 9pt">Data vem</font></b></td>
	</tr>
<%
for u = 1 to 5
Set c1 = conexao.execute("SELECT * FROM datas WHERE iddata = 'C"&Application("cat")&""&u&"' and mes = "&mes&" and datavai <> """" order by iddata DESC")
if not c1.eof then
%>
	<tr>
		<td width="40" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font face="Arial" color="#000000" style="font-size: 9pt">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="30" width="40">
		<font style="font-size: 9pt">
		<b>
		<%=c1("codcat")%></b></font></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="30" width="110">
		<font style="font-size: 9pt">
		<b>
		<%=c1("datavai")%></b></font></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="30" width="110">
		<font style="font-size: 9pt">
		<b>
		<%=c1("datavem")%></b></font></td>
	</tr>
<%
end if
c1.close
Set c1 = Nothing
next
%>
</table></td>
		<td width="33%" bgcolor="#F5F5F5" valign="top" align="center">
<table border="0" width="300" cellspacing="0" cellpadding="5" id="table1" style="border-width: 0px">
	<tr>
		<td width="40" height="35" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="2" face="Arial" color="#000000">Mês</font></b></td>
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
		<td width="40" height="20" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b>
		<font face="Arial" color="#000000" style="font-size: 9pt">Datas</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20" width="40">
		<b>
		<font face="Arial" color="#000000" style="font-size: 9pt">Código</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20" width="110">
		<b>
		<font face="Arial" color="#000000" style="font-size: 9pt">Data vai</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20" width="110">
		<b>
		<font face="Arial" color="#000000" style="font-size: 9pt">Data vem</font></b></td>
	</tr>
<%
for u = 1 to 5
Set c1 = conexao.execute("SELECT * FROM datas WHERE iddata = 'C"&Application("cat")&""&u&"' and mes = "&mes&" and datavai <> """" order by iddata DESC")
if not c1.eof then
%>
	<tr>
		<td width="40" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font face="Arial" color="#000000" style="font-size: 9pt">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="30" width="40">
		<font style="font-size: 9pt">
		<b>
		<%=c1("codcat")%></b></font></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="30" width="110">
		<font style="font-size: 9pt">
		<b>
		<%=c1("datavai")%></b></font></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="30" width="110">
		<font style="font-size: 9pt">
		<b>
		<%=c1("datavem")%></b></font></td>
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
<table border="0" width="300" cellspacing="0" cellpadding="5" id="table1" style="border-width: 0px">
	<tr>
		<td width="40" height="35" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font size="2" face="Arial" color="#000000">Mês</font></b></td>
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
		<td width="40" height="20" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b>
		<font face="Arial" color="#000000" style="font-size: 9pt">Datas</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20" width="40">
		<b>
		<font face="Arial" color="#000000" style="font-size: 9pt">Código</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20" width="110">
		<b>
		<font face="Arial" color="#000000" style="font-size: 9pt">Data vai</font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="20" width="110">
		<b>
		<font face="Arial" color="#000000" style="font-size: 9pt">Data vem</font></b></td>
	</tr>
<%
for u = 1 to 5
Set c1 = conexao.execute("SELECT * FROM datas WHERE iddata = 'C"&Application("cat")&""&u&"' and mes = "&mes&" and datavai <> """" order by iddata DESC")
if not c1.eof then
%>
	<tr>
		<td width="40" height="30" style="border-style:none; border-width:medium; " bgcolor="#F0F0F0">
		<b><font face="Arial" color="#000000" style="font-size: 9pt">Data <%=u%></font></b></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="30" width="40">
		<font style="font-size: 9pt">
		<b>
		<%=c1("codcat")%></b></font></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="30" width="110">
		<font style="font-size: 9pt">
		<b>
		<%=c1("datavai")%></b></font></td>
		<td style="border-style:none; border-width:medium; " align="left" bgcolor="#F0F0F0" height="30" width="110">
		<font style="font-size: 9pt">
		<b>
		<%=c1("datavem")%></b></font></td>
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
</table>
					</td>
				</tr>
				<tr>
					<td colspan="4" align="right">
					<table border="0" width="988" cellspacing="0" cellpadding="0" id="table4">
						<tr>
							<td width="27">
							<img border="0" src="Imagem/canto_cinza3.jpg" width="27" height="27"></td>
							<td bgcolor="#F5F5F5">&nbsp;</td>
							<td align="right" width="27">
							<img border="0" src="Imagem/canto_cinza4.jpg" width="27" height="27"></td>
						</tr>
					</table>
					</td>
				</tr>
				<tr>
					<td colspan="2" height="70" align="right">&nbsp;</td>
					<td height="70" align="right">&nbsp;</td>
					<td height="70" align="right">&nbsp;</td>
				</tr>
				<tr>
					<td colspan="4">&nbsp;</td>
				</tr>
			</table>
<%
case "central.cr"
%>
<table border="0" width="988" cellspacing="0" cellpadding="0" id="table1">
	<tr>
		<td align="center">
		<table border="0" width="929" cellspacing="0" cellpadding="0" id="table13" class="conteudo">
			<tr>
				<td>
				<p align="justify"><b><font size="1">&nbsp;</font><font size="4"><br>
				Veja alguns benefícios!</font><br>
				<br><br>
				</b>- As revendedoras podem passar seus pedidos por telefone, 
				pelo site ou poderá também trazer em nossa loja.<br>
				<br>
				- Entregamos suas caixas em sua residência para sua maior 
				comodidade.<br>
				<br>
				- As revendedoras juntam pontos para trocar por lindos prêmios.<br>
				<br>
				- Ganham brindes mandando seus pedidos todos os meses.<br>
				<br>
				- Pagamento no boleto bancário ou através do cartão de crédito 
				em seus pedidos.<br>
				<br>
				E muito mais! Clique <a href="?h=central.cr1">aqui</a>, cadastre-se e tenha você também 
				todas essas vantagens!!! <br>
				<br>
&nbsp;</td>
			</tr>
		</table>
		</td>
	</tr>
</table>
<%
case "central.pp"
%>
<table border="0" width="988" id="table1" cellspacing="0" cellpadding="0">
<form action="index.asp" method="POST">
<input type="hidden" name="h" id="h" value="central.pp1">
	<%if Application("erro") <> "" then%>
	<tr>
		<td width="35" height="25" bgcolor="#FFFF99">&nbsp;</td>
		<td height="25" colspan="3" bgcolor="#FFFF99" align="left">
<br>
				<b><font size="4">Erro!!!<br>
				</font><font style="font-size: 3pt">&nbsp;</font><font size="4"><br>
				</font><span style="font-size: 9pt">
 				<%if Application("erro2") = "1019831" then%>- Digite seu nome!<br><%end if%>
				<%if Application("erro3") = "1019832" then%>- Nome inválido!<br><%end if%>
 				<%if Application("erro10") = "10110831" then%>- Digite seu e-mail!<br><%end if%>
				<%if Application("erro11") = "10110832" then%>- E-mail inválido!<br><%end if%>
				<%if Application("erro4") = "1029831" then%>- Digite a data de seu nascimento!<br><%end if%>
				<%if Application("erro5") = "1029832" then%>- Data de nascimento inválida!<br><%end if%>
				<%if Application("erro6") = "1039831" then%>- Digite o DDD!<br><%end if%>
				<%if Application("erro7") = "1039832" then%>- DDD inválido!<br><%end if%>
				<%if Application("erro8") = "1049831" then%>- Digite seu telefone!<br><%end if%>
				<%if Application("erro9") = "1049832" then%>- Telefone inválido!<br><%end if%>
				<%if Application("erro1") = "1009833" then%>- Selecione o catálogo que deseja enviar pedido!<br><%end if%>
                <%if Application("erro13") = "1009834" then%>- Digite o nome do catálogo!<br><%end if%>
				</span>
&nbsp;</b>
		</td>
		<td height="25" bgcolor="#FFFF99">&nbsp;</td>
	</tr>
	<tr>
		<td width="35" height="25">&nbsp;</td>
		<td height="25" colspan="4">&nbsp;</td>
	</tr>
	<%end if%>
	<tr>
		<td width="35" height="50">&nbsp;</td>
		<td height="50" colspan="4"><b><font size="4">Digite seus dados</font></b></td>
	</tr>
						<tr>
							<td align="left" width="27" height="25">&nbsp;</td>
							<td width="935" height="25" colspan="4">
							<table border="0" width="100%" cellspacing="0" cellpadding="0" id="table9" class="conteudo">
								<tr>
									<td height="25" colspan="3"><b>Nome do(a) 
									revendedor(a)</b></td>
								</tr>
								<tr>
									<td height="25" colspan="3">
							<input type="text" name="nome" id="descr<%=i%>4" tabindex="1" size="20" style="width: 435;" class="input" value="<%=Application("nome")%>"></td>
								</tr>
								<tr>
									<td height="25" colspan="3"><b>E-mail do(a) 
									revendedor(a)</b></td>
								</tr>
								<tr>
									<td height="35" colspan="3">
							<input type="text" name="email" id="descr<%=i%>0" tabindex="1" size="20" style="width: 435;" class="input" value="<%=Application("mail")%>"></td>
								</tr>
								<tr>
									<td height="25" width="23%"><b>Data de Nascimento <i>
									<font size="1">(Ex: DDMMAAAA)</font></i></b></td>
									<td height="25" colspan="2"><b>DDD e Telefone
									<i><font size="1">(somente os números)</font></i></b></td>
								</tr>
								<tr>
									<td height="25" width="23%">
							<input type="text" name="cpf" id="descr<%=i%>1" tabindex="2" size="11" style="width: 170;" class="input" maxlength="10" value="<%=Application("cpf")%>" OnKeyPress="return mascaraGenerica(event, this, '##/##/####');"></td>
									<td height="25" width="4%">
							<input type="text" name="ddd" id="descr<%=i%>3" tabindex="3" size="10" style="width: 35;" class="input" maxlength="2" value="<%=Application("ddd")%>"></td>
									<td height="25" width="75%">
							<input type="text" name="telefone" id="descr<%=i%>2" tabindex="4" size="10" style="width: 170;" class="input" value="<%=Application("telefone")%>" maxlength="8"></td>
								</tr>
								</table>
							</td>
						</tr>
	<tr>
		<td width="35" height="25">&nbsp;</td>
		<td height="25" colspan="4">&nbsp;</td>
	</tr>
	<tr>
		<td width="35" height="50">&nbsp;</td>
		<td height="50" colspan="4"><b><font size="4">Selecione qual o catálogo para passar seu 
		pedido</font></b></td>
	</tr>
	<tr>
		<td width="35" height="25">&nbsp;</td>
		<td width="37" height="25">
		<input type="radio" value="C2" name="catalogo" tabindex="6" <%if Application("catalogo") = "C2" then%>checked<%end if%>></td>
		<td height="25" width="916" colspan="3">
		<span style="font-size: 9pt; font-weight: 700">Quatro Estações / Quintess / Bazar Brasil / 
		Clarissa / Actual / Natubelly / Marcyn</span></td>
	</tr>
	<tr>
		<td width="35" height="25">&nbsp;</td>
		<td width="37" height="25">
		<input type="radio" value="C1" name="catalogo" tabindex="5" <%if Application("catalogo") = "C1" then%>checked<%end if%>></td>
		<td height="25" width="916" colspan="3"><b><span style="font-size: 9pt">Hiroshima / Essencial / IVETE / L'Arc en 
		Ciel</span></b></td>
	</tr>
	<tr>
		<td width="35" height="25">&nbsp;</td>
		<td width="37" height="25">
		<input type="radio" value="C7" name="catalogo" tabindex="11" <%if Application("catalogo") = "C7" then%>checked<%end if%>></td>
		<td height="25" width="458">
		<span style="font-size: 9pt; font-weight: 700">Brasil Tropical / Biodany / Sexydany</span></td>
		<td height="25" width="37">&nbsp;</td>
	</tr>
	<tr>
		<td width="35" height="25">&nbsp;</td>
		<td width="37" height="25">
		<input type="radio" value="C4" name="catalogo" tabindex="8" <%if Application("catalogo") = "C4" then%>checked<%end if%>></td>
		<td height="25" width="916" colspan="3">
		<span style="font-size: 9pt; font-weight: 700">Luzon</span></td>
	</tr>
	<tr>
		<td width="35" height="25">&nbsp;</td>
		<td width="37" height="25">
		<input type="radio" value="C5" name="catalogo" tabindex="9" <%if Application("catalogo") = "C5" then%>checked<%end if%>></td>
		<td height="25" width="916" colspan="3">
		<span style="font-size: 9pt; font-weight: 700">Miro Star</span></td>
	</tr>
	<tr>
		<td width="35" height="25">&nbsp;</td>
		<td width="37" height="25">
		<input type="radio" value="C6" name="catalogo" tabindex="10" <%if Application("catalogo") = "C6" then%>checked<%end if%>></td>
		<td height="25" width="916" colspan="3">
		<span style="font-size: 9pt; font-weight: 700">Abelha Rainha</span></td>
	</tr>
	<tr>
		<td width="35" height="25">&nbsp;</td>
		<td width="37" height="25">
		<input type="radio" value="C9" name="catalogo" tabindex="10" <%if Application("catalogo") = "C9" then%>checked<%end if%>></td>
		<td height="25" width="916" colspan="3">
		<span style="font-size: 9pt; font-weight: 700">Vitória Régia</span></td>
	</tr>
	<tr>
		<td width="35" height="25">&nbsp;</td>
		<td width="37" height="25">
		<input type="radio" value="C10" name="catalogo" tabindex="10" <%if Application("catalogo") = "C10" then%>checked<%end if%>></td>
		<td height="25" width="458">
		<span style="font-size: 9pt; font-weight: 700">Expressão Feminina</span></td>
		<td height="25" width="421" rowspan="2" align="right">
		<span style="font-size: 9pt; font-weight: 700">
		<input type="image" src="Imagem/proximo.jpg" id="mminvite" name="Ifa2fdaekjhhft.904$6y9kgx"></span></td>
		<td height="25" width="37">&nbsp;</td>
	</tr>
	<tr>
		<td width="35" height="25">&nbsp;</td>
		<td width="37" height="25">
		<input type="radio" value="C8" name="catalogo" tabindex="12" <%if Application("catalogo") = "C8" then%>checked<%end if%>></td>
		<td height="25" width="458">
		<span style="font-size: 9pt; font-weight: 700">Outros&nbsp;&nbsp;<input type="text" id="qualcat" name="qualcat" value="<%=Application("qualcat")%>">&nbsp;&nbsp;&nbsp;Digite o nome do catálogo.</span></td>
		<td height="25" width="37">&nbsp;</td>
	</tr>
	<tr>
		<td width="35" height="25">&nbsp;</td>
		<td width="37" height="25">&nbsp;</td>
		<td height="25" width="916" colspan="3">&nbsp;</td>
	</tr>
	</form>
</table>
<%
case "central.pp1"
Application("nome") = request.form("nome")
Application("mail") = request.form("email")
Application("cpf") = request.form("cpf")
Application("ddd") = request.form("ddd")
Application("telefone") = request.form("telefone")
Application("revista") = request.form("catalogo")
Application("qualcat") = request.form("qualcat")
erros = ""
error = 0
if Application("nome") = "" then
erros = erros&"&error2=1019831"
error = error+1
end if
if Application("mail") = "" then
erros = erros&"&error10=10110831"
error = error+1
else
if Application("mail") = "" or instr(Application("mail"),"@") < 1 or instr(Application("mail"),".com") < 1 and instr(Application("mail"),".jur") < 1 and instr(Application("mail"),".net") < 1 and instr(Application("mail"),".org") < 1 and instr(Application("mail"),".info") < 1 and instr(Application("mail"),".gov") < 1 then
erros = erros&"&error11=10110832"
error = error+1
end if
end if
if isnumeric(Application("nome")) = true and trim(Application("nome")) <> ""  or len(Application("nome")) < 3 and trim(Application("nome")) <> "" then
erros = erros&"&error3=1019832"
error = error+1
end if
if Application("cpf") = "" then
erros = erros&"&error4=1029831"
error = error+1
else
if len(replace(Application("cpf"),"/","")) < 8 or isnumeric(replace(Application("cpf"),"/","")) = false then
erros = erros&"&error5=1029832"
error = error+1
end if
end if
if Application("ddd") = "" then
erros = erros&"&error6=1039831"
error = error+1
end if
if isNumeric(Application("ddd")) = false and trim(Application("ddd")) <> "" or len(Application("ddd")) <> 2 and trim(Application("ddd")) <> "" then
erros = erros&"&error7=1039832"
error = error+1
end if
if Application("telefone") = "" then
erros = erros&"&error8=1049831"
error = error+1
end if
if isnumeric(Application("telefone")) = false and trim(Application("telefone")) <> ""  or len(Application("telefone")) <> 8 and trim(Application("telefone")) <> "" then
erros = erros&"&error9=1049832"
error = error+1
end if
if Application("revista") <> "C1" and Application("revista") <> "C2" and Application("revista") <> "C3" and Application("revista") <> "C4" and Application("revista") <> "C5" and Application("revista") <> "C6" and Application("revista") <> "C7" and Application("revista") <> "C8" and Application("revista") <> "C9" and Application("revista") <> "C10" then
erros = erros&"&error1=1009833"
error = error+1
end if
if Application("revista") = "C8" and Application("qualcat") = "" then
erros = erros&"&error13=1009834"
error = error+1
end if
if error > 0 then
response.redirect("?h=central.pp&revista="&Application("revista")&"&error=sI&nome="&application("nome")&"&email="&application("mail")&"&cpf="&Application("cpf")&"&ddd="&Application("ddd")&"&telefone="&application("telefone")&"&qualcat="&Application("qualcat")&""&erros&"&x=adybia103nov")
end if
%>
<div id="somatotal" name="somatotal">
	<table border="0" width="100%" id="table7" cellspacing="0" cellpadding="0" height="100%">
		<tr>
			<td background="imagem/total.png">
			
			<table border="0" width="100%" cellspacing="0" cellpadding="0" height="100%" id="table8">
				<tr>
					<td height="23">&nbsp;</td>
				</tr>
				<tr>
					<td align="center"><div id="total" class="soma"></div></td>
				</tr>
			</table>
			
			</td>
		</tr>
	</table>
</div>
			<table border="0" width="988" cellspacing="0" cellpadding="0" id="table3">
			<form name="form" id="form" action="index.asp" method="POST">
			<input type="hidden" name="h" id="h" value="central.pp2">
			<input type="hidden" name="nome" id="nome" value="<%=Application("nome")%>">
			<input type="hidden" name="email" id="email" value="<%=Application("mail")%>">
			<input type="hidden" name="cpf" id="cpf" value="<%=Application("cpf")%>">
			<input type="hidden" name="ddd" id="ddd" value="<%=Application("ddd")%>">
			<input type="hidden" name="telefone" id="telefone" value="<%=Application("telefone")%>">
			<input type="hidden" name="revista" id="revista" value="<%=Application("revista")%>">
            <input type="hidden" name="qualcat" id="qualcat" value="<%=Application("qualcat")%>">
				<tr>
					<td width="30" height="70">&nbsp;</td>
					<td width="479" height="70" valign="bottom">
					<img border="0" src="Imagem/botao_ped.jpg"></td>
					<td width="452" align="right" height="70">
					<input type="image" src="Imagem/proximo.jpg" name="I1"></td>
					<td width="27" height="70">&nbsp;</td>
				</tr>
				<tr>
					<td align="center" valign="top" colspan="4">
					<table border="0" width="988" cellspacing="0" cellpadding="0" id="table4">
						<tr>
							<td width="27">
							<img border="0" src="Imagem/canto_cinza1.jpg" width="27" height="27"></td>
							<td bgcolor="#F5F5F5">&nbsp;</td>
							<td align="right" width="27">
							<img border="0" src="Imagem/canto_cinza2.jpg" width="27" height="27"></td>
						</tr>
					</table>
					</td>
				</tr>
				<tr>
					<td align="center" valign="top" colspan="4">
					<table border="0" width="988" cellspacing="0" cellpadding="0" id="table5" class="conteudo">
						<tr>
							<td align="left" width="27" bgcolor="#F5F5F5" height="25">&nbsp;</td>
							<td bgcolor="#F5F5F5" width="935" height="25" colspan="9" <%if Application("revista") <> "C8" then%>align="center"<%else%>align="left"<%end if%>>
							<%if Application("revista") <> "C8" then%><img src="imagem/<%=Application("revista")%>.jpg"><%else%><font color="#000" size="14pt"><%=Application("qualcat")%></font><%end if%></td>
							<td align="right" width="27" bgcolor="#F5F5F5" height="25">&nbsp;</td>
						</tr>
						<tr>
							<td align="left" width="27" bgcolor="#F5F5F5" height="25">&nbsp;</td>
							<td bgcolor="#F5F5F5" width="935" height="25" colspan="9">&nbsp;</td>
							<td align="right" width="27" bgcolor="#F5F5F5" height="25">&nbsp;</td>
						</tr>
						<tr>
							<td align="left" width="27" bgcolor="#F5F5F5" height="25">&nbsp;</td>
							<td bgcolor="#F5F5F5" width="45" height="25"><b>Quant.</b></td>
							<td bgcolor="#F5F5F5" width="67" height="25"><b>Referência </b></td>
							<td bgcolor="#F5F5F5" width="22" height="25">&nbsp;
							</td>
							<td bgcolor="#F5F5F5" width="45" height="25"><b>Tam.</b></td>
							<td bgcolor="#F5F5F5" width="445" height="25"><b>Descrição</b></td>
							<td bgcolor="#F5F5F5" width="51" height="25"><b>Página</b></td>
							<td bgcolor="#F5F5F5" width="90" height="25"><b>Cor</b></td>
							<td bgcolor="#F5F5F5" width="90" height="25"><b>Preço Unit.</b></td>
							<td bgcolor="#F5F5F5" width="80" height="25"><b>Preço Total</b></td>
							<td align="right" width="27" bgcolor="#F5F5F5" height="25">&nbsp;</td>
						</tr>
						<%
						x = 0
						for i = 1 to 50
						%>
			<tr>
							<td align="left" width="27" bgcolor="#F5F5F5" height="35">&nbsp;</td>
							<td bgcolor="#F5F5F5" width="45" height="35"><%x = x+1%>
							<input type="text" name="quant<%=i%>" id="quant<%=i%>" size="2" style="width: 35;" class="input" tabindex="<%=x%>" maxlength="2" onKeyPress="return SomenteNumero(event)" onKeyUp="Calc<%=i%>()" onFocus="Totalizar();"></td>
							<td bgcolor="#F5F5F5" width="90" height="35" colspan="2"><%x = x+1%>
							<input type="text" name="codigo<%=i%>" id="codigo<%=i%>" size="20" style="width: 80;" class="input" tabindex="<%=x%>" onKeyPress="return SomenteNumero(event);" onKeyDown="autocomp<%=i%>();" onKeyUp="Calc<%=i%>()" onFocus="Totalizar();" maxlength="10"></td>
							<td bgcolor="#F5F5F5" width="45" height="35"><%x = x+1%>
							<input type="text" name="tam<%=i%>" id="tam<%=i%>" size="20" style="width: 35;" class="input" tabindex="<%=x%>" onkeyup="alta(this);" onkeyup="Calc<%=i%>()" onfocus="Totalizar();" maxlength="3"></td>
							<td bgcolor="#F5F5F5" width="445" height="35"><%x = x+1%>
							<input type="text" name="descr<%=i%>" id="descr<%=i%>" tabindex="<%=x%>" size="20" style="width: 435;" class="input" onKeyPress="return SomenteLetras(event);" onKeyUp="alta(this);" maxlength="40"></td>
							<td bgcolor="#F5F5F5" width="51" height="35"><%x = x+1%>
							<input type="text" name="pag<%=i%>" id="pag<%=i%>" tabindex="<%=x%>" size="20" style="width: 40;" class="input" maxlength="3" onKeyPress="return SomenteNumero(event);"></td>
							<td bgcolor="#F5F5F5" width="90" height="35"><%x = x+1%>
							<input type="text" name="cor<%=i%>" id="cor<%=i%>" tabindex="<%=x%>" size="20" style="width: 80;" class="input" maxlength="7" onKeyUp="alta(this);" onKeyPress="return SomenteLetras(event);"></td>
							<td bgcolor="#F5F5F5" width="90" height="35"><%x = x+1%>
							<input dir="rtl" type="text" name="punit<%=i%>" id="punit<%=i%>" size="20" style="width: 80;" class="input" tabindex="<%=x%>" onKeyPress="return(MascaraMoeda(this,'.',',',event))" onKeyUp="Calc<%=i%>()" onFocus="Totalizar();"></td>
							<td bgcolor="#F5F5F5" width="80" height="35">
							<input dir="rtl" type="text" name="ptotal<%=i%>" id="ptotal<%=i%>" size="20" style="width: 80;" class="input" disabled="disabled"></td>
							<td align="right" width="27" bgcolor="#F5F5F5" height="35">&nbsp;</td>
						</tr>
						<%
						next
						%>
						<tr>
							<td align="left" width="27">
							<img border="0" src="Imagem/canto_cinza3.jpg" width="27" height="27"></td>
							<td bgcolor="#F5F5F5" colspan="9" width="934" align="right">&nbsp;
							</td>
							<td align="right" width="27" background="Imagem/canto_cinza4.jpg">&nbsp;
							</td>
						</tr>
					</table>
					</td>
				</tr>
				<tr>
					<td colspan="2" height="70" align="right">&nbsp;</td>
					<td height="70" align="right"><input type="image" src="Imagem/proximo.jpg"></td>
					<td height="70" align="right">&nbsp;</td>
				</tr>
				<tr>
					<td colspan="4">&nbsp;</td>
				</tr>
			</form>
			</table>
<%
case "central.pp2"
nome = request.form("nome")
email = request.form("email")
cpf = request.form("cpf")
cpf = mid(cpf,1,2)&"/"&mid(cpf,3,2)&"/"&mid(cpf,5,4)
ddd = request.form("ddd")
telefone = request.form("telefone")
revista = request.form("revista")
qualcat = request.form("qualcat")
select case revista
case "C1"
RR = "Hiroshima / Essencial / IVETE / L'Arc en Ciel"
case "C2"
RR = "Quatro Estações / Quintess / Bazar Brasil / Clarissa / Actual / Natubelly / Marcyn"
case "C3"
RR = "Biodany / Brasil Tropical / Sexydany"
case "C4"
RR = "Luzon"
case "C5"
RR = "Miro Star"
case "C6"
RR = "Abelha Rainha"
case "C7"
RR = "Brasil Tropical / Biodany / Sexydany"
case "C8"
RR = qualcat
case "C9"
RR = "Vitória Régia"
case "C10"
RR = "Expressão Feminina"
end select
MSN = "					<table border=0 width=988 cellspacing=0 cellpadding=0 id=table5 class=conteudo>"
MSN = MSN&"						<tr>"
MSN = MSN&"							<td align=left width=27 height=25>"
MSN = MSN&"							&nbsp;</td>"
MSN = MSN&"							<td width=935 height=25 colspan=8>"
MSN = MSN&"							&nbsp;</td>"
MSN = MSN&"							<td align=right width=27 height=25>"
MSN = MSN&"							&nbsp;</td>"
MSN = MSN&"						</tr>"
MSN = MSN&"						<tr>"
MSN = MSN&"							<td align=left width=27 height=25>&nbsp;</td>"
MSN = MSN&"							<td width=935 height=25 colspan=8>"
MSN = MSN&"								<table border=0 width=935 cellspacing=0 cellpadding=0 id=table9 class=conteudo>"
MSN = MSN&"								<tr>"
MSN = MSN&"									<td height=25 width=453>"
MSN = MSN&"									<font face=Arial><b>Nome</b></font></td>"
MSN = MSN&"									<td height=25><font face=Arial><b>Data de nascimento</b></font></td>"
MSN = MSN&"									<td height=25 width=243>"
MSN = MSN&"									<font face=Arial><b>Telefone / Celular</b></font></td>"
MSN = MSN&"								</tr>"
MSN = MSN&"								<tr>"
MSN = MSN&"									<td height=25 width=453>"&nome&"</td>"
MSN = MSN&"									<td height=25>"&cpf&"</td>"
MSN = MSN&"									<td height=25 width=243>("&ddd&") "&telefone&"</td>"
MSN = MSN&"								</tr>"
MSN = MSN&"								<tr><td height=35 colspan=3><b>E-mail do revendedor(a)</b></td></tr>"
MSN = MSN&"								<tr><td height=25 colspan=3>"&email&"</td></tr>	"
MSN = MSN&"								</table>"
MSN = MSN&"							</td>"
MSN = MSN&"							<td align=right width=27 height=25>&nbsp;</td>"
MSN = MSN&"						</tr>"
MSN = MSN&"						<tr>"
MSN = MSN&"							<td align=left width=27 height=25>&nbsp;</td>"
MSN = MSN&"							<td width=935 height=25 colspan=8>&nbsp;</td>"
MSN = MSN&"							<td align=right width=27 height=25>&nbsp;</td>"
MSN = MSN&"						</tr>"
MSN = MSN&"						<tr>"
MSN = MSN&"							<td align=left width=27 height=25>&nbsp;</td>"
MSN = MSN&"							<td width=935 height=25 colspan=8>"
MSN = MSN&"							<font face=Arial><b>Catálogos:</b> "&RR&"</font></td>"
MSN = MSN&"							<td align=right width=27 height=25>&nbsp;</td>"
MSN = MSN&"						</tr>"
MSN = MSN&"						<tr>"
MSN = MSN&"							<td align=left width=985 height=15 colspan=10>"
MSN = MSN&"							<span style=font-size: 1pt>&nbsp;</span></td>"
MSN = MSN&"						</tr>"
MSN = MSN&"						<tr>"
MSN = MSN&"							<td align=left width=27 height=25>&nbsp;</td>"
MSN = MSN&"							<td width=59 height=25>"
MSN = MSN&"							<font face=Arial><b>Quant.</b></font></td>"
MSN = MSN&"							<td width=91 height=25>"
MSN = MSN&"							<font face=Arial><b>Referência </b></font></td>"
MSN = MSN&"							<td width=48 height=25>"
MSN = MSN&"							<font face=Arial><b>Tam.</b></font></td>"
MSN = MSN&"							<td width=383 height=25>"
MSN = MSN&"							<font face=Arial><b>Descrição</b></font></td>"
MSN = MSN&"							<td width=64 height=25>"
MSN = MSN&"							<font face=Arial><b>Página</b></font></td>"
MSN = MSN&"							<td width=110 height=25>"
MSN = MSN&"							<font face=Arial><b>Cor</b></font></td>"
MSN = MSN&"							<td width=88 height=25>"
MSN = MSN&"							<font face=Arial><b>Preço Unit.</b></font></td>"
MSN = MSN&"							<td width=88 height=25>"
MSN = MSN&"							<font face=Arial><b>Preço Total</b></font></td>"
MSN = MSN&"							<td align=right width=27 height=25>&nbsp;</td>"
MSN = MSN&"						</tr>"
						total = 0
						x = 0
						for u = 1 to 50
						q = request.form("quant"&u&"")
						c = request.form("codigo"&u&"")
						p = request.form("punit"&u&"")
						if trim(q) <> "" and trim(c) <> "" and trim(p) <> "" then
						x = x+1
						end if
						next
						for i = 1 to x
						quant = request.form("quant"&i&"")
						codigo = request.form("codigo"&i&"")
						tam = request.form("tam"&i&"")
						descr = request.form("descr"&i&"")
						pag = request.form("pag"&i&"")
						cor = request.form("cor"&i&"")
						punit = request.form("punit"&i&"")
						ptotal = quant*punit
						total = total + ptotal
						if i MOD 2 = 0 Then
						Application("ccor") = "#F9F9F9" 
						else
						Application("ccor") = "#FFFFFF"
						end if
						if codigo <> "" then
MSN = MSN&"						<tr>"
MSN = MSN&"							<td align=left width=27 height=25>&nbsp;</td>"
MSN = MSN&"							<td bgcolor="&Application("ccor")&" width=59 height=25>"
MSN = MSN&"							<font face=Arial size=2 color=#000000>&nbsp;"&quant&"</font></td>"
MSN = MSN&"							<td bgcolor="&Application("ccor")&" width=91 height=25>"
MSN = MSN&"							<font face=Arial size=2 color=#000000>"&codigo&"</font></td>"
MSN = MSN&"							<td bgcolor="&Application("ccor")&" width=48 height=25>"
MSN = MSN&"							<font face=Arial size=2 color=#000000>"&tam&"</font></td>"
MSN = MSN&"							<td bgcolor="&Application("ccor")&" width=383 height=25>"
MSN = MSN&"							<font face=Arial size=2 color=#000000>"&descr&"</font></td>"
MSN = MSN&"							<td bgcolor="&Application("ccor")&" width=64 height=25>"
MSN = MSN&"							<font face=Arial size=2 color=#000000>"&pag&"</font></td>"
MSN = MSN&"							<td bgcolor="&Application("ccor")&" width=110 height=25>"
MSN = MSN&"							<font face=Arial size=2 color=#000000>"&cor&"</font></td>"
MSN = MSN&"							<td bgcolor="&Application("ccor")&" width=88 height=25>"
MSN = MSN&"							<font face=Arial size=2 color=#000000>R$ "&formatnumber(punit,2)&"</font></td>"
MSN = MSN&"							<td bgcolor="&Application("ccor")&" width=88 height=25>"
MSN = MSN&"							<font face=Arial size=2 color=#000000>R$ "&formatnumber(ptotal,2)&"</font></td>"
MSN = MSN&"							<td align=right width=27 height=25>&nbsp;</td>"
MSN = MSN&"						</tr>"
						end if
						next
MSN = MSN&"						<tr>"
MSN = MSN&"							<td align=left width=27 height=40 valign=bottom>&nbsp;</td>"
MSN = MSN&"							<td width=935 height=40 colspan=8 valign=bottom>"
MSN = MSN&"							<table border=0 width=100% cellspacing=0 cellpadding=0 id=table10>"
MSN = MSN&"								<tr>"
MSN = MSN&"									<td width=820 align=right><font face=Arial><b>Valor total do pedido:&nbsp;&nbsp;</b></font></td>"
MSN = MSN&"									<td width=115 align=left>"
MSN = MSN&"									<font size=4 color=#000000 face=Arial>R$ "&formatnumber(total,2)&"</font></td>"
MSN = MSN&"								</tr>"
MSN = MSN&"							</table>"
MSN = MSN&"							</td>"
MSN = MSN&"							<td align=right width=27 height=40 valign=bottom>&nbsp;</td>"
MSN = MSN&"						</tr>"
MSN = MSN&"						<tr>"
MSN = MSN&"							<td align=left width=27>"
MSN = MSN&"							&nbsp;</td>"
MSN = MSN&"							<td colspan=8 width=935 align=right>"
MSN = MSN&"							&nbsp;</td>"
MSN = MSN&"							<td align=right width=27>"
MSN = MSN&"							&nbsp;</td>"
MSN = MSN&"						</tr>"
Session("MSN") = MSN&"					</table>"
Response.AddHeader "Refresh","599 ; URL='javascript:history.back(-1)'"
%>
					<table border="0" width="988" cellspacing="0" cellpadding="0" id="table5" class="conteudo">
						<tr>
							<td align="left" width="27" bgcolor="#F5F5F5" height="25">
							<img border="0" src="http://www.centraldoscatalogos.com.br/site/imagem/canto_cinza1.jpg" width="27" height="27"></td>
							<td bgcolor="#F5F5F5" width="935" height="25" colspan="8">&nbsp;
							</td>
							<td align="right" width="27" bgcolor="#F5F5F5" height="25">
							<img border="0" src="http://www.centraldoscatalogos.com.br/site/imagem/canto_cinza2.jpg" width="27" height="27"></td>
						</tr>
						<tr>
							<td align="left" width="27" bgcolor="#F5F5F5" height="25">&nbsp;</td>
							<td bgcolor="#F5F5F5" width="935" height="25" colspan="9" <%if revista <> "C8" then%>align="center"<%else%>align="left"<%end if%>>
							<%if revista <> "C8" then%><img src="imagem/<%=revista%>.jpg"><%else%><font color="#000" size="14pt"><%=qualcat%></font><%end if%></td>
							<td align="right" width="27" bgcolor="#F5F5F5" height="25">&nbsp;</td>
						</tr>
						<tr>
							<td align="left" width="27" bgcolor="#F5F5F5" height="25">&nbsp;</td>
							<td bgcolor="#F5F5F5" width="935" height="25" colspan="8">
							<table border="0" width="100%" cellspacing="0" cellpadding="0" id="table9" class="conteudo">
								<tr>
									<td height="25" width="453">&nbsp;
									</td>
									<td height="25">&nbsp;</td>
									<td height="25" width="243">&nbsp;
									</td>
								</tr>
								<tr>
									<td height="25" colspan="3">
									<b><font size="4">Confira os dados e clique 
									no botão &quot;Enviar&quot; abaixo do formúlário para 
									enviar seu pedido.</font></b></td>
								</tr>
								<tr>
									<td height="15" width="453">
									</td>
									<td height="15"></td>
									<td height="15" width="243">
									</td>
								</tr>
								<tr>
									<td height="35" width="453">
									<font face="Arial"><b>Nome do(a) 
									revendedor(a)</b></font></td>
									<td height="35"><font face="Arial"><b>CPF</b></font></td>
									<td height="35" width="243">
									<font face="Arial"><b>DDD e Telefone</b></font></td>
								</tr>
								<tr>
									<td height="25" width="453">
									<font color="#000000"><%=nome%></font></td>
									<td height="25"><font color="#000000"><%=cpf%></font></td>
									<td height="25" width="243">
									<font color="#000000">(<%=ddd%>)&nbsp;<%=telefone%></font></td>
								</tr>
								<tr><td height="35" colspan="3"><b>E-mail do revendedor(a)</b></td></tr>
								<tr><td height="25" colspan="3"><font color="#000000"><%=email%></font></td></tr>								
								</table>
							</td>
							<td align="right" width="27" bgcolor="#F5F5F5" height="25">&nbsp;</td>
						</tr>
						<tr>
							<td align="left" width="989" bgcolor="#F5F5F5" height="10" colspan="10">
							<span style="font-size: 1pt">&nbsp;</span></td>
						</tr>
						<tr>
							<td align="left" width="27" bgcolor="#F5F5F5" height="25">&nbsp;</td>
							<td bgcolor="#F5F5F5" width="59" height="25">
							<font face="Arial"><b>Quant.</b></font></td>
							<td bgcolor="#F5F5F5" width="91" height="25">
							<font face="Arial"><b>Referência </b></font></td>
							<td bgcolor="#F5F5F5" width="48" height="25">
							<font face="Arial"><b>Tam.</b></font></td>
							<td bgcolor="#F5F5F5" width="383" height="25">
							<font face="Arial"><b>Descrição</b></font></td>
							<td bgcolor="#F5F5F5" width="64" height="25">
							<font face="Arial"><b>Página</b></font></td>
							<td bgcolor="#F5F5F5" width="110" height="25">
							<font face="Arial"><b>Cor</b></font></td>
							<td bgcolor="#F5F5F5" width="88" height="25">
							<font face="Arial"><b>Preço Unit.</b></font></td>
							<td bgcolor="#F5F5F5" width="88" height="25">
							<font face="Arial"><b>Preço Total</b></font></td>
							<td align="right" width="27" bgcolor="#F5F5F5" height="25">&nbsp;</td>
						</tr>
						<%
						total = 0
						x = 0
						for u = 1 to 50
						q = request.form("quant"&u&"")
						c = request.form("codigo"&u&"")
						p = request.form("punit"&u&"")
						if trim(q) <> "" and trim(c) <> "" and trim(p) <> "" then
						x = x+1
						end if
						next
						for i = 1 to x
						quant = request.form("quant"&i&"")
						codigo = request.form("codigo"&i&"")
						tam = request.form("tam"&i&"")
						descr = request.form("descr"&i&"")
						pag = request.form("pag"&i&"")
						cor = request.form("cor"&i&"")
						punit = request.form("punit"&i&"")
						ptotal = quant*punit
						total = total + ptotal
						if i MOD 2 = 0 Then
						Application("ccor") = "#F9F9F9" 
						else
						Application("ccor") = "#FFFFFF"
						end if
						if codigo <> "" then
						%>
						<tr>
							<td align="left" width="27" bgcolor="#F5F5F5" height="25">&nbsp;</td>
							<td bgcolor=<%=Application("ccor")%> width="59" height="25">
							<font face="Arial" size="2" color="#000000">&nbsp;<%=quant%></font></td>
							<td bgcolor=<%=Application("ccor")%> width="91" height="25">
							<font face="Arial" size="2" color="#000000"><%=codigo%></font></td>
							<td bgcolor=<%=Application("ccor")%> width="48" height="25">
							<font face="Arial" size="2" color="#000000"><%=tam%></font></td>
							<td bgcolor=<%=Application("ccor")%> width="383" height="25">
							<font face="Arial" size="2" color="#000000"><%=descr%></font></td>
							<td bgcolor=<%=Application("ccor")%> width="64" height="25">
							<font face="Arial" size="2" color="#000000"><%=pag%></font></td>
							<td bgcolor=<%=Application("ccor")%> width="110" height="25">
							<font face="Arial" size="2" color="#000000"><%=cor%></font></td>
							<td bgcolor=<%=Application("ccor")%> width="88" height="25">
							<font face="Arial" size="2" color="#000000">R$ <%=formatnumber(punit,2)%></font></td>
							<td bgcolor=<%=Application("ccor")%> width="88" height="25">
							<font face="Arial" size="2" color="#000000">R$ <%=formatnumber(ptotal,2)%></font></td>
							<td align="right" width="27" bgcolor="#F5F5F5" height="25">&nbsp;</td>
						</tr>
						<%
						end if
						next
						%>
						<tr>
							<td align="left" width="27" bgcolor="#F5F5F5" height="40" valign="bottom">&nbsp;</td>
							<td bgcolor="#F5F5F5" width="935" height="40" colspan="8" valign="bottom">
							<table border="0" width="100%" cellspacing="0" cellpadding="0" id="table10" class="conteudo">
								<tr>
									<td width="820" align="right"><b>Valor total do pedido:</b></td>
									<td width="115" align="right">
									<font size="4" color="#000000">R$ <%=formatnumber(total,2)%>&nbsp;</font></td>
								</tr>
							</table>
							</td>
							<td align="right" width="27" bgcolor="#F5F5F5" height="40" valign="bottom">&nbsp;</td>
						</tr>
						<tr>
							<td align="left" width="27">
							<img border="0" src="http://www.centraldoscatalogos.com.br/site/imagem/canto_cinza3.jpg" width="27" height="27"></td>
							<td bgcolor="#F5F5F5" colspan="8" width="934" align="right">&nbsp;
							</td>
							<td align="right" width="27" background="http://www.centraldoscatalogos.com.br/site/imagem/canto_cinza4.jpg">&nbsp;
							</td>
						</tr>
						<form action="index.asp" method="POST">
						<input type="hidden" name="action" id="action" value="gomail">
						<input type="hidden" name="nome" id="nome" value="<%=nome%>">
						<input type="hidden" name="email" id="email" value="<%=email%>">
						<input type="hidden" name="cpf" id="cpf" value="<%=cpf%>">
						<input type="hidden" name="ddd" id="ddd" value="<%=ddd%>">
						<input type="hidden" name="telefone" id="telefone" value="<%=telefone%>">
						<input type="hidden" name="revista" id="revista" value="<%=RR%>">
						<tr>
							<td  height="70" align="right">&nbsp;</td>
							<td height="70" colspan="8" align="right">
							<input type="image" src="Imagem/botao_enviar.jpg" name="x.input" width="133" height="43"></td>
							<td height="70" align="right">&nbsp;</td>
						</tr>
						</form>
				</table>
<%
case "central.pp.ok"
%>			
<table border="0" width="988" cellspacing="0" cellpadding="0" id="table1" height="450">
	<tr>
		<td align="center">
		<%if Application("status") = "ok" then%>
		<b><font face="Arial" color="#009900">Pedido enviado com sucesso!!!</font></b>
		<%else%>
		<b><font face="Arial" color="#CC0000">Houve um erro no envio de seu pedido, favor enviar novamente.</font></b>
		<%end if%>
		</td>
	</tr>
</table>

<%
case "central.cr1"
%>
				<table border="0" width="988" cellspacing="0" cellpadding="0" id="table3">
				<%if Application("erroconfirm") > 0 then%>
					<tr>
						<td align="center" colspan="2" bgcolor="#FFFF99">
						<table border="0" width="94%" cellspacing="0" cellpadding="0" id="table11">
							<tr>
								<td colspan="5"><br>
								<b><font size="4">Erro!!!<br>
								</font><font style="font-size: 3pt">&nbsp;</font><font size="4"><br>
								</font><span style="font-size: 9pt">
								<%=Session("erro_nomec")%>
								<%=Session("erro_emailc")%>
								<%=Session("erro_dddc")%>
								<%=Session("erro_telc")%>
								<%=Session("erro_enderecoc")%>
								<%=Session("erro_cidadec")%>
								<%=Session("erro_estadoc")%><br>
								</span></b></td>
							</tr>
						</table>			        
				        </td>
					</tr>
					<%end if%>
					<tr>
					<form action="index.asp" method="POST">
					<input type="hidden" name="action" id="action" value="crevendedora">
						<td width="48%" align="center">
					<table border="0" width="420" cellspacing="0" cellpadding="0" id="table3" class="conteudo">
					<tr>
						<td height="50" colspan="3"><b><font size="4">Formulário 
						de Solicitação de Cadastro</font></b>
						</td>
					</tr>
					<tr>
						<td width="84" height="30"><font face="Arial"><b>Nome:</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						</font></td>
						<td width="336" colspan="2" height="30">
						<input type="text" name="nome" class="form input" size="40" tabindex="1" value="<%=Application("nome")%>"></td>
					</tr>
					<tr>
						<td width="84" height="30"><b><font face="Arial">E-mail:</font></b></td>
						<td width="336" colspan="2" height="30">
						<input type="text" name="email" class="form input" size="40" tabindex="2" value="<%=Application("mail")%>"></td>
					</tr>
					<tr>
						<td width="84" height="30"><b><font face="Arial">Telefone:</font></b></td>
						<td width="30" height="30">
						<input type="text" name="ddd" size="2" style="width: 30" class="input" maxlength="2" tabindex="3" value="<%=Application("ddd")%>"></td>
						<td width="306" height="30" align="right">
						&nbsp;<input type="text" name="telefone" style="width: 300"  class="input" size="8" tabindex="4" maxlength="8" value="<%=Application("fone")%>"></td>
					</tr>
					<tr>
						<td width="84" height="30"><b><font face="Arial">
						Endereço:</font></b></td>
						<td width="336" colspan="2" height="30">
						<input type="text" name="endereco" class="form input" size="40" tabindex="5" value="<%=Application("endereco")%>"></td>
					</tr>
					<tr>
						<td width="84" height="30"><b><font face="Arial">Cidade:</font></b></td>
						<td width="336" colspan="2" height="30">
						<input type="text" name="cidade" class="form input" size="40" tabindex="6" value="<%=Application("cidade")%>"></td>
					</tr>
					<tr>
						<td width="84" height="30"><b><font face="Arial">Estado:</font></b></td>
						<td width="336" colspan="2" height="30">
						<select size="1" name="estado" class="form input" tabindex="7">
						<option value="" <%if Application("estado") = "" then%>selected<%end if%>>Selecione o Estado</option>
						<option value="ac" <%if Application("estado") = "ac" then%>selected<%end if%>>Acre</option>
						<option value="al" <%if Application("estado") = "al" then%>selected<%end if%>>Alagoas</option>
						<option value="ap" <%if Application("estado") = "ap" then%>selected<%end if%>>Amapá</option>
						<option value="am" <%if Application("estado") = "am" then%>selected<%end if%>>Amazonas</option>
						<option value="ba" <%if Application("estado") = "ba" then%>selected<%end if%>>Bahia</option>
						<option value="ce" <%if Application("estado") = "ce" then%>selected<%end if%>>Ceará</option>
						<option value="df" <%if Application("estado") = "df" then%>selected<%end if%>>Distrito Federal</option>
						<option value="es" <%if Application("estado") = "es" then%>selected<%end if%>>Espirito Santo</option>
						<option value="go" <%if Application("estado") = "go" then%>selected<%end if%>>Goiás</option>
						<option value="ma" <%if Application("estado") = "ma" then%>selected<%end if%>>Maranhão</option>
						<option value="ms" <%if Application("estado") = "ms" then%>selected<%end if%>>Mato Grosso do Sul</option>
						<option value="mt" <%if Application("estado") = "mt" then%>selected<%end if%>>Mato Grosso</option>
						<option value="mg" <%if Application("estado") = "mg" then%>selected<%end if%>>Minas Gerais</option>
						<option value="pa" <%if Application("estado") = "pa" then%>selected<%end if%>>Pará</option>
						<option value="pb" <%if Application("estado") = "pb" then%>selected<%end if%>>Paraíba</option>
						<option value="pr" <%if Application("estado") = "pr" then%>selected<%end if%>>Paraná</option>
						<option value="pe" <%if Application("estado") = "pe" then%>selected<%end if%>>Pernambuco</option>
						<option value="pi" <%if Application("estado") = "pi" then%>selected<%end if%>>Piauí</option>
						<option value="rj" <%if Application("estado") = "rj" then%>selected<%end if%>>Rio de Janeiro</option>
						<option value="rn" <%if Application("estado") = "rn" then%>selected<%end if%>>Rio Grande do Norte</option>
						<option value="rs" <%if Application("estado") = "rs" then%>selected<%end if%>>Rio Grande do Sul</option>
						<option value="ro" <%if Application("estado") = "ro" then%>selected<%end if%>>Rondônia</option>
						<option value="rr" <%if Application("estado") = "rr" then%>selected<%end if%>>Roraima</option>
						<option value="sc" <%if Application("estado") = "sc" then%>selected<%end if%>>Santa Catarina</option>
						<option value="sp" <%if Application("estado") = "sp" then%>selected<%end if%>>São Paulo</option>
						<option value="se" <%if Application("estado") = "se" then%>selected<%end if%>>Sergipe</option>
						<option value="to" <%if Application("estado") = "to" then%>selected<%end if%>>Tocantins</option>
						</select></td>
					</tr>
										<tr>
						<td colspan="3" height="60" valign="bottom" align="right">
						<input type="image" src="Imagem/botao_enviar.jpg"></td>
					</form>
					</tr>
				</table></td>
						<td>
						<p align="center"><font size="6">Atenção!!!</font><br>
						<br>
						<font size="4">Assim que possível entraremos em contato
						<br>
						para 
						agendar uma visita
						para realizar o cadastro oficial<br>
						<br>
&nbsp;</font></td>
					</tr>
				</table>
			
<%
Application("erroconfirm") = 0
Session("erro_nomec") = ""
Session("erro_emailc") = ""
Session("erro_dddc") = ""
Session("erro_telc") = ""
Session("erro_enderecoc") = ""
Session("erro_cidadec") = ""
Session("erro_estadoc") = ""
case "central.fc"
%>
				<table border="0" width="988" cellspacing="0" cellpadding="0" id="table3">
				<%if Application("erroconfirm") > 0 then%>
					<tr>
						<td align="center" colspan="2" bgcolor="#FFFF99">
						<table border="0" width="94%" cellspacing="0" cellpadding="0" id="table11">
							<tr>
								<td colspan="5"><br>
								<b><font size="4">Erro!!!<br>
								</font><font style="font-size: 3pt">&nbsp;</font><font size="4"><br>
								</font><span style="font-size: 9pt">
								<%=Session("erro_nomec")%>
								<%=Session("erro_emailc")%>
								<%=Session("erro_dddc")%>
								<%=Session("erro_telc")%>
								<%=Session("erro_assuntoc")%>
								<%=Session("erro_mensagemc")%><br>
								</span></b></td>
							</tr>
						</table>			        
				        </td>
					</tr>
					<%end if%>
					<tr>
					<form action="index.asp" method="POST">
					<input type="hidden" name="action" id="action" value="contato">
						<td width="48%" align="center">
					<table border="0" width="420" cellspacing="0" cellpadding="0" id="table3" class="conteudo">
					<tr>
						<td height="50" colspan="3"><b><font size="4">Fale Conosco</font></b>
						</td>
					</tr>
					<tr>
						<td width="84" height="30"><font face="Arial"><b>Nome:</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						</font></td>
						<td width="336" colspan="2" height="30">
						<input type="text" name="nome" class="form input" size="40" tabindex="1" value="<%=Application("nome")%>"></td>
					</tr>
					<tr>
						<td width="84" height="30"><b><font face="Arial">E-mail:</font></b></td>
						<td width="336" colspan="2" height="30">
						<input type="text" name="email" class="form input" size="40" tabindex="2" value="<%=Application("mail")%>"></td>
					</tr>
					<tr>
						<td width="84" height="30"><b><font face="Arial">Telefone:</font></b></td>
						<td width="30" height="30">
						<input type="text" name="ddd" size="2" style="width: 30" class="input" maxlength="2" tabindex="3" value="<%=Application("ddd")%>"></td>
						<td width="306" height="30" align="right">
						&nbsp;<input type="text" name="telefone" style="width: 300"  class="input" size="8" tabindex="4" maxlength="8" value="<%=Application("fone")%>"></td>
					</tr>
					<tr>
						<td width="84" height="30"><b><font face="Arial">Assunto:</font></b></td>
						<td width="336" colspan="2" height="30">
						<input type="text" name="assunto" class="form input" size="40" tabindex="5" value="<%=Application("assu")%>"></td>
					</tr>
					<tr>
						<td width="400" colspan="3" height="30"><b><font face="Arial">Mensagem</font>:</b></td>
					</tr>

					<tr>
						<td colspan="3">
						<textarea rows="2" name="mensagem" cols="20" class="area" tabindex="6"><%=Application("mens")%></textarea></td>
					</tr>
										<tr>
						<td colspan="3" height="60" valign="bottom" align="right">
						<input type="image" src="Imagem/botao_enviar.jpg"></td>
					</form>
					</tr>
				</table></td>
						<td>
						<p align="center"><font size="6">Tem 
						alguma dúvida ou sugestão?</font></p>
						<p align="center"><font size="4">Fale 
						conosco através dos telefones<br>
&nbsp;(22) 2724-2857 / (22) 99800-9061<br>
&nbsp;ou através do formulário ao lado
						e <br>
						breve entraremos em contato.</font></td>
					</tr>
				</table>
<%
Application("erroconfirm") = 0
Session("erro_nomec") = ""
Session("erro_emailc") = ""
Session("erro_dddc") = ""
Session("erro_telc") = ""
Session("erro_mensagemc") = ""
End Select
%>
			</td>
		</tr>
		<%if Application("pag") <> "central.qs" then%>
		<tr>
			<td height="20" align="center" valign="top" width="50%">&nbsp;
			</td>
		</tr>
		<%end if%>
		<tr>
			<td background="imagem/barra_down.jpg" height="113" align="center" valign="top" width="50%">
			<table border="0" width="988" cellspacing="0" cellpadding="0" id="table5" class="rodape">
				<tr>
					<td width="50%">&nbsp;</td>
					<td>&nbsp;</td>
				</tr>
				<tr>
					<td height="80" colspan="2" align="center">Rua: Saldanha 
					Marinho, 457 - loja 5 - Centro<br>
					Próximo da Rodoviária – Nas proximidades do Hospital Pronto 
					Cárdio e Padaria Pavê.<br>
					E-mail: atendimento@centraldoscatalogos.com.br<br>
					Contato: (22) 2724-2857 / (22) 99800-9061
                    <div style="position:absolute;width:988px;text-align:right;">
                    <%
					Set bcounter = conexao.execute("SELECT * FROM counter WHERE id = 1")
					If not bcounter.eof then
						Application("counternum") = bcounter("contador")
						if Session("contadores") <> "ContPlus" then
							Application("up") = Application("counternum") + 1
							Set upcounter = conexao.execute("update counter set contador="&Application("up")&" where id=1")
							response.write Application("up")
							Session("contadores") = "ContPlus"
						else
							response.write Application("counternum")
						end if
					end if
					%>
                    </div></td>
				</tr>
			</table>
			</td>
		</tr>
		</table>
</div>
<%
End If
conexao.close
Set conexao = Nothing
%>
</body>
</html>