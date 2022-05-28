<%
Set conexao = Server.CreateObject("ADODB.Connection")
conexao.open="Driver={MySQL ODBC 5.1 Driver};server=186.202.13.31;database=centraldoscata;uid=centraldoscata;pwd=art@central"
Application("mes") = 0
f = request.form("revista")
for x = 0 to 2
if Application("mes") = 0 then
Application("mes") = month(Now)
else
Application("mes") = Application("mes")+1
end if
for u = 1 to 5
sidcat = f&u&x
scodcat = request.form(""&f&u&x&"cod")
sdatavai = request.form(""&f&u&x&"vai")
sdatavem = request.form(""&f&u&x&"vem")
smes = Application("mes") 
Set up = conexao.execute("update datas set codcat='"&scodcat&"', mes='"&smes&"', datavai='"&sdatavai&"', datavem='"&sdatavem&"' where idcat='"&sidcat&"'")
next
next
conexao.close
set conexao = nothing
response.redirect("update_datas.asp?save=ok")
%>