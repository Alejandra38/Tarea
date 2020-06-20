<html>
<body>
<%@ LANGUAGE="VBSCRIPT" %>
<%
option explicit
Dim conn,fol,date,rfc,codp,cpcs

fol=Request.Form("Folio")
date=request.Form("Fecha_C")
rfc=Request.Form("RFC")
codp=request.Form("Codigo_producto")
cpcs=Request.Form("Cantidad_producto_comprado")

set conn=Server.CreateObject("ADODB.connection")
Conn.open "provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\osori\OneDrive\Escritorio\Carrillo\Personal.mdb"
Conn.execute "INSERT INTO PEDIDOS(Folio,Fecha_C,RFC,Codigo_producto,Cantidad_producto_comprado) values('"& fol & "','"& date & "','"& rfc & "','"& codp & "','"& cpcs & "')"
Conn.close
set conn=nothing



%>
</body>
</html>