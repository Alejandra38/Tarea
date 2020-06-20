<html>
<body>
<%@ LANGUAGE="VBSCRIPT" %>
<%
option explicit
Dim conn,rfc,nom,dir,tel,Ciu,codpos,notar

rfc=Request.Form("RFC")
nom=Request.Form("Nombre")
dir=Request.Form("Direccion")
tel=Request.Form("Telefono")
Ciu=Request.Form("Ciudad")
codpos=Request.Form("[Codigo postal]")
notar=Request.Form("No_Tarjeta")

set conn=Server.CreateObject("ADODB.connection")
Conn.open "provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\osori\OneDrive\Escritorio\Carrillo\Personal.mdb"
Conn.execute "INSERT INTO CLIENTES(RFC,Nombre,Direccion,Telefono,Ciudad,[Codigo postal],No_Tarjeta) values('"& rfc & "','"& nom & "','"& dir & "','"& tel & "','"& Ciu & "','"& codpos &"','"& notar & "')"
Conn.close
set conn=nothing

%>
</body>  
</html>