<%@ LANGUAGE = VBScript%>

<%
Response.Buffer=false
	Set G_DBConn = Server.CreateObject("ADODB.Connection")

			G_DBConn.Open Application("ConnectionString")

Response.ContentType="image/*"
SelStr=Request("SelStr")
TabStr=Request("TabStr")
FldStr=Request("FldStr")
ValStr=Request("ValStr")
SqlStr="SELECT "&SelStr&" FROM "&TabStr&" WHERE "&FldStr&"='"&ValStr&"'"
if SelStr<>"" then
  Set rs = G_DBConn.Execute(SqlStr)
  if not RS.EOF then
	PhotoSize = rs(0).ActualSize
	if PhotoSize <> 0 then
		do while Wrt < PhotoSize
			Photo = rs(0).GetChunk(4194304)
			Response.BinaryWrite Photo
			Response.Flush
			Wrt = Wrt + 4194304
		loop
	end if
  end if
  rs.Close
  G_DBConn.Close
  Response.end
  g_DBConn.Close
end if
%>