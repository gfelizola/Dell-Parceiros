<%
'Response.ContentType = "application/vnd.ms-excel;"
'Response.AddHeader "content-disposition" , "attachment;filename=planilha_dell_parceiros_" & year(now) & "_" & month(now) & "_" & day(now) & ".xls"
'Response.Charset = "utf-8"

%>
<!-- #INCLUDE file="conexao.asp" -->
<%

SQL = 	"SELECT * FROM Cadastros ORDER BY DataCadastro DESC"
Set RS = Conexao.execute(SQL,3)

Dim u_title
u_title = "CADASTRO"
Response.Write "<html xmlns:x=""urn:schemas-microsoft-com:office:excel"">"
Response.Write "<head>"
Response.Write "<!--[if gte mso 9]><xml>"
Response.Write "<x:ExcelWorkbook>"
Response.Write "<x:ExcelWorksheets>"
Response.Write "<x:ExcelWorksheet>"
Response.Write "<x:Name>"& u_title &"</x:Name>"
Response.Write "<x:WorksheetOptions>"
Response.Write "<x:Print>"
Response.Write "<x:ValidPrinterInfo/>"
Response.Write "</x:Print>"
Response.Write "</x:WorksheetOptions>"
Response.Write "</x:ExcelWorksheet>"
Response.Write "</x:ExcelWorksheets>"
Response.Write "</x:ExcelWorkbook>"
Response.Write "</xml>"
Response.Write "<![endif]--> "
Response.Write "</head>"
Response.Write "<body>"
Response.Write "<!--- <table boder='1'>"
%>
    <TR>
		<% For i = 0 to RS.fields.count - 1 %>
        <TD><strong><%= Replace( ucase(RS(i).name), "'", "") %></strong></TD>
        <% next %>
		<TD><strong>FILIAIS ATENDIDAS</strong></TD>
    </TR>
    <% Do Until RS.eof %>
        <TR>
			<% 
			For i = 0 to RS.fields.count - 1
				If RS(i).name = "Empresa" Then
					SQL = "SELECT * FROM Empresas WHERE Codigo = " & RS(i)
					Set RSE = Conexao.execute(SQL,3)
					
					Response.Write("<TD VALIGN=TOP>" & RSE("nome") & "</TD>")
					
					RSE.Close
					Set RSE = Nothing
				Else
				%>
                <TD VALIGN=TOP><%= RS(i)%></TD>
            	<% 
				End If
			Next 
			
			StrFiliais = ""
			SQL = "SELECT * FROM Cadastro_Estados WHERE Cadastro = " & RS("Codigo") & " ORDER BY Estado"
			Set RSF = Conexao.execute(SQL,3)
			
			Do Until RSF.Eof
				StrFiliais = StrFiliais & RSF("Estado") & ": " & RSF("Qtde") & ", "
				RSF.MoveNext
			Loop
			If Len( StrFiliais ) > 0 Then StrFiliais = Left( StrFiliais, Len( StrFiliais ) - 2 )
			
			RSF.Close
			Set RSF = Nothing 
			
			Response.Write("<TD VALIGN=TOP>" & StrFiliais & "</TD>")
			%>
        </TR>
    <%
    RS.MoveNext
    
    Loop
    
    RS.Close
    Conexao.close
    %>
<%
Response.Write "</table --->"
Response.Write "</body>"
Response.Write "</html>"
%>