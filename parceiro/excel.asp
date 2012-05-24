<%
Response.ContentType = "application/vnd.ms-excel;"
Response.AddHeader "content-disposition" , "attachment;filename=planilha_dell_parceiros_" & year(now) & "_" & month(now) & "_" & day(now) & ".xls"
Response.Charset = "utf-8"

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
%>
<table border="1" bordercolor="#FFFFFF">

	<TR>
		<TD bgcolor="#538ed5">&nbsp;</TD>
		<TD bgcolor="#538ed5">&nbsp;</TD>
		<TD bgcolor="#538ed5">&nbsp;</TD>
		<TD bgcolor="#538ed5">&nbsp;</TD>
		<TD bgcolor="#538ed5">&nbsp;</TD>
		<TD bgcolor="#538ed5">&nbsp;</TD>
		<TD bgcolor="#538ed5">&nbsp;</TD>
		<TD bgcolor="#538ed5">&nbsp;</TD>
		<TD bgcolor="#538ed5">&nbsp;</TD>
		<TD bgcolor="#538ed5">&nbsp;</TD>
		<TD bgcolor="#538ed5">&nbsp;</TD>
		<TD bgcolor="#538ed5" align="center" colspan="<%=Ubound(Estados)+1%>" style="border-left-width:4px;"><font color="#FFFFFF"><strong>Qtde de filiais por Estado</strong></font></TD>
		<TD bgcolor="#538ed5" align="center" colspan="<%=Ubound(Estados)+1%>" style="border-left-width:4px;"><font color="#FFFFFF"><strong>Estados de Atendimento</strong></font></TD>
		<TD bgcolor="#538ed5" align="center" colspan="7"><font color="#FFFFFF"><strong>Setor Foco</strong></font></TD>
		<TD bgcolor="#538ed5">&nbsp;</TD>
		<TD bgcolor="#538ed5">&nbsp;</TD>
		<TD bgcolor="#538ed5">&nbsp;</TD>
		<TD bgcolor="#538ed5">&nbsp;</TD>
		<TD bgcolor="#538ed5" align="center" colspan="4"><font color="#FFFFFF"><strong>Faturamento de 2011</strong></font></TD>
		<TD bgcolor="#538ed5" align="center" colspan="7"><font color="#FFFFFF"><strong>Nível de Capacitação da Equipe comercial</strong></font></TD>
		<TD bgcolor="#538ed5">&nbsp;</TD>
		<TD bgcolor="#538ed5">&nbsp;</TD>
		<TD bgcolor="#538ed5">&nbsp;</TD>
		<TD bgcolor="#538ed5">&nbsp;</TD>
		<TD bgcolor="#538ed5" align="center" colspan="<%=Ubound(Marcas)+1%>"><font color="#FFFFFF"><strong>Marcas Comercializadas </strong></font></TD>
		<TD bgcolor="#538ed5">&nbsp;</TD>
		<TD bgcolor="#538ed5">&nbsp;</TD>
    </TR>
    <TR>
		<TD bgcolor='#376091'><font color='#ffffff'><strong>Prioridade</strong></font></TD>
		<TD bgcolor='#376091'><font color='#ffffff'><strong>Id</strong></font></TD>
		<TD bgcolor='#376091'><font color='#ffffff'><strong>Nome do parceiro</strong></font></TD>
		<TD bgcolor='#376091'><font color='#ffffff'><strong>Estado</strong></font></TD>
		<TD bgcolor='#376091'><font color='#ffffff'><strong>Telefone</strong></font></TD>
		<TD bgcolor='#376091'><font color='#ffffff'><strong>Contato</strong></font></TD>
		<TD bgcolor='#376091'><font color='#ffffff'><strong>E-mail</strong></font></TD>
		<TD bgcolor='#376091'><font color='#ffffff'><strong>Nível de Certificação</strong></font></TD>
		<TD bgcolor='#376091'><font color='#ffffff'><strong>Nº de Funcionários</strong></font></TD>
		<TD bgcolor='#376091'><font color='#ffffff'><strong>Estado da Matriz</strong></font></TD>
		<TD bgcolor='#376091'><font color='#ffffff'><strong>Quantas filiais</strong></font></TD>
		
		<%
		UFA = 1
		For Each UF In Estados
			borda = ""
			If UFA = 1 Then borda = "style='border-left-width:4px;'"
			Response.Write("<TD bgcolor='#376091' " & borda & "><font color='#ffffff'><strong>" & UF & "</strong></font></TD>")
			UFA = UFA + 1
		Next
		UFA = 1
		For Each UF In Estados
			borda = ""
			If UFA = 1 Then borda = "style='border-left-width:4px;'"
			Response.Write("<TD bgcolor='#376091' " & borda & "><font color='#ffffff'><strong>" & UF & "</strong></font></TD>")
			UFA = UFA + 1
		Next
		%>
		
		<TD bgcolor='#376091'><font color='#ffffff'><strong>Governo</strong></font></TD>
		<TD bgcolor='#376091'><font color='#ffffff'><strong>Educação</strong></font></TD>
		<TD bgcolor='#376091'><font color='#ffffff'><strong>Manufatura</strong></font></TD>
		<TD bgcolor='#376091'><font color='#ffffff'><strong>Finanças</strong></font></TD>
		<TD bgcolor='#376091'><font color='#ffffff'><strong>Oil & gás</strong></font></TD>
		<TD bgcolor='#376091'><font color='#ffffff'><strong>Bem de Consumo</strong></font></TD>
		<TD bgcolor='#376091'><font color='#ffffff'><strong>Outros</strong></font></TD>
		
		<TD bgcolor='#376091'><font color='#ffffff'><strong>Possui setor de vendas?</strong></font></TD>
		<TD bgcolor='#376091'><font color='#ffffff'><strong>Possui setor de marketing?</strong></font></TD>
		<TD bgcolor='#376091'><font color='#ffffff'><strong>Possui setor de financiamento?</strong></font></TD>
		<TD bgcolor='#376091'><font color='#ffffff'><strong>Realiza treinamento aos funcionários?</strong></font></TD>
		
		<TD bgcolor='#376091'><font color='#ffffff'><strong>De R$ 100.000,00 a R$ 500.000,00</strong></font></TD>
		<TD bgcolor='#376091'><font color='#ffffff'><strong>De 500.000,00 a 1 milhão</strong></font></TD>
		<TD bgcolor='#376091'><font color='#ffffff'><strong>De R$ 1 milhão a R$ 5 milhões </strong></font></TD>
		<TD bgcolor='#376091'><font color='#ffffff'><strong>Mais de R$ 5 milhões</strong></font></TD>
		
		<TD bgcolor='#376091'><font color='#ffffff'><strong>Client</strong></font></TD>
		<TD bgcolor='#376091'><font color='#ffffff'><strong>Storage</strong></font></TD>
		<TD bgcolor='#376091'><font color='#ffffff'><strong>Servidores</strong></font></TD>
		<TD bgcolor='#376091'><font color='#ffffff'><strong>Networking</strong></font></TD>
		<TD bgcolor='#376091'><font color='#ffffff'><strong>Software</strong></font></TD>
		<TD bgcolor='#376091'><font color='#ffffff'><strong>Telecom</strong></font></TD>
		<TD bgcolor='#376091'><font color='#ffffff'><strong>Serviços</strong></font></TD>
		
		<TD bgcolor='#376091'><font color='#ffffff'><strong>Vende para Setor Publico</strong></font></TD>
		<TD bgcolor='#376091'><font color='#ffffff'><strong>Vende para Grandes Empresas +500</strong></font></TD>
		<TD bgcolor='#376091'><font color='#ffffff'><strong>Vende para SMB (-500 Funcionários)</strong></font></TD>
		<TD bgcolor='#376091'><font color='#ffffff'><strong>Vende para Consumer</strong></font></TD>
		<%
		For Each Mar In Marcas
			Response.Write("<TD bgcolor='#376091'><font color='#ffffff'><strong>" & Mar & "</strong></font></TD>")
		Next
		%>
		<TD bgcolor='#376091'><font color='#ffffff'><strong>Possui Site?</strong></font></TD>
		<TD bgcolor='#376091'><font color='#ffffff'><strong>Data Cadastro</strong></font></TD>
    </TR>
    <% 
	Do Until RS.eof 
		Response.Write "<tr>"
		Response.Write "<td bgcolor='#b8cce4'>1</td>"
		Response.Write "<td bgcolor='#b8cce4'>" & RS("Codigo") & "</td>"
		
		SQL = "SELECT * FROM Empresas WHERE Codigo = " & RS("Empresa")
		Set RSE = Conexao.execute(SQL,3)
		
		Response.Write("<TD bgcolor='#b8cce4' VALIGN=TOP>" & RSE("Nome") & "</TD>")
		
		RSE.Close
		Set RSE = Nothing
		
		Response.Write "<td bgcolor='#b8cce4'>" & RS("EstadoMatriz") & "</td>"
		Response.Write "<td bgcolor='#b8cce4'>&nbsp;</td>"
		Response.Write "<td bgcolor='#b8cce4'>" & RS("Contato") & "</td>"
		Response.Write "<td bgcolor='#b8cce4'>" & RS("Email") & "</td>"
		Response.Write "<td bgcolor='#b8cce4'>&nbsp;</td>"
		Response.Write "<td bgcolor='#b8cce4'>" & RS("QtdeFuncionarios") & "</td>"
		Response.Write "<td bgcolor='#b8cce4'>" & RS("EstadoMatriz") & "</td>"
		
		SQL = "SELECT SUM(Qtde) as Soma FROM Cadastro_Estados WHERE Cadastro = " & RS("Codigo")
		Set RSE = Conexao.execute(SQL,3)
		
		Response.Write("<TD bgcolor='#b8cce4' VALIGN=TOP>" & RSE("Soma") & "</TD>")
		
		RSE.Close
		Set RSE = Nothing
		
		Set RSE = Server.CreateObject("ADODB.Recordset")
		RSE.CursorLocation = 3
		RSE.CursorType = 1
		RSE.LockType = 3
		
		UFA = 1
		
		For Each UF In Estados
			SQL = "SELECT Qtde FROM Cadastro_Estados WHERE Estado = '" & UF & "' AND Cadastro = " & RS("Codigo")
			RSE.Open SQL , Conexao, 1, 2
			
			borda = ""
			If UFA = 1 Then borda = "style='border-left-width:4px;'"
			
			If RSE.RecordCount > 0 Then
				Response.Write("<TD bgcolor='#b8cce4' " & borda & "><strong>" & RSe("Qtde") & "</strong></TD>")
			Else
				Response.Write "<td bgcolor='#b8cce4' " & borda & ">&nbsp;</td>"
			End IF
			RSE.Close
			UFA = UFA + 1
		Next
		
		Set RSE = Nothing
		
		OutrosEstados = RS("OutrosEstados")
		
		UFA = 1
		
		For Each UF In Estados
			borda = ""
			If UFA = 1 Then borda = "style='border-left-width:4px;'"
		
			If InStr( OutrosEstados, UF ) > 0 Then
				Response.Write("<TD bgcolor='#b8cce4' " & borda & ">X</TD>")
			Else
				Response.Write "<td bgcolor='#b8cce4' " & borda & ">&nbsp;</td>"
			End IF
			UFA = UFA + 1
		Next
		
		SetorFoco = RS("SetorFoco")
		
		For Each Setor In Setores
			If InStr( UCase(SetorFoco), UCase(Setor) ) > 0 Then
				Response.Write("<TD bgcolor='#b8cce4'>X</TD>")
			Else
				Response.Write "<td bgcolor='#b8cce4'>&nbsp;</td>"
			End IF
		Next
		
		Response.Write "<td bgcolor='#b8cce4'>" & RS("PossuiVendas") & "</td>"
		Response.Write "<td bgcolor='#b8cce4'>" & RS("PossuiMarketing") & "</td>"
		Response.Write "<td bgcolor='#b8cce4'>" & RS("PossuiFinanciamento") & "</td>"
		Response.Write "<td bgcolor='#b8cce4'>" & RS("TreinaFuncionarios") & "</td>"
		
		Faturamento2011 = RS("Faturamento2011")
		
		For Each Fat In Faturamentos
			If InStr( UCase(Faturamento2011), UCase(Fat) ) > 0 Then
				Response.Write("<TD bgcolor='#b8cce4'>X</TD>")
			Else
				Response.Write "<td bgcolor='#b8cce4'>&nbsp;</td>"
			End IF
		Next
		
		Response.Write "<td bgcolor='#b8cce4'>" & RS("NivelClient") & "</td>"
		Response.Write "<td bgcolor='#b8cce4'>" & RS("NivelStorage") & "</td>"
		Response.Write "<td bgcolor='#b8cce4'>" & RS("NivelServidores") & "</td>"
		Response.Write "<td bgcolor='#b8cce4'>" & RS("NivelNetworking") & "</td>"
		Response.Write "<td bgcolor='#b8cce4'>" & RS("NivelSoftware") & "</td>"
		Response.Write "<td bgcolor='#b8cce4'>" & RS("NivelTelecom") & "</td>"
		Response.Write "<td bgcolor='#b8cce4'>" & RS("NivelServicos") & "</td>"
		Response.Write "<td bgcolor='#b8cce4'>" & RS("VendePublico") & "</td>"
		Response.Write "<td bgcolor='#b8cce4'>" & RS("VendeGrande") & "</td>"
		Response.Write "<td bgcolor='#b8cce4'>" & RS("VendeSMB") & "</td>"
		Response.Write "<td bgcolor='#b8cce4'>" & RS("VendeConsumer") & "</td>"
		
		OutrasMarcas = RS("OutrasMarcas")
		
		For Each Mar In Marcas
			If InStr( UCase(OutrasMarcas), UCase(Mar) ) > 0 Then
				Response.Write("<TD bgcolor='#b8cce4'>X</TD>")
			Else
				Response.Write "<td bgcolor='#b8cce4'>&nbsp;</td>"
			End IF
		Next
		
		Response.Write "<td bgcolor='#b8cce4'>" & RS("Site") & "</td>"
		Response.Write "<td bgcolor='#b8cce4'>" & RS("DataCadastro") & "</td>"
		
		Response.Write "</tr>"
	
		RS.MoveNext
    Loop
    
    RS.Close
    Conexao.close

Response.Write "</table>"
Response.Write "</body>"
Response.Write "</html>"
%>