<%
Response.ContentType = "application/vnd.ms-excel;"
Response.AddHeader "content-disposition" , "attachment;filename=planilha_dell_clientes_" & year(now) & "_" & month(now) & "_" & day(now) & ".xls"
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
		<TD bgcolor="#9e3232">&nbsp;</TD>
		<TD bgcolor="#9e3232">&nbsp;</TD>
		<TD bgcolor="#9e3232">&nbsp;</TD>
		<TD bgcolor="#9e3232">&nbsp;</TD>
		<TD bgcolor="#9e3232">&nbsp;</TD>
		<TD bgcolor="#9e3232">&nbsp;</TD>
		<TD bgcolor="#9e3232">&nbsp;</TD>
		<TD bgcolor="#9e3232">&nbsp;</TD>
		<TD bgcolor="#9e3232">&nbsp;</TD>
		<TD bgcolor="#9e3232">&nbsp;</TD>
		<TD bgcolor="#9e3232">&nbsp;</TD>
		<TD bgcolor="#9e3232">&nbsp;</TD>
		<TD bgcolor="#9e3232">&nbsp;</TD>
		<TD bgcolor="#9e3232">&nbsp;</TD>
		<TD bgcolor="#9e3232">&nbsp;</TD>
		<TD bgcolor="#9e3232">&nbsp;</TD>
		<TD bgcolor="#9e3232">&nbsp;</TD>
		<TD bgcolor="#9e3232">&nbsp;</TD>
		<TD bgcolor="#9e3232">&nbsp;</TD>
		<TD bgcolor="#9e3232">&nbsp;</TD>
		<TD bgcolor="#9e3232">&nbsp;</TD>
		<TD bgcolor="#9e3232">&nbsp;</TD>
		<TD bgcolor="#9e3232">&nbsp;</TD>
		<TD bgcolor="#9e3232" align="center" colspan="<%=Ubound(Estados)+2%>" style="border-left-width:4px;border-right-width:4px;"><font color="#FFFFFF"><strong>Localização das Filiais</strong></font></TD>
		<TD bgcolor="#9e3232">&nbsp;</TD>
		<TD bgcolor="#9e3232" align="center" colspan="<%=Ubound(Setores)+1%>"><font color="#FFFFFF"><strong>Setor Foco</strong></font></TD>
		<%
		For Each Infra In Infras
			Response.Write("<TD bgcolor='#9e3232'>&nbsp;</TD>")
			Response.Write("<TD bgcolor='#9e3232' align='center' colspan='" & (Ubound(Marcas)+1) & "'><font color='#FFFFFF'><strong>Marcas</strong></font></TD>")
		Next
		%>
		<TD bgcolor="#9e3232">&nbsp;</TD>
		<TD bgcolor="#9e3232">&nbsp;</TD>
		<TD bgcolor="#9e3232">&nbsp;</TD>
		<TD bgcolor="#9e3232">&nbsp;</TD>
    </TR>
	
	
    <TR>
		<TD bgcolor='#cf6969'><font color='#ffffff'><strong>account_party_id</strong></font></TD>
		<TD bgcolor='#cf6969'><font color='#ffffff'><strong>account</strong></font></TD>
		<TD bgcolor='#cf6969'><font color='#ffffff'><strong>CNPJ</strong></font></TD>
		<TD bgcolor='#cf6969'><font color='#ffffff'><strong>Cidade</strong></font></TD>
		<TD bgcolor='#cf6969'><font color='#ffffff'><strong>Estado</strong></font></TD>
		<TD bgcolor='#cf6969'><font color='#ffffff'><strong>DDD</strong></font></TD>
		<TD bgcolor='#cf6969'><font color='#ffffff'><strong>Telefone da empresa</strong></font></TD>
		<TD bgcolor='#cf6969'><font color='#ffffff'><strong>A empresa tem setor de TI?</strong></font></TD>
		<TD bgcolor='#cf6969'><font color='#ffffff'><strong>Contato 1</strong></font></TD>
		<TD bgcolor='#cf6969'><font color='#ffffff'><strong>Cargo 1</strong></font></TD>
		<TD bgcolor='#cf6969'><font color='#ffffff'><strong>E-mail 1</strong></font></TD>
		<TD bgcolor='#cf6969'><font color='#ffffff'><strong>Telefone 1</strong></font></TD>
		<TD bgcolor='#cf6969'><font color='#ffffff'><strong>Contato 2</strong></font></TD>
		<TD bgcolor='#cf6969'><font color='#ffffff'><strong>Cargo 2</strong></font></TD>
		<TD bgcolor='#cf6969'><font color='#ffffff'><strong>E-mail 2</strong></font></TD>
		<TD bgcolor='#cf6969'><font color='#ffffff'><strong>Telefone 2</strong></font></TD>
		<TD bgcolor='#cf6969'><font color='#ffffff'><strong>Contato 3</strong></font></TD>
		<TD bgcolor='#cf6969'><font color='#ffffff'><strong>Cargo 3</strong></font></TD>
		<TD bgcolor='#cf6969'><font color='#ffffff'><strong>E-mail 3</strong></font></TD>
		<TD bgcolor='#cf6969'><font color='#ffffff'><strong>Telefone 3</strong></font></TD>
		<TD bgcolor='#cf6969'><font color='#ffffff'><strong>A Empresa faz Parte de algum grupo?</strong></font></TD>
		<TD bgcolor='#cf6969'><font color='#ffffff'><strong>Nome do grupo</strong></font></TD>
		<TD bgcolor='#cf6969'><font color='#ffffff'><strong>Quantidades de Filiais</strong></font></TD>
		
		<%
		UFA = 1
		For Each UF In Estados
			borda = ""
			If UFA = 1 Then borda = "style='border-left-width:4px;'"
			'If UFA = Ubound(Estados) + 1 Then borda = "style='border-right-width:4px;'"
			
			Response.Write("<TD bgcolor='#cf6969' " & borda & "><font color='#ffffff'><strong>" & UF & "</strong></font></TD>")
			UFA = UFA + 1
		Next
		%>
		<TD bgcolor='#cf6969' style="border-right-width:4px;"><font color='#ffffff'><strong>Todos os estados do Brasil</strong></font></TD>
		
		<TD bgcolor='#cf6969'><font color='#ffffff'><strong>Nº de Empregados</strong></font></TD>
		
		<%
		For Each Setor In Setores
			Response.Write("<TD bgcolor='#cf6969'><font color='#ffffff'><strong>" & Setor & "</strong></font></TD>")
		Next
		
		For Each Infra In Infras
			Response.Write("<TD bgcolor='#cf6969'><font color='#ffffff'><strong>Quantidade de " & Infra & "</strong></font></TD>")
			For Each Mar In Marcas
				Response.Write("<TD bgcolor='#cf6969'><font color='#ffffff'><strong>" & Mar & "</strong></font></TD>")
			Next
		Next
		%>
		
		<TD bgcolor='#cf6969'><font color='#ffffff'><strong>Tem previsão de investimento para o setor de TI?</strong></font></TD>
		<TD bgcolor='#cf6969'><font color='#ffffff'><strong>Prazo</strong></font></TD>
		<TD bgcolor='#cf6969'><font color='#ffffff'><strong>Gostaria de receber contato da Dell?</strong></font></TD>
		<TD bgcolor='#cf6969'><font color='#ffffff'><strong>Data de Cadastro</strong></font></TD>
    </TR>
    <% 
	
	ArrayPush Estados, "BRASIL"
	
	Do Until RS.eof 
		Response.Write "<tr>"
		
		SQL = "SELECT * FROM Empresas WHERE Codigo = " & RS("Empresa")
		Set RSE = Conexao.execute(SQL,3)
		
		Response.Write("<TD bgcolor='#f6e2e2' VALIGN=TOP>" & RSE("Account") & "</TD>")
		Response.Write("<TD bgcolor='#f6e2e2' VALIGN=TOP>" & RSE("Nome") & "</TD>")
		
		RSE.Close
		Set RSE = Nothing
		
		Response.Write "<td bgcolor='#f6e2e2'>" & RS("CNPJ") & "</td>"
		Response.Write "<td bgcolor='#f6e2e2'>" & RS("Cidade") & "</td>"
		Response.Write "<td bgcolor='#f6e2e2'>" & RS("Estado") & "</td>"

		DDD = ""
		Telefone = RS("Telefone")
		If Telefone <> "" Then
			TelSpl = Split(Telefone,")")
			DDD = Right( TelSpl(0), 2 )
			Telefone = TelSpl(1)
		End If
		
		Response.Write("<td bgcolor='#f6e2e2' VALIGN=TOP>" & DDD & "</td>")
		Response.Write("<td bgcolor='#f6e2e2' VALIGN=TOP>" & Telefone & "</td>")
		Response.Write "<td bgcolor='#f6e2e2'>" & RS("SetorTecnologia") & "</td>"
		
		SQL = "SELECT * FROM Contatos WHERE Codigo_Cadastro = " & RS("Codigo")
		Set RSE = Conexao.execute(SQL,3)
		
		QtdeContatos = 1	 
		
		Do Until RSE.eof 
			Response.Write("<TD bgcolor='#f6e2e2' VALIGN=TOP>" & RSE("Nome") & "</TD>")
			Response.Write("<TD bgcolor='#f6e2e2' VALIGN=TOP>" & RSE("Cargo") & "</TD>")
			Response.Write("<TD bgcolor='#f6e2e2' VALIGN=TOP>" & RSE("Email") & "</TD>")
			Response.Write("<TD bgcolor='#f6e2e2' VALIGN=TOP>" & RSE("Telefone") & "</TD>")
			
			RSE.MoveNext
			QtdeContatos = QtdeContatos + 1
		Loop
		
		RSE.Close
		Set RSE = Nothing
		
		For i = QtdeContatos To 3
			Response.Write "<TD bgcolor='#f6e2e2' VALIGN=TOP>&nbsp;</TD>"
			Response.Write "<TD bgcolor='#f6e2e2' VALIGN=TOP>&nbsp;</TD>"
			Response.Write "<TD bgcolor='#f6e2e2' VALIGN=TOP>&nbsp;</TD>"
			Response.Write "<TD bgcolor='#f6e2e2' VALIGN=TOP>&nbsp;</TD>"
		Next
		
		Response.Write "<td bgcolor='#f6e2e2'>" & RS("ParticipaGrupo") & "</td>"
		Response.Write "<td bgcolor='#f6e2e2'>" & RS("Grupo") & "</td>"
		Response.Write "<td bgcolor='#f6e2e2'>" & RS("QtdeFiliais") & "</td>"
		
		UFA = 1
		
		EstadosFiliais = RS("EstadosFiliais")
		
		
		
		For Each UF In Estados
			borda = ""
			If UFA = 1 Then borda = "style='border-left-width:4px;'"
			If UFA = Ubound(Estados) + 1 Then borda = "style='border-right-width:4px;'"
		
			If InStr( EstadosFiliais, UF ) > 0 Then
				Response.Write("<TD bgcolor='#f6e2e2' " & borda & ">X</TD>")
			Else
				Response.Write "<td bgcolor='#f6e2e2' " & borda & ">&nbsp;</td>"
			End IF
			UFA = UFA + 1
		Next
		
		Response.Write "<td bgcolor='#f6e2e2'>" & RS("QtdeFuncionarios") & "</td>"
		
		SetorFoco = RS("SetorFoco")
		
		For Each Setor In Setores
			If InStr( UCase(SetorFoco), UCase(Setor) ) > 0 Then
				Response.Write("<TD bgcolor='#f6e2e2'>X</TD>")
			Else
				Response.Write "<td bgcolor='#f6e2e2'>&nbsp;</td>"
			End IF
		Next
		
		For Each Infra In Infras
			QtdeInfra = RS("Qtde" & Infra)
			MarcasInfra = RS("Marcas" & Infra)
			
			Response.Write "<td bgcolor='#f6e2e2'>" & QtdeInfra & "</td>"
		
			For Each Mar In Marcas
				If InStr( UCase(MarcasInfra), UCase(Mar) ) > 0 Then
					Response.Write("<TD bgcolor='#f6e2e2'>X</TD>")
				Else
					Response.Write "<td bgcolor='#f6e2e2'>&nbsp;</td>"
				End IF
			Next
		Next
		
		Response.Write "<td bgcolor='#f6e2e2'>" & RS("PrevisaoCompras") & "</td>"
		Response.Write "<td bgcolor='#f6e2e2'>" & RS("PrazoInvestimento") & "</td>"
		Response.Write "<td bgcolor='#f6e2e2'>" & RS("QuerContato") & "</td>"
		Response.Write "<td bgcolor='#f6e2e2'>" & RS("DataCadastro") & "</td>"
		
		Response.Write "</tr>"
	
		RS.MoveNext
    Loop
    
    RS.Close
    Conexao.close

Response.Write "</table>"
Response.Write "</body>"
Response.Write "</html>"
%>