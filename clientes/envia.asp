<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!-- #include file="conexao.asp" -->
<%
Dim Corpo, Nome, Email

Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3
rs.CursorType = 1
rs.LockType = 3

SQL = "SELECT * FROM Cadastros"
rs.Open SQL , Conexao, 1, 2

rs.AddNew

Corpo = "Dados enviados pelo formulário de clientes<br><br>"

Sub RW( msg )
	Response.Write msg & "<br>" & vbcrlf
	Response.Flush()
End Sub

Function AcertaCampo(nomeCampo, nomeBanco)
	valor = Ucase( Trim( Request.Form(NomeCampo) ) )
	valor = Replace( valor , "'" , "" )
	
	Corpo = Corpo & NomeCampo & ": " & Valor & "<br>"
	If NomeBanco = "" Then NomeBanco = NomeCampo
	
	'RW NomeBanco & ": " & Valor
	rs.Fields( NomeBanco ).Value = valor
	
	AcertaCampo = valor
End Function 

AcertaCampo "Empresa", ""
AcertaCampo "CNPJ", ""
AcertaCampo "ParticipaGrupo", ""
AcertaCampo "Grupo", ""
AcertaCampo "Cidade", ""
AcertaCampo "Estado", ""
AcertaCampo "PossuiFiliais", ""
AcertaCampo "QtdeFiliais", ""
AcertaCampo "EstadosFiliais", ""
AcertaCampo "QtdeFuncionarios", ""
AcertaCampo "SetorFoco", ""
AcertaCampo "SetorTecnologia", ""
AcertaCampo "QtdeDesktops", ""
AcertaCampo "MarcasDesktops", ""
AcertaCampo "QtdeNotebooks", ""
AcertaCampo "MarcasNotebooks", ""
AcertaCampo "QtdeServidores", ""
AcertaCampo "MarcasServidores", ""
AcertaCampo "QtdeStorages", ""
AcertaCampo "MarcasStorages", ""
AcertaCampo "PrevisaoCompras", ""
AcertaCampo "PrazoInvestimento", ""
AcertaCampo "QuerContato", ""

rs.Fields("DataCadastro").Value = Now
rs.Update

Codigo = rs("codigo")

rs.Close

Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3
rs.CursorType = 1
rs.LockType = 3

SQL = "SELECT * FROM Contatos"
rs.Open SQL , Conexao, 1, 2

QtdeF = Request.Form("qtdeContatos")

Nome = ""
Email = ""

If QtdeF <> "0" Then
	For i = 1 To CDbl(QtdeF)
		rs.AddNew
		NomeB = AcertaCampo( "Nome" & i, "Nome" )
		EmailB = AcertaCampo( "Email" & i, "Email" )
		AcertaCampo "Sobrenome" & i, "Sobrenome"
		AcertaCampo "Telefone" & i, "Telefone"
		
		If Nome = "" Then Nome = NomeB
		If Email = "" Then Email = EmailB
		
		rs.Fields("Codigo_Cadastro").Value = Codigo
		rs.Update
	Next
End IF

rs.Close
Set rs = Nothing

Conexao.close
Set Conexao = Nothing

RW Corpo

SET AspEmail = Server.CreateObject("Persits.MailSender")
AspEmail.Host = "localhost"
AspEmail.FromName = Nome
AspEmail.From = Email
 
'Configura os destinatários da mensagem
AspEmail.AddAddress "gfelizola@gmail.com", "Dell - PartnerDirect"
AspEmail.Subject = "Novo cadastro - Cliente"
AspEmail.IsHTML = True
AspEmail.Body = Corpo
 
'#Ativa o tratamento de erros
On Error Resume Next
 
'Envia a mensagem
AspEmail.Send
 
'Caso ocorra problemas no envio, descreve os detalhes do mesmo.
If Err <> 0 Then
	erro = "<b><font color='red'> Erro ao enviar a mensagem.</font></b><br>"
	erro = erro & "<b>Erro.Description:</b> " & Err.Description & "<br>"
	erro = erro & "<b>Erro.Number:</b> "      & Err.Number & "<br>"
	erro = erro & "<b>Erro.Source:</b> "      & Err.Source & "<br>"
	RW erro
Else
    RW "<font color='blue'><b>Mensagem enviada com sucesso</b></font> "
	Response.Redirect("default.asp?sucesso=true")
End If

SET AspEmail = Nothing
%>