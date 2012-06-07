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

Corpo = "Dados enviados pelo site Parceiros<br><br>"

Sub RW( msg )
	Response.Write msg & "<br>" & vbcrlf
End Sub

Function AcertaCampo(nomeCampo)
	valor = Ucase( Trim( Request.Form(nomeCampo) ) )
	valor = Replace( valor , "'" , "" )
	
	Corpo = Corpo & NomeCampo & ": " & Valor & "<br>"
	
	rs.Fields( nomeCampo ).Value = valor
	
	AcertaCampo = valor
End Function 

Nome = AcertaCampo("Contato")
Email = AcertaCampo("Email")

AcertaCampo "Telefone"
AcertaCampo "Empresa"
AcertaCampo "Cargo"
AcertaCampo "QtdeFuncionarios"
AcertaCampo "EstadoMatriz"
AcertaCampo "EstadosFiliais"
AcertaCampo "QtdeFiliais"
AcertaCampo "OutrosEstados"
AcertaCampo "SetorFoco"
AcertaCampo "PossuiVendas"
AcertaCampo "PossuiMarketing"
AcertaCampo "PossuiFinanciamento"
AcertaCampo "TreinaFuncionarios"
AcertaCampo "Faturamento2011"
AcertaCampo "NivelClient"
AcertaCampo "NivelServicos"
AcertaCampo "NivelStorage"
AcertaCampo "NivelSoftware"
AcertaCampo "NivelServidores"
AcertaCampo "NivelTelecom"
AcertaCampo "NivelNetworking"
AcertaCampo "VendePublico"
AcertaCampo "VendeGrande"
AcertaCampo "VendeSMB"
AcertaCampo "VendeConsumer"
AcertaCampo "OutrasMarcas"
AcertaCampo "Site"

rs.Fields("DataCadastro").Value = Now
rs.Update

rs.Close

Conexao.close
Set Conexao = Nothing

'Response.Write Corpo
'Response.Flush()

SET AspEmail = Server.CreateObject("Persits.MailSender")
AspEmail.Host = "localhost"
AspEmail.FromName = Nome
AspEmail.From = Email
 
'Configura os destinatários da mensagem
AspEmail.AddAddress "gfelizola@gmail.com", "Dell - PartnerDirect"
AspEmail.Subject = "Novo cadastro - Parceiro"
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
	Response.write erro
Else
    'Response.write "<font color='blue'><b>Mensagem enviada com sucesso para</b></font> "
	Response.Redirect("default.asp?sucesso=true")
End If

SET AspEmail = Nothing
%>