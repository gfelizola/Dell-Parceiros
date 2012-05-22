<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!-- #include file="conexao.asp" -->
<%
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3
rs.CursorType = 1
rs.LockType = 3

SQL = "SELECT e.nome as Nome, e.codigo as Codigo FROM Empresas as e, Cadastros as c WHERE e.codigo <> c.empresa ORDER BY nome"
rs.Open SQL , Conexao, 1, 2

If rs.RecordCount = 0 Then
	rs.Close

	SQL = "SELECT * FROM Empresas ORDER BY nome"
	rs.Open SQL , Conexao, 1, 2
End IF
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" 		content="text/html; charset=iso-8859-1" 											/>
   		
		<title>Dell - Formulário de Clientes</title>
		
		<meta http-equiv="Content-Style-Type"	content="text/css"																	/>
		<meta http-equiv="Content-Language" 	content="pt-br" 																	/>
		<meta http-equiv="pragma" 				content="no-cache"																	/>
		<meta http-equiv="cache-control" 		content="no-cache"																	/>
		<meta name="expires" 					content="0"																			/>
		<meta name="robots" 					content="all" 																		/>
		<meta name="robots" 					content="index,follow" 																/>
		<meta name="title" 						content=""																			/>
		<meta name="author" 					content="Gustavo Felizola"															/>
		<meta name="description" 				content="" 																			/>
		<meta name="keywords" 					content=""																			/>
		<meta name="distribution" 				content="Global"		 															/>
		
		<link rel="shortcut icon" 				type="image/ico" 								href="img/common/favicon.ico" 		/>
		
		<link rel="stylesheet" 					type="text/css" 	media="screen, projection" 	href="css/screen.css" 				/>
		<link rel="stylesheet" 					type="text/css" 	media="print" 				href="css/print.css" 				/>

		<!--[if lte IE 6]>
		<link rel="stylesheet" 					type="text/css" 	media="screen, projection" 	href="css/hackIe6.css" 				/>
		<![endif]-->
		<!--[if IE 7]>
		<link rel="stylesheet" 					type="text/css" 	media="screen, projection" 	href="css/hackIe7.css" 				/>
		<![endif]-->

		<script type="text/javascript" 			src="js/lib/jquery/jquery-1.7.1.js"						></script>
		<script type="text/javascript" 			src="js/lib/jquery/jquery.maskedinput-1.3.js"			></script>
		<script type="text/javascript" 			src="js/lib/jquery/jquery.metadata.js"					></script>
		<script type="text/javascript" 			src="js/lib/jquery/jquery.validate.js"					></script>
		
		<script type="text/javascript" 			src="js/lib/cufon/cufon-1.09.js"						></script>
		<script type="text/javascript" 			src="js/lib/cufon/Museo_For_Dell.font.js"				></script>
		
		<script type="text/javascript" 			src="js/main.js"										></script>
		
	</head>
	<body>
		<div id="site">
        	<div id="header">
				<h1 class="repfl">Dell - PartnerDirect</h1>
			</div>
			<div id="content">
				<h2 class="cuf titulo">Formulário de clientes</h2>
				<h3 class="cuf subtitulo">Registre abaixo seus dados.</h3>
				
				<div class="form fl">
					<form name="formulario" id="formulario" action="envia.asp" method="post">
					
						<fieldset>
							<label for="Empresa" class="cuf Empresa">Empresa</label>
							<select name="Empresa" id="Empresa" title="" validate="required:true">
								<option value=""></option>
								<%
								Do Until rs.Eof
									Response.Write("<option value='" & rs("Codigo") & "'>" & rs("Nome") & "</option>")
									rs.MoveNext
								Loop
								%>
							</select>
						</fieldset>
					
						<fieldset>
							<label for="CNPJ" class="cuf CNPJ">CNPJ</label>
							<input type="text" name="CNPJ" id="CNPJ" value="" title="" validate="required:true" />
						</fieldset>
						
						<div class="contatos_template">
							<div class="template">
								<fieldset class="Contatos">
									<label for="Nome$" class="cuf Nome">Nome do contato $</label>
									<input type="text" name="Nome$" id="Nome$" value="" title="" validate="required:true" />
								</fieldset>
								<fieldset class="Contatos">
									<label for="Sobrenome$" class="cuf Sobrenome">Sobrenome</label>
									<input type="text" name="Sobrenome$" id="Sobrenome$" value="" title="" validate="required:true" />
								</fieldset>
								<fieldset class="Contatos">
									<label for="Email$" class="cuf Email">E-mail</label>
									<input type="text" name="Email$" id="Email$" value="" title="" validate="required:true" />
								</fieldset>
								<fieldset class="Contatos">
									<label for="Telefone$" class="cuf Telefone">Telefone</label>
									<input type="text" name="Telefone$" id="Telefone$" class="MaskedTel" value="" title="" validate="required:true" />
								</fieldset>
							</div>
						</div>
						
						<input type="image" name="AddContato" id="AddContato" src="img/adicionar_contato.jpg" width="135" height="28" />
						
						<fieldset class="simnao">
							<span class="cuf ParticipaGrupo">A empresa faz parte de algum grupo?</span>
							<label for="ParticipaGrupoSim"><input type="radio" class="radiobutton" name="ParticipaGrupo" id="ParticipaGrupoSim" value="Sim" title="" validate="required:true" /> Sim</label>
							<label for="ParticipaGruposNao"><input type="radio" class="radiobutton" name="ParticipaGrupo" id="ParticipaGruposNao" value="Não" /> Não</label>
						</fieldset>
						
						<fieldset>
							<label for="Grupo" class="cuf Grupo">Grupo</label>
							<input type="text" name="Grupo" id="Grupo" value="" title="" validate="required:'#ParticipaGrupoSim:checked'" />
						</fieldset>
						
						<fieldset>
							<label for="Cidade" class="cuf Cidade">Cidade e Estado da Matriz</label>
							<input type="text" name="Cidade" id="Cidade" value="" title="" validate="required:true" />
							<select name="Estado" id="Estado" title="" validate="required:true">
								<option value="0"></option>
								<%
								For Each UF In Estados
									Response.Write("<option value='" & UF & "'>" & UF & "</option>")
								Next
								%>
							</select>
						</fieldset>
						
						<fieldset class="simnao">
							<span class="cuf PossuiFiliais">Tem filiais?</span>
							<label for="PossuiFiliaisSim"><input type="radio" class="radiobutton" name="PossuiFiliais" id="PossuiFiliaisSim" value="Sim" title="" validate="required:true" /> Sim</label>
							<label for="PossuiFiliaisNao"><input type="radio" class="radiobutton" name="PossuiFiliais" id="PossuiFiliaisNao" value="Não" /> Não</label>
						</fieldset>
						
						<fieldset>
							<label for="QtdeFiliais" class="cuf QtdeFiliais">Quantidade de filiais:</label>
							<input type="text" name="QtdeFiliais" id="QtdeFiliais" value="" title="" validate="required:'#PossuiFiliaisSim:checked'" />
						</fieldset>
						
						<fieldset class="EstadosFiliais Outros">
							<span class="cuf">Em quais estados estão as filiais:</span>
							<%
							For Each UF In Estados
								Response.Write("<label for='EstadosFiliais" & UF & "'>")
								Response.Write("<input type='checkbox' name='EstadosFiliais' id='EstadosFiliais" & UF & "' value='" & UF & "' class='radiobutton' /> ")
								Response.Write( UF & "</label>")
							Next
							%>
						</fieldset>
						
						<fieldset>
							<label for="QtdeFuncionarios" class="cuf QtdeFuncionarios">Quantidades de funcionário:</label>
							<input type="text" name="QtdeFuncionarios" id="QtdeFuncionarios" value="" title="" validate="required:true" />
						</fieldset>
						
						<fieldset>
							<span class="cuf SetorFoco">Setor foco:</span>
							<label for="SetorFocoGoverno"><input type="radio" class="radiobutton" name="SetorFoco" id="SetorFocoGoverno" value="Governo" title="" validate="required:true" /> Governo</label>
							<label for="SetorFocoEducacao"><input type="radio" class="radiobutton" name="SetorFoco" id="SetorFocoEducacao" value="Educação" /> Educação</label>
							<label for="SetorFocoManufatura"><input type="radio" class="radiobutton" name="SetorFoco" id="SetorFocoManufatura" value="Manufatura" /> Manufatura</label>
							<label for="SetorFocoFinancas"><input type="radio" class="radiobutton" name="SetorFoco" id="SetorFocoFinancas" value="Finanças" /> Finanças</label>
							<label for="SetorFocoOil"><input type="radio" class="radiobutton" name="SetorFoco" id="SetorFocoOil" value="Oil & Gás" /> Oil & Gás</label>
							<label for="SetorFocoBem"><input type="radio" class="radiobutton" name="SetorFoco" id="SetorFocoBem" value="Bem de consumo" /> Bem de consumo</label>
							<label for="SetorFocoOutros"><input type="radio" class="radiobutton" name="SetorFoco" id="SetorFocoOutros" value="Outros" /> Outros</label>
						</fieldset>
						
						<fieldset class="simnao">
							<span class="cuf">A empresa possui setor de tecnologia?</span>
							<label for="SetorTecnologiaSim"><input type="radio" class="radiobutton" name="SetorTecnologia" id="SetorTecnologiaSim" value="Sim" title="" validate="required:true" /> Sim</label>
							<label for="SetorTecnologiaNao"><input type="radio" class="radiobutton" name="SetorTecnologia" id="SetorTecnologiaNao" value="Não" /> Não</label>
						</fieldset>
						
						<fieldset>
							<label for="QtdeDesktops" class="cuf QtdeDesktops">Quantidades de Desktops:</label>
							<input type="text" name="QtdeDesktops" id="QtdeDesktops" value="" title="" validate="required:true" />
						</fieldset>
						
						<fieldset class="QtdeDesktop Marcas">
							<span class="cuf">De qual marca?</span>
							<%
							For Each Mar In Marcas
								Response.Write("<label for='MarcasDesktops" & Mar & "'>")
								Response.Write("<input type='checkbox' name='MarcasDesktops' id='MarcasDesktops" & Mar & "' value='" & Mar & "' class='radiobutton' /> ")
								Response.Write( Mar & "</label>")
							Next
							%>
						</fieldset>
						
						<fieldset>
							<label for="QtdeNotebooks" class="cuf QtdeNotebooks">Quantidades de Notebooks:</label>
							<input type="text" name="QtdeNotebooks" id="QtdeNotebooks" value="" title="" validate="required:true" />
						</fieldset>
						
						<fieldset class="MarcasNotebooks Marcas">
							<span class="cuf">De qual marca?</span>
							<%
							For Each Mar In Marcas
								Response.Write("<label for='MarcasNotebooks" & Mar & "'>")
								Response.Write("<input type='checkbox' name='MarcasNotebooks' id='MarcasNotebooks" & Mar & "' value='" & Mar & "' class='radiobutton' /> ")
								Response.Write( Mar & "</label>")
							Next
							%>
						</fieldset>
						
						<fieldset>
							<label for="QtdeServidores" class="cuf QtdeServidores">Quantidades de Servidores:</label>
							<input type="text" name="QtdeServidores" id="QtdeServidores" value="" title="" validate="required:true" />
						</fieldset>
						
						<fieldset class="MarcasServidores Marcas">
							<span class="cuf">De qual marca?</span>
							<%
							For Each Mar In Marcas
								Response.Write("<label for='MarcasServidores" & Mar & "'>")
								Response.Write("<input type='checkbox' name='MarcasServidores' id='MarcasServidores" & Mar & "' value='" & Mar & "' class='radiobutton' /> ")
								Response.Write( Mar & "</label>")
							Next
							%>
						</fieldset>
						
						<fieldset>
							<label for="QtdeStorages" class="cuf QtdeStorages">Quantidades de Storages:</label>
							<input type="text" name="QtdeStorages" id="QtdeStorages" value="" title="" validate="required:true" />
						</fieldset>
						
						<fieldset class="MarcasStorages Marcas">
							<span class="cuf">De qual marca?</span>
							<%
							For Each Mar In Marcas
								Response.Write("<label for='MarcasStorages" & Mar & "'>")
								Response.Write("<input type='checkbox' name='MarcasStorages' id='MarcasStorages" & Mar & "' value='" & Mar & "' class='radiobutton' /> ")
								Response.Write( Mar & "</label>")
							Next
							%>
						</fieldset>
						
						<fieldset>
							<label for="PrevisaoCompras" class="cuf PrevisaoCompras">Previsão para compras no setor de T.I:</label>
							<input type="text" name="PrevisaoCompras" id="PrevisaoCompras" value="" title="" validate="required:true" />
						</fieldset>
						
						<fieldset>
							<label for="PrazoInvestimento" class="cuf PrazoInvestimento">Qual o prazo de investimento?</label>
							<input type="text" name="PrazoInvestimento" id="PrazoInvestimento" value="" title="" validate="required:true" />
						</fieldset>
						
						<fieldset class="simnao">
							<span class="cuf QuerContato">Gostaria de receber um contato telefônico da Dell, para ajudá-lo em projetos futuros <br />
								ou em andamento?</span>
							<label for="QuerContatoSim"><input type="radio" class="radiobutton" name="QuerContato" id="QuerContatoSim" value="Sim" title="" validate="required:true" /> Sim</label>
							<label for="QuerContatoNao"><input type="radio" class="radiobutton" name="QuerContato" id="QuerContatoNao" value="Não" /> Não</label>
						</fieldset>
						
						<input type="hidden" name="qtdeContatos" id="qtdeContatos" value="0" />
						<input type="image" src="img/bt_enviar.png" width="105" height="28" id="Enviar" name="Enviar" value="Enviar >" class="cuf fr" />
					</form>
				</div>
				<div class="form_footer fl"></div>
        	</div>
        </div>
    
		<script type="text/javascript">
		    $(document).ready(function() {
				Site.Init();
				Site.Home();
			});
			
			<%
			If Request.QueryString("sucesso") = "true" Then
				Response.Write("alert('Seus dados foram enviados. Obrigado');")
			End IF
			%>
		</script>
	</body>
</html>