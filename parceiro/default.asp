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
   		
		<title>Dell - Formulário de Parceiros</title>
		
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
				<h2 class="cuf titulo">Formulário de parceiros</h2>
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
							<label for="Contato" class="cuf Contato">Contato</label>
							<input type="text" name="Contato" id="Contato" value="" title="" validate="required:true" />
						</fieldset>
						
						<fieldset>
							<label for="Email" class="cuf Email">E-mail</label>
							<input type="text" name="Email" id="Email" value="" title="" validate="required:true,email:true" />
						</fieldset>
						
						<fieldset>
							<label for="Cargo" class="cuf Cargo">Cargo</label>
							<input type="text" name="Cargo" id="Cargo" value="" title="" validate="required:true" />
						</fieldset>
						
						<fieldset>
							<label for="QtdeFuncionarios" class="cuf QtdeFuncionarios">Número de funcionários da empresa (considerando matriz e filial)</label>
							<input type="text" name="QtdeFuncionarios" id="QtdeFuncionarios" value="" title="" validate="required:true" />
						</fieldset>
						
						<fieldset>
							<label for="EstadoMatriz" class="cuf EstadoMatriz">Estado da matriz?</label>
							<select name="EstadoMatriz" id="EstadoMatriz" title="" validate="required:true">
								<option value="0"></option>
								<%
								For Each UF In Estados
									Response.Write("<option value='" & UF & "'>" & UF & "</option>")
								Next
								%>
							</select>
						</fieldset>
						
						<fieldset class="Filiais">
							<label class="cuf">Quantas filiais a empresa possui e em que estados elas estão localizadas?</label>
							<%
							For Each UF In Estados
								Response.Write("<div><select name='Filiais" & UF & "' id='Filiais" & UF & "'>")
								For i = 0 To 10
									Response.Write("<option value='" & i & "'>" & i & "</option>")
								Next
								Response.Write("</select> " & UF & "</div>")
							Next
							%>
						</fieldset>
						
						<fieldset class="Outros">
							<span class="cuf">Estados que são atendidos, além de matriz e filiais?</span>
							<%
							For Each UF In Estados
								Response.Write("<label for='OutrosEstados" & UF & "'>")
								Response.Write("<input type='checkbox' name='OutrosEstados' id='OutrosEstados" & UF & "' value='" & UF & "' class='radiobutton' /> ")
								Response.Write( UF & "</label>")
							Next
							%>
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
							<span class="cuf PossuiVendas">Possui setor de vendas?</span>
							<label for="PossuiVendasSim"><input type="radio" class="radiobutton" name="PossuiVendas" id="PossuiVendasSim" value="Sim" title="" validate="required:true" /> Sim</label>
							<label for="PossuiVendasNao"><input type="radio" class="radiobutton" name="PossuiVendas" id="PossuiVendasNao" value="Não" /> Não</label>
						</fieldset>
						
						<fieldset class="simnao">
							<span class="cuf PossuiMarketing">Possui setor de marketing?</span>
							<label for="PossuiMarketingSim"><input type="radio" class="radiobutton" name="PossuiMarketing" id="PossuiMarketingSim" value="Sim" title="" validate="required:true" /> Sim</label>
							<label for="PossuiMarketingNao"><input type="radio" class="radiobutton" name="PossuiMarketing" id="PossuiMarketingNao" value="Não" /> Não</label>
						</fieldset>
						
						<fieldset class="simnao">
							<span class="cuf PossuiFinanciamento">Possui setor de financiamento?</span>
							<label for="PossuiFinanciamentoSim"><input type="radio" class="radiobutton" name="PossuiFinanciamento" id="PossuiFinanciamentoSim" value="Sim" title="" validate="required:true" /> Sim</label>
							<label for="PossuiFinanciamentoNao"><input type="radio" class="radiobutton" name="PossuiFinanciamento" id="PossuiFinanciamentoNao" value="Não" /> Não</label>
						</fieldset>
						
						<fieldset class="simnao">
							<span class="cuf TreinaFuncionarios">Realiza treinamento com os funcionários?</span>
							<label for="TreinaFuncionariosSim"><input type="radio" class="radiobutton" name="TreinaFuncionarios" id="TreinaFuncionariosSim" value="Sim" title="" validate="required:true" /> Sim</label>
							<label for="TreinaFuncionariosNao"><input type="radio" class="radiobutton" name="TreinaFuncionarios" id="TreinaFuncionariosNao" value="Não" /> Não</label>
						</fieldset>
						
						<fieldset>
							<span class="cuf Faturamento2011">Faturamento de 2011:</span>
							<label for="Faturamento2011_100"><input type="radio" class="radiobutton" name="Faturamento2011" id="Faturamento2011_100" value="De R$ 100.000,00 a R$ 500.000,00" title="" validate="required:true" /> De R$ 100.000,00 a R$ 500.000,00</label>
							<label for="Faturamento2011_500"><input type="radio" class="radiobutton" name="Faturamento2011" id="Faturamento2011_500" value="De R$ 500.000,00 a 1 milhão" /> De R$ 500.000,00 a 1 milhão</label>
							<label for="Faturamento2011_1000"><input type="radio" class="radiobutton" name="Faturamento2011" id="Faturamento2011_1000" value="De 1 milhão a 5 milhões" /> De 1 milhão a 5 milhões</label>
							<label for="Faturamento2011_5000"><input type="radio" class="radiobutton" name="Faturamento2011" id="Faturamento2011_5000" value="Mais de 5 milhões" /> Mais de 5 milhões</label>
						</fieldset>
						
						<fieldset>
							<span class="cuf NivelLinhas">Qual o nível de capacitação da equipe comercial com relação as linhas:</span>
						</fieldset>
						<fieldset class="nivel">
							<span class="cuf">Linha Serviços:</span>
							<label for="NivelServicosAlto"><input  type="radio" class="radiobutton" name="NivelServicos" id="NivelServicosAlto"  value="Alto" title="" validate="required:true" />  Alto</label>
							<label for="NivelServicosMedio"><input type="radio" class="radiobutton" name="NivelServicos" id="NivelServicosMedio" value="Médio" /> Médio</label>
							<label for="NivelServicosBaixo"><input type="radio" class="radiobutton" name="NivelServicos" id="NivelServicosBaixo" value="Baixo" /> Baixo</label>
						</fieldset>
						
						<fieldset class="nivel">
							<span class="cuf">Linha Client:</span>
							<label for="NivelClientAlto"><input  type="radio" class="radiobutton" name="NivelClient" id="NivelClientAlto"  value="Alto" title="" validate="required:true" />  Alto</label>
							<label for="NivelClientMedio"><input type="radio" class="radiobutton" name="NivelClient" id="NivelClientMedio" value="Médio" /> Médio</label>
							<label for="NivelClientBaixo"><input type="radio" class="radiobutton" name="NivelClient" id="NivelClientBaixo" value="Baixo" /> Baixo</label>
						</fieldset>
						
						<fieldset class="nivel">
							<span class="cuf">Linha Storage:</span>
							<label for="NivelStorageAlto"><input  type="radio" class="radiobutton" name="NivelStorage" id="NivelStorageAlto"  value="Alto" title="" validate="required:true" />  Alto</label>
							<label for="NivelStorageMedio"><input type="radio" class="radiobutton" name="NivelStorage" id="NivelStorageMedio" value="Médio" /> Médio</label>
							<label for="NivelStorageBaixo"><input type="radio" class="radiobutton" name="NivelStorage" id="NivelStorageBaixo" value="Baixo" /> Baixo</label>
						</fieldset>
						
						<fieldset class="nivel">
							<span class="cuf">Linha Software:</span>
							<label for="NivelSoftwareAlto"><input  type="radio" class="radiobutton" name="NivelSoftware" id="NivelSoftwareAlto"  value="Alto" title="" validate="required:true" />  Alto</label>
							<label for="NivelSoftwareMedio"><input type="radio" class="radiobutton" name="NivelSoftware" id="NivelSoftwareMedio" value="Médio" /> Médio</label>
							<label for="NivelSoftwareBaixo"><input type="radio" class="radiobutton" name="NivelSoftware" id="NivelSoftwareBaixo" value="Baixo" /> Baixo</label>
						</fieldset>
						
						<fieldset class="nivel">
							<span class="cuf">Linha Servidores:</span>
							<label for="NivelServidoresAlto"><input  type="radio" class="radiobutton" name="NivelServidores" id="NivelServidoresAlto"  value="Alto" title="" validate="required:true" />  Alto</label>
							<label for="NivelServidoresMedio"><input type="radio" class="radiobutton" name="NivelServidores" id="NivelServidoresMedio" value="Médio" /> Médio</label>
							<label for="NivelServidoresBaixo"><input type="radio" class="radiobutton" name="NivelServidores" id="NivelServidoresBaixo" value="Baixo" /> Baixo</label>
						</fieldset>
						
						<fieldset class="nivel">
							<span class="cuf">Linha Telecom:</span>
							<label for="NivelTelecomAlto"><input  type="radio" class="radiobutton" name="NivelTelecom" id="NivelTelecomAlto"  value="Alto" title="" validate="required:true" />  Alto</label>
							<label for="NivelTelecomMedio"><input type="radio" class="radiobutton" name="NivelTelecom" id="NivelTelecomMedio" value="Médio" /> Médio</label>
							<label for="NivelTelecomBaixo"><input type="radio" class="radiobutton" name="NivelTelecom" id="NivelTelecomBaixo" value="Baixo" /> Baixo</label>
						</fieldset>
						
						<fieldset class="nivel">
							<span class="cuf">Linha Networking:</span>
							<label for="NivelNetworkingAlto"><input  type="radio" class="radiobutton" name="NivelNetworking" id="NivelNetworkingAlto"  value="Alto" title="" validate="required:true" />  Alto</label>
							<label for="NivelNetworkingMedio"><input type="radio" class="radiobutton" name="NivelNetworking" id="NivelNetworkingMedio" value="Médio" /> Médio</label>
							<label for="NivelNetworkingBaixo"><input type="radio" class="radiobutton" name="NivelNetworking" id="NivelNetworkingBaixo" value="Baixo" /> Baixo</label>
						</fieldset>
						
						<fieldset class="simnao">
							<span class="cuf VendePublico">Vende para setor público?</span>
							<label for="VendePublicoSim"><input type="radio" class="radiobutton" name="VendePublico" id="VendePublicoSim" value="Sim" title="" validate="required:true" /> Sim</label>
							<label for="VendePublicoNao"><input type="radio" class="radiobutton" name="VendePublico" id="VendePublicoNao" value="Não" /> Não</label>
						</fieldset>
						
						<fieldset class="simnao">
							<span class="cuf VendeGrande">Vende para empresas de setor grande? (+ de 500 funcionários)</span>
							<label for="VendeGrandeSim"><input type="radio" class="radiobutton" name="VendeGrande" id="VendeGrandeSim" value="Sim" title="" validate="required:true" /> Sim</label>
							<label for="VendeGrandeNao"><input type="radio" class="radiobutton" name="VendeGrande" id="VendeGrandeNao" value="Não" /> Não</label>
						</fieldset>
						
						<fieldset class="simnao">
							<span class="cuf VendeSMB">Vende para SBM? (- de 500 funcionários)</span>
							<label for="VendeSMBSim"><input type="radio" class="radiobutton" name="VendeSMB" id="VendeSMBSim" value="Sim" title="" validate="required:true" /> Sim</label>
							<label for="VendeSMBNao"><input type="radio" class="radiobutton" name="VendeSMB" id="VendeSMBNao" value="Não" /> Não</label>
						</fieldset>
						
						<fieldset class="simnao">
							<span class="cuf VendeConsumer">Vende para consumer?</span>
							<label for="VendeConsumerSim"><input type="radio" class="radiobutton" name="VendeConsumer" id="VendeConsumerSim" value="Sim" title="" validate="required:true" /> Sim</label>
							<label for="VendeConsumerNao"><input type="radio" class="radiobutton" name="VendeConsumer" id="VendeConsumerNao" value="Não" /> Não</label>
						</fieldset>
						
						<fieldset class="marcas">
							<span class="cuf OutrasMarcas">Além de produtos Dell, quais marcas a sua empresa trabalha:</span>
							<label for="OutrasMarcasHP"><input 			type="checkbox" class="radiobutton" name="OutrasMarcas" id="OutrasMarcasHP" 			value="HP" /> 			HP</label>
							<label for="OutrasMarcasLenovo"><input 		type="checkbox" class="radiobutton" name="OutrasMarcas" id="OutrasMarcasLenovo" 		value="Lenovo" /> 		Lenovo</label>
							<label for="OutrasMarcasIBM"><input 		type="checkbox" class="radiobutton" name="OutrasMarcas" id="OutrasMarcasIBM" 			value="IBM" /> 			IBM</label>
							<label for="OutrasMarcasPositivo"><input 	type="checkbox" class="radiobutton" name="OutrasMarcas" id="OutrasMarcasPositivo" 		value="Positivo" /> 	Positivo</label>
							<label for="OutrasMarcasAccer"><input 		type="checkbox" class="radiobutton" name="OutrasMarcas" id="OutrasMarcasHP" 			value="Accer" /> 		Accer</label>
							<label for="OutrasMarcasItautec"><input 	type="checkbox" class="radiobutton" name="OutrasMarcas" id="OutrasMarcasItautec" 		value="Itautec" /> 		Itautec</label>
							<label for="OutrasMarcasApple"><input 		type="checkbox" class="radiobutton" name="OutrasMarcas" id="OutrasMarcasApple" 			value="Apple" /> 		Apple</label>
							<label for="OutrasMarcasEMC"><input 		type="checkbox" class="radiobutton" name="OutrasMarcas" id="OutrasMarcasEMC" 			value="EMC" /> 			EMC</label>
							<label for="OutrasMarcasMicrosoft"><input 	type="checkbox" class="radiobutton" name="OutrasMarcas" id="OutrasMarcasMicrosoft" 		value="Microsoft" /> 	Microsoft</label>
							<label for="OutrasMarcasOracle"><input 		type="checkbox" class="radiobutton" name="OutrasMarcas" id="OutrasMarcasOracle" 		value="Oracle" /> 		Oracle</label>
							<label for="OutrasMarcasSAP"><input 		type="checkbox" class="radiobutton" name="OutrasMarcas" id="OutrasMarcasSAP" 			value="SAP" /> 			SAP</label>
							<label for="OutrasMarcasSun"><input 		type="checkbox" class="radiobutton" name="OutrasMarcas" id="OutrasMarcasSun" 			value="Sun" /> 			Sun</label>
							<label for="OutrasMarcasCisco"><input 		type="checkbox" class="radiobutton" name="OutrasMarcas" id="OutrasMarcasCisco" 			value="Cisco" /> 		Cisco</label>
							<label for="OutrasMarcasCA"><input 			type="checkbox" class="radiobutton" name="OutrasMarcas" id="OutrasMarcasCA" 			value="CA" /> 			CA</label>
						</fieldset>
						
						<fieldset>
							<label for="Site" class="cuf Site">Qual o site da empresa:</label>
							<input type="text" name="Site" id="Site" value="" title="" validate="required:true" />
						</fieldset>
						
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
			Cufon.replace('.cuf', { fontFamily: 'Museo For Dell 300', hover: true });
	        Cufon.replace('.cufBold', { fontFamily: 'Museo For Dell 700', hover: true });
			
			<%
			If Request.QueryString("sucesso") = "true" Then
				Response.Write("alert('Seus dados foram enviados. Obrigado');")
			End IF
			
			rs.Close
			Set rs = Nothing
			%>
		</script>
	</body>
</html>