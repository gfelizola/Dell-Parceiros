//var verificarEstudio = false ;

var Site = {
    Init: function() {
		String.prototype.linkar = function( url ){
			return '<a href="' + url + '" target="_blank">' + this + '</a>' ;
		};
		
		String.prototype.zeros = function( qtde ){
			var texto = this ;
			while( texto.length < qtde ) texto = '0' + texto ;
			return texto ;
		};
		
		String.prototype.parseURL = function() {
			return this.replace(/[A-Za-z]+:\/\/[A-Za-z0-9-_]+\.[A-Za-z0-9-_:%&~\?\/.=]+/g, function(url) {
				return url.linkar(url);
			});
		};
		
		String.prototype.parseUsername = function() {
			return this.replace(/[@]+[A-Za-z0-9-_]+/g, function(u) {
				var username = u.replace("@","")
				return u.linkar("http://twitter.com/"+username);
			});
		};
		
		String.prototype.parseHashtag = function() {
			return this.replace(/[#]+[A-Za-z0-9-_]+/g, function(t) {
				var tag = t.replace("#","%23")
				return t.linkar("http://search.twitter.com/search?q="+tag);
			});
		};
    },
    Generics: {
        OpenExternalModal: function(id, source, w, h, content) {
			var conteudo = '<iframe width=\"' + w + '\" height=\"' + h + '\" frameborder=\"0\" scrolling=\"no\" allowtransparency=\"true\" src=\"' + source + '\"></iframe>';
			if( content != null ) conteudo = content ;
			
            ModalWindow.windowId = id;
            ModalWindow.width = w;
            ModalWindow.height = h;
            ModalWindow.content = conteudo;
            ModalWindow.Open();
        },
        OpenInternalModal: function(id) {
            $(id).jqmShow({ toTop: true });
        },
        FormsEffects: function() {
            $('input[type=text], textarea, select').focus(function() {
                $(this).addClass('on');
            });
            $('input[type=text], textarea, select').blur(function() {
                $(this).removeClass('on');
            });
        },
        ChangeFonts: function() {
            var elements = "#content #main p, #content #main p strong, #content #main p strong span, #content #main li, #content #main a, #content #main h1, #content #main h2, #content #main h3, #content #main h4, #content #main h5, #content #main h6";
            $('.aumentar-fonte').unbind().bind('click', function() {
                var currentFontSize = $(elements).css('font-size');
                var currentFontSizeNum = parseFloat(currentFontSize, 2000);
                var newFontSize = currentFontSizeNum * 1.2;
                $(elements).css('font-size', newFontSize);
                return false;
            });
            $('.diminuir-fonte').unbind().bind('click', function() {
                var currentFontSize = $(elements).css('font-size');
                var currentFontSizeNum = parseFloat(currentFontSize, 2000);
                var newFontSize = currentFontSizeNum * 0.9;
                $(elements).css('font-size', newFontSize);
                return false;
            });
        }
	},
	
	Home: function(){
		jQuery.validator.messages.required = "";
		jQuery.validator.messages.email = "";
		
		$.metadata.setType("attr", "validate");
		
		$('#formulario').validate({
			/*
			rules: {
				Empresa: "required",
				Contato: "required",
				Email: {
					required: true,
					email: true
				},
				Cargo: "required",
				QtdeFuncionarios: "required",
				EstadoMatriz: "required",
				SetorFoco: "required",
				PossuiVendas: "required",
				PossuiMarketing: "required",
				PossuiFinanciamento: "required",
				TreinaFuncionarios: "required",
				Faturamento2011: "required",
				NivelServicos: "required",
				NivelClient: "required",
				NivelStorage: "required",
				NivelSoftware: "required",
				NivelServidores: "required",
				NivelTelecom: "required",
				NivelNetworking: "required",
				VendePublico: "required",
				VendeGrande: "required",
				VendeSMB: "required",
				VendeConsumer: "required",
				Site: "required"
			},
			
			errorPlacement: function(error, element) {
				
			},
			*/
			
			highlight: function(element) {
				$(element).parent('fieldset').children('.cuf:not(.errored)').addClass('errored');
				$(element).parent('label').parent('fieldset').children('.cuf:not(.errored)').addClass('errored');
				Cufon.refresh();
			},
			unhighlight: function(element) {
				$(element).parent('fieldset').children('.errored').removeClass('errored');
				$(element).parent('label').parent('fieldset').children('.errored').removeClass('errored');
				Cufon.refresh();
			},
			
			messages: {
				/*
				Nome: "Por favor, preencha o campo Nome<br>",
				Email: "Por favor, preencha o campo E-mail com um e-mail valido<br>",
				Telefone: "Por favor, preencha o campo Telefone<br>",
				Revenda: "Por favor, preencha o campo Nome da Revenda<br>",
				Gerente: "Por favor, preencha o campo Gerente de Atendimento Dell:<br>",
				Cliente: "Por favor, preencha o campo Nome do cliente<br>",
				Contato: "Por favor, preencha o campo Contato no cliente<br>",
				EmailCliente: "Por favor, preencha o campo E-mail do cliente com um e-mail valido<br>",
				RO: "Por favor, preencha o campo Número da RO<br>",
				BRC: "Por favor, preencha o campo Quantidade de BRC vendida<br>",
				*/
				
				Nome: "",
				Email: "",
				Telefone: "",
				Revenda: "",
				Gerente: "",
				Cliente: "",
				Contato: "",
				EmailCliente: "",
				RO: "",
				BRC: ""
				
			}
		});
	}
}