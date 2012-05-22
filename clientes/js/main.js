var qtdeContatos = 0 ;
var camposHTML = "";

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
		camposHTML = $('.contatos_template').html() ;
		$('.contatos_template').html('');
		addContato();
		
		$('#AddContato').click(function(e) {
			e.preventDefault();
			addContato();
		});
		
		Cufon.replace('.cuf', { fontFamily: 'Museo For Dell 300', hover: true });
	    Cufon.replace('.cufBold', { fontFamily: 'Museo For Dell 700', hover: true });
		
		jQuery.validator.messages.required = "";
		jQuery.validator.messages.email = "";
		
		$.metadata.setType("attr", "validate");
		
		$('#formulario').validate({
			highlight: function(element) {
				$(element).parent('fieldset').children('.cuf:not(.errored)').addClass('errored');
				$(element).parent('label').parent('fieldset').children('.cuf:not(.errored)').addClass('errored');
				Cufon.refresh();
			},
			unhighlight: function(element) {
				$(element).parent('fieldset').children('.errored').removeClass('errored');
				$(element).parent('label').parent('fieldset').children('.errored').removeClass('errored');
				Cufon.refresh();
			}
		});
	}
}

function addContato(){
	qtdeContatos++ ;
	$('.contatos_template').append( replaceAll(camposHTML,"$",qtdeContatos) );
	$('#qtdeContatos').val(qtdeContatos);
	
	if( qtdeContatos >= 3 ) $('#AddContato').hide();
	Cufon.refresh();
	
	$('.MaskedTel').mask('(99) 9999-9999');
}

function replaceAll(string, token, newtoken) {
	while (string.indexOf(token) != -1) {
 		string = string.replace(token, newtoken);
	}
	return string;
}