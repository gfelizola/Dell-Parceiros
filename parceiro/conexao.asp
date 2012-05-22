<%
Function ArrayPush(mArray, mValue)
	Dim mValEl
	
	If IsArray(mArray) Then
		If IsArray(mValue) Then
			For Each mValEl In mValue
				Redim Preserve mArray(UBound(mArray) + 1)
				mArray(UBound(mArray)) = mValEl
			Next
		Else
			Redim Preserve mArray(UBound(mArray) + 1)
			If TypeName( mValue ) = "Dictionary" Then
				Set mArray(UBound(mArray)) = mValue
			Else
				mArray(UBound(mArray)) = mValue
			End If
		End If
	Else
		If IsArray(mValue) Then
			mArray = mValue
		Else
			mArray = Array(mValue)
		End If
	End If
	Push = UBound(mArray)
End Function

Estados = Array()
ArrayPush Estados, "AC"
ArrayPush Estados, "AL"
ArrayPush Estados, "AP"
ArrayPush Estados, "AM"
ArrayPush Estados, "BA"
ArrayPush Estados, "CE"
ArrayPush Estados, "DF"
ArrayPush Estados, "ES"
ArrayPush Estados, "GO"
ArrayPush Estados, "MA"
ArrayPush Estados, "MT"
ArrayPush Estados, "MS"
ArrayPush Estados, "MG"
ArrayPush Estados, "PA"
ArrayPush Estados, "PB"
ArrayPush Estados, "PE"
ArrayPush Estados, "PI"
ArrayPush Estados, "PR"
ArrayPush Estados, "RJ"
ArrayPush Estados, "RN"
ArrayPush Estados, "RO"
ArrayPush Estados, "RR"
ArrayPush Estados, "RS"
ArrayPush Estados, "SC"
ArrayPush Estados, "SE"
ArrayPush Estados, "SP"
ArrayPush Estados, "TO"

Marcas = Array()
ArrayPush Marcas, "HP"
ArrayPush Marcas, "Lenovo"
ArrayPush Marcas, "IBM"
ArrayPush Marcas, "Positivo"
ArrayPush Marcas, "Accer"
ArrayPush Marcas, "Itautec"
ArrayPush Marcas, "Apple"
ArrayPush Marcas, "EMC"
ArrayPush Marcas, "Microsoft"
ArrayPush Marcas, "Oracle"
ArrayPush Marcas, "SAP"
ArrayPush Marcas, "Sun"
ArrayPush Marcas, "Cisco"
ArrayPush Marcas, "CA"

Setores = Array()
ArrayPush Setores, "Governo"
ArrayPush Setores, "Educação"
ArrayPush Setores, "Manufatura"
ArrayPush Setores, "Finanças"
ArrayPush Setores, "Oil & Gás"
ArrayPush Setores, "Bem de consumo"
ArrayPush Setores, "Outros"

Faturamentos = Array()
ArrayPush Faturamentos, "De R$ 100.000,00 a R$ 500.000,00"
ArrayPush Faturamentos, "De R$ 500.000,00 a 1 milhão"
ArrayPush Faturamentos, "De 1 milhão a 5 milhões"
ArrayPush Faturamentos, "Mais de 5 milhões"

TestServers = Array("localhost")
CurrentServer = LCase( Request.ServerVariables("SERVER_NAME") )
BaseFileName = "dell_parceiros.mdb"
BaseFilePathFinal = "e:\home\agente\Dados\"

For i = 0 To Ubound( TestServers )
	If TestServers(i) = CurrentServer Then
		BaseFilePathFinal = Server.MapPath( "/Bancos/" ) & "\"
	End If
Next

BaseFullPathFinal = BaseFilePathFinal & BaseFileName
StringConn = "Provider=Microsoft.jet.OLEDB.4.0;Data Source=" & BaseFullPathFinal

Set Conexao = Server.CreateObject ("ADODB.Connection")
Conexao.Open StringConn
%>