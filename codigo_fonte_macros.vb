REM *********************************************************************************************************
REM *****  FLAGRANTES  *****  MACROS PARA OPENOFFICE  *******************************************************
REM *********************************************************************************************************
REM
REM Autor: Fabio Minami 
REM
REM Biblioteca de macros para confecção automática de Inquéritos Policiais em OpenOffice.org.
REM 
REM Modelos de documentos pertencem a "Flagrantes Jatinhos" para MS Word de Escrivão Jorge - 9227-0192
REM 
REM Direitos Autorais Reservados (c) 2010 Fábio Minami <fminami@gmail.com>
REM 
REM Esta biblioteca é Software Livre; você pode redistribuí-la e/ou modificá-la sob os termos da
REM Licença Pública Geral Menor do GNU conforme publicada pela Free Software Foundation; tanto a
REM versão 2.1 da Licença, ou (a seu critério) qualquer versão posterior.
REM
REM Esta biblioteca é distribuída na expectativa de que seja útil, porém, SEM NENHUMA GARANTIA; 
REM nem mesmo a garantia implícita de COMERCIABILIDADE OU ADEQUAÇÃO A UMA FINALIDADE ESPECÍFICA. 
REM Consulte a Licença Pública Geral Menor do GNU para mais detalhes.
REM
REM Você deve ter recebido uma cópia da Licença Pública Geral Menor do GNU junto com esta biblioteca; 
REM se não, escreva para a Free Software Foundation, Inc., no endereço 59 Temple Street, Suite 330,
REM Boston, MA 02111-1307 USA.
REM
REM *********************************************************************************************************
REM BASIC

Option Explicit

' Dimensionar variáveis
Dim Doc As Object
Dim DatabaseContext As Object
Dim DataSource As Object
Dim Connection As Object
Dim Statement As Object
Dim PrepStatement As Object
Dim rs As Object

Dim intVitimas As Integer
Dim intConduzidos As Integer
Dim i As Integer, n as integer

Dim strDIR As String, strURL As String
Dim strIP As String, strBO As String, strDelitoTipo As String
Dim strConduzidos As String
Dim strConduzidosRgPais As String
Dim strConduzidosLista As String
Dim strConduzidosListaRelatorio As String
Dim strVitimas As String
Dim strVitimasLista As String
Dim strVitimasListaRelatorio As String
Dim S_cond As String, EM_cond As String, INS_cond As String, L_cond As String, I_cond As String
Dim U_cond As String, ES_cond As String, AO_cond As String, S_vit As String, strRelatorioFim As String
Dim strMaiusCond As String, strMaiusVitim As String, strCapaCond As String, strCapaVitim As String
Dim strTiraBarra As String

Dim strDocCapa As String
Dim strDocDespacho As String
Dim strDocVitimas As String
Dim strDocConduzidos As String
Dim strDocEspecialistas As String


REM **********
REM * MACROS
REM *


Sub Sobre_Click

Dim strLicenca As String

strLicenca = "* * *     FLAGRANTES - OpenOffice.org     * * *" & chr$(13) & chr$(13) & chr$(13) & _
"Formulários e biblioteca de macros para confecção automática de Inquéritos Policiais no OpenOffice.org. " & _
"Os modelos de documentos pertencem a ""Flagrantes Jatinhos"" para MS Word de Escrivão Jorge." & chr$(13) & chr$(13) & _
"Direitos Autorais Reservados" & chr$(13) & _
"(c) 2010 Fábio Minami" & chr$(13) & chr$(13) & _
"Esta biblioteca é Software Livre; você pode redistribuí-la e/ou modificá-la sob os termos da Licença Pública Geral Menor do GNU conforme publicada pela Free Software Foundation; tanto a versão 2.1 da Licença, ou (a seu critério) qualquer versão posterior." & chr$(13) & chr$(13) & _
"Esta biblioteca é distribuída na expectativa de que seja útil, porém, SEM NENHUMA GARANTIA; nem mesmo a garantia implícita de COMERCIABILIDADE OU ADEQUAÇÃO A UMA FINALIDADE ESPECÍFICA. Consulte a Licença Pública Geral Menor do GNU para mais detalhes." & chr$(13) & chr$(13) & _
"Você deve ter recebido uma cópia da Licença Pública Geral Menor do GNU junto com esta biblioteca; se não, escreva para a Free Software Foundation, Inc., no endereço 59 Temple Street, Suite 330, Boston, MA 02111-1307 USA."

Msgbox (strLicenca, 64, "MinamiSoft")

End Sub

Sub Main

On Error GoTo Erro:

' Questionar usuario
Dim intResposta As Integer
	intResposta = MsgBox("Gerar Inquérito Policial?" & chr$(13) & chr$(13) & "Execute este passo após preencher todos os formulários.     " & chr$(13) & chr$(13), 32 + 4, "Flagrantes OpenOffice.org - Versão BETA")
	If intResposta = 7 Then exit Sub
	' obter diretorio de trabalho
	GetDIR	
	If FileExists(strDIR & "TempMM") Then RemoverTempMM			
	CriarTempMM
	' atribuir valor as variaveis, popular tblAuxiliar
	PreencherTabelaAuxiliar	
	' verificar se as tabelas Vitimas e Conduzidos tem registros
	If intConduzidos = 0 Then ' intVitimas = 0 Or 
		MsgBox("Nenhum conduzido encontrado. Impossível gerar documento.", 48, "Flagrantes OpenOffice.org - Versão BETA")
		GoTo Sair:
	End If
	' dir de origem dos templates
	strDocCapa = strDir & "Modelos/FlagranteDoc.ott"
	strDocVitimas = strDir & "Modelos/FlagranteVitimas.ott"
	strDocConduzidos = strDir & "Modelos/FlagranteConduzidos.ott"
	If intVitimas = 0 Then
		'nenhuma vitima, criar arquivo vazio
		FileCopy(strDIR & "Modelos/Vazio.ott", strDIR & "TempMM/" & strDelitoTipo & "2.odt")
	End If	
	
	' Executar mala direta mostrando barra de progresso
Dim oProgressBar as Object, oProgressBarModel As Object, oDialog as Object
Dim ProgressValue As Long
	DialogLibraries.loadLibrary("Standard")
	oDialog = CreateUnoDialog(DialogLibraries.Standard.Dialog1)
	oProgressBarModel = oDialog.getModel().getByName( "ProgressBar1" )
	oProgressBarModel.setPropertyValue( "ProgressValueMin", 0)
	oProgressBarModel.setPropertyValue( "ProgressValueMax", 7)
	oDialog.setVisible( True )
	For ProgressValue = 0 To 7 Step 1
		oProgressBarModel.setPropertyValue( "ProgressValue", ProgressValue )
		Wait 250 '0 é para aguardar o dialogo ser desenhado na tela
		If ProgressValue = 1 Then	GerarMalaDiretaFlagrante( strDocCapa, "consulta_DADOS", strDIR & "TempMM") 'DelitoTipo & 0
		If ProgressValue = 2 Then	GerarMalaDiretaFlagrante( strDocDespacho, "consulta_DADOS", strDIR & "TempMM") 'DelitoTipo & 1	 
		If ProgressValue = 3 Then	GerarMalaDiretaFlagrante( strDocVitimas, "consulta_VitimasTodos", strDIR & "TempMM") ' DelitoTipo & 2
		If ProgressValue = 4 Then	GerarMalaDiretaFlagrante( strDocConduzidos, "consulta_ConduzidosTodos", strDIR & "TempMM") ' DelitoTipo & 3
		If ProgressValue = 5 Then	GerarMalaDiretaFlagrante( strDocEspecialistas, "consulta_DADOS", strDIR & "TempMM") ' DelitoTipo & 4
		If ProgressValue = 6 Then	ConcatenarDocs
		If ProgressValue = 7 Then   AbreFlagrante
		Wait 250
	Next ProgressValue
	GoTo Sair:


Erro:
	Dim sMsgErro As String
	sMsgErro = "Desculpe, ocorreu um erro inesperado!    " & chr$(13) & chr$(13)
	MsgBox ( sMsgErro, 16, "Erro!")
	Goto Sair:

Sair:
	RemoverTempMM
	Exit Sub
	
End Sub


Function GerarMalaDiretaFlagrante(Documento As String, OrigemDados As String, DirDestino As String) As Variant

Dim oMailMerge As Object
	'abrir servico mail merge
	oMailMerge = CreateUnoService("com.sun.star.text.MailMerge")  
	'nome do documento
   	oMailMerge.DocumentURL =  Documento
   	oMailMerge.DataSourceName =  "FLAGRANTES" '' nome do banco de dados registrado 
   	oMailMerge.CommandType = 1  '' 0 = table, 1 = query, 2 = SQL-statement
   	oMailMerge.Command =  OrigemDados
   	oMailMerge.OutputType = 2   '' 1 = Printer, 2 = file, 3 = mail
   	oMailMerge.OutputURL = DirDestino
'    oMailMerge.SaveFilter = "writer_pdf_Export"
'    oMailMerge.SaveFilter = "MS Word 97"
	oMailMerge.SaveFilter = "writer8"
   	oMailMerge.SaveAsSingleFile = TRUE
   	oMailMerge.FileNameFromColumn = TRUE
   	oMailMerge.FileNamePrefix = "DelitoTipo" '' nome da coluna da tabela/consulta
	oMailMerge.execute(Array())
   	oMailMerge.dispose
   
End Function


Sub PreencherTabelaAuxiliar

Rem Macro para listar todos [Conduzido|Vitimas] das tabelas "Conduzidos" e "Vitimas" e preencher a tabela "tblAuxiliar"
Rem autor: Fabio Minami <fminami@gmail.com>

	' Abrir Banco de Dados
	DatabaseContext = createUnoService("com.sun.star.sdb.DatabaseContext")
	' Escolher tabela
	DataSource = DatabaseContext.getByName("FLAGRANTES")	
	Connection = DataSource.GetConnection("","")
	' Conectar SQL
	Statement = Connection.createStatement()
	
	' Consultar o numero de registros - ResultSet
	rs = Statement.executeQuery("SELECT COUNT(Conduzido) FROM Conduzidos;")
	If Not IsNull(rs) Then
		While rs.next			
			' Capturar o numero de registros
			intConduzidos = rs.getString(1)
		Wend
	End if
	rs.close
	rs.dispose()
	
	' Capturar o número do IP (inquérito policial?)
	rs = Statement.executeQuery("SELECT IP FROM Delito;")
	If Not IsNull(rs) Then
		While rs.next			
			strIP = rs.getString(1)
		Wend
	End if
	rs.close
	rs.dispose()

	' Capturar o número do BO
	rs = Statement.executeQuery("SELECT Número_BO FROM Delito;")
	If Not IsNull(rs) Then
		While rs.next			
			strBO = rs.getString(1)
		Wend
	End if
	rs.close
	rs.dispose()

	' Capturar o Tipo de Delito
	rs = Statement.executeQuery("SELECT DelitoTipo FROM Delito;")
	If Not IsNull(rs) Then
		While rs.next			
			strDelitoTipo = rs.getString(1)
		Wend
	End if
	rs.close
	rs.dispose()
	
	' Consultar os [Conduzido] - ResultSet 
	rs = Statement.executeQuery("SELECT Conduzido, RG, Pai, Mãe, Depoimento_Conduzido FROM Conduzidos;")	
	' Dimensionar variavel matriz para capturar os registros
	Dim mtrConduzidos(intConduzidos, 4) As String

	' Capturar ResultSet em uma matriz	
	If Not IsNull(rs) Then
		i = 0
		While rs.next
			' Capturar registros			
			mtrConduzidos(i, 0) = trim(rs.getString(1))
			mtrConduzidos(i, 1) = trim(rs.getString(2))
			mtrConduzidos(i, 2) = trim(rs.getString(3))
			mtrConduzidos(i, 3) = trim(rs.getString(4))
			mtrConduzidos(i, 4) = trim(rs.getString(5))
			i = i + 1					
		Wend
	End If
	
	' Se não existir nenhum registro, sair da macro
	If intConduzidos = 0 Then 
		MsgBox ("Erro. Nenhum registro de ""Conduzidos""! Não há conduzidos registrados.", 48, "Flagrantes OpenOffice.org - Versão BETA")
		Exit Sub
	ElseIf intConduzidos = 1 Then 	' Se exister somente um unico registro...
		'If Not IsNull(rs) Then
		'	While rs.next
		'		' Capturar registro
		'		strConduzidos = trim(rs.getString(1))
		'		strConduzidosRgPais = trim(rs.getString(1)) & " RG " & trim(rs.getString(2)) & ", filho de " & trim(rs.getString(3)) & " e de " & trim(rs.getString(4))
		'		strConduzidosLista = strConduzidos
		'		strConduzidosListaRelatorio = trim(rs.getString(1)) & ", em seu interrogatório, disse que " & trim(rs.getString(5)) & chr$(13) & chr$(13)
		'
		'	Wend
		'End If
		strConduzidos = mtrConduzidos(0, 0)
		strConduzidosRgPais = mtrConduzidos(0, 0) & " RG " & mtrConduzidos(0, 1) & ", filho de " & mtrConduzidos(0, 2) & " e de " & mtrConduzidos(0, 3)
		strConduzidosLista = strConduzidos
		strConduzidosListaRelatorio = mtrConduzidos(0 ,0) & ", em seu interrogatório, disse que " & mtrConduzidos(0, 4) & chr$(13) & chr$(10)
		strMaiusCond = "INDICIADO"
		strCapaCond = strConduzidos
	' Se existir 2 registros, entao...
	ElseIf intConduzidos = 2 Then	
		strConduzidos = mtrConduzidos(0, 0) & " e " & mtrConduzidos(1, 0) '& "."
		strConduzidosRgPais = mtrConduzidos(0, 0) & " RG " & mtrConduzidos(0, 1) & ", filho de " & mtrConduzidos(0, 2) & " e de " & mtrConduzidos(0, 3) & "; e " & mtrConduzidos(1, 0) & ", RG. " & mtrConduzidos(1, 1) & ", filho de " & mtrConduzidos(1, 2) & " e de " & mtrConduzidos(1, 3) 
		strConduzidosLista = mtrConduzidos(0, 0) & "," & chr$(13) & chr$(13) & mtrConduzidos(1, 0) '& chr$(13) & chr$(13)
		' FALTA O SEGUNDO
		strConduzidosListaRelatorio = mtrConduzidos(0, 0) & ", em seu interrogatório, disse que " & mtrConduzidos(0, 4) & chr$(13) & chr$(10) & chr$(9) & chr$(9) & chr$(9) & chr$(9) & mtrConduzidos(1, 0) & ", em seu interrogatório, disse que " & mtrConduzidos(1, 4) & chr$(13) & chr$(10)	 
		strMaiusCond = "INDICIADOS"
		strCapaCond = mtrConduzidos(0, 0) & " E OUTRO"
	' Se existir 3 ou mais registros, entao...
	Else 'if intConduzidos >= 3 Then
	'	n = 0
	'	Do until n  = intConduzidos - 1
		For n = 0 To intConduzidos - 2
			' Formatar linha com registros capturados, passo 1
			strConduzidos = strConduzidos & mtrConduzidos(n, 0) & ", "
			strConduzidosLista = strConduzidosLista & mtrConduzidos(n, 0) & "," & chr$(13) & chr$(10)
			strConduzidosListaRelatorio = strConduzidosListaRelatorio & chr$(9) & chr$(9) & chr$(9) & chr$(9) & mtrConduzidos(n, 0) & ", em seu interrogatório, disse que " & mtrConduzidos(n, 4) & chr$(13) & chr$(10)
			strConduzidosRgPais = strConduzidosRgPais & mtrConduzidos(n, 0) & " RG " & mtrConduzidos(n, 1) & ", filho de " & mtrConduzidos(n, 2) & " e de " & mtrConduzidos(n, 3) & "; "
			'n = n + 1
		Next n
		'Loop
		' Formatar linha com registros capturados, passo 2
		strConduzidos = Left(strConduzidos, Len(strConduzidos) - 2)
		strConduzidos = strConduzidos & " e " & mtrConduzidos(intConduzidos - 1, 0) ' & "."
		'strConduzidosRgPais = Left(strConduzidosRgPais, Len(strConduzidosRgPais) - 2)	
		strConduzidosRgPais = strConduzidosRgPais & "e " & mtrConduzidos(intConduzidos - 1, 0) & " RG " & mtrConduzidos(intConduzidos - 1, 1) & ", filho de " & mtrConduzidos(intConduzidos - 1, 2) & " e de " & mtrConduzidos(intConduzidos - 1, 3) 
		strConduzidosLista = strConduzidosLista & mtrConduzidos(intConduzidos - 1, 0) '& chr$(13) & chr$(10)
		strConduzidosListaRelatorio = Mid(strConduzidosListaRelatorio, 5) & chr$(9) & chr$(9) & chr$(9) & chr$(9) & mtrConduzidos(intConduzidos - 1, 0) & ", em seu interrogatório, disse que " & mtrConduzidos(intConduzidos - 1, 4) & chr$(13) & chr$(10)
		strMaiusCond = "INDICIADOS"
		strCapaCond = mtrConduzidos(0, 0) & " E OUTROS"
	End if	
	rs.close
	rs.dispose()

	' Formatar linha Cúmplices da página "Nota de Culpa"
	' atualizar o campo "cumplices" da tabela "Conduzidos" 
	' se houver algum nome("jose da silva") dentro de outro nome("jose da silva sauro") provavelmente
	' o nome("jose da silva sauro") tambem sera apagado e vai sobrar somento o nome("sauro")
	Dim sCumplices As String
	If intConduzidos = 1 Then
		rs = Statement.executeQuery("UPDATE ""Conduzidos"" SET ""Cumplices"" = '' WHERE ""Conduzido"" = '" & mtrConduzidos (i, 0) & "';")
		rs.close
		rs.dispose()	
	ElseIf intConduzidos = 2 Then
		For i = 0 To UBound(mtrConduzidos) - 1
			sCumplices = Replace(strConduzidos, " e ", " ")
			sCumplices = Replace(sCumplices, mtrConduzidos(i, 0), "")
			sCumplices = ", juntamente com " & Trim(sCumplices)
			rs = Statement.executeQuery("UPDATE ""Conduzidos"" SET ""Cumplices"" = '" & sCumplices  & "' WHERE ""Conduzido"" = '" & mtrConduzidos (i, 0) & "';")
			rs.close
			rs.dispose()
		Next i
	Elseif intConduzidos > 2 Then		
		For i = 0 To UBound(mtrConduzidos) - 1
			If i = intConduzidos - 2 Then
				sCumplices = Replace(strConduzidos, mtrConduzidos(i, 0), "")
				For n = Len( sCumplices ) To 1 Step -1
					If Mid( sCumplices, n, 1 ) = "," Then Exit For
				Next n   
				sCumplices = Left( sCumplices, n - 1 ) & Mid( sCumplices, n + 2)
			ElseIf i = intConduzidos - 1 Then
				sCumplices = Replace(strConduzidos, "e " & mtrConduzidos(i, 0), "")
				For n = Len( sCumplices ) To 1 Step -1
					If Mid( sCumplices, n, 1 ) = "," Then Exit For
				Next n   
				sCumplices = Left( sCumplices, n - 1 ) & " e " & Trim(Mid( sCumplices, n + 1))
			Else				
				sCumplices = Replace(strConduzidos, mtrConduzidos(i, 0) & ", ", "")
			End if	
			sCumplices = ", juntamente com " & sCumplices
			rs = Statement.executeQuery("UPDATE ""Conduzidos"" SET ""Cumplices"" = '" & sCumplices  & "' WHERE ""Conduzido"" = '" & mtrConduzidos (i, 0) & "';")
			rs.close
			rs.dispose()
		Next i
	End if
		
	' Vitimas	
	rs = Statement.executeQuery("SELECT COUNT(Nome_Vítima) FROM Vitimas;")
	If Not IsNull(rs) Then
		While rs.next			
			' Capturar o numero de registros
			intVitimas = rs.getString(1)
		Wend
	End if
	
	' Dimensionar variavel matriz para capturar os registros
	Dim mtrVitimas(intVitimas, 1) As String
	
	' Consultar as [Vitimas] - rs 
	rs = Statement.executeQuery("SELECT Nome_Vítima, Depoimento_Vitima FROM Vitimas;")	
	If Not IsNull(rs) Then
		i = 0
		While rs.next
			' Capturar registros			
			mtrVitimas(i, 0) = trim(rs.getString(1))
			mtrVitimas(i, 1) = trim(rs.getString(2))
			i = i + 1
		Wend
	End If	
	
	If intVitimas = 0 Then 
	'	MsgBox("Erro. Nenhum registro de ""Vitimas""! Não há vítimas registradas.", 48, "Flagrantes OpenOffice.org - Versão BETA")
	'	Exit sub	
	ElseIf intVitimas = 1 Then
		strVitimas = mtrVitimas(0, 0)
		strVitimasLista = strVitimas
		strVitimasListaRelatorio = mtrVitimas(0, 0) & ", ratificou o BO " & strBO & " e prestou declarações dizendo que " & mtrVitimas(0, 1) & chr$(13) & chr$(10) 
		strMaiusVitim = "VÍTIMA:"
		strCapaVitim = strVitimas	
	' Se existir 2 registros, entao...
	ElseIf intVitimas = 2 Then	
		strVitimas = mtrVitimas(0, 0) & " e " & mtrVitimas(1, 0) '& "."
		strVitimasLista = mtrVitimas(0, 0) & "," & chr$(13) & chr$(10) & mtrVitimas(1, 0) '& chr$(13) & chr$(10)
		' FALTA O SEGUNDO
		strVitimasListaRelatorio = mtrVitimas(0, 0) & ", ratificou o BO " & strBO & " e prestou declarações dizendo que " & mtrVitimas(0, 1) & chr$(13) & chr$(10) & chr$(9) & chr$(9) & chr$(9) & chr$(9) & mtrVitimas(1, 0) & ", ratificou o BO " & strBO & " e prestou declarações dizendo que " & mtrVitimas(1, 1) & chr$(13) & chr$(10)
		strMaiusVitim = "VÍTIMAS:"
		strCapaVitim = mtrVitimas(0, 0) & " E OUTRO"
	' Se existir 3 ou mais registros, entao...
	Else ' >= 3 Then
		'n = 0
		'Do until n  = intVitimas - 1
		For n = 0 To intVitimas - 2
			' Formatar linha com registros capturados, passo 1
			strVitimas = strVitimas & mtrVitimas(n, 0) & ", "
			strVitimasLista = strVitimasLista & mtrVitimas(n, 0) & "," & chr$(13) & chr$(10)
			strVitimasListaRelatorio = strVitimasListaRelatorio & chr$(9) & chr$(9) & chr$(9) & chr$(9) & mtrVitimas(n, 0) & ", ratificou o BO " & strBO & " e prestou declarações dizendo que " & mtrVitimas(n, 1) & chr$(13) & chr$(10)
			'n = n + 1
		Next n
		'Loop
		' Formatar linha com registros capturados, passo 2
		strVitimas = Left(strConduzidos, Len(strConduzidos) - 2)
		strVitimas = strVitimas & " e " & mtrVitimas(intVitimas - 1, 0) ' & "."
		strVitimasLista = strVitimasLista & mtrVitimas(intVitimas - 1, 0) '& chr$(13) & chr$(10)
		strVitimasListaRelatorio = Mid(strVitimasListaRelatorio, 5) & chr$(9) & chr$(9) & chr$(9) & chr$(9) & mtrVitimas(intVitimas - 1, 0) & ", ratificou o BO " & strBO & " e prestou declarações dizendo que " & mtrVitimas(intVitimas - 1, 1) & chr$(13) & chr$(10)
		strMaiusVitim = "VÍTIMAS:"
		strCapaVitim = mtrVitimas(0, 0) & " E OUTROS"		
	End if	
	rs.close
	rs.dispose()

	' definir se algumas palavras no texto serão no plural ou no singular
	If intConduzidos = 1 Then 
		S_cond = ""
		EM_cond = ""
		INS_cond = "im"
		L_cond = "l"
		I_cond = "i"
		U_cond = "u"
		ES_cond = " "
		AO_cond = "ão"
	Else
		S_cond = "s"
		EM_cond = "em"
		INS_cond = "ins"
		L_cond = "is"
		I_cond = "ram"
		U_cond = "ram"
		ES_cond = "es"
		AO_cond = "ões"
	End If
	If intVitimas = 1 Then
		S_vit = ""
	Else
		S_vit = "s"
	End if
	'GetDir
	' definir mais variaveis
	' escolher o arquivo especialistas de acordo com o tipo do delito
	' escolher o paragrafo final do relatorio
	If strDelitoTipo = "Roubo" Then 
		strDocDespacho = strDir & "Modelos/FlagranteRouboDespacho.ott"
		strDocEspecialistas = strDir & "Modelos/FlagranteRoubo.ott"
		If intConduzidos = 1 Then
			strRelatorioFim = "Na seqüência, juntou-se aos autos o Boletim de Ocorrência," & _
			" Autos de Exibição e Apreensão, Auto de Avaliação e de Entrega da “res furtiva”," & _
			" Boletim Individual do acusado, Oficio encaminhando a(s) arma(s) ao Instituto" & _
			" de Criminalística e ao Instituto de Identificação comunicando quanto à prisão" & _
			" em flagrante delito do acusado e demais peças."
		Else
			strRelatorioFim = " Na seqüência, juntou-se aos autos o Boletim de Ocorrência," & _
			" Autos de Exibição e Apreensão, Auto de Avaliação e de Entrega da “res furtiva”," & _
			" Boletins Individuais dos acusados, Oficio encaminhando a(s) arma(s) ao Instituto" & _
			" de Criminalística e ao Instituto de Identificação comunicando quanto às prisões" & _
			" em flagrante delito dos acusados e demais peças."
		End If
	ElseIf strDelitoTipo = "Furto" Then
		strDocDespacho = strDir & "Modelos/FlagranteFurtoDespacho.ott"
		strDocEspecialistas = strDir & "Modelos/FlagranteFurto.ott"
		If intConduzidos = 1 Then
			strRelatorioFim = "Na seqüência, juntou-se aos autos o Boletim de Ocorrência," & _
			" Autos de Exibição e Apreensão, Auto de Avaliação e de Entrega da “res furtiva”," & _
			" Boletim Individual do acusado, Oficio ao Instituto de Identificação comunicando" & _
			" quanto à prisão em flagrante delito do acusado e demais peças. "
		Else
			strRelatorioFim = "Na seqüência, juntou-se aos autos o Boletim de Ocorrência," & _
			" Autos de Exibição e Apreensão, Auto de Avaliação e de Entrega da “res furtiva”," & _
			" Boletins Individuais dos acusados, Oficio ao Instituto de Identificação comunicando" & _
			" quanto às prisões em flagrante delito dos acusados e demais peças. "
		End If					
	ElseIf strDelitoTipo = "Armas" Then
		strDocDespacho = strDir & "Modelos/FlagranteArmasDespacho.ott"
		strDocEspecialistas = strDir & "Modelos/FlagranteArmas.ott"
		If intConduzidos = 1 Then
			strRelatorioFim = "Na seqüência, juntou-se aos autos o Boletim de Ocorrência," & _
			" o Auto de Exibição e Apreensão, Auto de Constatação, Boletim Individual," & _
			" Ofícios encaminhando a(s) arma(s) apreendida(s) e comunicando ao Instituto" & _
			" de Identificação, quanto a prisão do acusado, e demais peças."
		Else
			strRelatorioFim = "Na seqüência, juntou-se aos autos o Boletim de Ocorrência," & _
			" o Auto de Exibição e Apreensão, Auto de Constatação, Boletins Individuais," & _
			" Ofícios encaminhando a(s) arma(s) apreendida(s) e comunicando ao Instituto" & _
			" de Identificação, quanto aos Indiciamentos dos acusados, e demais peças."
		End If
	ElseIf strDelitoTipo = "Tóxicos" Then
		strDocDespacho = strDir & "Modelos/FlagranteToxicosDespacho.ott"
		strDocEspecialistas = strDir & "Modelos/FlagranteToxicos.ott"
		If intConduzidos = 1 Then
			strRelatorioFim = "Na seqüência, juntou-se aos autos o Boletim de Ocorrências," & _
			" o Auto de Exibição e Apreensão, Auto de Constatação, Boletim Individual do autuado," & _
			" Ofícios encaminhando a(s) substância(s) apreendida(s) e comunicando ao Instituto" & _
			" de Identificação, quanto ao indiciamento do acusado, e demais peças."
		Else
			strRelatorioFim = "Na seqüência, juntou-se aos autos o Boletim de Ocorrências," & _
			" o Auto de Exibição e Apreensão, Auto de Constatação, Boletins Individuais dos autuados," & _
			" Ofícios encaminhando a(s) substância(s) apreendida(s) e comunicando ao Instituto" & _
			" de Identificação, quanto aos Indiciamentos dos acusados, e demais peças."
		End If
	End If
	
'   Exibir resultadoS
'	msgbox intConduzidos :	msgbox intVitimas :	msgbox strIP : msgbox strBO	
'	msgbox S_cond : msgbox EM_cond : msgbox INS_cond : msgbox L_cond : msgbox I_cond : msgbox U_cond : msgbox ES_cond : msgbox AO_cond	
'	msgbox strConduzidos : msgbox strConduzidosRgPais : msgbox strConduzidosLista :	msgbox strConduzidosListaRelatorio 
'	msgbox strVitimas : msgbox strVitimasLista : msgbox strVitimasListaRelatorio

	rs = Statement.executeQuery ("DELETE FROM ""tblAuxiliar"";")
	rs.close
	rs.dispose()

	PrepStatement = Connection.prepareStatement("INSERT INTO ""tblAuxiliar"" (""Id"", ""txtConduzidos"", ""txtConduzidosRgPais"", ""txtConduzidosLista"", ""txtConduzidosListaRelatorio"", ""txtVitimas"", ""txtVitimasLista"", ""txtVitimasListaRelatorio"", ""S_cond"", ""EM_cond"", ""INS_cond"", ""L_cond"", ""I_cond"", ""U_cond"", ""ES_cond"", ""AO_cond"", ""S_vit"", ""txtRelatorioFim"", ""txtMaiusCond"", ""txtMaiusVitim"", ""txtCapaCond"", ""txtCapaVitim"" ) VALUES ( ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ? );")   
	'xray PrepStatement
	PrepStatement.setInt(1, "1")
	PrepStatement.setString(2, strConduzidos)
	PrepStatement.setString(3, strConduzidosRgPais)
	PrepStatement.setString(4, strConduzidosLista)
	PrepStatement.setString(5, strConduzidosListaRelatorio)
	PrepStatement.setString(6, strVitimas)
	PrepStatement.setString(7, strVitimasLista)
	PrepStatement.setString(8, strVitimasListaRelatorio)
	PrepStatement.setString(9, S_cond)
	PrepStatement.setString(10, EM_cond)
	PrepStatement.setString(11, INS_cond)
	PrepStatement.setString(12, L_cond)
	PrepStatement.setString(13, I_cond)
	PrepStatement.setString(14, U_cond)
	PrepStatement.setString(15, ES_cond)
	PrepStatement.setString(16, AO_cond)
	PrepStatement.setString(17, S_vit)
	PrepStatement.setString(18, strRelatorioFim)	
	PrepStatement.setString(19, strMaiusCond) 
	PrepStatement.setString(20, strMaiusVitim)
	PrepStatement.setString(21, strCapaCond)
	PrepStatement.setString(22, strCapaVitim)
	PrepStatement.execute()	
	PrepStatement.clearParameters()	
	PrepStatement.close
	PrepStatement.dispose()	
'	rs = Statement.executeQuery("CHECKPOINT DEFRAG")
'	rs = Statement.executeQuery("SHUTDOWN COMPACT")
	Connection.close
	Connection.dispose()
	
End Sub

Sub ConcatenarDocs

Dim FileURL As String
Dim Texto As Object
Dim Cursor As Object
Dim BookMark As Object
Dim Bookmarks As Object
Dim BookmarkNames() As String
Dim FileProperties(0) As New com.sun.star.beans.PropertyValue
	FileProperties(0).Name = "Hidden"
	FileProperties(0).Value = True	
	GetDir
	FileURL = strDIR & "TempMM/" & strDelitoTipo & "0.odt"
	Doc = StarDesktop.loadComponentFromURL(FileURL,"_blank", 0, FileProperties())
'	Doc.CurrentController.Frame.ContainerWindow.setVisible(False)	
	BookmarkNames = Array("Doc01Despacho", "Doc02Vitimas", "Doc03Conduzidos", "Doc04Oficios")
	Texto = Doc.getText
	Cursor = Texto.createTextCursor()
	Bookmarks = Doc.Bookmarks	
	For i = 0 To UBound(BookmarkNames)
		Bookmark = Bookmarks.getByName(BookmarkNames(i))
		Cursor.gotoRange(Bookmark.Anchor, False)
		Cursor.InsertDocumentFromURL (strDIR & "TempMM/" & strDelitoTipo & (i + 1) & ".odt", Array())		
	Next i
	Wait 1000
	strTiraBarra = Replace (strIP, "/", "_")	
Dim FileProperties2(0) As New com.sun.star.beans.PropertyValue
	FileProperties2(0).Name = "Overwrite"
	FileProperties2(0).Value = True
	Doc.storeAsURL( strDIR & "Saida/" & strDelitoTipo & "-IP" & strTiraBarra & ".odt", FileProperties2())
	Doc.dispose()
	Doc.Close(True)
	' sem esta pausa o programa quebra quando apaga o dir temporario
	Wait 2000	
	' fechar o doc tres vezes para ter certeza de que nao vai travar (e mesmo assim trava...)
	Doc.Close(True)
	Doc.Close(True)
	Doc.Close(True)		
End Sub

Sub AbreFlagrante
	Wait 1000
	Doc = StarDesktop.loadComponentFromURL(strDIR & "Saida/" & strDelitoTipo & "-IP" & strTiraBarra & ".odt","_blank", 0, Array())
End Sub

Sub GetDIR 
Dim oDoc As Object
	oDoc = ThisComponent
	strURL = oDoc.URL
   	For n = Len( strURL ) To 1 Step -1
    	If Mid( strURL, n, 1 ) = "/" Then Exit For
   	Next n   
   	strDIR = Left( strURL, n )	
   	'msgbox strDIR
End Sub
'criar e apagar diretorio temporario para Mail Merge
Sub CriarTempMM
	GetDIR
	MkDir strDIR & "TempMM"
End Sub

Sub RemoverTempMM
	GetDIR
	RmDir strDIR & "TempMM"	
End Sub
