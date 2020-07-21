Const ForReading = 1 
Const ForWriting = 2 

'Etapa de loop pelos arquivos nas pastas e subpastas
sFolder = "C:\Users\Robvic\Desktop\Resultado"	'Endereço da pasta a ser lida
Set oFSO = CreateObject("Scripting.FileSystemObject")  
For Each oFile In oFSO.GetFolder(sFolder).Files  
	ProcessFile
Next
For Each oSubFolder In oFSO.GetFolder(sFolder).SubFolders  
	For Each oFile In oFSO.GetFolder(oSubFolder).Files  
		ProcessFile
	Next
Next


Sub ProcessFile
'Abertura e leitura do arquivo
	If oFSO.GetExtensionName(oFile.Name) = "txt" Then    'Verificação da extensão desejada do arquivo
		'msgbox(oFSO.GetAbsolutePathName(oFile))	'Linha de debug
		Set openFile = oFSO.OpenTextFile(oFSO.GetAbsolutePathName(oFile), ForReading) 
		strText = openFile.ReadAll
'Substituição dos tags
		tagInicial = "<Tag1>"	'Tag a ser substituída
		tagFinal = "</Tag1>"
		pos1 = InStr(strText, tagInicial)
		pos2 = InStr(strText, tagFinal)
		conteudoTag = Mid(strText, pos1+Len(tagInicial), pos2-(pos1+Len(tagInicial)))
		strNewText = Replace(strText, conteudoTag, "[TEXTO A SUBSTITUIR]") 'Conteúdo fixo que substituirá o original
'Overwrite do arquivo original
		Set openFile = oFSO.OpenTextFile(oFSO.GetAbsolutePathName(oFile), ForWriting) 
		openFile.WriteLine strNewText 
		openFile.Close
	End if  
End Sub