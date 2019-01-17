#INCLUDE RTF.H

*/=======================================================================================
*/  Класс хариктеристик страницы
*/
Define Class CRtfPage As Custom
	m_nHeight = 0
	m_nWidth  = 0
	m_nMLeft  = 0
	m_nMRight = 0
	m_nMTop   = 0
	m_nMBottom = 0
	m_bLandscape = .F.

	Procedure SetValue
		Lparameters nHeight_, nWidth_, nMLeft_, nMRight_, nMTop_, nMBottom_, bLandscape_
		With (This)
			.m_nHeight    = nHeight_
			.m_nWidth     = nWidth_
			.m_nMLeft     = nMLeft_
			.m_nMRight    = nMRight_
			.m_nMTop      = nMTop_
			.m_nMBottom   = nMBottom_
			.m_bLandscape = bLandscape_
		Endwith

	Endproc

Enddefine

Define Class CRtfFile As Custom
	m_bDOSText  = .F.
	m_bTextMode = .F.
	m_nSubSize  = 0
	m_nSize     = 0
	m_nFont     = 0
	m_nCPage    = 0
	m_nLang     = 0
	m_cStyle    = ""
	m_nAlign    = 0
	m_cTableAlign = ""
	m_cTableFont = ""
	m_nFirstIndent = 0
	m_nLeftIndent = 0

	m_cFileName    = ""   && имя файла
	m_bOpened      = .F.  && открыт ли файл файл в настоящий момент
	m_bInitialized = .F.  && инициализирован ли файл
	m_nFile        = 0    && Дескриптор файла
	m_bParOpen     = .F.
	m_bTextMode    = .F.
	m_bHexCodes    = .F.
	Dimension m_arFonts(100, 2)
	m_nFontsCount = 0
	Dimension m_arColAttr(100, 2)

&& Своиства доля работы с таблицами
	m_nColumnsCount = 100
	Dimension m_arTblValues(100)
	Dimension m_arTblWidths(100)
	Dimension m_arTblAlign(100)

	Dimension aTwoDigitHexArray(256)
	Add Object m_oRtfPage As CRtfPage

*/-------------------------------------------------------------------------------------
*/   Конструктор.
*/
	Procedure Init
		Lparameters cFileName_, bNew_
		Local aHexFF[256], nFirstChar, nSecondChar, cOldErrBlock, bErrorFlag, cFileName, bOk, nCounter, nPos
		This.m_bOpened = .F.
		This.m_bInitialized = .F.

		nCounter = 0
		bOk = .F.
		Do While (nCounter <= 100) .And. (!bOk)
			If (nCounter != 0)
				nPos = At(".RTF", Upper(cFileName_))
				If (nPos != 0)
					cFileName = Substr(cFileName_, 1, nPos-1)+"["+Alltrim(Str(nCounter,3,0))+"].RTF"
				Else
					cFileName = cFileName_+"["+Alltrim(Str(nCounter,3,0))+"].RTF"
				Endif
			Else
				cFileName = cFileName_
			Endif
			If File(cFileName)
				bErrorFlag = .F.
				cOldErrBlock = On("ERROR")
				On Error bErrorFlag = .T.
				Erase &cFileName
				If (!bErrorFlag)
					bOk = .T.
				Endif
				On Error &cOldErrBlock
*        Erase &cFileName_
			Else
				bOk = .T.
			Endif
			nCounter = nCounter+1
		Enddo

		This.m_cFileName = cFileName
		This.m_nFile = Fcreate(cFileName)

		This.m_bOpened = .T.
		This.m_bInitialized = .T.
		This.ClearParams
		This.m_bParOpen = .F.
		This.m_bTextMode = .F.
&& Инициализируем массив шестнадцатиричных значений
		For nFirstChar = 0 To 15
			For nSecondChar = 0 To 15
				aHexFF[nFirstChar*16 + nSecondChar + 1] = ;
					SubStr("0123456789ABCDEF",nFirstChar + 1,1) + ;
					SubStr("0123456789ABCDEF",nSecondChar + 1,1)
			Next nSecondChar
		Next nFirstChar
		=Acopy(aHexFF, This.aTwoDigitHexArray)
	Endproc

	Procedure ClearParams
		With (This)
*      .m_oRtfPage = CreateObject("CRtfPage")
			.m_bHexCodes = .F.
			.DefaultAttr
		Endwith
	Endproc

	Procedure DefaultAttr
		With (This)
			.m_nLang   = CODEPAGE_RUSSION
			.m_nAlign = raLeft
			.m_nSize  = 20
			.m_cStyle = ""
		Endwith
	Endproc

	Procedure DefaultInit
		With (This)
			.ClearParams
			.m_nCPage = 1251
			.m_nLang  = CODEPAGE_RUSSION
			.m_arFonts(1, 1) = rfgRoman
			.m_arFonts(1, 2) = 'Times New Roman Cyr'
			.m_arFonts(2, 1) = rfgModern
			.m_arFonts(2, 2) = 'Courier New'
			.m_arFonts(3, 1) = rfgRoman
			.m_arFonts(3, 2) = 'Arial Cyr'
			.m_arFonts(4, 1) = rfgRoman
			.m_arFonts(4, 2) = 'Tahoma'
			.m_arFonts(5, 1) = rfgRoman
			.m_arFonts(5, 2) = 'Verdana'
			.m_arFonts(6, 1) = rfgRoman
			.m_arFonts(6, 2) = 'Symbol'
			.m_nFontsCount = 6
			.PageA4
		Endwith
	Endproc

	Procedure PageA4
		With (This)
			.PageSetup(.Twips(29.7), .Twips(21), .Twips(1.5), .Twips(1.5), .Twips(2), .Twips(2), .F.)
		Endwith
	Endproc

	Procedure PageA4LandScape
		With (This)
			.PageSetup(.Twips(21), .Twips(29.7), .Twips(2.5), .Twips(2.5), .Twips(2), .Twips(2), .T.)
		Endwith
	Endproc

	Procedure PageSetup
		Lparameters nHeight_, nWidth_, nMLeft_, nMRight_, nMTop_, nMBottom_, bLandscape_
		This.m_oRtfPage.SetValue(nHeight_, nWidth_, nMLeft_, nMRight_, nMTop_, nMBottom_,;
			bLandscape_)
	Endproc

*/--------------------------------------------------
*/  Перевод сантиметров в twip-ы
*/
	Function Twips
		Lparameters nCm_
		Return Int(nCm_ * 1440 / 2.54)

*/--------------------------------------------------
*/  Перевод дюймов в twip-ы
*/
	Function InchTwips
		Lparameters nInch_
		Return Int(nInch_ * 1440)


	Function GroupName
		Lparameters nNumFont_
		Local cResult
		Do Case
			Case nNumFont_ == rfgRoman
				cResult = 'roman'
			Case nNumFont_ == rfgDecor
				cResult = 'decor'
			Case nNumFont_ == rfgTech
				cResult = 'tech'
			Case nNumFont_ == rfgScript
				cResult = 'script'
			Case nNumFont_ == rfgSwiss
				cResult = 'swiss'
			Case nNumFont_ ==  rfgModern
				cResult = 'modern'
			Otherwise
				cResult = 'nil'
		Endcase
		Return cResult

	Function WriteHeader
		Local i
		With (This)
			If .m_bTextMode
				Return
			Endif
			=Fputs(.m_nFile, '{\rtf1\ansi\ansicpg' + Alltrim(Str(.m_nCPage, 4, 0)) + '\deflang' +;
				AllTrim(Str(.m_nLang)))
			If .m_nFontsCount>0
				=Fputs(.m_nFile, '{\fonttbl{')
				For i = 1 To .m_nFontsCount
					=Fputs(.m_nFile, '\f' + Alltrim(Str(i, 2, 0)) + '\f' + ;
						.GroupName(.m_arFonts(i, 1)) + ' ' + .m_arFonts(i, 2)+';')
				Next i
				=Fputs(.m_nFile, '}}')
			Endif
			=Fwrite(.m_nFile, '\paperw' + Alltrim(Str(.m_oRtfPage.m_nWidth, 10, 0)) + '\paperh' +;
				AllTrim(Str(.m_oRtfPage.m_nHeight, 10, 0))+;
				'\margl' + Alltrim(Str(.m_oRtfPage.m_nMLeft, 10, 0)) + '\margr' +;
				AllTrim(Str(.m_oRtfPage.m_nMRight, 10, 0)) +;
				'\margt'+Alltrim(Str(.m_oRtfPage.m_nMTop, 10, 0)) + '\margb' +;
				AllTrim(Str(.m_oRtfPage.m_nMBottom, 10, 0)))
			=Fputs(.m_nFile, Iif(.m_oRtfPage.m_bLandscape, "\landscape", ""))
		Endwith
		Return

	Procedure BeginParagraph
		Lparameters nFirstIndent_, nLeftIndent_, nAl_
		With (This)
			If !(.m_bTextMode)
				If .m_bParOpen
					.EndParagraph
				Endif
				=Fputs(.m_nFile, "")
				.m_nAlign = nAl_
				=Fwrite(.m_nFile, '{' + .AlignConvert(.m_nAlign))
				If (nFirstIndent_ != 0)
					=Fwrite(.m_nFile, '\fi' + Alltrim(Str(nFirstIndent_)))
				Endif
				If (nLeftIndent_ != 0)
					=Fwrite(.m_nFile, '\li' + Alltrim(Str(nLeftIndent_)))
				Endif
			Endif
			=Fputs(.m_nFile, "")
			.m_bParOpen = .T.
		Endwith
	Endproc

	Procedure EndParagraph
		With (This)
			If !(.m_bTextMode)
				If .m_bParOpen
					=Fputs(.m_nFile, "\par}")
					.m_bParOpen = .F.
				Endif
			Endif
		Endwith
	Endproc


	Procedure WriteTag
		Lparameters cString_
		If !This.m_bTextMode
			=Fputs(This.m_nFile, This.Convert(cString_))
		Endif
	Endproc

	Procedure WriteString
		Lparameters cString_
		=Fputs(This.m_nFile, This.Convert(cString_))
	Endproc

	Procedure WriteLine
		Lparameters cString_
		=Fputs(This.m_nFile, This.Convert(cString_)+'\line')
	Endproc

	Procedure SetAlignment
		Lparameters nAlign_
		.m_nAlign = nAlign_
	Endproc

	Procedure SetTableAlignment
		Lparameters nAlign_
		.m_cTableAlign = .AlignConvert(nAlign_)
	Endproc

	Function AlignConvert
		Lparameters nAlign_
		Local cResult
		cResult = ''
		Do Case
			Case nAlign_ == raLeft
				cResult = '\ql'
			Case nAlign_ == raRight
				cResult = '\qr'
			Case nAlign_ == raCenter
				cResult = '\qc'
			Case nAlign_ == raJustify
				cResult = '\qj'
		Endcase
		Return cResult

	Function Dec2Hex
		Lparameter nDecimal
		Return This.aTwoDigitHexArray(nDecimal + 1)

	Function Convert
		Lparameters cString_
		Local i, cResult
		With (This)
			If .m_bTextMode
				cResult = cString_
				If .m_bDOSText
					cResult = Ansitooem(Result)
				Endif
				Return cResult
			Endif
			cResult = ""
			For i = 1 To Len(cString_)
				cChar = Substr(cString_, i, 1)
				Do Case
					Case Asc(cChar) >= 192
						If .m_bHexCodes
							cResult = cResult + "\'" + Dec2Hex(Asc(cChar))
						Else
							cResult = cResult + cChar
						Endif
					Case cChar == "'"
						cResult = cResult+'\rquote '
					Case cChar == "}"
						cResult = cResult+'\}'
					Case cChar == "{"
						cResult = cResult+'\{'
					Case cChar == "\"
						cResult = cResult+'\\'
					Case Asc(cChar) == 9
						cResult = cResult+'\tab'
					Otherwise
						cResult = cResult+cChar
				Endcase
			Next i
		Endwith
		Return cResult

	Procedure NewLine
		If !This.m_bTextMode
			=Fputs(This.m_nFile, '\line')
		Endif
	Endproc

	Procedure NewPage
		If !This.m_bTextMode
			=Fputs(This.m_nFile, '\page')
		Endif
	Endproc


	Procedure SetFont
		Lparameters nNum_, nSize_, cStyle_
		With (This)
			If !(.m_bTextMode)
				=Fputs(.m_nFile, "")
				.m_cStyle = cStyle_
				=Fputs(.m_nFile, .StyleConvert(.m_cStyle))
				If nNum_ <= .m_nFontsCount
					=Fputs(.m_nFile, '\f' + Alltrim(Str(nNum_)))
					.m_nFont = nNum_
				Endif
				.m_nSize = nSize_
				=Fputs(.m_nFile, '\fs' + Alltrim(Str(.m_nSize)))
			Endif
		Endwith
	Endproc

	Procedure SetTableFont
		Lparameters nNum_, nSize_, cStyle_
		With (This)
			.m_cTableFont = .StyleConvert(cStyle_) + '\f' + Alltrim(Str(nNum_)) +;
				'\fs' + Alltrim(Str(.m_nSize))
		Endwith
	Endproc

	Function StyleConvert
		Lparameters cStyle_
		Local cResult, cChar, i
		cResult = ''
		For i = 1 To Len(cStyle_)
			cChar = Substr(cStyle_, i, 1)
			Do Case
				Case cChar == rfsBold
					cResult = cResult + '\b'
				Case cChar == rfsItalic
					cResult = cResult + '\i'
				Case cChar == rfsStrike
					cResult = cResult + '\strike'
				Case cChar == rfsUnderline
					cResult = cResult + '\ul'
				Case cChar == rfsUnderword
					cResult = cResult + '\ulw'
				Case cChar == rfsUnderdot
					cResult = cResult + '\uld'
				Case cChar == rfsUnderdouble
					cResult = cResult + '\uldb'
				Case cChar == rfsSuperScript
					cResult = cResult + '\super'
				Case cChar == rfsSubScript
*           cResult = cResult + '\dn' + AllTrim(Str(.m_nSubSize, 10, 0))
					cResult = cResult + '\sub'
				Case cChar == rfsDefault
					cResult = cResult + '\plain'

			Endcase
		Next i
		Return cResult

	Procedure BeginTable
		With (This)
			If !(.m_bTextMode)
				=Fputs(.m_nFile, "")
				=Fputs(.m_nFile, '{')
			Endif
		Endwith
	Endproc

	Procedure EndTable
		With (This)
			If !(.m_bTextMode)
				If (.m_nColumnsCount > 0)
					=Fputs(.m_nFile, "")
					=Fputs(.m_nFile, '\pard}')
				Endif
			Endif
			.m_nColumnsCount = 0
		Endwith
	Endproc

	Procedure SetColumnsCount
		Lparameters nCount_
		With (This)
			Dimension .m_arTblValues(nCount_)
			Dimension .m_arTblWidths(nCount_)
			Dimension .m_arTblAlign(nCount_)
			.m_nColumnsCount = nCount_
		Endwith
	Endproc

	Procedure SetupColumns
		Lparameters nLeftInd_
		Local nTmp, i, nc
		With (This)
			If !(.m_bTextMode)
				=Fputs(.m_nFile, "")
				=Fputs(.m_nFile, '\trowd')

				If !Empty(nLeftInd_)
					=Fputs(.m_nFile, '\trleft'+Alltrim(Str(nLeftInd_)))
				Endif

				nTmp = 0
				nc = .m_nColumnsCount
				For i = 1 To nc
					nTmp = nTmp + .m_arTblWidths(i)
					=Fputs(.m_nFile, '\clbrdrt\brdrs'+;
						'\clbrdrl\brdrs'+;
						'\clbrdrr\brdrs'+;
						'\clbrdrb\brdrs')
					=Fputs(.m_nFile, '\cellx'+Alltrim(Str(nTmp, 10, 0)))
					.m_arTblAlign(i) = .m_nAlign
					.m_arTblValues(i) = ""
				Next i
				=Fputs(.m_nFile, "")
			Endif
		Endwith
	Endproc

	Procedure WriteRow
		Local i, cString, cFormat, ttt

		With (This)
			cString = Iif(.m_bTextMode, '', '{\cell}')
			If .m_nColumnsCount>0
				=Fputs(.m_nFile, "")
				If !(.m_bTextMode)
					=Fwrite(.m_nFile, '\intbl{')
					=Fputs(.m_nFile, .StyleConvert(.m_cStyle)+'\f'+Alltrim(Str(.m_nFont))+'\fs'+;
						AllTrim(Str(.m_nSize, 10, 0))+.AlignConvert(.m_nAlign))

				Endif
				cText = ""
				For i = 1 To .m_nColumnsCount
					cString = "{"+.AlignConvert(.m_arTblAlign(i))+"\li30\ri30\cell}"
					cText = cText + Iif(i == 1, "", "{") + .m_arTblValues(i) + "}" + cString
				Next i
				=Fwrite(.m_nFile, cText)
				If !(.m_bTextMode)
					=Fputs(.m_nFile, '{\row}')
				Endif
			Endif
		Endwith
	Endproc

	Procedure CloseFile
		With (This)
			If .m_bOpened
				If !.m_bTextMode
					If .m_bParOpen
						=Fputs(.m_nFile, '\par}')
					Endif
					=Fputs(.m_nFile, '')
					=Fputs(.m_nFile, '\par}')
				Endif
				=Fclose(.m_nFile)
				.m_bOpened = .F.
			Endif
		Endwith
	Endproc

	Procedure WriteParagraph
		Lparameters cText_, nAlign_, cFontStyle_, nFirstIndent_, nIndent_, nFont_, nFontSize_
		With (This)
			nAlign_ = Iif(nAlign_ >= 0, nAlign_, .m_nAlign)
			cFontStyle_ = Iif(!Empty(cFontStyle_), cFontStyle_, .m_cStyle)
			nFirstIndent_ = Iif(nFirstIndent_ >= 0, .Twips(nFirstIndent_), .m_nFirstIndent)
			nIndent_ = Iif(nIndent_ >= 0, .Twips(nIndent_), .m_nLeftIndent)
			nFont_ = Iif(nFont_ >= 0, nFont_, .m_nFont)
			nFontSize_ = Iif(nFontSize_ >= 0, nFontSize_, .m_nSize)
			.BeginParagraph(nFirstIndent_, nIndent_, nAlign_)
			.SetFont(nFont_, nFontSize_, cFontStyle_)
			.WriteString(cText_)
			.EndParagraph
		Endwith
	Endproc

	Procedure WriteSpecParagraph
		Lparameters cText_, nAlign_, cFontStyle_, nFirstIndent_, nIndent_, nFont_, nFontSize_
		Local nStartPos, nEndPos, cSpecText
		With (This)
			nAlign_ = Iif(nAlign_ >= 0, nAlign_, .m_nAlign)
			cFontStyle_ = Iif(!Empty(cFontStyle_), cFontStyle_, .m_cStyle)
			nFirstIndent_ = Iif(nFirstIndent_ >= 0, .Twips(nFirstIndent_), .m_nFirstIndent)
			nIndent_ = Iif(nIndent_ >= 0, .Twips(nIndent_), .m_nLeftIndent)
			nFont_ = Iif(nFont_ >= 0, nFont_, .m_nFont)
			nFontSize_ = Iif(nFontSize_ >= 0, nFontSize_, .m_nSize)
			.BeginParagraph(nFirstIndent_, nIndent_, nAlign_)
			.SetFont(nFont_, nFontSize_, cFontStyle_)
			If (At("~", cText_) = 0)
				.WriteString(cText_)
			Else
				Do While (At("~", cText_) != 0)
					nStartPos = At("~", cText_)
					nEndPos   = At("~", cText_, 2)
					.WriteTag(Substr(cText_, 1, nStartPos-1))
					cSpecText = Substr(cText_, nStartPos+1, nEndPos-nStartPos-1)
					If (At("@", cSpecText) = 0)
						.SetFont(6, nFontSize_, cFontStyle_)
						.WriteTag(cSpecText)
					Else
						.SetFont(6, nFontSize_, cFontStyle_)
						.WriteTag(Substr(cSpecText, 1, At("@", cSpecText)-1))
						.SetFont(nFont_, nFontSize_, cFontStyle_+rfsSubScript)
						.WriteTag(Substr(cSpecText, At("@", cSpecText)+1))
					Endif
					.SetFont(nFont_, nFontSize_, rfsDefault+cFontStyle_)
					cText_ = Substr(cText_, nEndPos+1)
				Enddo
				.WriteTag(cText_)
			Endif
			.EndParagraph
		Endwith
	Endproc

Enddefine

Function DelSpecChar(cString_)
	Local i, cResult
	cResult = ""
	For i = 1 To Len(cString_)
		nChar = Asc(Substr(cString_, i, 1))
		If (nChar != 13) .And. (nChar != 9) .And. (nChar != 10)
			cResult = cResult + Chr(nChar)
		Else
			cResult = cResult + " "
		Endif
	Next i
	Return cResult


*==============================================================================================
*   Сокращает строку, содержащую число, до последней значащей цифры.
*==============================================================================================
Function DelEndZero(cString_)
	Local nLen, i, cChar
	If ((At(".", cString_) != 0) .Or. (At(",", cString_) != 0))
		nLen = Len(cString_)
		For i = nLen To 1 Step (-1)
			cChar = Substr(cString_, i, 1)
			If (cChar == "0") .Or. (cChar == ".") .Or. (cChar == ",")
				cString_ = Stuff(cString_, i, 1, " ")
				If (cChar == ",") .Or. (cChar == ".")
					Exit
				Endif
			Else
				Exit
			Endif
		Next i
		cResult = Padl(Alltrim(cString_), nLen)
	Else
		cResult = Alltrim(cString_)
	Endif
	Return cResult
