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
  LParameters nHeight_, nWidth_, nMLeft_, nMRight_, nMTop_, nMBottom_, bLandscape_
    With (THIS)
      .m_nHeight    = nHeight_
      .m_nWidth     = nWidth_
      .m_nMLeft     = nMLeft_
      .m_nMRight    = nMRight_
      .m_nMTop      = nMTop_
      .m_nMBottom   = nMBottom_
      .m_bLandscape = bLandscape_
    EndWith  
  
  EndProc  
   
EndDefine



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
  LParameters cFileName_, bNew_
  Local aHexFF[256], nFirstChar, nSecondChar, cOldErrBlock, bErrorFlag, cFileName, bOk, nCounter, nPos
    THIS.m_bOpened = .F.
    THIS.m_bInitialized = .F.
    
    nCounter = 0
    bOk = .F.
    Do While (nCounter <= 100) .And. (!bOk)
      If (nCounter != 0)
        nPos = At(".RTF", Upper(cFileName_))
        If (nPos != 0)
          cFileName = SubStr(cFileName_, 1, nPos-1)+"["+Alltrim(Str(nCounter,3,0))+"].RTF"
        Else
          cFileName = cFileName_+"["+Alltrim(Str(nCounter,3,0))+"].RTF"
        EndIf
      Else
        cFileName = cFileName_
      EndIf
      If File(cFileName)
        bErrorFlag = .F.
        cOldErrBlock = On("ERROR")
        On Error bErrorFlag = .T.
        Erase &cFileName
        If (!bErrorFlag)
          bOk = .T.
        EndIf
        On Error &cOldErrBlock
*        Erase &cFileName_
      Else
        bOk = .T.
      EndIf
      nCounter = nCounter+1
    EndDo
    
    THIS.m_cFileName = cFileName
    THIS.m_nFile = FCreate(cFileName)

    THIS.m_bOpened = .T.
    THIS.m_bInitialized = .T.
    THIS.ClearParams
    THIS.m_bParOpen = .F.
    THIS.m_bTextMode = .F.
    && Инициализируем массив шестнадцатиричных значений
    For nFirstChar = 0 To 15
      For nSecondChar = 0 To 15
 		    aHexFF[nFirstChar*16 + nSecondChar + 1] = ;
		  	SubStr("0123456789ABCDEF",nFirstChar + 1,1) + ;
	  		SubStr("0123456789ABCDEF",nSecondChar + 1,1)
   	  Next nSecondChar
    Next nFirstChar
    =ACopy(aHexFF, THIS.aTwoDigitHexArray)
  EndProc
   
  Procedure ClearParams
    With (THIS)
*      .m_oRtfPage = CreateObject("CRtfPage")
      .m_bHexCodes = .F.
      .DefaultAttr
    EndWith  
  EndProc

  Procedure DefaultAttr
    With (THIS)
      .m_nLang   = CODEPAGE_RUSSION
      .m_nAlign = raLeft
      .m_nSize  = 20
      .m_cStyle = ""
    EndWith  
  EndProc

  Procedure DefaultInit
    With (THIS)
      .ClearParams
      .m_nCpage = 1251
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
    EndWith   
  EndProc

  Procedure PageA4
    With (THIS)
      .PageSetup(.Twips(29.7), .Twips(21), .Twips(1.5), .Twips(1.5), .Twips(2), .Twips(2), .F.)
    EndWith
  EndProc   

  Procedure PageA4LandScape
    With (THIS)
      .PageSetup(.Twips(21), .Twips(29.7), .Twips(2.5), .Twips(2.5), .Twips(2), .Twips(2), .T.)
    EndWith
  EndProc   
  
  Procedure PageSetup
  LParameters nHeight_, nWidth_, nMLeft_, nMRight_, nMTop_, nMBottom_, bLandscape_
     THIS.m_oRtfPage.SetValue(nHeight_, nWidth_, nMLeft_, nMRight_, nMTop_, nMBottom_,;
                              bLandscape_)
  EndProc

  */--------------------------------------------------
  */  Перевод сантиметров в twip-ы
  */
  Function Twips
  LParameters nCm_
  Return Int(nCm_ * 1440 / 2.54)
  
  */--------------------------------------------------
  */  Перевод дюймов в twip-ы
  */
  Function InchTwips
  LParameters nInch_
  Return Int(nInch_ * 1440)


  Function GroupName
  LParameters nNumFont_
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
     EndCase
  Return cResult 

  Function WriteHeader
  Local i
    With (THIS)
      If .m_bTextMode
        Return
      EndIf  
      =FPuts(.m_nFile, '{\rtf1\ansi\ansicpg' + AllTrim(Str(.m_nCPage, 4, 0)) + '\deflang' +;
                      AllTrim(Str(.m_nLang)))
      If .m_nFontsCount>0
        =FPuts(.m_nFile, '{\fonttbl{')
        For i = 1 to .m_nFontsCount
          =FPuts(.m_nFile, '\f' + AllTrim(Str(i, 2, 0)) + '\f' + ;
                          .GroupName(.m_arFonts(i, 1)) + ' ' + .m_arFonts(i, 2)+';')
        Next i       
        =FPuts(.m_nFile, '}}')
      EndIf
      =FWrite(.m_nFile, '\paperw' + AllTrim(Str(.m_oRtfPage.m_nWidth, 10, 0)) + '\paperh' +;
                AllTrim(Str(.m_oRtfPage.m_nHeight, 10, 0))+;
                '\margl' + AllTrim(Str(.m_oRtfPage.m_nMLeft, 10, 0)) + '\margr' +;
                AllTrim(Str(.m_oRtfPage.m_nMRight, 10, 0)) +;
                '\margt'+AllTrim(Str(.m_oRtfPage.m_nMTop, 10, 0)) + '\margb' +;
                AllTrim(Str(.m_oRtfPage.m_nMBottom, 10, 0)))
      =FPuts(.m_nFile, IIF(.m_oRtfPage.m_bLandscape, "\landscape", ""))
    EndWith
  Return

  Procedure BeginParagraph
  LParameters nFirstIndent_, nLeftIndent_, nAl_
    With (THIS)
      If !(.m_bTextMode)
        If .m_bParOpen
          .EndParagraph
        EndIf  
        =FPuts(.m_nFile, "")
        .m_nAlign = nAl_
        =FWrite(.m_nFile, '{' + .AlignConvert(.m_nAlign))
        If (nFirstIndent_ != 0)
          =FWrite(.m_nFile, '\fi' + AllTrim(Str(nFirstIndent_)))
        EndIf  
        If (nLeftIndent_ != 0)
          =FWrite(.m_nFile, '\li' + AllTrim(Str(nLeftIndent_)))
        EndIf  
      EndIf
      =FPuts(.m_nFile, "")
      .m_bParOpen = .T.
    EndWith 
  EndProc

  Procedure EndParagraph
    With (THIS)
      If !(.m_bTextMode)
        If .m_bParOpen
          =FPuts(.m_nFile, "\par}")
          .m_bParOpen = .F.
       EndIf
     EndIf
   EndWith  
 EndProc


  Procedure WriteTag
  LParameters cString_
    If !THIS.m_bTextMode
      =FPuts(THIS.m_nFile, THIS.Convert(cString_))
    EndIf  
  EndProc

  Procedure WriteString
  LParameters cString_
    =FPuts(THIS.m_nFile, THIS.Convert(cString_))
  EndProc
  
  Procedure WriteLine
  LParameters cString_
    =FPuts(THIS.m_nFile, THIS.Convert(cString_)+'\line')
  EndProc

  Procedure SetAlignment
  LParameters nAlign_
     .m_nAlign = nAlign_
  EndProc
 
  Procedure SetTableAlignment
  LParameters nAlign_
     .m_cTableAlign = .AlignConvert(nAlign_)
  EndProc

 Function AlignConvert
 LParameters nAlign_
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
   EndCase
  Return cResult

  Function Dec2Hex
  LPARAMETER nDecimal
  Return THIS.aTwoDigitHexArray(nDecimal + 1)

  Function Convert
  LParameters cString_
  Local i, cResult
    With (THIS)
      If .m_bTextMode
         cResult = cString_
         If .m_bDOSText 
           cResult = AnsiToOem(Result)
         EndIf  
         Return cResult
      EndIf
      cResult = ""
      For i = 1 To Len(cString_) 
        cChar = SubStr(cString_, i, 1)
        Do Case 
          Case Asc(cChar) >= 192
                 If .m_bHexCodes 
                   cResult = cResult + "\'" + Dec2Hex(Asc(cChar))
                 Else
                   cResult = cResult + cChar
                 EndIf  
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
          OtherWise  
            cResult = cResult+cChar
        EndCase  
      Next i
    EndWith 
  Return cResult

  Procedure NewLine
    If !THIS.m_bTextMode 
      =FPuts(THIS.m_nFile, '\line')
    EndIf  
  EndProc

  Procedure NewPage
    If !THIS.m_bTextMode 
      =FPuts(THIS.m_nFile, '\page')
    EndIf  
  EndProc


  Procedure SetFont
  LParameters nNum_, nSize_, cStyle_
    With (THIS)
      If !(.m_bTextMode)
        =FPuts(.m_nFile, "")
        .m_cStyle = cStyle_
        =FPuts(.m_nFile, .StyleConvert(.m_cStyle))
        If nNum_ <= .m_nFontsCount
           =FPuts(.m_nFile, '\f' + AllTrim(Str(nNum_)))
           .m_nFont = nNum_
        EndIf
        .m_nSize = nSize_
        =FPuts(.m_nFile, '\fs' + AllTrim(Str(.m_nSize)))
      EndIf
    EndWith 
  EndProc

  Procedure SetTableFont
  LParameters nNum_, nSize_, cStyle_
    With (THIS)
        .m_cTableFont = .StyleConvert(cStyle_) + '\f' + AllTrim(Str(nNum_)) +;
                        '\fs' + AllTrim(Str(.m_nSize))
    EndWith 
  EndProc

  Function StyleConvert
  LParameters cStyle_
  Local cResult, cChar, i
     cResult = ''
     For i = 1 To Len(cStyle_)
       cChar = SubStr(cStyle_, i, 1)
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
      
       EndCase
     Next i  
  Return cResult

  Procedure BeginTable
    With (THIS)
      If !(.m_bTextMode)
        =FPuts(.m_nFile, "")
        =FPuts(.m_nFile, '{')
      EndIf
    EndWith 
  EndProc

  Procedure EndTable
    With (THIS)
      If !(.m_bTextMode) 
        If (.m_nColumnsCount > 0) 
          =FPuts(.m_nFile, "")
          =FPuts(.m_nFile, '\pard}')
        EndIf
      EndIf
      .m_nColumnsCount = 0
    EndWith 
  EndProc

  Procedure SetColumnsCount
  LParameters nCount_
    With (THIS)
      Dimension .m_arTblValues(nCount_)
      Dimension .m_arTblWidths(nCount_)
      Dimension .m_arTblAlign(nCount_)
      .m_nColumnsCount = nCount_
    EndWith
  EndProc

  Procedure SetupColumns
  LParameters nLeftInd_
  Local nTmp, i, nc
    With (THIS)
      If !(.m_bTextMode) 
        =FPuts(.m_nFile, "")
        =FPuts(.m_nFile, '\trowd')

        If !Empty(nLeftInd_)
           =FPuts(.m_nFile, '\trleft'+AllTrim(Str(nLeftInd_)))
        EndIf

        nTmp = 0
        nc = .m_nColumnsCount
        For i = 1 To nc
          nTmp = nTmp + .m_arTblWidths(i)
          =FPuts(.m_nFile, '\clbrdrt\brdrs'+;
                           '\clbrdrl\brdrs'+;
                           '\clbrdrr\brdrs'+;
                           '\clbrdrb\brdrs')
          =FPuts(.m_nFile, '\cellx'+AllTrim(Str(nTmp, 10, 0)))
          .m_arTblAlign(i) = .m_nAlign
          .m_arTblValues(i) = ""
        Next i
        =FPuts(.m_nFile, "")
      EndIf
    EndWith 
  EndProc

  Procedure WriteRow
  Local i, cString, cFormat, ttt
    
    With (THIS)
     cString = IIF(.m_bTextMode, '', '{\cell}')
     If .m_nColumnsCount>0
        =FPuts(.m_nFile, "")
        If !(.m_bTextMode)
           =FWrite(.m_nFile, '\intbl{')
           =FPuts(.m_nFile, .StyleConvert(.m_cStyle)+'\f'+AllTrim(Str(.m_nFont))+'\fs'+;
                           AllTrim(Str(.m_nSize, 10, 0))+.AlignConvert(.m_nAlign))

        EndIf
        cText = ""
        For i = 1 To .m_nColumnsCount
          cString = "{"+.AlignConvert(.m_arTblAlign(i))+"\li30\ri30\cell}"
          cText = cText + IIF(i == 1, "", "{") + .m_arTblValues(i) + "}" + cString
        Next i    
        =FWrite(.m_nFile, cText) 
        If !(.m_bTextMode)
          =FPuts(.m_nFile, '{\row}')
        EndIf  
      EndIf
    EndWith 
  EndProc

  Procedure CloseFile
    With (THIS)
      If .m_bOpened
        If !.m_bTextMode
           If .m_bParopen 
             =FPuts(.m_nFile, '\par}')
           EndIf
           =FPuts(.m_nFile, '')
           =FPuts(.m_nFile, '\par}')
        EndIf
        =FClose(.m_nFile)
        .m_bOpened = .F.
      EndIf
    EndWith 
  EndProc

  Procedure WriteParagraph
  LParameters cText_, nAlign_, cFontStyle_, nFirstIndent_, nIndent_, nFont_, nFontSize_
    With (THIS)
      nAlign_ = IIF(nAlign_ >= 0, nAlign_, .m_nAlign)  
      cFontStyle_ = IIF(!Empty(cFontStyle_), cFontStyle_, .m_cStyle) 
      nFirstIndent_ = IIF(nFirstIndent_ >= 0, .Twips(nFirstIndent_), .m_nFirstIndent)
      nIndent_ = IIF(nIndent_ >= 0, .Twips(nIndent_), .m_nLeftIndent)
      nFont_ = IIF(nFont_ >= 0, nFont_, .m_nFont)
      nFontSize_ = IIF(nFontSize_ >= 0, nFontSize_, .m_nSize)
      .BeginParagraph(nFirstIndent_, nIndent_, nAlign_)
      .SetFont(nFont_, nFontSize_, cFontStyle_)
      .WriteString(cText_)
      .EndParagraph
    EndWith
  EndProc

  Procedure WriteSpecParagraph
  LParameters cText_, nAlign_, cFontStyle_, nFirstIndent_, nIndent_, nFont_, nFontSize_
  Local nStartPos, nEndPos, cSpecText
    With (THIS)
      nAlign_ = IIF(nAlign_ >= 0, nAlign_, .m_nAlign)  
      cFontStyle_ = IIF(!Empty(cFontStyle_), cFontStyle_, .m_cStyle) 
      nFirstIndent_ = IIF(nFirstIndent_ >= 0, .Twips(nFirstIndent_), .m_nFirstIndent)
      nIndent_ = IIF(nIndent_ >= 0, .Twips(nIndent_), .m_nLeftIndent)
      nFont_ = IIF(nFont_ >= 0, nFont_, .m_nFont)
      nFontSize_ = IIF(nFontSize_ >= 0, nFontSize_, .m_nSize)
      .BeginParagraph(nFirstIndent_, nIndent_, nAlign_)
      .SetFont(nFont_, nFontSize_, cFontStyle_)
      If (At("~", cText_) = 0)
        .WriteString(cText_)
      Else
        Do While (At("~", cText_) != 0)
          nStartPos = At("~", cText_)
          nEndPos   = At("~", cText_, 2)
          .WriteTag(SubStr(cText_, 1, nStartPos-1))
          cSpecText = SubStr(cText_, nStartPos+1, nEndPos-nStartPos-1)
          If (At("@", cSpecText) = 0)
            .SetFont(6, nFontSize_, cFontStyle_)
            .WriteTag(cSpecText)
          Else
            .SetFont(6, nFontSize_, cFontStyle_)
            .WriteTag(SubStr(cSpecText, 1, At("@", cSpecText)-1))
            .SetFont(nFont_, nFontSize_, cFontStyle_+rfsSubScript)
            .WriteTag(SubStr(cSpecText, At("@", cSpecText)+1))
          EndIf
          .SetFont(nFont_, nFontSize_, rfsDefault+cFontStyle_)
          cText_ = SubStr(cText_, nEndPos+1)
        EndDo
        .WriteTag(cText_)
      EndIf
      .EndParagraph
    EndWith
  EndProc
  
EndDefine

Function DelSpecChar(cString_)
Local i, cResult
  cResult = ""
  For i = 1 To Len(cString_)
    nChar = Asc(SubStr(cString_, i, 1))
    If (nChar != 13) .And. (nChar != 9) .And. (nChar != 10)
      cResult = cResult + Chr(nChar)
    Else  
      cResult = cResult + " "
    EndIf
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
      cChar = SubStr(cString_, i, 1)
      If (cChar == "0") .Or. (cChar == ".") .Or. (cChar == ",")
        cString_ = Stuff(cString_, i, 1, " ")
        If (cChar == ",") .Or. (cChar == ".")
          Exit        
        EndIf
      Else 
        Exit
      EndIf
    Next i
    cResult = PadL(AllTrim(cString_), nLen)
  Else  
    cResult = AllTrim(cString_)
  EndIf
Return cResult
