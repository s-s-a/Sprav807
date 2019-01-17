#INCLUDE "RTF.H"
  * Процедура удаления объекта Mutex  
  Procedure CloseMutex  
  LParameters IsExists	&& существует ли другое приложение  
    
 * Если другое приложение существует, то удалять объект Mutex не надо  
 * Удаление выполняется только если объект был создан именно в этом приложении  
  If IsExists = .f.  
 	* Удаление объекта Mutex  
  	Declare integer ReleaseMutex IN Win32API Integer hMutex  
  	ReleaseMutex(m.gnMutex)  
  EndIf  
    
 * Закрытие уже не нужного хендла объекта Mutex  
  Declare integer CloseHandle IN Kernel32 Integer hObject  
  CloseHandle(m.gnMutex)  
    
  EndProc
*-----------------------------------------------------------------------------------------------
 
 * ПОДКЛЮЧЕН ЛИ КОМПЬЮТЕР К ИНТЕРНЕТУ ?  
  FUNCTION IsInternetConnected  
  LOCAL lnFlags AS Integer  
  DECLARE SHORT InternetGetConnectedState IN WININET LONG @, LONG  
  lnFlags = 0  
  InternetGetConnectedState(@lnFlags, 0)  
  CLEAR DLLS 'InternetGetConnectedState'  
  RETURN !INLIST(lnFlags, 0, 16, 32, 48)  
*-----------------------------------------------------------------------------------------------

 * ЗАГРУЗИМ ФАЙЛ И СОХРАНИМ ЕГО ЛОКАЛЬНО 
  FUNCTION IsFileDownloaded  
  LPARAMETERS tcSourceFile AS String, tcTargetFile AS String  
  IF !FILE(tcTargetFile)  
  	DECLARE INTEGER URLDownloadToFile IN URLMON.DLL LONG, STRING, STRING, LONG, LONG  
  	URLDownloadToFile(0, tcSourceFile, tcTargetFile, 0, 0)  
  	CLEAR DLLS 'URLDownloadToFile'  
  	RETURN FILE(tcTargetFile)  
  ENDIF  
  RETURN .F.  
*-----------------------------------------------------------------------------------------------    
 * СООБЩЕНИЕ ОБ ОШИБКЕ  
  PROCEDURE ShowError  
  LPARAMETERS toException AS Exception  
  LOCAL lcErrorNo AS String, lcMessage AS String, lcStackLevel AS String,;  
  	lcProcedure AS String, lcLineNo AS String, lcLineContents AS String  
  TRY  
  	lcErrorNo = 'Номер ошибки' + CHR_TAB + ': ' + TRANSFORM(toException.ErrorNo) + CHR_CR  
  	lcMessage = 'Сообщение' + CHR_TAB + ': ' + toException.Message + CHR_CR  
  	lcStackLevel = 'Уровень стека' + CHR_TAB + ': ' + TRANSFORM(toException.StackLevel) + CHR_CR  
  	lcProcedure = 'Процедура' + CHR_TAB + ': ' + toException.Procedure + CHR_CR  
  	lcLineNo = 'Номер строки' + CHR_TAB + ': ' + TRANSFORM(toException.LineNo)  
  	lcLineContents = IIF(Application.Startmode = 0,;  
  		CHR_CR + 'Содержимое' + CHR_TAB + ': ' + toException.LineContents, '')  
  	MESSAGEBOX(lcErrorNo + lcMessage + lcStackLevel + lcProcedure + lcLineNo + lcLineContents, 16,'Sprav807')  
  CATCH  
  	MESSAGEBOX('Ошибка при попытке вывести сообщение об ошибке', 16, 'Sprav807')
  ENDTRY  
  RETURN  

*-----------------------------------------------------------------------------------------------
* ЕЩЁ ОДНО СООБЩЕНИЕ ОБ ОШИБКЕ (вызывается это)
PROCEDURE errHandler
   PARAMETER merror, mess, mess1, mprog, mlineno
   CLEAR
   err1 = 'Номер ошибки: ' + LTRIM(STR(merror))+ CHR(13)
   err2 = 'Сообщение об ошибке: ' + mess + CHR(13)
   err3 = 'Строка кода с ошибкой: ' + mess1 + CHR(13)
   err4 = 'Номер строки с ошибкой: ' + LTRIM(STR(mlineno)) + CHR(13)
   err5 = 'Программа с ошибкой: ' + mprog + CHR(13)
   MESSAGEBOX(err1 + err2 + err3 + err4 + err5, 16,'Sprav807')  
ENDPROC
*-----------------------------------------------------------------------------------------------
PROCEDURE poisk

*=MESSAGEBOX(_SCREEN.ActiveForm.ActiveControl.Name)
IF UPPER(_SCREEN.ActiveForm.ActiveControl.Name)='GRID2'
 fr_2.Grid1.SetFocus()
ENDIF 

IF UPPER(_SCREEN.ActiveForm.ActiveControl.Name)='GRID1'
 activ_col = _SCREEN.ActiveForm.ActiveControl.ActiveColumn
 
 IF activ_col = 1 && BIC
  vact1 = act_poisk()
  IF !vact1 
   DO FORM w_poisk NAME frm_poisk NOSHOW 
   frm_poisk.Show(1)
  ENDIF  
 ENDIF 
 
 IF activ_col = 2 && NameP
  vact2 = act_poisk2()
  IF !vact2 
   DO FORM w_poisk2 NAME frm_poisk2
   frm_poisk2.Hide 
   frm_poisk2.Show(1)
  ENDIF  
 ENDIF 

 IF activ_col = 13 && UID
  vact3 = act_poisk3()
  IF !vact3
   DO FORM w_poisk3 NAME frm_poisk3 NOSHOW 
   frm_poisk3.Show(1)
  ENDIF  
 ENDIF 
 
 IF activ_col = 16 && Regn
  vact4 = act_poisk4()
  IF !vact4
   DO FORM w_poisk4 NAME frm_poisk4 NOSHOW 
   frm_poisk4.Show(1)
  ENDIF  
 ENDIF 

 IF activ_col = 18 && SWBIC
  vact5 = act_poisk5()
  IF !vact5
   DO FORM w_poisk5 NAME frm_poisk5 NOSHOW 
   frm_poisk5.Show(1)
  ENDIF  
 ENDIF 

 IF activ_col = 4 && Ind
  vact6 = act_poisk6()
  IF !vact6
   DO FORM w_poisk6 NAME frm_poisk6 NOSHOW 
   frm_poisk6.Show(1)
  ENDIF  
 ENDIF 

 
ELSE
* =MESSAGEBOX('',0,'',3000)
ENDIF 
 
RETURN 
*-----------------------------------------------------------------------------------------------
PROCEDURE poisk_men
 HIDE POPUP _3mp

 DO poisk

 DEACTIVATE POPUP _3mp
 RELEASE POPUPS _3mp
RETURN 
*-----------------------------------------------------------------------------------------------
PROCEDURE p1menu

 DEFINE POPUP _3mp FROM y_p_my,x_p_my MARGIN RELATIVE SHADOW FONT 'Arial', 10   && FONT 'Courier New', 10 STYLE 'B'  
 DEFI BAR 1 OF _3mp PROMPT " Просмотр " COLOR SCHEME 3
 DEFI BAR 2 OF _3mp PROMPT " Отбор по фильтру "   COLOR SCHEME 3 
 DEFI BAR 3 OF _3mp PROMPT " Сброс фильтра "   COLOR SCHEME 3 
 DEFI BAR 4 OF _3mp PROMPT " Поиск в таблице БИК "   COLOR SCHEME 3 
 DEFI BAR 5 OF _3mp PROMPT " Копировать значение в буфер обмена "   COLOR SCHEME 3
 DEFI BAR 6 OF _3mp PROMPT " Сравнить с датой "   COLOR SCHEME 3
 DEFI BAR 7 OF _3mp PROMPT " Список клиентов "   COLOR SCHEME 3   
 ON SELEC BAR 1 OF _3mp do p2p
 ON SELEC BAR 2 OF _3mp do p3p 
 ON SELEC BAR 3 OF _3mp do p4p 
 ON SELEC BAR 4 OF _3mp do poisk_men
 ON SELEC BAR 5 OF _3mp do clipmy
 ON SELEC BAR 6 OF _3mp do pcompare 
 ON SELEC BAR 7 OF _3mp do pcallrtf1 && plstl
 ACTIVATE POPUP _3mp 
 RELEASE POPUP _3mp 

RETURN 
*--------------------------------------------------------------------------------------------------
PROCEDURE paccmenu
 DEFINE POPUP _7mp FROM y_p_my,x_p_my MARGIN RELATIVE SHADOW FONT 'Arial', 10   && FONT 'Courier New', 10 STYLE 'B'  
 DEFI BAR 1 OF _7mp PROMPT " Поиск счета в таблице счетов " COLOR SCHEME 3
 DEFI BAR 2 OF _7mp PROMPT " Копировать значение в буфер обмена"   COLOR SCHEME 3 
 ON SELEC BAR 1 OF _7mp do pacc7
 ON SELEC BAR 2 OF _7mp do clipmy2
 ACTIVATE POPUP _7mp 
 RELEASE POPUP _7mp 

RETURN 
*--------------------------------------------------------------------------------------------------
PROCEDURE vs_menu
 DEFINE POPUP _9mp FROM y_q_my,x_q_my MARGIN RELATIVE SHADOW FONT 'Arial', 10   && FONT 'Courier New', 10 STYLE 'B'  
 DEFI BAR 1 OF _9mp PROMPT " Вставить " COLOR SCHEME 3
 ON SELEC BAR 1 OF _9mp do pvs7
 ACTIVATE POPUP _9mp 
 RELEASE POPUP _9mp 

RETURN 
*--------------------------------------------------------------------------------------------------
PROCEDURE pvs7  && вставка из буфера обмена в текстбоксы
 HIDE POPUP _9mp
 _SCREEN.ActiveForm.ActiveControl.Value = _CLIPTEXT
 DEACTIVATE POPUP _9mp
 RELEASE POPUPS _9mp
RETURN 
*--------------------------------------------------------------------------------------------------
PROCEDURE p2p

 HIDE POPUP _3mp

 mya1=My_activate_frm('FORM3')
 IF !mya1
  DO FORM Form3 NAME fr_3 NOSHOW 
  fr_3.Show(1) 
 ENDIF  
  
 DEACTIVATE POPUP _3mp
 RELEASE POPUPS _3mp

RETURN 
*--------------------------------------------------------------------------------------------------
PROCEDURE pacc7
 HIDE POPUP _7mp

 mya1=My_activate_frm('FORM_ACC')
 IF !mya1
  DO FORM w_poisk_acc1 NAME fr_acc7 NOSHOW 
  fr_acc7.Show(1)
 ENDIF  
  
 DEACTIVATE POPUP _7mp
 RELEASE POPUPS _7mp


RETURN 
*--------------------------------------------------------------------------------------------------
PROCEDURE p3p  && установка фильтра
 HIDE POPUP _3mp
 PUSH KEY CLEAR
 
 WAIT CLEAR 
 _vfp.StatusBar=''
 
 mya2=My_activate_frm('FORM4')
 IF !mya2
  DO FORM Form4 NAME fr_4 NOSHOW 
  fr_4.Show(1) 
 ENDIF 
 
 POP KEY 
 WAIT 'Записей БИК = '+ALLTRIM(STR(k_filt, 18)) WINDOW NOWAIT 

 DEACTIVATE POPUP _3mp
 RELEASE POPUPS _3mp
 
RETURN 
*--------------------------------------------------------------------------------------------------
PROCEDURE p4p  && сброс фильтра
 HIDE POPUP _3mp
 SET FILTER TO 
 COUNT TO k_filt 
 GO TOP
 WAIT 'Записей БИК = '+ALLTRIM(STR(k_filt, 18)) WINDOW NOWAIT  && NOCLEAR 
 _vfp.StatusBar='Записей БИК = '+ALLTRIM(STR(k_filt, 18))

 tx1 = ''
 tx2 = ''
 tx3 = ''
 tx4 = ''
 tx5 = ''
 tx6 = ''
 tx7 = ''
 tx8 = ''
 tx9 = ''
 kus4 = ''
 DEACTIVATE POPUP _3mp
 RELEASE POPUPS _3mp 
RETURN 
*--------------------------------------------------------------------------------------------------
FUNCTION My_activate_frm 
LPARAMETERS tcFormName 

IF PCOUNT() > 0 
 IF VARTYPE(tcFormName) = 'C' 

 tcFormName = UPPER(tcFormName) 

 LOCAL lnForCounter 


 FOR lnForCounter = 1 TO _Screen.FormCount 

* WAIT _Screen.Forms(lnForCounter).Name WINDOW 

 IF UPPER(_Screen.Forms(lnForCounter).Name) = tcFormName && Если форма есть в массиве _Screen.Forms()
 
  IF TYPE('_SCREEN.FORMS(lnForCounter).NAME') = 'C' && Если _Screen.ActiveForm в данный момент является объектом и на неё можно ссылаться 


*WAIT tcFormName + STR(lnForCounter ,4)  WINDOW 

   IF UPPER(_SCREEN.FORMS(lnForCounter).NAME) == tcFormName && Если форма-параметр в данный момент активна 
    _SCREEN.FORMS(lnForCounter).Show()
    RETURN .T. 
   ENDIF 
  
  ENDIF 
 ENDIF 
 ENDFOR 


 ENDIF 
ENDIF 

RETURN .F. 
ENDFUNC
*------------------------------------------------------------------------
FUNCTION act_poisk
RETURN My_activate_frm('FORM_BIC')
*------------------------------------------------------------------------
FUNCTION act_poisk2
RETURN My_activate_frm('FORM_NAIM')
*------------------------------------------------------------------------
FUNCTION act_poisk3
RETURN My_activate_frm('FORM_UID')
*------------------------------------------------------------------------
FUNCTION act_poisk4
RETURN My_activate_frm('FORM_Regn')
*------------------------------------------------------------------------
FUNCTION act_poisk5
RETURN My_activate_frm('FORM_SWBIC')
*------------------------------------------------------------------------
FUNCTION act_poisk6
RETURN My_activate_frm('FORM_Ind')
*------------------------------------------------------------------------
FUNCTION act_poisk7
RETURN My_activate_frm('FORM_ACC')
*------------------------------------------------------------------------
PROCEDURE act_poisk_a
IF UPPER(_SCREEN.ActiveForm.ActiveControl.Name)='GRID1'
 fr_2.Grid2.SetFocus()
ENDIF 


  vactacc = act_poisk7()
  IF !vactacc
   DO FORM w_poisk_acc1 NAME frm_poisk_acc NOSHOW 
   frm_poisk_acc.Show(1)
  ENDIF  

RETURN 
*------------------------------------------------------------------------
PROCEDURE clipmy && !!!копирование в буфер обмена нужно делать только в русской раскладке!!!!!
 HIDE POPUP _3mp
 * Константы:
*  #DEFINE KEYBOARD_GERMAN_ST 	0x0407		&& Немецкий (Стандарт)  
  #DEFINE KEYBOARD_ENGLISH_US 	0x0409		&& Английский (Соединенные Штаты)  
*  #DEFINE KEYBOARD_FRENCH_ST 	0x040c		&& Французский (Стандарт)  
  #DEFINE KEYBOARD_RUSSIAN 		0x0419		&& Русский   
   
  lnCurrentKeyboard = GetKeyboardLayout(0)  
 * Считываем младшее слово (младшие 16 бит из 32)  
  lnCurrentKeyboard = BitRShift(m.lnCurrentKeyboard,16)
 
  IF m.lnCurrentKeyboard <> KEYBOARD_RUSSIAN
   =LoadKeyboardLayout("00000419",1) && Рус
  ENDIF 
 
 ccx='fr_2.Grid1.Column'+ALLTRIM(STR(_SCREEN.ActiveForm.ActiveControl.ActiveColumn,18))+'.Text1.Value'
 ccx=ALLTRIM(ccx)
 _CLIPTEXT=&ccx && !!!копирование в буфер обмена нужно делать только в русской раскладке!!!!!
 
 IF m.lnCurrentKeyboard=KEYBOARD_ENGLISH_US 
  =LoadKeyboardLayout("00000409",1) && Eng
 ENDIF 
 
 DEACTIVATE POPUP _3mp
 RELEASE POPUPS _3mp 
RETURN 
*------------------------------------------------------------------------
*------------------------------------------------------------------------
PROCEDURE clipmy2 && !!!копирование в буфер обмена нужно делать только в русской раскладке!!!!!
 HIDE POPUP _7mp
 * Константы:
*  #DEFINE KEYBOARD_GERMAN_ST 	0x0407		&& Немецкий (Стандарт)  
  #DEFINE KEYBOARD_ENGLISH_US 	0x0409		&& Английский (Соединенные Штаты)  
*  #DEFINE KEYBOARD_FRENCH_ST 	0x040c		&& Французский (Стандарт)  
  #DEFINE KEYBOARD_RUSSIAN 		0x0419		&& Русский   
 
  lnCurrentKeyboard = GetKeyboardLayout(0)  
 * Считываем младшее слово (младшие 16 бит из 32)  
  lnCurrentKeyboard = BitRShift(m.lnCurrentKeyboard,16)
 
  IF m.lnCurrentKeyboard <> KEYBOARD_RUSSIAN
   =LoadKeyboardLayout("00000419",1) && Рус
  ENDIF 
  
 ccx='fr_2.Grid2.Column'+ALLTRIM(STR(_SCREEN.ActiveForm.ActiveControl.ActiveColumn,18))+'.Text1.Value'
 ccx=ALLTRIM(ccx)
 _CLIPTEXT=&ccx && !!!копирование в буфер обмена нужно делать только в русской раскладке!!!!!
 
 IF m.lnCurrentKeyboard=KEYBOARD_ENGLISH_US 
  =LoadKeyboardLayout("00000409",1) && Eng
 ENDIF 
 
 DEACTIVATE POPUP _7mp
 RELEASE POPUPS _7mp 
RETURN 
*------------------------------------------------------------------------
PROCEDURE pimenu1
 DEFINE POPUP _1mq FROM y_i_my,x_i_my MARGIN RELATIVE SHADOW FONT 'Arial', 10   && FONT 'Courier New', 10 STYLE 'B'  
 DEFI BAR 1 OF _1mq PROMPT " Вывод в текстовый файл " COLOR SCHEME 3
 ON SELEC BAR 1 OF _1mq do pcallrtf2
 ACTIVATE POPUP _1mq
 RELEASE POPUP _1mq 
RETURN 
*------------------------------------------------------------------------
PROCEDURE pcallrtf2
 =pRTF2(.T., "Data\lst_record.RTF")
ENDPROC 
*--------------------------------------------------------------------
PROCEDURE p_lst

 HIDE POPUP _1mq 
 
 pal02=ALIAS()

 f02='Data\lst_record.txt' && файл вывода

 des1=FCREATE(f02)
 IF (des1<0)
  =MESSAGEBOX('Невозможно создать файл листинга!',16,'Внимание!',3000)
  RETURN 
 ENDIF 
 
 rr02=RECNO()
 GO TOP 
 
 DO WHILE !EOF()
  =FPUTS(des1, ALLTRIM(pNames)+' :    '+ALLTRIM(pZnach))
  SKIP 
 ENDDO 
 SELECT (al2)
 ror=RECNO()
 =FPUTS(des1,'---------СЧЕТ----------Дата откр.---Дата искл.--Статус---БИК ПБР--К.ключ-Тип сч.-Дата огран.-Тип ограничения-------')
 DO WHILE a807.BIC=BIC
  =FPUTS(des1,Account+' | '+DateIn+' | '+DateOut+' | '+AccountSta+' | '+AccountCBR+' | '+CK+' | '+RAccountT+' | '+ARDat+' | '+AccRs  )

  SKIP 
 ENDDO  
 =FPUTS(des1,'-------------------------------------------------------------------------------------------------------------------')  
 GO ror
 SELECT (pal02)
 GO rr02

 =FCLOSE(des1)
 
 LOCAL loWshShell as Wscript.Shell   
 
 parms = 'notepad.exe'+' '+f02


 loWshShell=CREATEOBJECT("WScript.Shell")
 loWshShell.Run(parms, 1, .F.) && .F. не ждать выполнения notepad.exe

Release loWshShell
DEACTIVATE POPUP _1mq 
RELEASE POPUPS _1mq 
*SELECT (pal02) 
RETURN 
*------------------------------------------------------------------------
PROCEDURE myHelp

IF !FILE('readme.txt')
 =MESSAGEBOX('Файл помощи не найден! ', 48, 'СПРАВОЧНИК БИК')
  RETURN .F. 
ENDIF 


 LOCAL loH as Wscript.Shell   &&, 1cApplicationRootFolder as String
 fH='readme.txt'
 parms = 'notepad.exe'+' '+fH


 loH=CREATEOBJECT("WScript.Shell")
 loH.Run(parms, 1, .F.) && .F. не ждать выполнения notepad.exe

Release loH



ENDPROC 
*------------------------------------------------------------------------
procedure UnZipFile
parameters pID, zTag
local I,J,K,L,BF,LBF
 L=65536
 I=space(1024) && Информация об файле
 J=space(100)   && Имя файла

 unzOpenCurrentFile(pID)
 unzGetCurrentFileInfo(pID,@I,@J,len(J),null,0,null,0)
 
 n_FileInZip = J
 
 K=fcreate(zTag+J)
 do while unzeof(pID)=0
  BF=space(L)
  LBF=unzReadCurrentFile(pID,@BF,L)
  fwrite(K,BF,LBF)
 enddo
 fclose(K)
 unzCloseCurrentFile(pID)
return

*--------------------------------------------------------------------
 Procedure url_download  
  PARAMETERS  lcRemoteFile, lcLocalFile   
    
 *lcRemoteFile -откуда скачать  
 *lcLocalFile  -где сохранить  
    
  DECLARE INTEGER URLDownloadToFile IN urlmon.dll;   
      INTEGER pCaller, STRING szURL, STRING szFileName,;   
      INTEGER dwReserved, INTEGER lpfnCB   
    
  WAIT "Идет закачка файла!" WINDOW NOWAIT   
    
  =URLDownloadToFile (0, lcRemoteFile, lcLocalFile, 0, 0)  
    
  WAIT "Закачка файла завершена!" WINDOW NOWAIT   
     
  endproc
*--------------------------------------------------------------------
PROCEDURE Kopi
LPARAMETERS how_copy

 IF FILE(pathdata+'a807'+dat77+'.dbf').OR.;
    FILE(pathdata+'acc807'+dat77+'.dbf').OR.;
    FILE(pathdata+'h807'+dat77+'.dbf')
    IF how_copy='вручную'
     =MESSAGEBOX('DBF-файлы справочника уже существуют.'+CHR(13)+'Копирование невозможно!',0+48,'Сохранение справочника в DBF')
    ENDIF  
  RETURN .F.
 ENDIF 

 tmp_al = ALIAS()
 tektmp=RECNO()
 WAIT 'Копирование начато...' WINDOW NOWAIT 
 SELECT (al)
 to1='Data\a807'+dat77+'.dbf'
 COPY TO &to1
 SELECT (al2)
 to2='Data\acc807'+dat77+'.dbf'
 COPY TO &to2
 to3='Data\h807'+dat77+'.dbf'
 SELECT 0 && переключаемся в область, где нет таблтцы
 CREATE DBF &to3 (EDNo C(9), EDDate C(10), EDAuthor C(10),  EDReceiver C(10),;
                  CreationRe C(4), CreationDT C(20), InfoTypeCo C(4), BusinessDa C(10),;
                  DirectoryV C(2))
 APPEND BLANK 
 
 
 
 REPLACE EDNo WITH m_EDNo, EDDate WITH m_EDDate, EDAuthor WITH m_EDAuthor, EDReceiver WITH m_ED11,;
         CreationRe WITH m_CreationReason, CreationDT WITH m_CreationDateTime, InfoTypeCo WITH m_InfoTypeCode,;
         BusinessDa WITH m_Bus11, DirectoryV WITH m_Dir11                  

 WAIT 'Копирование DBF завершено!' WINDOW NOWAIT 
 
 dat77=SUBSTR(DTOC(fr_start.Text1.Value),7,4)+SUBSTR(DTOC(fr_start.Text1.Value),4,2)+SUBSTR(DTOC(fr_start.Text1.Value),1,2)
 IF FILE(pathdata+'a807'+dat77+'.dbf').AND.;
   FILE(pathdata+'acc807'+dat77+'.dbf').AND.;
   FILE(pathdata+'h807'+dat77+'.dbf')
  fr_start.Command2.ForeColor = RGB(0,128,0)
 ELSE 
  fr_start.Command2.ForeColor = RGB(255,128,0)
 ENDIF 
 
 SELECT (tmp_al)
 IF !EOF()
  GO tektmp
 ENDIF  
ENDPROC 
*--------------------------------------------------------------------
PROCEDURE pcompare 

  HIDE POPUP _3mp 
  DO FORM w_compare_dat NAME w_com_d NOSHOW 
  w_com_d.Show(1)
  DEACTIVATE POPUP _3mp 
  RELEASE POPUPS _3mp 

ENDPROC 
*--------------------------------------------------------------------
PROCEDURE p_com1

SELECT (al)
GO TOP 
kuch1=0
DO WHILE !EOF()
 SELECT d807
 SEEK a807.BIC
 IF !FOUND()
 
  .m_arTblValues(1) = a807.BIC
  .m_arTblValues(2) = a807.NAMEP
  .WriteRow()               && занести значения .m_arTblValues в графы таблицы  
*  =FPUTS(f01,a807.BIC+' '+a807.NAMEP)
  kuch1=kuch1+1
 ENDIF 
 SELECT (al)
 SKIP 
ENDDO 

ENDPROC 
*--------------------------------------------------------------------
PROCEDURE p_com2

SELECT d807
GO TOP 
kuch2=0
DO WHILE !EOF()
 SELECT (al)
 SEEK d807.BIC
 IF !FOUND()
  .m_arTblValues(1) = d807.BIC
  .m_arTblValues(2) = d807.NAMEP
  .WriteRow()               && занести значения .m_arTblValues в графы таблицы   
*  =FPUTS(f01,d807.BIC+' '+d807.NAMEP)
  kuch2=kuch2+1
 ENDIF 
 SELECT d807
 SKIP 
ENDDO 

ENDPROC 
*--------------------------------------------------------------------
PROCEDURE p_exp_bnkseek

 WAIT 'Экспорт начат...' WINDOW NOWAIT 

 al33=ALIAS()
 tmp_bnkseek = pathdata+'bnkseek.dbf'
 SELECT 0
 CREATE TABLE &tmp_bnkseek CODEPAGE = 866;
(VKEY C(8),;
REAL C(4),;  
PZN C(2),;  
UER C(1),;  
RGN C(2),;  
IND C(6),;  
TNP C(20),;  
NNP C(25),;  
ADR C(160),;  
RKC C(9),;  
NAMEP C(160),;  
NAMEN C(30),;  
NEWNUM C(9),;  
NEWKS C(9),;  
PERMFO C(6),;  
SROK C(2),;  
AT1 C(7),;  
AT2 C(7),;  
TELEF C(25),;  
REGN C(9),;  
OKPO C(8),;  
DT_IZM D,; 
CKS C(6),;  
KSNP C(20),;  
DATE_IN D,;  
DATE_CH D,;  
VKEYDEL C(8),;  
DT_IZMR D)
 
 al_bnks=ALIAS()
 
 SELECT (al)
 GO TOP 
 DO WHILE !EOF()
  SELECT (al_bnks)
  APPEND BLANK 
  REPLACE PZN WITH a807.PTTYPE, NNP WITH a807.NNP, IND WITH a807.IND, ADR WITH a807.ADR, RGN WITH a807.RGN, REGN WITH a807.REGN,;
          NAMEP WITH a807.NAMEP, NEWNUM WITH a807.BIC, TNP WITH a807.TNP
 
  SELECT (al)
  SKIP 
 ENDDO 
 
 SELECT (al_bnks)
 USE 
 
 
 SELECT (al33)
 GO TOP 
 WAIT 'Экспорт закончен!' WINDOW NOWAIT  
ENDPROC 
*--------------------------------------------------------------------
PROCEDURE plstl
 HIDE POPUP _3mp
 SELECT (al) 
 i = 0
 trr = RECNO()
 fl = 'Data\'+'lstclnt.txt'
 
 IF FILE(fl)
  DELETE FILE &fl
 ENDIF 
  
 GO TOP  
 STRTOFILE('  Список участников справочника БИК за дату '+m_EDDate+' '+CHR(13)+CHR(10), fl, 1)  
 ffi = FILTER()
 IF LEN(ffi)>0
  STRTOFILE('  Фильтр: '+CHR(13)+CHR(10), fl, 1)
  
  IF ATC('tx1', ffi)>0
   STRTOFILE('  Наименование участника = '+tx1+CHR(13)+CHR(10), fl, 1)
  ENDIF 

  IF ATC('tx2', ffi)>0
   STRTOFILE('  Наименование населенного пункта = '+tx2+CHR(13)+CHR(10), fl, 1)
  ENDIF 

  IF ATC('tx3', ffi)>0
   STRTOFILE('  Адрес = '+tx3+CHR(13)+CHR(10), fl, 1)
  ENDIF 

  IF ATC('tx4', ffi)>0
   STRTOFILE('  Код территории = '+tx4+CHR(13)+CHR(10), fl, 1)
  ENDIF 

  IF ATC('tx5', ffi)>0
   STRTOFILE('  Тип участника перевода = '+tx5+CHR(13)+CHR(10), fl, 1)
  ENDIF 
 
  IF ATC('tx6', ffi)>0
   STRTOFILE('  Наименование участника на английском яз. = '+tx6+CHR(13)+CHR(10), fl, 1)
  ENDIF 
  
  IF ATC('tx7', ffi)>0
   STRTOFILE('  БИК головной орг. = '+tx7+CHR(13)+CHR(10), fl, 1)
  ENDIF 

  IF ATC('kus4', ffi)>0
   STRTOFILE('  Дата вкл. в состав уч. перевода = '+tx7+CHR(13)+CHR(10), fl, 1)
  ENDIF 

  IF ATC('tx9', ffi)>0
   STRTOFILE('  Участник обмена (0 - нет) (1 - да) = '+tx9+CHR(13)+CHR(10), fl, 1)
  ENDIF 


 ENDIF 
 STRTOFILE(' '+CHR(13)+CHR(10), fl, 1)
    
  DO WHILE !EOF()
   i = i+1
   STRTOFILE(BIC+' '+ALLTRIM(NameP)+' '+CHR(13)+CHR(10), fl, 1)  
   
   
   SKIP 
  ENDDO 

  GOTO trr
  STRTOFILE('--------------------------------'+CHR(13)+CHR(10), fl, 1)  
  STRTOFILE('ИТОГО:  '+ALLTRIM(STR(i,18))+' '+CHR(13)+CHR(10), fl, 1) 
  parl = 'notepad.exe'+' '+fl
  loWshShell=CREATEOBJECT("WScript.Shell")
  loWshShell.Run(parl, 1, .F.) && .F. не ждать выполнения notepad.exe

  Release loWshShell  
  
  DEACTIVATE POPUP _3mp
  RELEASE POPUPS _3mp
ENDPROC 
*--------------------------------------------------------------------
PROCEDURE pcallrtf1
 =pRTF1(.T., "Data\lstclnt.RTF")
ENDPROC 
*--------------------------------------------------------------------
FUNCTION prtf1(bWordStart_, cFileName_) 
 HIDE POPUP _3mp
 
WAIT 'Начало формирования листинга.... ' WINDOW NOWAIT 
 
 SELECT (al) 
 i = 0
 trr = RECNO()
 IF FILE(cFileName_)
  DELETE FILE &cFileName_
 ENDIF 
   
 GO TOP  
 

  oFile = CreateObject("CRtfFile", cFileName_,.T.)
  With (oFile) 
    .DefaultInit
    .WriteHeader
    .PageA4
    

   .WriteParagraph("  Список участников справочника БИК за дату "+m_EDDate,;
                   raCenter, rfsBold+rfsItalic, 0, 0, 3, 30)
    
    

 ffi = FILTER()
 IF LEN(ffi)>0
      
 .WriteParagraph("  Фильтр:", raLeft, rfsBold, 0, 0, 2, 24)      
      
  IF ATC('tx1', ffi)>0
   .WriteParagraph("    Наименование участника = "+tx1, raLeft, rfsDefault, 0, 0, 2, 18)      
  ENDIF 

  IF ATC('tx2', ffi)>0
   .WriteParagraph("    Наименование населенного пункта = "+tx2, raLeft, rfsDefault, 0, 0, 2, 18)  
  ENDIF 

  IF ATC('tx3', ffi)>0
   .WriteParagraph("    Адрес = "+tx3, raLeft, rfsDefault, 0, 0, 2, 18)     
  ENDIF 

  IF ATC('tx4', ffi)>0
   .WriteParagraph("    Код территории = "+tx4, raLeft, rfsDefault, 0, 0, 2, 18)     

  ENDIF 

  IF ATC('tx5', ffi)>0
   .WriteParagraph("    Тип участника перевода = "+tx5, raLeft, rfsDefault, 0, 0, 2, 18)     
  ENDIF 
 
  IF ATC('tx6', ffi)>0
   .WriteParagraph("    Наименование участника на английском яз. = "+tx6, raLeft, rfsDefault, 0, 0, 2, 18)     
  ENDIF 
  
  IF ATC('tx7', ffi)>0
   .WriteParagraph("    БИК головной орг. = "+tx7, raLeft, rfsDefault, 0, 0, 2, 18)     
  ENDIF 

  IF ATC('kus4', ffi)>0
   .WriteParagraph("    Дата вкл. в состав уч. перевода = "+kus4, raLeft, rfsDefault, 0, 0, 2, 18)     
  ENDIF 

  IF ATC('tx9', ffi)>0
   .WriteParagraph("    Участник обмена (0 - нет) (1 - да) = "+tx9, raLeft, rfsDefault, 0, 0, 2, 18)     
  ENDIF 


 ENDIF 
 
    .WriteParagraph("", raLeft, rfsDefault, 0, 0, 2, 18)     
    .SetAlignment(raCenter)
    .BeginTable                           && начало таблицы
    .SetColumnsCount(2)
    .m_arTblWidths(1) = .Twips(2)         && ширины колонок (в скобках - см)
    .m_arTblWidths(2) = .Twips(12)         && ширины колонок (в скобках - см)

    .SetFont(3, 20, rfsBold)
    .SetupColumns()
    .m_arTblValues(1) = "БИК"    
    .m_arTblValues(2) = "Наименование участника"
    .WriteRow()               && занести значения .m_arTblValues в графы таблицы 
     
  DO WHILE !EOF()
   i = i+1

    WAIT 'Вывод участников в таблицу: '+STR(i,18) WINDOW NOWAIT 
  
    For x = 1 To 2
      .m_arTblAlign(x) = raLeft
    Next x

    .SetFont(3, 18, rfsDefault)
    .SetupColumns()
    .m_arTblAlign(1) = raRight
    .m_arTblAlign(2) = raLeft
    .m_arTblValues(1) = BIC
    .m_arTblValues(2) = ALLTRIM(NameP)
    .WriteRow()               && занести значения .m_arTblValues в графы таблицы
   
   
   SKIP 
  ENDDO && -------- конец цикла по записям dbf
 
  GOTO trr
    
    .SetFont(3, 18, rfsBold)
    .m_arTblValues(1) = 'ИТОГО:  '
    .m_arTblValues(2) = ALLTRIM(STR(i,18))
    .WriteRow()               && занести значения .m_arTblValues в графы таблицы

    .EndTable

    .CloseFile 

* --------- Рабочий КОД !!!
*    If(bWordStart_) 
*      DECLARE Integer GetFocus IN WIN32API
*      DECLARE Integer ShellExecute IN SHELL32 INTEGER, STRING, STRING, STRING, STRING, INTEGER
*      hWnd = GetFocus()
*      If (hWnd != 0)
*        result=ShellExecute(hWnd, "open", cFileName_, "", "", 5)
*      Else 
*        =MessageBox("Файл отчета используется другим приложением!", 48,"Ошибка!")
*      EndIf
*    EndIf  
* ----------

  EndWith 
  
  parl = 'wordpad.exe'+' '+cFileName_
  loWshShell=CREATEOBJECT("WScript.Shell")
  loWshShell.Run(parl, 1, .F.) && .F. не ждать выполнения notepad.exe
  Release loWshShell  
  
  WAIT CLEAR 
  DEACTIVATE POPUP _3mp
  RELEASE POPUPS _3mp
RETURN  
*--------------------------------------------------------------------
FUNCTION prtf2(bWordStart_, cFileName_) 

 WAIT 'Начало формирования листинга.... ' WINDOW NOWAIT 
 HIDE POPUP _1mq 
 
 pal02=ALIAS()

* f02 = cFileName_ && файл вывода

 IF FILE(cFileName_)
  DELETE FILE &cFileName_
 ENDIF 
 
 rr02=RECNO()
 GO TOP 
 
   oFile = CreateObject("CRtfFile", cFileName_,.T.)
 WITH (oFile) 
    .DefaultInit
    .WriteHeader
*    .PageA4
    .PageA4LandScape
    
    .WriteParagraph("", raLeft, rfsDefault, 0, 0, 2, 24)

    .BeginTable                           && начало таблицы
    .SetColumnsCount(2)
    .m_arTblWidths(1) = .Twips(8)         && ширины колонок (в скобках - см)
    .m_arTblWidths(2) = .Twips(5)         && ширины колонок (в скобках - см)
 
 
    .SetFont(3, 20, rfsDefault)
    .SetupColumns()
 tt=0
 DO WHILE !EOF()
  tt=tt+1
 
    .m_arTblValues(1) = ALLTRIM(pNames)
    IF tt=2
     .SetFont(3, 20, rfsBold) 
    ELSE 
     .SetFont(3, 20, rfsDefault)  
    ENDIF     
    .m_arTblValues(2) = ALLTRIM(pZnach)
    .WriteRow()               && занести значения .m_arTblValues в графы таблицы 
   
  SKIP 
 ENDDO 

 .EndTable 
 
 .WriteParagraph("", raLeft, rfsDefault, 0, 0, 2, 24)
 
 
 SELECT acc807
 ror=RECNO()
 
    .BeginTable                           && начало таблицы
    .SetColumnsCount(9)
    .m_arTblWidths(1) = .Twips(4.5)         && ширины колонок (в скобках - см)
    .m_arTblWidths(2) = .Twips(2)         && ширины колонок (в скобках - см)
    .m_arTblWidths(3) = .Twips(2)
    .m_arTblWidths(4) = .Twips(1.2)    
    .m_arTblWidths(5) = .Twips(2.2)
    .m_arTblWidths(6) = .Twips(1)        
    .m_arTblWidths(7) = .Twips(1.2)
    .m_arTblWidths(8) = .Twips(2)
    .m_arTblWidths(9) = .Twips(1.5)            
    
    .SetFont(3, 20, rfsBold)
    .SetupColumns()
    
    FOR zz=1 TO 9
     .m_arTblAlign(zz) = raCenter
    NEXT zz
    
    .m_arTblValues(1) = 'СЧЕТ'    
    .m_arTblValues(2) = 'Дата откр.'
    .m_arTblValues(3) = 'Дата искл.'
    .m_arTblValues(4) = 'Статус'
    .m_arTblValues(5) = 'БИК ПБР'
    .m_arTblValues(6) = 'К.ключ'
    .m_arTblValues(7) = 'Тип сч.'
    .m_arTblValues(8) = 'Дата огран.'
    .m_arTblValues(9) = 'Тип ограничения'
    .WriteRow()               && занести значения .m_arTblValues в графы таблицы 
    .SetFont(3, 20, rfsDefault)

 kk=0   
 DO WHILE a807.BIC=BIC
  kk=kk+1

    .m_arTblValues(1) = Account    
    .m_arTblValues(2) = DateIn
    .m_arTblValues(3) = DateOut
    .m_arTblValues(4) = AccountSta
    .m_arTblValues(5) = AccountCBR
    .m_arTblValues(6) = CK
    .m_arTblValues(7) = RAccountT
    .m_arTblValues(8) = ARDat
    .m_arTblValues(9) = AccRs
    .WriteRow()               && занести значения .m_arTblValues в графы таблицы 


  SKIP 
 ENDDO  
 .m_arTblValues(1) = 'ИТОГО: '
 .m_arTblValues(2) = ' '+ALLTRIM(STR(kk,18))
 .m_arTblValues(3) = ''
 .m_arTblValues(4) = ''
 .m_arTblValues(5) = ''
 .m_arTblValues(6) = ''
 .m_arTblValues(7) = ''
 .m_arTblValues(8) = ''
 .m_arTblValues(9) = ''
 .SetFont(3, 20, rfsBold) 
 .WriteRow()               && занести значения .m_arTblValues в графы таблицы  
 .EndTable
 
 .WriteParagraph("", raLeft, rfsDefault, 0, 0, 2, 24) 
  
* GO ror
 SELECT (pal02)
 GO rr02

  .CloseFile  && закрытие файла 

 
 ENDWITH 
 
 WAIT CLEAR 
 
 LOCAL loWshShell as Wscript.Shell   
 parms = 'wordpad.exe'+' '+cFileName_
 loWshShell=CREATEOBJECT("WScript.Shell")
 loWshShell.Run(parms, 1, .F.) && .F. не ждать выполнения wordpad.exe
 Release loWshShell
 
 DEACTIVATE POPUP _1mq 
 RELEASE POPUPS _1mq 





RETURN  
*--------------------------------------------------------------------
FUNCTION frtfcompare(bWordStart_, cFileName_) 

 IF FILE(cFileName_)
  DELETE FILE &cFileName_
 ENDIF 

al44=ALIAS()

fltt=FILTER(al)


SELECT 0
USE (qqo)
INDEX ON BIC TAG BIC

USE (qqo) ALIAS d807 ORDER TAG BIC 
IF LEN(fltt)>0
 SET FILTER TO &fltt
ENDIF 

*f01=FCREATE(pathdata+'compdat.txt')

   oFile = CreateObject("CRtfFile", cFileName_,.T.)
 WITH (oFile) 
    .DefaultInit
    .WriteHeader
    .PageA4
*    .PageA4LandScape
    tqtmp = 'Сравнение содержания справочника БИК (ED807) за даты: '+DTOC(fr_start.Text1.Value)+' и '+DTOC(w_com_d.Text1.Value)
    .WriteParagraph(tqtmp, raLeft, rfsBold, 0, 0, 2, 30)
    .WriteParagraph("", raLeft, rfsDefault, 0, 0, 2, 24)


 
*=FPUTS(f01,'       Сравнение содержания справочника БИК (ED807) за даты: '+DTOC(ThisForm.Text1.Value)+' и '+DTOC(fr_start.Text1.Value))
*=FPUTS(f01,' ')
 
 ffi = FILTER()
 IF LEN(ffi)>0 && Если есть фильтр, то сведения о нём выводятся в файл

 .WriteParagraph("  Фильтр:", raLeft, rfsBold, 0, 0, 2, 24)      
      
  IF ATC('tx1', ffi)>0
   .WriteParagraph("    Наименование участника = "+tx1, raLeft, rfsDefault, 0, 0, 2, 18)      
  ENDIF 

  IF ATC('tx2', ffi)>0
   .WriteParagraph("    Наименование населенного пункта = "+tx2, raLeft, rfsDefault, 0, 0, 2, 18)  
  ENDIF 

  IF ATC('tx3', ffi)>0
   .WriteParagraph("    Адрес = "+tx3, raLeft, rfsDefault, 0, 0, 2, 18)     
  ENDIF 

  IF ATC('tx4', ffi)>0
   .WriteParagraph("    Код территории = "+tx4, raLeft, rfsDefault, 0, 0, 2, 18)     

  ENDIF 

  IF ATC('tx5', ffi)>0
   .WriteParagraph("    Тип участника перевода = "+tx5, raLeft, rfsDefault, 0, 0, 2, 18)     
  ENDIF 
 
  IF ATC('tx6', ffi)>0
   .WriteParagraph("    Наименование участника на английском яз. = "+tx6, raLeft, rfsDefault, 0, 0, 2, 18)     
  ENDIF 
  
  IF ATC('tx7', ffi)>0
   .WriteParagraph("    БИК головной орг. = "+tx7, raLeft, rfsDefault, 0, 0, 2, 18)     
  ENDIF 

  IF ATC('kus4', ffi)>0
   .WriteParagraph("    Дата вкл. в состав уч. перевода = "+kus4, raLeft, rfsDefault, 0, 0, 2, 18)     
  ENDIF 

  IF ATC('tx9', ffi)>0
   .WriteParagraph("    Участник обмена (0 - нет) (1 - да) = "+tx9, raLeft, rfsDefault, 0, 0, 2, 18)     
  ENDIF 


 ENDIF  && IF LEN(ffi)>0
 
 
 
 
 
IF w_com_d.Text1.Value > fr_start.Text1.Value

    .SetFont(2, 28, rfsBold)
    .WriteParagraph(" Выбывшие участники расчетов: ", raLeft, rfsUnderline, 0, 0, 2, 28)  && rfsBold
    .WriteParagraph(" ", raLeft, rfsDefault, 0, 0, 2, 24)
    .BeginTable                           && начало таблицы
    .SetColumnsCount(2)
    .m_arTblWidths(1) = .Twips(2)         && ширины колонок (в скобках - см)
    .m_arTblWidths(2) = .Twips(12)         && ширины колонок (в скобках - см)
    .SetFont(3, 20, rfsDefault)
    .SetupColumns()
*~~~~
 DO p_com1
*~~~~
    .SetFont(3, 20, rfsBold)
    .m_arTblValues(1) = "ИТОГО:"
    .m_arTblValues(2) = ALLTRIM(STR(kuch1,18))
    .WriteRow()               && занести значения .m_arTblValues в графы таблицы  
    .EndTable
    .WriteParagraph("", raLeft, rfsDefault, 0, 0, 2, 24) 
 
    .SetFont(2, 28, rfsBold)
    .WriteParagraph(" Новые участники расчетов: ", raLeft, rfsUnderline, 0, 0, 2, 28)
    .WriteParagraph(" ", raLeft, rfsDefault, 0, 0, 2, 24)
    .BeginTable                           && начало таблицы
    .SetColumnsCount(2)
    .m_arTblWidths(1) = .Twips(2)         && ширины колонок (в скобках - см)
    .m_arTblWidths(2) = .Twips(12)         && ширины колонок (в скобках - см)
    .SetFont(3, 20, rfsDefault)
    .SetupColumns()
*~~~~
 DO p_com2
*~~~~
    .SetFont(3, 20, rfsBold)
    .m_arTblValues(1) = "ИТОГО:"
    .m_arTblValues(2) = ALLTRIM(STR(kuch2,18))
    .WriteRow()               && занести значения .m_arTblValues в графы таблицы  
    .EndTable
    .WriteParagraph("", raLeft, rfsDefault, 0, 0, 2, 24) 
 
ELSE 

    .SetFont(2, 28, rfsBold)
    .WriteParagraph(" Новые участники расчетов: ", raLeft, rfsUnderline, 0, 0, 2, 28)
    .WriteParagraph(" ", raLeft, rfsDefault, 0, 0, 2, 24)
    .BeginTable                           && начало таблицы
    .SetColumnsCount(2)
    .m_arTblWidths(1) = .Twips(2)         && ширины колонок (в скобках - см)
    .m_arTblWidths(2) = .Twips(12)         && ширины колонок (в скобках - см)
    .SetFont(3, 20, rfsDefault)
    .SetupColumns()
*~~~~
 DO p_com1
*~~~~
    .SetFont(3, 20, rfsBold)
    .m_arTblValues(1) = "ИТОГО:"
    .m_arTblValues(2) = ALLTRIM(STR(kuch1,18))
    .WriteRow()               && занести значения .m_arTblValues в графы таблицы  
    .EndTable
    .WriteParagraph("", raLeft, rfsDefault, 0, 0, 2, 24) 


    .SetFont(2, 28, rfsBold)
    .WriteParagraph(" Выбывшие участники расчетов: ", raLeft, rfsUnderline, 0, 0, 2, 28)
    .WriteParagraph(" ", raLeft, rfsDefault, 0, 0, 2, 24)
    .BeginTable                           && начало таблицы
    .SetColumnsCount(2)
    .m_arTblWidths(1) = .Twips(2)         && ширины колонок (в скобках - см)
    .m_arTblWidths(2) = .Twips(12)         && ширины колонок (в скобках - см)
    .SetFont(3, 20, rfsDefault)
    .SetupColumns()
*~~~~
 DO p_com2
*~~~~

    .SetFont(3, 20, rfsBold)
    .m_arTblValues(1) = "ИТОГО:"
    .m_arTblValues(2) = ALLTRIM(STR(kuch2,18))
    .WriteRow()               && занести значения .m_arTblValues в графы таблицы  
    .EndTable
    .WriteParagraph("", raLeft, rfsDefault, 0, 0, 2, 24) 

ENDIF 

  .CloseFile  && закрытие файла 
ENDWITH 

SELECT d807
USE 

SELECT (al44)
GO TOP 
*=FPUTS(f01,' --------------------------- ')
*=FCLOSE(f01)

 LOCAL loWshShell as Wscript.Shell   
 parms = 'wordpad.exe'+' '+cFileName_
 loWshShell=CREATEOBJECT("WScript.Shell")
 loWshShell.Run(parms, 1, .F.) && .F. не ждать выполнения wordpad.exe
 Release loWshShell



RETURN  
*--------------------------------------------------------------------

