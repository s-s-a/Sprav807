SET ESCAPE ON
SET TALK OFF
SET DATE german
SET CENTURY ON 
SET SAFETY OFF
SET PROCEDURE TO procfile.prg, make_rtf.prg
SET SYSMENU OFF
*SET ALTERNATE TO testalter.txt
*SET ALTERNATE ON
*SET CONSOLE OFF 

DECLARE INTEGER ShellExecute IN SHELL32.DLL ;
    INTEGER nWinHandle, ;
    STRING cOperation, ;
    STRING cFileName, ;
    STRING cParameters, ;
    STRING cDirectory, ;
    INTEGER nShowWindow
    
DECLARE INTEGER LoadKeyboardLayout IN win32api STRING, INTEGER
DECLARE INTEGER GetKeyboardLayout IN Win32API Integer  

*----------------------------------------------------------------------
 declare integer unzOpen in ZLib string FileName
 declare integer unzClose in ZLib integer FileID
* declare integer unzGoToFirstFile in ZLib integer FileID
 declare integer unzGoToNextFile in ZLib integer FileID
 declare integer unzOpenCurrentFile in ZLib integer FileID
 declare integer unzGetCurrentFileInfo in ZLib;
         integer FileID,      string @file_info,    string @namefile,;
         integer LenNameFile, string @extraField,   integer extraFieldBufferSize,;
         string  @szComment,  integer ComBufSize
 declare integer unzCloseCurrentFile in ZLib integer FileID
 declare integer unzReadCurrentFile in ZLib integer FileID, string @buf, integer lbuf
 declare integer unzeof in ZLib integer FileID

*----------------------------------------------------------------------

PUBLIC al,al2, x_p_my, y_p_my, x_q_my, y_q_my, y_i_my, x_i_my, tx1, tx2, tx3, tx4, tx5, tx6, tx7, tx8, tx9, k_filt, fr_2,;
       fr_start, g_filt, g_naimenovanie, f_p_n, g_opt1, g_opt2, g_kolac, ccx, flag_sd, gnMutex,;
       INIFILE, lcEntry1, lcValue1, MyValue, n_FileInZip, dat77, kuch1, kuch2,;
       m_EDNo, m_EDDate, m_EDAuthor, m_CreationReason, m_CreationDateTime, m_InfoTypeCode,;
       m_EDReceiver, m_BusinessDay, m_DirectoryVersion, m_ED11, m_Bus11, m_Dir11, pathdata,;
       lcEntry2, lcValue2, lcEntry3, lcValue3, lcEntry4, lcValue4, lcEntry5, lcValue5, MyValue2, MyValue3,;
       MyValue4, MyValue5, MyValue6, MyValue7, kus4, lcEntry6, lcValue6, lcEntry7, lcValue7,;
       g1_Width, g1_Height, g2_Height, g2_Top, g2_Width
       
*        koef, koef2, koef3, koef4    


       
       
         
ON SHUTDOWN DO ExProg
ON ERROR DO errHandler WITH ERROR( ), MESSAGE( ), MESSAGE(1), PROGRAM( ), LINENO( )

 

flag_sd = 0 && flag shutdown
*flag_sd2 = 0  && flag shutdown для закрытия второго экземпляра программы
k_filt = 0 && Инициализация (Количество записей в фильтре)
g_filt = 0 && flag k_filt
f_p_n = 0 && флаг повторения поиска по наименованию
g_opt1 = 1
g_opt2 = 0
pWidth = SYSMETRIC(1)    && 1530 для FHD
pHeight = 82.4*(SYSMETRIC(2)/100)  && -135    && 720  для FHD
g_naimenovanie ='' && наименование учатника (в поиске)

_Screen.Left = 0
*_Screen.Top = 0
_Screen.Width = pWidth
_Screen.Height = pHeight 
_Screen.Caption = ''

 * Формируем идентификатор данного приложения  
  LOCAL lcApplicationName  
  lcApplicationName = GetEnv("SessionName") + "#"+ SYS(0)  
    
 * Формируем ссылку на объект Mutex  
  Declare Integer CreateMutex In Win32API ;  
  	Integer lpMutexAttributes, ;  
  	Integer bInitialOwner, ;  
  	String lpName  
    
  gnMutex = CreateMutex(0,1,m.lcApplicationName)  
  
 * Проверяем факт существования объекта Mutex с тем же именем  
  #DEFINE ERROR_ALREADY_EXISTS 183  
  Declare integer GetLastError In Win32API  
    
  If GetLastError() = ERROR_ALREADY_EXISTS  
  
 
 	* Приложение уже запущено  
 	* Надо вывести ранее запущенное приложение на передний план  
 	* или сообщить об этом факте пользователю  
 	* и закрыть текущее приложение 
 	loWS=Createobject("WScript.Shell") 
    lcCaption = 'Справочник БИК (на основе ED807)'
    loWS.AppActivate(lcCaption) && Выводит окно первого экземпляра программы на передний план!!! 
*    loWS.SendKeys("% "+CHR(13))   
    loWS.SendKeys("{F11}") && посылает приложению нажатие F11 (если оно свёрнуто, распахивается)
    
*     IF !(loWS.AppActivate(lcCaption) = .T.) && Выводит окно первого экземпляра программы на передний план!!!
*      MessageBox("Приложение "+lcCaption+" не найдено")  
*     ENDIF

*    flag_sd2 = 1
    ON SHUTDOWN 

 	RELEASE loWS 
 	
  	Do CloseMutex with .T.  
  	Return  
  EndIf  





  lcBuffer = SPACE(100)+ CHR(0)

  INIFILE="sprav807.INI"
  lcEntry1="URL1"
  lcValue1="http://www.cbr.ru/VFS/mcirabis/BIKNew/"
  lcEntry2="AutoSave"
  lcValue2="NO"
  lcEntry3="NumberButton"
  lcValue3="1"
  lcEntry4="AfterDays"
  lcValue4="365"
  lcEntry5="NodeDocument"
  lcValue5="ED807"
  lcEntry6="NodeBIK"
  lcValue6="ED807/BICDirectoryEntry"
  lcEntry7="NodeAccount"
  lcValue7="ED807/BICDirectoryEntry/Accounts"
  
  
 
   *-- DECLARE DLL statements for reading/writing to private INI files
  DECLARE INTEGER GetPrivateProfileString IN Win32API AS GetPrivStr ;
  String cSection, String cKey, String cDefault, String @cBuffer, ;
  Integer nBufferSize, String cINIFile
  
  DECLARE INTEGER WritePrivateProfileString IN Win32API AS WritePrivStr ;
  String cSection, String cKey, String cValue, String cINIFile
  
  IF !FILE(CURDIR() + INIFILE)
  *-- Write the entry to the INI file
   =WritePrivStr("Source", lcEntry1, lcValue1, CURDIR() + INIFILE)
   =WritePrivStr("SaveDBFonLoad", lcEntry2, lcValue2, CURDIR() + INIFILE)
   =WritePrivStr("RadioButton", lcEntry3, lcValue3, CURDIR() + INIFILE)
   =WritePrivStr("ClearHistory", lcEntry4, lcValue4, CURDIR() + INIFILE) 
*   =WritePrivStr("Nodes", lcEntry5, lcValue5, CURDIR() + INIFILE) 
*   =WritePrivStr("Nodes", lcEntry6, lcValue6, CURDIR() + INIFILE) 
*   =WritePrivStr("Nodes", lcEntry7, lcValue7, CURDIR() + INIFILE) 

     
  ENDIF   
  
  *-- Read the window position from the INI file
  IF GetPrivStr("Source", lcEntry1, "", @lcBuffer, LEN(lcBuffer), CURDIR() + INIFILE) > 0
    lnPos = AT(CHR(0), lcBuffer) 
    MyValue=LEFT(lcBuffer,lnPos - 1)
  ENDIF 

  IF GetPrivStr("SaveDBFonLoad", lcEntry2, "", @lcBuffer, LEN(lcBuffer), CURDIR() + INIFILE) > 0
    lnPos2 = AT(CHR(0), lcBuffer) 
    MyValue2=LEFT(lcBuffer,lnPos2 - 1)
  ENDIF 

  IF GetPrivStr("RadioButton", lcEntry3, "", @lcBuffer, LEN(lcBuffer), CURDIR() + INIFILE) > 0
    lnPos3 = AT(CHR(0), lcBuffer) 
    MyValue3=LEFT(lcBuffer,lnPos3 - 1)
  ENDIF 

  IF GetPrivStr("ClearHistory", lcEntry4, "", @lcBuffer, LEN(lcBuffer), CURDIR() + INIFILE) > 0
    lnPos4 = AT(CHR(0), lcBuffer) 
    MyValue4=LEFT(lcBuffer,lnPos4 - 1)
  ENDIF 

*  IF GetPrivStr("Nodes", lcEntry5, "", @lcBuffer, LEN(lcBuffer), CURDIR() + INIFILE) > 0
*    lnPos5 = AT(CHR(0), lcBuffer) 
*    MyValue5=LEFT(lcBuffer,lnPos5 - 1)
*  ENDIF 
*
*  IF GetPrivStr("Nodes", lcEntry6, "", @lcBuffer, LEN(lcBuffer), CURDIR() + INIFILE) > 0
*    lnPos6 = AT(CHR(0), lcBuffer) 
*    MyValue6=LEFT(lcBuffer,lnPos6 - 1)
*  ENDIF 
*
*  IF GetPrivStr("Nodes", lcEntry7, "", @lcBuffer, LEN(lcBuffer), CURDIR() + INIFILE) > 0
*    lnPos7 = AT(CHR(0), lcBuffer) 
*    MyValue7=LEFT(lcBuffer,lnPos7 - 1)
*  ENDIF 




_Screen.Caption = 'Справочник БИК (на основе ED807)'
Zoom Window Screen Max

ON KEY LABEL F11 Zoom Window Screen Max && распахивает главное окно на полный экран

pathcur = SYS(5)+ADDBS(CURDIR())
pathdata = pathcur+'Data\'
path_tmp = pathcur+'TMP\'
path_zip = pathcur+'ZIP\'

*-----------------Получение списка файлов данных в папке Data и ZIP и очистка старых----------------------
LOCAL loWshShell as Wscript.Shell   
IF FILE(path_tmp+'flst')
 DELETE FILE path_tmp+'flst'
ENDIF 
IF FILE(path_tmp+'tmp4tmp.cmd')
 DELETE FILE path_tmp+'tmp4tmp.cmd'
ENDIF 
STRTOFILE('dir /b '+pathdata+'a807*.* >' +path_tmp+'flst'+CHR(13), path_tmp+'tmp4tmp.cmd')
STRTOFILE('dir /b '+pathdata+'h807*.* >>'+path_tmp+'flst'+CHR(13), path_tmp+'tmp4tmp.cmd',.T.)
STRTOFILE('dir /b '+pathdata+'acc807*.* >>'+path_tmp+'flst'+CHR(13), path_tmp+'tmp4tmp.cmd',.T.)
STRTOFILE('dir /b '+pathdata+'20*.xml >>'+path_tmp+'flst'+CHR(13), path_tmp+'tmp4tmp.cmd',.T.)
STRTOFILE('dir /b '+path_zip+'*.zip >>'+path_tmp+'flst'+CHR(13), path_tmp+'tmp4tmp.cmd',.T.)

parms = path_tmp+'tmp4tmp.cmd' 
loWshShell=CREATEOBJECT("WScript.Shell")
loWshShell.Run(parms, 0, .T.)
Release loWshShell

TRY 
flag_qh=0

WAIT 'Очистка устаревших копий файлов... ' WINDOW NOWAIT

datdel = DATE()-INT(VAL(MyValue4))
*=MESSAGEBOX(DTOC(datdel))
hnd1 = FCREATE(path_tmp+'tmp7tmp.cmd')
hnd2 = FOPEN(path_tmp+'flst')

DO WHILE !FEOF(hnd2) && цикл чтения листинга файлов
 bufe = FGETS(hnd2)

 IF (ATC('a807',bufe)=1).OR.(ATC('h807',bufe)=1)
  yqq1 = SUBSTR(bufe,5,4)
  mqq1 = SUBSTR(bufe,9,2)
  dqq1 = SUBSTR(bufe,11,2)
  datqq1 = CTOD(dqq1+'.'+mqq1+'.'+yqq1)
  IF datqq1<datdel 
   flag_qh=flag_qh+1
   =FPUTS(hnd1,'del '+pathdata+ bufe)
  ENDIF 
 ENDIF 

 IF ATC('acc807',bufe)=1
  yqq1 = SUBSTR(bufe,7,4)
  mqq1 = SUBSTR(bufe,11,2)
  dqq1 = SUBSTR(bufe,13,2)
  datqq1 = CTOD(dqq1+'.'+mqq1+'.'+yqq1)
  IF datqq1<datdel 
   flag_qh=flag_qh+1
   =FPUTS(hnd1,'del '+pathdata+ bufe)
  ENDIF 
 ENDIF 

 IF RIGHT(bufe,12)='807_full.xml'
  yqq1 = SUBSTR(bufe,1,4)
  mqq1 = SUBSTR(bufe,5,2)
  dqq1 = SUBSTR(bufe,7,2)
  datqq1 = CTOD(dqq1+'.'+mqq1+'.'+yqq1)
  IF datqq1<datdel 
   flag_qh=flag_qh+1
   =FPUTS(hnd1,'del '+pathdata+ bufe)
  ENDIF 
 ENDIF 

 IF RIGHT(bufe,7)='SBR.zip'
  yqq1 = SUBSTR(bufe,1,4)
  mqq1 = SUBSTR(bufe,5,2)
  dqq1 = SUBSTR(bufe,7,2)
  datqq1 = CTOD(dqq1+'.'+mqq1+'.'+yqq1)
  IF datqq1<datdel 
   flag_qh=flag_qh+1
   =FPUTS(hnd1,'del '+path_zip+ bufe)
  ENDIF 
 ENDIF 

ENDDO && Конец цикла чтения листинга файлов

=FCLOSE(hnd1)
=FCLOSE(hnd2)

IF flag_qh>1
 parmsk = path_tmp+'tmp7tmp.cmd' 
 loWshShell=CREATEOBJECT("WScript.Shell")
 loWshShell.Run(parmsk, 0, .T.)
 Release loWshShell
ENDIF 

IF FILE(path_tmp+'tmp4tmp.cmd')
 DELETE FILE path_tmp+'tmp4tmp.cmd'
ENDIF 
IF FILE(path_tmp+'tmp7tmp.cmd')
 DELETE FILE path_tmp+'tmp7tmp.cmd'
ENDIF 

WAIT CLEAR 

CATCH 
 STRTOFILE(DTOC(DATE())+' '+TIME()+' '+'Ошибка удаления устаревших файлов из папок Data и ZIP!','sprav_err.log')
ENDTRY 


IF FILE(path_tmp+'flst')
 DELETE FILE path_tmp+'flst'
ENDIF 
 
*---------------------------------------------------------------------------------------------------



    
DO FORM Form1 NAME fr_start

READ EVENTS


*-----------------------------------------------------------------------

