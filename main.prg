Set Escape On
Set Talk Off
Set Date German
Set Century On
Set Safety Off
Set Procedure To procfile.prg, make_rtf.prg
Set Sysmenu Off
*SET ALTERNATE TO testalter.txt
*SET ALTERNATE ON
*SET CONSOLE OFF

Declare Integer ShellExecute In SHELL32.Dll ;
	INTEGER nWinHandle, ;
	STRING cOperation, ;
	STRING cFileName, ;
	STRING cParameters, ;
	STRING cDirectory, ;
	INTEGER nShowWindow

Declare Integer LoadKeyboardLayout In win32api String, Integer
Declare Integer GetKeyboardLayout In Win32API Integer

*----------------------------------------------------------------------
Declare Integer unzOpen In ZLib String FileName
Declare Integer unzClose In ZLib Integer FileID
* declare integer unzGoToFirstFile in ZLib integer FileID
Declare Integer unzGoToNextFile In ZLib Integer FileID
Declare Integer unzOpenCurrentFile In ZLib Integer FileID
Declare Integer unzGetCurrentFileInfo In ZLib;
	integer FileID,      String @file_info,    String @namefile,;
	integer LenNameFile, String @extraField,   Integer extraFieldBufferSize,;
	string  @szComment,  Integer ComBufSize
Declare Integer unzCloseCurrentFile In ZLib Integer FileID
Declare Integer unzReadCurrentFile In ZLib Integer FileID, String @buf, Integer lbuf
Declare Integer unzeof In ZLib Integer FileID

*----------------------------------------------------------------------

Public al,al2, x_p_my, y_p_my, x_q_my, y_q_my, y_i_my, x_i_my, tx1, tx2, tx3, tx4, tx5, tx6, tx7, tx8, tx9, k_filt, fr_2,;
	fr_start, g_filt, g_naimenovanie, f_p_n, g_opt1, g_opt2, g_kolac, ccx, flag_sd, gnMutex,;
	INIFILE, lcEntry1, lcValue1, MyValue, n_FileInZip, dat77, kuch1, kuch2,;
	m_EDNo, m_EDDate, m_EDAuthor, m_CreationReason, m_CreationDateTime, m_InfoTypeCode,;
	m_EDReceiver, m_BusinessDay, m_DirectoryVersion, m_ED11, m_Bus11, m_Dir11, pathdata,;
	lcEntry2, lcValue2, lcEntry3, lcValue3, lcEntry4, lcValue4, lcEntry5, lcValue5, MyValue2, MyValue3,;
	MyValue4, MyValue5, MyValue6, MyValue7, kus4, lcEntry6, lcValue6, lcEntry7, lcValue7,;
	g1_Width, g1_Height, g2_Height, g2_Top, g2_Width, tipu, uch01, okspr, numtext

*        koef, koef2, koef3, koef4

On Shutdown Do ExProg
*ssa*	On Error Do errHandler With Error( ), Message( ), Message(1), Program( ), Lineno( )

flag_sd = 0 && flag shutdown
*flag_sd2 = 0  && flag shutdown для закрытия второго экземпляра программы
k_filt = 0 && Инициализация (Количество записей в фильтре)
g_filt = 0 && flag k_filt
f_p_n = 0 && флаг повторения поиска по наименованию
g_opt1 = 1
g_opt2 = 0
pWidth = Sysmetric(1)    && 1530 для FHD
pHeight = 82.4*(Sysmetric(2)/100)  && -135    && 720  для FHD
g_naimenovanie ='' && наименование учатника (в поиске)
numtext = ''

With _Screen
	.Left = 0
*_Screen.Top = 0
	.Width = pWidth
	.Height = pHeight
	.Caption = ''
Endwith

* Формируем идентификатор данного приложения
Local lcApplicationName
lcApplicationName = Getenv("SessionName") + "#"+ Sys(0)

* Формируем ссылку на объект Mutex
Declare Integer CreateMutex In Win32API ;
	Integer lpMutexAttributes, ;
	Integer bInitialOwner, ;
	String lpName

gnMutex = CreateMutex(0,1,m.lcApplicationName)

* Проверяем факт существования объекта Mutex с тем же именем
#Define ERROR_ALREADY_EXISTS 183
Declare Integer GetLastError In Win32API

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
	On Shutdown

	Release loWS

	Do CloseMutex With .T.
	Return
Endif


lcBuffer = Space(100)+ Chr(0)

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
Declare Integer GetPrivateProfileString In Win32API As GetPrivStr ;
	String cSection, String cKey, String cDefault, String @cBuffer, ;
	Integer nBufferSize, String cINIFile

Declare Integer WritePrivateProfileString In Win32API As WritePrivStr ;
	String cSection, String cKey, String cValue, String cINIFile

If !File(Curdir() + INIFILE)
*-- Write the entry to the INI file
	=WritePrivStr("Source", lcEntry1, lcValue1, Curdir() + INIFILE)
	=WritePrivStr("SaveDBFonLoad", lcEntry2, lcValue2, Curdir() + INIFILE)
	=WritePrivStr("RadioButton", lcEntry3, lcValue3, Curdir() + INIFILE)
	=WritePrivStr("ClearHistory", lcEntry4, lcValue4, Curdir() + INIFILE)
*   =WritePrivStr("Nodes", lcEntry5, lcValue5, CURDIR() + INIFILE)
*   =WritePrivStr("Nodes", lcEntry6, lcValue6, CURDIR() + INIFILE)
*   =WritePrivStr("Nodes", lcEntry7, lcValue7, CURDIR() + INIFILE)


Endif

*-- Read the window position from the INI file
If GetPrivStr("Source", lcEntry1, "", @lcBuffer, Len(lcBuffer), Curdir() + INIFILE) > 0
	lnPos = At(Chr(0), lcBuffer)
	MyValue=Left(lcBuffer,lnPos - 1)
Endif

If GetPrivStr("SaveDBFonLoad", lcEntry2, "", @lcBuffer, Len(lcBuffer), Curdir() + INIFILE) > 0
	lnPos2 = At(Chr(0), lcBuffer)
	MyValue2=Left(lcBuffer,lnPos2 - 1)
Endif

If GetPrivStr("RadioButton", lcEntry3, "", @lcBuffer, Len(lcBuffer), Curdir() + INIFILE) > 0
	lnPos3 = At(Chr(0), lcBuffer)
	MyValue3=Left(lcBuffer,lnPos3 - 1)
Endif

If GetPrivStr("ClearHistory", lcEntry4, "", @lcBuffer, Len(lcBuffer), Curdir() + INIFILE) > 0
	lnPos4 = At(Chr(0), lcBuffer)
	MyValue4=Left(lcBuffer,lnPos4 - 1)
Endif

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

On Key Label F11 Zoom Window Screen Max && распахивает главное окно на полный экран

pathcur = Sys(5)+Addbs(Curdir())
pathdata = pathcur+'Data\'
path_tmp = pathcur+'TMP\'
path_zip = pathcur+'ZIP\'

*ssa*	создание нужных каталогов
Try
	Md (pathdata)
	Md (path_tmp)
	Md (path_zip)
Catch
EndTry

*-----------------Получение списка файлов данных в папке Data и ZIP и очистка старых----------------------
Local loWshShell As Wscript.Shell
If File(path_tmp+'flst')
	Delete File path_tmp+'flst'
Endif
If File(path_tmp+'tmp4tmp.cmd')
	Delete File path_tmp+'tmp4tmp.cmd'
Endif
If File(path_tmp+'tmp7tmp.cmd')
	Delete File path_tmp+'tmp7tmp.cmd'
Endif

Strtofile('dir /b '+pathdata+'a807*.* >' +path_tmp+'flst'+Chr(13), path_tmp+'tmp4tmp.cmd')
Strtofile('dir /b '+pathdata+'h807*.* >>'+path_tmp+'flst'+Chr(13), path_tmp+'tmp4tmp.cmd',.T.)
Strtofile('dir /b '+pathdata+'acc807*.* >>'+path_tmp+'flst'+Chr(13), path_tmp+'tmp4tmp.cmd',.T.)
Strtofile('dir /b '+pathdata+'20*.xml >>'+path_tmp+'flst'+Chr(13), path_tmp+'tmp4tmp.cmd',.T.)
Strtofile('dir /b '+path_zip+'*.zip >>'+path_tmp+'flst'+Chr(13), path_tmp+'tmp4tmp.cmd',.T.)

parms = path_tmp+'tmp4tmp.cmd'
loWshShell=Createobject("WScript.Shell")
loWshShell.Run(parms, 0, .T.)
Release loWshShell

Try
	flag_qh=0

	Wait 'Очистка устаревших копий файлов... ' Window Nowait

	datdel = Date()-Int(Val(MyValue4))
*	hnd1 = Fcreate(path_tmp+'tmp7tmp.cmd')
	hnd2 = Fopen(path_tmp+'flst')

	Do While !Feof(hnd2) && цикл чтения листинга файлов
		bufe = Fgets(hnd2)

		If (Atc('a807',bufe)=1).Or.(Atc('h807',bufe)=1)
			yqq1 = Substr(bufe,5,4)
			mqq1 = Substr(bufe,9,2)
			dqq1 = Substr(bufe,11,2)
*ssa*				datqq1 = Ctod(dqq1+'.'+mqq1+'.'+)
			datqq1 = date(VAL(yqq1), VAL(mqq1), VAL(dqq1))
			If datqq1<datdel
				flag_qh=flag_qh+1
*				=Fputs(hnd1,'del '+pathdata+ bufe)
				Strtofile('del '+pathdata+ bufe+Chr(10)+Chr(13), path_tmp+'tmp7tmp.cmd',.T.)


			Endif
		Endif

		If Atc('acc807',bufe)=1
			yqq1 = Substr(bufe,7,4)
			mqq1 = Substr(bufe,11,2)
			dqq1 = Substr(bufe,13,2)
*ssa*				datqq1 = Ctod(dqq1+'.'+mqq1+'.'+)
			datqq1 = date(VAL(yqq1), VAL(mqq1), VAL(dqq1))
			If datqq1<datdel
				flag_qh=flag_qh+1
*				=Fputs(hnd1,'del '+pathdata+ bufe)
				Strtofile('del '+pathdata+ bufe+Chr(10)+Chr(13), path_tmp+'tmp7tmp.cmd',.T.)
			Endif
		Endif

		If Right(bufe,12)='807_full.xml'
			yqq1 = Substr(bufe,1,4)
			mqq1 = Substr(bufe,5,2)
			dqq1 = Substr(bufe,7,2)
*ssa*				datqq1 = Ctod(dqq1+'.'+mqq1+'.'+)
			datqq1 = date(VAL(yqq1), VAL(mqq1), VAL(dqq1))
			If datqq1<datdel
				flag_qh=flag_qh+1
*				=Fputs(hnd1,'del '+pathdata+ bufe)
				Strtofile('del '+pathdata+ bufe+Chr(10)+Chr(13), path_tmp+'tmp7tmp.cmd',.T.)
			Endif
		Endif

		If Right(bufe,7)='SBR.zip'
			yqq1 = Substr(bufe,1,4)
			mqq1 = Substr(bufe,5,2)
			dqq1 = Substr(bufe,7,2)
*ssa*				datqq1 = Ctod(dqq1+'.'+mqq1+'.'+)
			datqq1 = date(VAL(yqq1), VAL(mqq1), VAL(dqq1))
			If datqq1<datdel
				flag_qh=flag_qh+1
*				=Fputs(hnd1,'del '+path_zip+ bufe)
				Strtofile('del '+path_zip+ bufe+Chr(10)+Chr(13), path_tmp+'tmp7tmp.cmd',.T.)
			Endif
		Endif

	Enddo && Конец цикла чтения листинга файлов

*	=Fclose(hnd1)
	=Fclose(hnd2)
	Try
		If flag_qh>1
			parmsk = path_tmp+'tmp7tmp.cmd'
			loWshShell=Createobject("WScript.Shell")
			loWshShell.Run(parmsk, 0, .T.)
			Release loWshShell
		ENDIF
	CATCH

	Endtry


	Wait Clear

Catch
	Strtofile(Dtoc(Date())+' '+Time()+' '+'Ошибка удаления устаревших файлов из папок Data и ZIP!','sprav_err.log')
	MESSAGEBOX('Ошибка удаления устаревших файлов из папок Data и ZIP!', 16, 'Sprav807')
Endtry

If File(path_tmp+'flst')
	Delete File path_tmp+'flst'
Endif
If File(path_tmp+'tmp4tmp.cmd')
	Delete File path_tmp+'tmp4tmp.cmd'
Endif
If File(path_tmp+'tmp7tmp.cmd')
	Delete File path_tmp+'tmp7tmp.cmd'
Endif


*-------от 19.01.2019 alex2sign--Создание справочников вынесено в отдельный файл----->
Do CreateRefs    
*  <---------- от 19.01.2019 alex2sign ---------------

Do Form Form1 Name fr_start

Read Events

*-----------------------------------------------------------------------
