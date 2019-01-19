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
*=MESSAGEBOX(DTOC(datdel))
	hnd1 = Fcreate(path_tmp+'tmp7tmp.cmd')
	hnd2 = Fopen(path_tmp+'flst')

	Do While !Feof(hnd2) && цикл чтения листинга файлов
		bufe = Fgets(hnd2)

		If (Atc('a807',bufe)=1).Or.(Atc('h807',bufe)=1)
			yqq1 = Substr(bufe,5,4)
			mqq1 = Substr(bufe,9,2)
			dqq1 = Substr(bufe,11,2)
*ssa*				datqq1 = Ctod(dqq1+'.'+mqq1+'.'+)
			datqq1 = date(yqq1, mqq1, dqq1)
			If datqq1<datdel
				flag_qh=flag_qh+1
				=Fputs(hnd1,'del '+pathdata+ bufe)
			Endif
		Endif

		If Atc('acc807',bufe)=1
			yqq1 = Substr(bufe,7,4)
			mqq1 = Substr(bufe,11,2)
			dqq1 = Substr(bufe,13,2)
*ssa*				datqq1 = Ctod(dqq1+'.'+mqq1+'.'+)
			datqq1 = date(yqq1, mqq1, dqq1)
			If datqq1<datdel
				flag_qh=flag_qh+1
				=Fputs(hnd1,'del '+pathdata+ bufe)
			Endif
		Endif

		If Right(bufe,12)='807_full.xml'
			yqq1 = Substr(bufe,1,4)
			mqq1 = Substr(bufe,5,2)
			dqq1 = Substr(bufe,7,2)
*ssa*				datqq1 = Ctod(dqq1+'.'+mqq1+'.'+)
			datqq1 = date(yqq1, mqq1, dqq1)
			If datqq1<datdel
				flag_qh=flag_qh+1
				=Fputs(hnd1,'del '+pathdata+ bufe)
			Endif
		Endif

		If Right(bufe,7)='SBR.zip'
			yqq1 = Substr(bufe,1,4)
			mqq1 = Substr(bufe,5,2)
			dqq1 = Substr(bufe,7,2)
*ssa*				datqq1 = Ctod(dqq1+'.'+mqq1+'.'+)
			datqq1 = date(yqq1, mqq1, dqq1)
			If datqq1<datdel
				flag_qh=flag_qh+1
				=Fputs(hnd1,'del '+path_zip+ bufe)
			Endif
		Endif

	Enddo && Конец цикла чтения листинга файлов

	=Fclose(hnd1)
	=Fclose(hnd2)

	If flag_qh>1
		parmsk = path_tmp+'tmp7tmp.cmd'
		loWshShell=Createobject("WScript.Shell")
		loWshShell.Run(parmsk, 0, .T.)
		Release loWshShell
	Endif

	If File(path_tmp+'tmp4tmp.cmd')
		Delete File path_tmp+'tmp4tmp.cmd'
	Endif
	If File(path_tmp+'tmp7tmp.cmd')
		Delete File path_tmp+'tmp7tmp.cmd'
	Endif

	Wait Clear

Catch
	Strtofile(Dtoc(Date())+' '+Time()+' '+'Ошибка удаления устаревших файлов из папок Data и ZIP!','sprav_err.log')
Endtry


If File(path_tmp+'flst')
	Delete File path_tmp+'flst'
Endif

*-------от 19.01.2019 alex2sign---------------------->
tipu = pathdata+'tipuch.dbf'

IF !FILE(tipu)
 DIMENSION AR_TU(18,2)
 AR_TU(1,1) ='00'
 AR_TU(2,1) ='10'
 AR_TU(3,1) ='12'
 AR_TU(4,1) ='15'
 AR_TU(5,1) ='16'
 AR_TU(6,1) ='20'
 AR_TU(7,1) ='30'
 AR_TU(8,1) ='40'
 AR_TU(9,1) ='51'
 AR_TU(10,1)='52'
 AR_TU(11,1)='60'
 AR_TU(12,1)='61'
 AR_TU(13,1)='65'
 AR_TU(14,1)='71'
 AR_TU(15,1)='75'
 AR_TU(16,1)='78'
 AR_TU(17,1)='90'
 AR_TU(18,1)='99'

 AR_TU(1,2) ='Главное управление Банка России'
 AR_TU(2,2) ='Расчетно-кассовый центр'
 AR_TU(3,2) ='Отделение, отделение – национальный банк главного управления Банка России'
 AR_TU(4,2) ='Структурное подразделение центрального аппарата Банка России'
 AR_TU(5,2) ='Кассовый центр'
 AR_TU(6,2) ='Кредитная организация'
 AR_TU(7,2) ='Филиал кредитной организации'
 AR_TU(8,2) ='Полевое учреждение Банка России'
 AR_TU(9,2) ='Федеральное казначейство'
 AR_TU(10,2)='Территориальный орган Федерального казначейства'
 AR_TU(11,2)='Иностранная кредитная организация'
 AR_TU(12,2)='Иностранный банк'
 AR_TU(13,2)='Иностранный центральный (национальный) банк'
 AR_TU(14,2)='Клиент кредитной организации, являющийся косвенным участником'
 AR_TU(15,2)='Клиринговая организация'
 AR_TU(16,2)='Внешняя платежная система'
 AR_TU(17,2)='Конкурсный управляющий (ликвидатор ликвидационная комиссия)'
 AR_TU(18,2)='Клиент Банка России, не являющийся участником платежной системы'
 
 SELECT 0
 CREATE TABLE &tipu (POLE1 C(2), POLE2 C(100))
 INSERT INTO &tipu FROM ARRAY AR_TU 
 USE 
ENDIF 
 
uch01 = pathdata+'uch.dbf' 
IF !FILE(uch01)
 DIMENSION AR_UCH(2,2)
 AR_UCH(1,1)='0'
 AR_UCH(2,1)='1'
 AR_UCH(1,2)='НЕТ'
 AR_UCH(2,2)='ДА'
 SELECT 0
 CREATE TABLE &uch01 (Pole1 C(1), Pole2 C(3))
 INSERT INTO &uch01 FROM ARRAY AR_UCH
 USE  
ENDIF    

okspr = pathdata+'okato.dbf'

IF !FILE(okspr)
 DIMENSION AR_OKATO(83,2)

 AR_OKATO(1,1)='01'
 AR_OKATO(1,2)='Алтайский край (г Барнаул)'
 AR_OKATO(2,1)='03'
 AR_OKATO(2,2)='Краснодарский край (г Краснодар)'
 AR_OKATO(3,1)='04'
 AR_OKATO(3,2)='Красноярский край (г Красноярск)'
 AR_OKATO(4,1)='05'
 AR_OKATO(4,2)='Приморский край (г Владивосток)'
 AR_OKATO(5,1)='07'
 AR_OKATO(5,2)='Ставропольский край (г Ставрополь)'
 AR_OKATO(6,1)='08'
 AR_OKATO(6,2)='Хабаровский край (г Хабаровск)'
 AR_OKATO(7,1)='10'
 AR_OKATO(7,2)='Амурская область (г Благовещенск)'
 AR_OKATO(8,1)='11'
 AR_OKATO(8,2)='Архангельская область (г Архангельск)'
 AR_OKATO(9,1)='12'
 AR_OKATO(9,2)='Астраханская область (г Астрахань)'
 AR_OKATO(10,1)='14'
 AR_OKATO(10,2)='Белгородская область (г Белгород)'
 AR_OKATO(11,1)='15'
 AR_OKATO(11,2)='Брянская область (г Брянск)'
 AR_OKATO(12,1)='17'
 AR_OKATO(12,2)='Владимирская область (г Владимир)'
 AR_OKATO(13,1)='18'
 AR_OKATO(13,2)='Волгоградская область (г Волгоград)'
 AR_OKATO(14,1)='19'
 AR_OKATO(14,2)='Вологодская область (г Вологда)'
 AR_OKATO(15,1)='20'
 AR_OKATO(15,2)='Воронежская область (г Воронеж)'
 AR_OKATO(16,1)='22'
 AR_OKATO(16,2)='Нижегородская область (г Нижний Новгород)'
 AR_OKATO(17,1)='24'
 AR_OKATO(17,2)='Ивановская область (г Иваново)'
 AR_OKATO(18,1)='25'
 AR_OKATO(18,2)='Иркутская область (г Иркутск)'
 AR_OKATO(19,1)='26'
 AR_OKATO(19,2)='Республика Ингушетия (г Магас)'
 AR_OKATO(20,1)='27'
 AR_OKATO(20,2)='Калининградская область (г Калининград)'
 AR_OKATO(21,1)='28'
 AR_OKATO(21,2)='Тверская область (г Тверь)'
 AR_OKATO(22,1)='29'
 AR_OKATO(22,2)='Калужская область (г Калуга)'
 AR_OKATO(23,1)='30'
 AR_OKATO(23,2)='Камчатский край (г Петропавловск-Камчатский)'
 AR_OKATO(24,1)='32'
 AR_OKATO(24,2)='Кемеровская область (г Кемерово)'
 AR_OKATO(25,1)='33'
 AR_OKATO(25,2)='Кировская область (г Киров)'
 AR_OKATO(26,1)='34'
 AR_OKATO(26,2)='Костромская область (г Кострома)'
 AR_OKATO(27,1)='35'
 AR_OKATO(27,2)='Республика Крым (г Симферополь)'
 AR_OKATO(28,1)='36'
 AR_OKATO(28,2)='Самарская область (г Самара)'
 AR_OKATO(29,1)='37'
 AR_OKATO(29,2)='Курганская область (г Курган)'
 AR_OKATO(30,1)='38'
 AR_OKATO(30,2)='Курская область (г Курск)'
 AR_OKATO(31,1)='40'
 AR_OKATO(31,2)='Город Санкт-Петербург город федерального значения'
 AR_OKATO(32,1)='41'
 AR_OKATO(32,2)='Ленинградская область (г Санкт-Петербург)'
 AR_OKATO(33,1)='42'
 AR_OKATO(33,2)='Липецкая область (г Липецк)'
 AR_OKATO(34,1)='44'
 AR_OKATO(34,2)='Магаданская область (г Магадан)'
 AR_OKATO(35,1)='45'
 AR_OKATO(35,2)='Город Москва столица Российской Федерации город федерального значения'
 AR_OKATO(36,1)='46'
 AR_OKATO(36,2)='Московская область (г Москва)'
 AR_OKATO(37,1)='47'
 AR_OKATO(37,2)='Мурманская область (г Мурманск)'
 AR_OKATO(38,1)='49'
 AR_OKATO(38,2)='Новгородская область (г Великий Новгород)'
 AR_OKATO(39,1)='50'
 AR_OKATO(39,2)='Новосибирская область (г Новосибирск)'
 AR_OKATO(40,1)='52'
 AR_OKATO(40,2)='Омская область (г Омск)'
 AR_OKATO(41,1)='53'
 AR_OKATO(41,2)='Оренбургская область (г Оренбург)'
 AR_OKATO(42,1)='54'
 AR_OKATO(42,2)='Орловская область (г Орёл)'
 AR_OKATO(43,1)='55'
 AR_OKATO(43,2)='Байконур'
 AR_OKATO(44,1)='56'
 AR_OKATO(44,2)='Пензенская область (г Пенза)'
 AR_OKATO(45,1)='57'
 AR_OKATO(45,2)='Пермский край (г Пермь)'
 AR_OKATO(46,1)='58'
 AR_OKATO(46,2)='Псковская область (г Псков)'
 AR_OKATO(47,1)='60'
 AR_OKATO(47,2)='Ростовская область (г Ростов-на-Дону)'
 AR_OKATO(48,1)='61'
 AR_OKATO(48,2)='Рязанская область (г Рязань)'
 AR_OKATO(49,1)='63'
 AR_OKATO(49,2)='Саратовская область (г Саратов)'
 AR_OKATO(50,1)='64'
 AR_OKATO(50,2)='Сахалинская область (г Южно-Сахалинск)'
 AR_OKATO(51,1)='65'
 AR_OKATO(51,2)='Свердловская область (г Екатеринбург)'
 AR_OKATO(52,1)='66'
 AR_OKATO(52,2)='Смоленская область (г Смоленск)'
 AR_OKATO(53,1)='67'
 AR_OKATO(53,2)='Город федерального значения Севастополь'
 AR_OKATO(54,1)='68'
 AR_OKATO(54,2)='Тамбовская область (г Тамбов)'
 AR_OKATO(55,1)='69'
 AR_OKATO(55,2)='Томская область (г Томск)'
 AR_OKATO(56,1)='70'
 AR_OKATO(56,2)='Тульская область (г Тула)'
 AR_OKATO(57,1)='71'
 AR_OKATO(57,2)='Тюменская область (г Тюмень)'
 AR_OKATO(58,1)='73'
 AR_OKATO(58,2)='Ульяновская область (г Ульяновск)'
 AR_OKATO(59,1)='75'
 AR_OKATO(59,2)='Челябинская область (г Челябинск)'
 AR_OKATO(60,1)='76'
 AR_OKATO(60,2)='Забайкальский край (г Чита)'
 AR_OKATO(61,1)='77'
 AR_OKATO(61,2)='Чукотский автономный округ (г Анадырь)'
 AR_OKATO(62,1)='78'
 AR_OKATO(62,2)='Ярославская область (г Ярославль)'
 AR_OKATO(63,1)='79'
 AR_OKATO(63,2)='Республика Адыгея (Адыгея) (г Майкоп)'
 AR_OKATO(64,1)='80'
 AR_OKATO(64,2)='Республика Башкортостан (г Уфа)'
 AR_OKATO(65,1)='81'
 AR_OKATO(65,2)='Республика Бурятия (г Улан-Удэ)'
 AR_OKATO(66,1)='82'
 AR_OKATO(66,2)='Республика Дагестан (г Махачкала)'
 AR_OKATO(67,1)='83'
 AR_OKATO(67,2)='Кабардино-Балкарская Республика (г Нальчик)'
 AR_OKATO(68,1)='84'
 AR_OKATO(68,2)='Республика Алтай (г Горно-Алтайск)'
 AR_OKATO(69,1)='85'
 AR_OKATO(69,2)='Республика Калмыкия (г Элиста)'
 AR_OKATO(70,1)='86'
 AR_OKATO(70,2)='Республика Карелия (г Петрозаводск)'
 AR_OKATO(71,1)='87'
 AR_OKATO(71,2)='Республика Коми (г Сыктывкар)'
 AR_OKATO(72,1)='88'
 AR_OKATO(72,2)='Республика Марий Эл (г Йошкар-Ола)'
 AR_OKATO(73,1)='89'
 AR_OKATO(73,2)='Республика Мордовия (г Саранск)'
 AR_OKATO(74,1)='90'
 AR_OKATO(74,2)='Республика Северная Осетия-Алания (г Владикавказ)'
 AR_OKATO(75,1)='91'
 AR_OKATO(75,2)='Карачаево-Черкесская Республика (г Черкесск)'
 AR_OKATO(76,1)='92'
 AR_OKATO(76,2)='Республика Татарстан (Татарстан) (г Казань)'
 AR_OKATO(77,1)='93'
 AR_OKATO(77,2)='Республика Тыва (г Кызыл)'
 AR_OKATO(78,1)='94'
 AR_OKATO(78,2)='Удмуртская Республика (г Ижевск)'
 AR_OKATO(79,1)='95'
 AR_OKATO(79,2)='Республика Хакасия (г Абакан)'
 AR_OKATO(80,1)='96'
 AR_OKATO(80,2)='Чеченская Республика (г Грозный)'
 AR_OKATO(81,1)='97'
 AR_OKATO(81,2)='Чувашская Республика - Чувашия (г Чебоксары)'
 AR_OKATO(82,1)='98'
 AR_OKATO(82,2)='Республика Саха (Якутия) (г Якутск)'
 AR_OKATO(83,1)='99'
 AR_OKATO(83,2)='Еврейская автономная область (г Биробиджан)'

 SELECT 0
 CREATE TABLE &okspr (POLE1 C(2), POLE2 C(150))
 INSERT INTO &okspr FROM ARRAY AR_OKATO
 USE  


ENDIF 
    
*  <-------------------- от 19.01.2019 alex2sign -------------------------

Do Form Form1 Name fr_start

Read Events

*-----------------------------------------------------------------------
