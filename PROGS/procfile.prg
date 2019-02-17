#INCLUDE "RTF.H"
* ��������� �������� ������� Mutex
Procedure CloseMutex
Lparameters IsExists  && ���������� �� ������ ����������

* ���� ������ ���������� ����������, �� ������� ������ Mutex �� ����
* �������� ����������� ������ ���� ������ ��� ������ ������ � ���� ����������
If IsExists = .F.
* �������� ������� Mutex
  Declare Integer ReleaseMutex In Win32API Integer hMutex
  ReleaseMutex(m.gnMutex)
Endif

* �������� ��� �� ������� ������ ������� Mutex
Declare Integer CloseHandle In Kernel32 Integer hObject
CloseHandle(m.gnMutex)

Endproc
*-----------------------------------------------------------------------------------------------

* ��������� �� ��������� � ��������� ?
Function IsInternetConnected
Local lnFlags As Integer
Declare SHORT InternetGetConnectedState In WININET Long @, Long
lnFlags = 0
InternetGetConnectedState(@lnFlags, 0)
Clear Dlls 'InternetGetConnectedState'
Return !Inlist(lnFlags, 0, 16, 32, 48)
*-----------------------------------------------------------------------------------------------

* �������� ���� � �������� ��� ��������
Function IsFileDownloaded
Lparameters tcSourceFile As String, tcTargetFile As String
If !File(tcTargetFile)
  Declare Integer URLDownloadToFile In URLMON.Dll Long, String, String, Long, Long
  URLDownloadToFile(0, tcSourceFile, tcTargetFile, 0, 0)
  Clear Dlls 'URLDownloadToFile'
  Return File(tcTargetFile)
Endif
Return .F.
*-----------------------------------------------------------------------------------------------
* ��������� �� ������
Procedure ShowError
Lparameters toException As Exception
Local lcErrorNo As String, lcMessage As String, lcStackLevel As String,;
  lcProcedure As String, lcLineNo As String, lcLineContents As String
Try
  lcErrorNo = '����� ������' + CHR_TAB + ': ' + Transform(toException.ErrorNo) + CHR_CR
  lcMessage = '���������' + CHR_TAB + ': ' + toException.Message + CHR_CR
  lcStackLevel = '������� �����' + CHR_TAB + ': ' + Transform(toException.StackLevel) + CHR_CR
  lcProcedure = '���������' + CHR_TAB + ': ' + toException.Procedure + CHR_CR
  lcLineNo = '����� ������' + CHR_TAB + ': ' + Transform(toException.Lineno)
  lcLineContents = Iif(Application.StartMode = 0,;
    CHR_CR + '����������' + CHR_TAB + ': ' + toException.LineContents, '')
  Messagebox(lcErrorNo + lcMessage + lcStackLevel + lcProcedure + lcLineNo + lcLineContents, 16,'Sprav807')
Catch
  Messagebox('������ ��� ������� ������� ��������� �� ������', 16, 'Sprav807')
Endtry
Return

*-----------------------------------------------------------------------------------------------
* �٨ ���� ��������� �� ������ (���������� ���)
Procedure errHandler
Parameter merror, Mess, mess1, mprog, mlineno
Clear
err1 = '����� ������: ' + Str(merror)+ Chr(13)
err2 = '��������� �� ������: ' + Mess + Chr(13)
err3 = '������ ���� � �������: ' + mess1 + Chr(13)
err4 = '����� ������ � �������: ' + Str(mlineno) + Chr(13)
err5 = '��������� � �������: ' + mprog + Chr(13)
Messagebox(err1 + err2 + err3 + err4 + err5, 16,'Sprav807')
Endproc
*-----------------------------------------------------------------------------------------------
Procedure poisk

*=MESSAGEBOX(_SCREEN.ActiveForm.ActiveControl.Name)
If Upper(_Screen.ActiveForm.ActiveControl.Name)='GRID2'
  fr_2.Grid1.SetFocus()
Endif

If Upper(_Screen.ActiveForm.ActiveControl.Name)='GRID1'
  activ_col = _Screen.ActiveForm.ActiveControl.ActiveColumn

  If activ_col = 1 && BIC
    vact1 = act_poisk()
    If !vact1
      Do Form w_poisk Name frm_poisk Noshow
      frm_poisk.Show(1)
    Endif
  Endif

  If activ_col = 2 && NameP
    vact2 = act_poisk2()
    If !vact2
      Do Form w_poisk2 Name frm_poisk2
      frm_poisk2.Hide
      frm_poisk2.Show(1)
    Endif
  Endif

  If activ_col = 13 && UID
    vact3 = act_poisk3()
    If !vact3
      Do Form w_poisk3 Name frm_poisk3 Noshow
      frm_poisk3.Show(1)
    Endif
  Endif

  If activ_col = 16 && Regn
    vact4 = act_poisk4()
    If !vact4
      Do Form w_poisk4 Name frm_poisk4 Noshow
      frm_poisk4.Show(1)
    Endif
  Endif

  If activ_col = 18 && SWBIC
    vact5 = act_poisk5()
    If !vact5
      Do Form w_poisk5 Name frm_poisk5 Noshow
      frm_poisk5.Show(1)
    Endif
  Endif

  If activ_col = 4 && Ind
    vact6 = act_poisk6()
    If !vact6
      Do Form w_poisk6 Name frm_poisk6 Noshow
      frm_poisk6.Show(1)
    Endif
  Endif

Else
* =MESSAGEBOX('',0,'',3000)
Endif

Return
*-----------------------------------------------------------------------------------------------
Procedure poisk_men
Hide Popup _3mp

Do poisk

Deactivate Popup _3mp
Release Popups _3mp
Return
*-----------------------------------------------------------------------------------------------
Procedure p1menu

Define Popup _3mp From y_p_my,x_p_my Margin Relative Shadow Font 'Arial', 10   && FONT 'Courier New', 10 STYLE 'B'
Defi Bar 1 Of _3mp Prompt " �������� "
Defi Bar 2 Of _3mp Prompt " ����� �� ������� "
Defi Bar 3 Of _3mp Prompt " ����� ������� "
Defi Bar 4 Of _3mp Prompt " ����� � ������� ��� "
Defi Bar 5 Of _3mp Prompt " ���������� �������� � ����� ������ "
Defi Bar 6 Of _3mp Prompt " �������� � ����� "
Defi Bar 7 Of _3mp Prompt " ������ �������� "
On Selec Bar 1 Of _3mp Do p2p
On Selec Bar 2 Of _3mp Do p3p
On Selec Bar 3 Of _3mp Do p4p
On Selec Bar 4 Of _3mp Do poisk_men
On Selec Bar 5 Of _3mp Do clipmy
On Selec Bar 6 Of _3mp Do pcompare
On Selec Bar 7 Of _3mp Do pcallrtf1 && plstl
Activate Popup _3mp
Release Popup _3mp

Return
*--------------------------------------------------------------------------------------------------
Procedure paccmenu
Define Popup _7mp From y_p_my,x_p_my Margin Relative Shadow Font 'Arial', 10   && FONT 'Courier New', 10 STYLE 'B'
Defi Bar 1 Of _7mp Prompt " ����� ����� � ������� ������ "
Defi Bar 2 Of _7mp Prompt " ���������� �������� � ����� ������"
On Selec Bar 1 Of _7mp Do pacc7
On Selec Bar 2 Of _7mp Do clipmy2
Activate Popup _7mp
Release Popup _7mp

Return
*--------------------------------------------------------------------------------------------------
Procedure vs_menu
Define Popup _9mp From y_q_my,x_q_my Margin Relative Shadow Font 'Arial', 10   && FONT 'Courier New', 10 STYLE 'B'
Defi Bar 1 Of _9mp Prompt " �������� "
On Selec Bar 1 Of _9mp Do pvs7
Activate Popup _9mp
Release Popup _9mp

Return
*--------------------------------------------------------------------------------------------------
Procedure pvs7  && ������� �� ������ ������ � ����������
Hide Popup _9mp
_Screen.ActiveForm.ActiveControl.Value = _Cliptext
Deactivate Popup _9mp
Release Popups _9mp
Return
*--------------------------------------------------------------------------------------------------
Procedure p2p

Hide Popup _3mp

mya1=My_activate_frm('FORM3')
If !mya1
  Do Form Form3 Name fr_3 Noshow
  fr_3.Show(1)
Endif

Deactivate Popup _3mp
Release Popups _3mp

Return
*--------------------------------------------------------------------------------------------------
Procedure pacc7
Hide Popup _7mp

mya1=My_activate_frm('FORM_ACC')
If !mya1
  Do Form w_poisk_acc1 Name fr_acc7 Noshow
  fr_acc7.Show(1)
Endif

Deactivate Popup _7mp
Release Popups _7mp

Return
*--------------------------------------------------------------------------------------------------
Procedure p3p  && ��������� �������
Hide Popup _3mp
Push Key Clear

Wait Clear
_vfp.StatusBar=''

mya2=My_activate_frm('FORM4')
If !mya2
  Do Form Form4 Name fr_4 Noshow
  fr_4.Show(1)
Endif

Pop Key
Wait '������� ��� = '+Str(k_filt) Window Nowait

Deactivate Popup _3mp
Release Popups _3mp

Return
*--------------------------------------------------------------------------------------------------
Procedure p4p  && ����� �������
Hide Popup _3mp
Set Filter To
Count To k_filt
Go Top
Wait '������� ��� = '+Str(k_filt) Window Nowait  && NOCLEAR
_vfp.StatusBar='������� ��� = '+Str(k_filt)

Store '' To tx1, tx2, tx3, tx4, tx5, tx6, tx7, tx8, tx9, kus4

Deactivate Popup _3mp
Release Popups _3mp
Return
*--------------------------------------------------------------------------------------------------
Function My_activate_frm
Lparameters tcFormName

If Pcount() > 0
  If Vartype(tcFormName) = 'C'

    tcFormName = Upper(tcFormName)

    Local lnForCounter

    For lnForCounter = 1 To _Screen.FormCount

* WAIT _Screen.Forms(lnForCounter).Name WINDOW

      If Upper(_Screen.Forms(lnForCounter).Name) = tcFormName && ���� ����� ���� � ������� _Screen.Forms()

        If Type('_SCREEN.FORMS(lnForCounter).NAME') = 'C' && ���� _Screen.ActiveForm � ������ ������ �������� �������� � �� �� ����� ���������

*WAIT tcFormName + STR(lnForCounter ,4)  WINDOW

          If Upper(_Screen.Forms(lnForCounter).Name) == tcFormName && ���� �����-�������� � ������ ������ �������
            _Screen.Forms(lnForCounter).Show()
            Return .T.
          Endif

        Endif
      Endif
    Endfor

  Endif
Endif

Return .F.
Endfunc
*------------------------------------------------------------------------
Function act_poisk
Return My_activate_frm('FORM_BIC')
*------------------------------------------------------------------------
Function act_poisk2
Return My_activate_frm('FORM_NAIM')
*------------------------------------------------------------------------
Function act_poisk3
Return My_activate_frm('FORM_UID')
*------------------------------------------------------------------------
Function act_poisk4
Return My_activate_frm('FORM_Regn')
*------------------------------------------------------------------------
Function act_poisk5
Return My_activate_frm('FORM_SWBIC')
*------------------------------------------------------------------------
Function act_poisk6
Return My_activate_frm('FORM_Ind')
*------------------------------------------------------------------------
Function act_poisk7
Return My_activate_frm('FORM_ACC')
*------------------------------------------------------------------------
Procedure act_poisk_a
If Upper(_Screen.ActiveForm.ActiveControl.Name)='GRID1'
  fr_2.Grid2.SetFocus()
Endif

vactacc = act_poisk7()
If !vactacc
  Do Form w_poisk_acc1 Name frm_poisk_acc Noshow
  frm_poisk_acc.Show(1)
Endif

Return
*------------------------------------------------------------------------
Procedure clipmy && !!!����������� � ����� ������ ����� ������ ������ � ������� ���������!!!!!
Hide Popup _3mp
* ���������:
*  #DEFINE KEYBOARD_GERMAN_ST   0x0407    && �������� (��������)
#Define KEYBOARD_ENGLISH_US   0x0409    && ���������� (����������� �����)
*  #DEFINE KEYBOARD_FRENCH_ST   0x040c    && ����������� (��������)
#Define KEYBOARD_RUSSIAN     0x0419    && �������

lnCurrentKeyboard = GetKeyboardLayout(0)
* ��������� ������� ����� (������� 16 ��� �� 32)
lnCurrentKeyboard = Bitrshift(m.lnCurrentKeyboard,16)

If m.lnCurrentKeyboard <> KEYBOARD_RUSSIAN
  =LoadKeyboardLayout("00000419",1) && ���
Endif

ccx='fr_2.Grid1.Column'+Transform(_Screen.ActiveForm.ActiveControl.ActiveColumn)+'.Text1.Value'
ccx=Alltrim(ccx)
_Cliptext=&ccx && !!!����������� � ����� ������ ����� ������ ������ � ������� ���������!!!!!

If m.lnCurrentKeyboard=KEYBOARD_ENGLISH_US
  LoadKeyboardLayout("00000409",1) && Eng
Endif

Deactivate Popup _3mp
Release Popups _3mp
Return
*------------------------------------------------------------------------
Procedure clipmy2 && !!!����������� � ����� ������ ����� ������ ������ � ������� ���������!!!!!
Hide Popup _7mp
* ���������:
*  #DEFINE KEYBOARD_GERMAN_ST   0x0407    && �������� (��������)
#Define KEYBOARD_ENGLISH_US   0x0409    && ���������� (����������� �����)
*  #DEFINE KEYBOARD_FRENCH_ST   0x040c    && ����������� (��������)
#Define KEYBOARD_RUSSIAN     0x0419    && �������

lnCurrentKeyboard = GetKeyboardLayout(0)
* ��������� ������� ����� (������� 16 ��� �� 32)
lnCurrentKeyboard = Bitrshift(m.lnCurrentKeyboard,16)

If m.lnCurrentKeyboard <> KEYBOARD_RUSSIAN
  LoadKeyboardLayout("00000419",1) && ���
Endif

ccx='fr_2.Grid2.Column'+Transform(_Screen.ActiveForm.ActiveControl.ActiveColumn)+'.Text1.Value'
ccx=Alltrim(ccx)
_Cliptext=&ccx && !!!����������� � ����� ������ ����� ������ ������ � ������� ���������!!!!!

If m.lnCurrentKeyboard=KEYBOARD_ENGLISH_US
  LoadKeyboardLayout("00000409",1) && Eng
Endif

Deactivate Popup _7mp
Release Popups _7mp
Return
*------------------------------------------------------------------------
Procedure pimenu1
Define Popup _1mq From y_i_my,x_i_my Margin Relative Shadow Font 'Arial', 10   && FONT 'Courier New', 10 STYLE 'B'
Defi Bar 1 Of _1mq Prompt " ����� � ��������� ���� "
On Selec Bar 1 Of _1mq Do pcallrtf2
Activate Popup _1mq
Release Popup _1mq
Return
*------------------------------------------------------------------------
Procedure pcallrtf2
=pRTF2(.T., "Data\lst_record.RTF")
Endproc
*--------------------------------------------------------------------
Procedure p_lst

Hide Popup _1mq

pal02=Alias()

f02='Data\lst_record.txt' && ���� ������
Set Textmerge To (f02) On Noshow

*!*  des1=Fcreate(f02)
*!*  If (des1<0)
*!*    Messagebox('���������� ������� ���� ��������!',16,'��������!',3000)
*!*    Return
*!*  Endif

*!*  rr02=Recno()
*!*  Go Top

Scan
*!*    Fputs(des1, Alltrim(pNames)+' :    '+Alltrim(pZnach))
  \<<Alltrim(pNames)>> :    <<Alltrim(pZnach)>>
Endscan
Select (al2)
ror=Recno()
*!*  Fputs(des1,'---------����----------���� ����.---���� ����.--������---��� ���--�.����-��� ��.-���� �����.-��� �����������-------')
\---------����----------���� ����.---���� ����.--������---��� ���--�.����-��� ��.-���� �����.-��� �����������-------
*!*  Do While a807.BIC=BIC
Scan While a807.BIC=BIC
  Fputs(des1,Account+' | '+DateIn+' | '+DateOut+' | '+AccountSta+' | '+AccountCBR+' | '+CK+' | '+RAccountT+' | '+ARDat+' | '+AccRs  )
*!*    Skip
*!*  Enddo
Endscan
*!*  Fputs(des1,'-------------------------------------------------------------------------------------------------------------------')
\-------------------------------------------------------------------------------------------------------------------
Go ror
Select (pal02)
Go rr02

Fclose(des1)
Set Textmerge To Off

Local loWshShell As Wscript.Shell

parms = 'notepad.exe'+' '+f02

loWshShell=Createobject("WScript.Shell")
loWshShell.Run(parms, 1, .F.) && .F. �� ����� ���������� notepad.exe

Release loWshShell
Deactivate Popup _1mq
Release Popups _1mq
*SELECT (pal02)
Return
*------------------------------------------------------------------------
Procedure myHelp

If !File('readme.txt')
  Messagebox('���� ������ �� ������! ', 48, '���������� ���')
  Return .F.
Endif

Local loH As Wscript.Shell   &&, 1cApplicationRootFolder as String
fH='readme.txt'
parms = 'notepad.exe'+' '+fH

loH=Createobject("WScript.Shell")
loH.Run(parms, 1, .F.) && .F. �� ����� ���������� notepad.exe

Release loH

Endproc
*------------------------------------------------------------------------
Procedure UnZipFile
Parameters pID, zTag
Local I,J,K,L,BF,LBF
L=65536
I=Space(1024) && ���������� �� �����
J=Space(100)   && ��� �����

unzOpenCurrentFile(pID)
unzGetCurrentFileInfo(pID,@I,@J,Len(J),Null,0,Null,0)

*!*  n_FileInZip = Rtrim(J)

K=Fcreate(zTag+J)
Do While unzeof(pID)=0
  BF=Space(L)
  LBF=unzReadCurrentFile(pID,@BF,L)
  Fwrite(K,BF,LBF)
Enddo
Fclose(K)
unzCloseCurrentFile(pID)
Return Rtrim(J)


*--------------------------------------------------------------------
Procedure url_download
Parameters  lcRemoteFile, lcLocalFile

*lcRemoteFile -������ �������
*lcLocalFile  -��� ���������

Declare Integer URLDownloadToFile In urlmon.Dll;
  INTEGER pCaller, String szURL, String szFileName,;
  INTEGER dwReserved, Integer lpfnCB

Wait "���� ������� �����!" Window Nowait

URLDownloadToFile (0, lcRemoteFile, lcLocalFile, 0, 0)

Wait "������� ����� ���������!" Window Nowait

Endproc
*--------------------------------------------------------------------
Procedure Kopi
Lparameters how_copy

dat77=Dtos(fr_start.Text1.Value)

If IsFileExists(fr_start.Text1.Value)
  If how_copy='�������'
    Messagebox('DBF-����� ����������� ��� ����������.'+Chr(13)+'����������� ����������!',0+48,'���������� ����������� � DBF')
  Endif
  Return .F.
Endif

tmp_al = Alias()
tektmp=Recno()
Wait '����������� ������...' Window Nowait
Select a807 && (al)
Copy To 'Data\a807'+dat77+'.dbf'
Select acc807 && (al2)
Copy To 'Data\acc807'+dat77+'.dbf'
to3='Data\h807'+dat77+'.dbf'
Select 0 && ������������� � �������, ��� ��� �������
Create Dbf (to3) (EDNo C(9), EDDate C(10), EDAuthor C(10),  EDReceiver C(10),;
  CreationRe C(4), CreationDT C(20), InfoTypeCo C(4), BusinessDa C(10),;
  DirectoryV C(2))
Append Blank

Replace EDNo With m_EDNo, EDDate With m_EDDate, EDAuthor With m_EDAuthor, EDReceiver With m_ED11,;
  CreationRe With m_CreationReason, CreationDT With m_CreationDateTime, InfoTypeCo With m_InfoTypeCode,;
  BusinessDa With m_Bus11, DirectoryV With m_Dir11

Wait '����������� DBF ���������!' Window Nowait

fr_start.Command2.ForeColor = Iif(IsFileExists(fr_start.Text1.Value), Rgb(255,128,0), Rgb(0,128,0))

Select (tmp_al)
If !Eof()
  Go tektmp
Endif
Endproc
*--------------------------------------------------------------------
Procedure pcompare

Hide Popup _3mp
Do Form w_compare_dat Name w_com_d Noshow
w_com_d.Show(1)
Deactivate Popup _3mp
Release Popups _3mp

Endproc
*--------------------------------------------------------------------
Procedure p_com1

Select (al)
Go Top
kuch1=0
Do While !Eof()
  Select d807
  Seek a807.BIC
  If !Found()

    .m_arTblValues(1) = a807.BIC
    .m_arTblValues(2) = a807.NAMEP
    .WriteRow()               && ������� �������� .m_arTblValues � ����� �������
    kuch1=kuch1+1
  Endif
  Select (al)
  Skip
Enddo

Endproc
*--------------------------------------------------------------------
Procedure p_com2

Select d807
Go Top
kuch2=0
Do While !Eof()
  Select (al)
  Seek d807.BIC
  If !Found()
    .m_arTblValues(1) = d807.BIC
    .m_arTblValues(2) = d807.NAMEP
    .WriteRow()               && ������� �������� .m_arTblValues � ����� �������
    kuch2=kuch2+1
  Endif
  Select d807
  Skip
Enddo

Endproc
*--------------------------------------------------------------------
Procedure p_exp_bnkseek

Wait '������� �����...' Window Nowait

al33=Alias()
tmp_bnkseek = pathdata+'bnkseek.dbf'
Select 0
Create Table &tmp_bnkseek Codepage = 866;
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

al_bnks=Alias()

Select (al)
ord1=Order()
Set Order To Tag BIC
Go Top

Do While !Eof()

  Select (al_bnks)
  Append Blank
  Replace PZN With a807.PTTYPE, NNP With a807.NNP, IND With a807.IND, ADR With a807.ADR, RGN With a807.RGN, REGN With a807.REGN,;
    NAMEP With a807.NAMEP, NEWNUM With a807.BIC, TNP With a807.TNP, KSNP With Iif(Atc('CRSA',acc807.RAccountT)=1, acc807.Account, '') && CRSA - ������� ������������������ �����

  Select (al)
  Skip
Enddo

Select (al_bnks)
Use

Select (al33)
Set Order To &ord1
Go Top
Wait '������� ��������!' Window Nowait
Endproc
*--------------------------------------------------------------------
Procedure plstl
Hide Popup _3mp
Select (al)
I = 0
trr = Recno()
fl = 'Data\'+'lstclnt.txt'
Set Textmerge To (fl) On Noshow

*!*  If File(fl)
*!*    Delete File &fl
*!*  Endif

Go Top
*!*  Strtofile('  ������ ���������� ����������� ��� �� ���� '+m_EDDate+' '+Chr(13)+Chr(10), fl, 1)
\\  ������ ���������� ����������� ��� �� ���� <<m_EDDate>>
ffi = Filter()
If Len(ffi)>0
*!*    Strtofile('  ������: '+Chr(13)+Chr(10), fl, 1)
  \  ������:

  If Atc('tx1', ffi)>0
*!*      Strtofile('  ������������ ��������� = '+tx1+Chr(13)+Chr(10), fl, 1)
    \  ������������ ��������� = <<tx1>>
  Endif

  If Atc('tx2', ffi)>0
*!*      Strtofile('  ������������ ����������� ������ = '+tx2+Chr(13)+Chr(10), fl, 1)
    \  ������������ ����������� ������ = <<tx2>>
  Endif

  If Atc('tx3', ffi)>0
*!*      Strtofile('  ����� = '+tx3+Chr(13)+Chr(10), fl, 1)
    \  ����� = <<tx3>>
  Endif

  If Atc('tx4', ffi)>0
*!*      Strtofile('  ��� ���������� = '+tx4+Chr(13)+Chr(10), fl, 1)
    \  ��� ���������� = '+tx4>>
  Endif

  If Atc('tx5', ffi)>0
*!*      Strtofile('  ��� ��������� �������� = '+tx5+Chr(13)+Chr(10), fl, 1)
    \  ��� ��������� �������� = '+tx5>>
  Endif

  If Atc('tx6', ffi)>0
*!*      Strtofile('  ������������ ��������� �� ���������� ��. = '+tx6+Chr(13)+Chr(10), fl, 1)
    \  ������������ ��������� �� ���������� ��. = '+tx6>>
  Endif

  If Atc('tx7', ffi)>0
*!*      Strtofile('  ��� �������� ���. = '+tx7+Chr(13)+Chr(10), fl, 1)
    \  ��� �������� ���. = '+tx7>>
  Endif

  If Atc('kus4', ffi)>0
*!*      Strtofile('  ���� ���. � ������ ��. �������� = '+tx7+Chr(13)+Chr(10), fl, 1)
    \  ���� ���. � ������ ��. �������� = '+tx7>>
  Endif

  If Atc('tx9', ffi)>0
*!*      Strtofile('  �������� ������ (0 - ���) (1 - ��) = '+tx9+Chr(13)+Chr(10), fl, 1)
    \  �������� ������ (0 - ���) (1 - ��) = '+tx9>>
  Endif

Endif
*!*  Strtofile(' '+Chr(13)+Chr(10), fl, 1)
\\
Scan
  I = I+1
*!*    Strtofile(BIC+' '+Alltrim(NAMEP)+' '+Chr(13)+Chr(10), fl, 1)
  \<<BIC>> <<Alltrim(NAMEP)>>
Endscan

Goto trr
*!*  Strtofile('--------------------------------'+Chr(13)+Chr(10), fl, 1)
\--------------------------------
*!*  Strtofile('�����:  '+Transform(I)+' '+Chr(13)+Chr(10), fl, 1)
\�����:  <<Transform(I)>>
Set Textmerge To Off
parl = 'notepad.exe'+' '+fl
loWshShell=Createobject("WScript.Shell")
loWshShell.Run(parl, 1, .F.) && .F. �� ����� ���������� notepad.exe

Release loWshShell

Deactivate Popup _3mp
Release Popups _3mp
Endproc
*--------------------------------------------------------------------
Procedure pcallrtf1
pRTF1(.T., "Data\lstclnt.RTF")
Endproc
*--------------------------------------------------------------------
Function pRTF1(bWordStart_, cFileName_)
Hide Popup _3mp

Wait '������ ������������ ��������.... ' Window Nowait

Select (al)
I = 0
trr = Recno()
Erase (cFileName_)

Go Top

oFile = Createobject("CRtfFile", cFileName_,.T.)
With (oFile)
  .DefaultInit
  .WriteHeader
  .PageA4

  .WriteParagraph("  ������ ���������� ����������� ��� �� ���� "+m_EDDate,;
    raCenter, rfsBold+rfsItalic, 0, 0, 3, 30)

  ffi = Filter()
  If Len(ffi)>0

    .WriteParagraph("  ������:", raLeft, rfsBold, 0, 0, 2, 24)

    If Atc('tx1', ffi)>0
      .WriteParagraph("    ������������ ��������� = "+tx1, raLeft, rfsDefault, 0, 0, 2, 18)
    Endif

    If Atc('tx2', ffi)>0
      .WriteParagraph("    ������������ ����������� ������ = "+tx2, raLeft, rfsDefault, 0, 0, 2, 18)
    Endif

    If Atc('tx3', ffi)>0
      .WriteParagraph("    ����� = "+tx3, raLeft, rfsDefault, 0, 0, 2, 18)
    Endif

    If Atc('tx4', ffi)>0
      .WriteParagraph("    ��� ���������� = "+tx4, raLeft, rfsDefault, 0, 0, 2, 18)
    Endif

    If Atc('tx5', ffi)>0
      .WriteParagraph("    ��� ��������� �������� = "+tx5, raLeft, rfsDefault, 0, 0, 2, 18)
    Endif

    If Atc('tx6', ffi)>0
      .WriteParagraph("    ������������ ��������� �� ���������� ��. = "+tx6, raLeft, rfsDefault, 0, 0, 2, 18)
    Endif

    If Atc('tx7', ffi)>0
      .WriteParagraph("    ��� �������� ���. = "+tx7, raLeft, rfsDefault, 0, 0, 2, 18)
    Endif

    If Atc('kus4', ffi)>0
      .WriteParagraph("    ���� ���. � ������ ��. �������� = "+kus4, raLeft, rfsDefault, 0, 0, 2, 18)
    Endif

    If Atc('tx9', ffi)>0
      .WriteParagraph("    �������� ������ (0 - ���) (1 - ��) = "+tx9, raLeft, rfsDefault, 0, 0, 2, 18)
    Endif
  Endif

  .WriteParagraph("", raLeft, rfsDefault, 0, 0, 2, 18)
  .SetAlignment(raCenter)
  .BeginTable                           && ������ �������
  .SetColumnsCount(2)
  .m_arTblWidths(1) = .Twips(2)         && ������ ������� (� ������� - ��)
  .m_arTblWidths(2) = .Twips(12)         && ������ ������� (� ������� - ��)

  .SetFont(3, 20, rfsBold)
  .SetupColumns()
  .m_arTblValues(1) = "���"
  .m_arTblValues(2) = "������������ ���������"
  .WriteRow()               && ������� �������� .m_arTblValues � ����� �������

  Scan
    I = I+1

    Wait '����� ���������� � �������: '+Str(I,18) Window Nowait

    For x = 1 To 2
      .m_arTblAlign(x) = raLeft
    Next x

    .SetFont(3, 18, rfsDefault)
    .SetupColumns()
    .m_arTblAlign(1) = raRight
    .m_arTblAlign(2) = raLeft
    .m_arTblValues(1) = BIC
    .m_arTblValues(2) = Alltrim(NAMEP)
    .WriteRow()               && ������� �������� .m_arTblValues � ����� �������


  Endscan && -------- ����� ����� �� ������� dbf

  Goto trr

  .SetFont(3, 18, rfsBold)
  .m_arTblValues(1) = '�����:  '
  .m_arTblValues(2) = Str(I) && ssa Alltrim(Str(I,18))
  .WriteRow()               && ������� �������� .m_arTblValues � ����� �������

  .EndTable

  .CloseFile

* --------- ������� ��� !!!
*    If(bWordStart_)
*      DECLARE Integer GetFocus IN WIN32API
*      DECLARE Integer ShellExecute IN SHELL32 INTEGER, STRING, STRING, STRING, STRING, INTEGER
*      hWnd = GetFocus()
*      If (hWnd != 0)
*        result=ShellExecute(hWnd, "open", cFileName_, "", "", 5)
*      Else
*        Messagebox("���� ������ ������������ ������ �����������!", 48,"������!")
*      EndIf
*    EndIf
* ----------

Endwith

parl = 'wordpad.exe'+' '+cFileName_
loWshShell=Createobject("WScript.Shell")
loWshShell.Run(parl, 1, .F.) && .F. �� ����� ���������� notepad.exe
Release loWshShell

Wait Clear
Deactivate Popup _3mp
Release Popups _3mp
Return
*--------------------------------------------------------------------
Function pRTF2(bWordStart_, cFileName_)

Wait '������ ������������ ��������.... ' Window Nowait
Hide Popup _1mq

pal02=Alias()

* f02 = cFileName_ && ���� ������

Erase (cFileName_)

rr02=Recno()
Go Top

oFile = Createobject("CRtfFile", cFileName_,.T.)
With (oFile)
  .DefaultInit
  .WriteHeader
*    .PageA4
  .PageA4LandScape

  .WriteParagraph("", raLeft, rfsDefault, 0, 0, 2, 24)

  .BeginTable                           && ������ �������
  .SetColumnsCount(2)
  .m_arTblWidths(1) = .Twips(8)         && ������ ������� (� ������� - ��)
  .m_arTblWidths(2) = .Twips(5)         && ������ ������� (� ������� - ��)

  .SetFont(3, 20, rfsDefault)
  .SetupColumns()
  tt=0
  Scan
    tt=tt+1

    .m_arTblValues(1) = Alltrim(pNames)
    If tt=2
      .SetFont(3, 20, rfsBold)
    Else
      .SetFont(3, 20, rfsDefault)
    Endif
    .m_arTblValues(2) = Alltrim(pZnach)
    .WriteRow()               && ������� �������� .m_arTblValues � ����� �������
  Endscan

  .EndTable

  .WriteParagraph("", raLeft, rfsDefault, 0, 0, 2, 24)

  Select acc807
  ror=Recno()

  .BeginTable                           && ������ �������
  .SetColumnsCount(9)
  .m_arTblWidths(1) = .Twips(4.5)         && ������ ������� (� ������� - ��)
  .m_arTblWidths(2) = .Twips(2)         && ������ ������� (� ������� - ��)
  .m_arTblWidths(3) = .Twips(2)
  .m_arTblWidths(4) = .Twips(1.2)
  .m_arTblWidths(5) = .Twips(2.2)
  .m_arTblWidths(6) = .Twips(1)
  .m_arTblWidths(7) = .Twips(1.2)
  .m_arTblWidths(8) = .Twips(2)
  .m_arTblWidths(9) = .Twips(1.5)

  .SetFont(3, 20, rfsBold)
  .SetupColumns()

  For zz=1 To 9
    .m_arTblAlign(zz) = raCenter
  Next zz

  .m_arTblValues(1) = '����'
  .m_arTblValues(2) = '���� ����.'
  .m_arTblValues(3) = '���� ����.'
  .m_arTblValues(4) = '������'
  .m_arTblValues(5) = '��� ���'
  .m_arTblValues(6) = '�.����'
  .m_arTblValues(7) = '��� ��.'
  .m_arTblValues(8) = '���� �����.'
  .m_arTblValues(9) = '��� �����������'
  .WriteRow()               && ������� �������� .m_arTblValues � ����� �������
  .SetFont(3, 20, rfsDefault)

  kk=0
  Do While a807.BIC=BIC
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
    .WriteRow()               && ������� �������� .m_arTblValues � ����� �������
    Skip
  Enddo
  .m_arTblValues(1) = '�����: '
  .m_arTblValues(2) = ' '+Str(kk) && ssa Alltrim(Str(kk,18))
  .m_arTblValues(3) = ''
  .m_arTblValues(4) = ''
  .m_arTblValues(5) = ''
  .m_arTblValues(6) = ''
  .m_arTblValues(7) = ''
  .m_arTblValues(8) = ''
  .m_arTblValues(9) = ''
  .SetFont(3, 20, rfsBold)
  .WriteRow()               && ������� �������� .m_arTblValues � ����� �������
  .EndTable

  .WriteParagraph("", raLeft, rfsDefault, 0, 0, 2, 24)

* GO ror
  Select (pal02)
  Go rr02

  .CloseFile  && �������� �����

Endwith

Wait Clear

Local loWshShell As Wscript.Shell
parms = 'wordpad.exe'+' '+cFileName_
loWshShell=Createobject("WScript.Shell")
loWshShell.Run(parms, 1, .F.) && .F. �� ����� ���������� wordpad.exe
Release loWshShell

Deactivate Popup _1mq
Release Popups _1mq

Return
*--------------------------------------------------------------------
Function frtfcompare(bWordStart_, cFileName_)

Erase (cFileName_)

al44=Alias()

fltt=Filter(al)

Select 0
Use (qqo) Alias d807
Index On BIC Tag BIC
If Len(fltt)>0
  Set Filter To &fltt
Endif

oFile = Createobject("CRtfFile", cFileName_,.T.)
With (oFile)
  .DefaultInit
  .WriteHeader
  .PageA4
*    .PageA4LandScape
  tqtmp = '��������� ���������� ����������� ��� (ED807) �� ����: '+Dtoc(fr_start.Text1.Value)+' � '+Dtoc(w_com_d.Text1.Value)
  .WriteParagraph(tqtmp, raLeft, rfsBold, 0, 0, 2, 30)
  .WriteParagraph("", raLeft, rfsDefault, 0, 0, 2, 24)

  ffi = Filter()
  If Len(ffi)>0 && ���� ���� ������, �� �������� � �� ��������� � ����

    .WriteParagraph("  ������:", raLeft, rfsBold, 0, 0, 2, 24)

    If Atc('tx1', ffi)>0
      .WriteParagraph("    ������������ ��������� = "+tx1, raLeft, rfsDefault, 0, 0, 2, 18)
    Endif

    If Atc('tx2', ffi)>0
      .WriteParagraph("    ������������ ����������� ������ = "+tx2, raLeft, rfsDefault, 0, 0, 2, 18)
    Endif

    If Atc('tx3', ffi)>0
      .WriteParagraph("    ����� = "+tx3, raLeft, rfsDefault, 0, 0, 2, 18)
    Endif

    If Atc('tx4', ffi)>0
      .WriteParagraph("    ��� ���������� = "+tx4, raLeft, rfsDefault, 0, 0, 2, 18)

    Endif

    If Atc('tx5', ffi)>0
      .WriteParagraph("    ��� ��������� �������� = "+tx5, raLeft, rfsDefault, 0, 0, 2, 18)
    Endif

    If Atc('tx6', ffi)>0
      .WriteParagraph("    ������������ ��������� �� ���������� ��. = "+tx6, raLeft, rfsDefault, 0, 0, 2, 18)
    Endif

    If Atc('tx7', ffi)>0
      .WriteParagraph("    ��� �������� ���. = "+tx7, raLeft, rfsDefault, 0, 0, 2, 18)
    Endif

    If Atc('kus4', ffi)>0
      .WriteParagraph("    ���� ���. � ������ ��. �������� = "+kus4, raLeft, rfsDefault, 0, 0, 2, 18)
    Endif

    If Atc('tx9', ffi)>0
      .WriteParagraph("    �������� ������ (0 - ���) (1 - ��) = "+tx9, raLeft, rfsDefault, 0, 0, 2, 18)
    Endif

  Endif  && IF LEN(ffi)>0

  If w_com_d.Text1.Value > fr_start.Text1.Value

    .SetFont(2, 28, rfsBold)
    .WriteParagraph(" �������� ��������� ��������: ", raLeft, rfsUnderline, 0, 0, 2, 28)  && rfsBold
    .WriteParagraph(" ", raLeft, rfsDefault, 0, 0, 2, 24)
    .BeginTable                           && ������ �������
    .SetColumnsCount(2)
    .m_arTblWidths(1) = .Twips(2)         && ������ ������� (� ������� - ��)
    .m_arTblWidths(2) = .Twips(12)         && ������ ������� (� ������� - ��)
    .SetFont(3, 20, rfsDefault)
    .SetupColumns()
*~~~~
    Do p_com1
*~~~~
    .SetFont(3, 20, rfsBold)
    .m_arTblValues(1) = "�����:"
    .m_arTblValues(2) = Str(kuch1)  && ssa Alltrim(Str(kuch1,18))
    .WriteRow()               && ������� �������� .m_arTblValues � ����� �������
    .EndTable
    .WriteParagraph("", raLeft, rfsDefault, 0, 0, 2, 24)

    .SetFont(2, 28, rfsBold)
    .WriteParagraph(" ����� ��������� ��������: ", raLeft, rfsUnderline, 0, 0, 2, 28)
    .WriteParagraph(" ", raLeft, rfsDefault, 0, 0, 2, 24)
    .BeginTable                           && ������ �������
    .SetColumnsCount(2)
    .m_arTblWidths(1) = .Twips(2)         && ������ ������� (� ������� - ��)
    .m_arTblWidths(2) = .Twips(12)         && ������ ������� (� ������� - ��)
    .SetFont(3, 20, rfsDefault)
    .SetupColumns()
*~~~~
    Do p_com2
*~~~~
    .SetFont(3, 20, rfsBold)
    .m_arTblValues(1) = "�����:"
    .m_arTblValues(2) = Str(kuch2)  && ssa Alltrim(Str(kuch2,18))
    .WriteRow()               && ������� �������� .m_arTblValues � ����� �������
    .EndTable
    .WriteParagraph("", raLeft, rfsDefault, 0, 0, 2, 24)

  Else

    .SetFont(2, 28, rfsBold)
    .WriteParagraph(" ����� ��������� ��������: ", raLeft, rfsUnderline, 0, 0, 2, 28)
    .WriteParagraph(" ", raLeft, rfsDefault, 0, 0, 2, 24)
    .BeginTable                           && ������ �������
    .SetColumnsCount(2)
    .m_arTblWidths(1) = .Twips(2)         && ������ ������� (� ������� - ��)
    .m_arTblWidths(2) = .Twips(12)         && ������ ������� (� ������� - ��)
    .SetFont(3, 20, rfsDefault)
    .SetupColumns()
*~~~~
    Do p_com1
*~~~~
    .SetFont(3, 20, rfsBold)
    .m_arTblValues(1) = "�����:"
    .m_arTblValues(2) = Str(kuch1)  && ssa Alltrim(Str(kuch1,18))
    .WriteRow()               && ������� �������� .m_arTblValues � ����� �������
    .EndTable
    .WriteParagraph("", raLeft, rfsDefault, 0, 0, 2, 24)


    .SetFont(2, 28, rfsBold)
    .WriteParagraph(" �������� ��������� ��������: ", raLeft, rfsUnderline, 0, 0, 2, 28)
    .WriteParagraph(" ", raLeft, rfsDefault, 0, 0, 2, 24)
    .BeginTable                           && ������ �������
    .SetColumnsCount(2)
    .m_arTblWidths(1) = .Twips(2)         && ������ ������� (� ������� - ��)
    .m_arTblWidths(2) = .Twips(12)         && ������ ������� (� ������� - ��)
    .SetFont(3, 20, rfsDefault)
    .SetupColumns()
*~~~~
    Do p_com2
*~~~~

    .SetFont(3, 20, rfsBold)
    .m_arTblValues(1) = "�����:"
    .m_arTblValues(2) = Str(kuch2)  && ssa Alltrim(Str(kuch2,18))
    .WriteRow()               && ������� �������� .m_arTblValues � ����� �������
    .EndTable
    .WriteParagraph("", raLeft, rfsDefault, 0, 0, 2, 24)

  Endif

  .CloseFile  && �������� �����
Endwith

Use In d807

Select (al44)
Go Top
*=FPUTS(f01,' --------------------------- ')
*=FCLOSE(f01)

Local loWshShell As Wscript.Shell
parms = 'wordpad.exe'+' '+cFileName_
loWshShell=Createobject("WScript.Shell")
loWshShell.Run(parms, 1, .F.) && .F. �� ����� ���������� wordpad.exe
Release loWshShell

Return
*--------------------------------------------------------------------

Function IsFileExists
Lparameters ldDate
Local dat77 As Date
dat77 = Dtos(ldDate)
Return File(pathdata+'a807'+dat77+'.dbf') And File(pathdata+'acc807'+dat77+'.dbf') And File(pathdata+'h807'+dat77+'.dbf')

Define Class SerchTxt As TextBox
  ForntSize = 11
  Format = 'T'
  Height = 1.44
  Width  = 80

  Procedure RightClick
  Do vs_menu

Enddefine
