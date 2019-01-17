Parameters xdt
Local _returndate
_changeDate = Createobject("cdate",xdt)
_changeDate.Show(1)
Return

Define Class daylabel As Label
	FontSize=13
	FontBold=.T.
	Height=16*1.4
	Width=18*1.4
	Alignment=2
	BorderStyle=1
	BackColor=Rgb(250,250,250)
	Caption=""
	Procedure Click
		With This
			Local vdt
*ssa*			vdt=ctod(this.caption+"."+Tran(this.parent.vmonth)+"."+Tran(this.parent.vyear))
			vdt=Date(.Parent.vyear, .Parent.vmonth, Val(.Caption))
			If ! Empty(vdt)
				Do Case
					Case Type("thisform.odat")="D"
						Thisform.odat=vdt
					Case Type("thisform.odat")="O"
						Thisform.odat.Value=vdt
						Thisform.odat.Refresh
				Endcase
			Endif
			.Parent.Release
		Endwith
	Endproc
	Procedure MouseMove
		Lparameters nButton, nShift, nXCoord, nYCoord
		If ! Empty(This.Caption)
			If This.Parent.vday!=Val(This.Caption)
				This.Parent.vday=Val(This.Caption)
				This.Parent.olm.Caption=Allt(Str(This.Parent.vday))+" "+This.Parent.mns[this.parent.vmonth]+" "+Allt(Str(This.Parent.vyear))
			Endif
		Endif
	Endproc
	Procedure RightClick
		Thisform.Release
	Endproc
Enddefine

Define Class cdcommandbutton As CommandButton
	FontBold=.T.
	FontSize=13
	Top=26 && 21
	Width=30*1.4
	Height=16*1.4
	Visible=.T.
	TabStop=.F.
Enddefine

Define Class dowlabel As Label
	FontSize=13
	FontBold=.F.
	Height=16*1.4
	Width=18*1.4
	Alignment=2
	BorderStyle=1
	Top=47 && 37
	BackColor=Rgb(0,0,254)
	ForeColor=Rgb(254,254,254)
Enddefine

Define Class cDate As Form
	odat=.F.
	Height=206 && 161 - основная форма
	Width=185  && 135 - основная форма
	BorderStyle=1
	TitleBar=0
	ShowTips=.T.
	vday=0
	vmonth=0
	vyear=0
	vd=Date()
	Declare mns[12]
	mns[1]="Января"
	mns[2]="Февраля"
	mns[3]="Марта"
	mns[4]="Апреля"
	mns[5]="Мая"
	mns[6]="Июня"
	mns[7]="Июля"
	mns[8]="Августа"
	mns[9]="Сентября"
	mns[10]="Октября"
	mns[11]="Ноября"
	mns[12]="Декабря"
	Add Object "obml" As "cdcommandbutton"
	Add Object "obmr" As "cdcommandbutton"
	Add Object "obyl" As "cdcommandbutton"
	Add Object "obyr" As "cdcommandbutton"
	Add Object "tmer" As "timer"
	tmer.Cycl=0
	Procedure Init
		Parameters  cDate
		If Parameters()=0
			cDate=Date()
		Endif
		_scw=_Screen.Width
		_sch=_Screen.Height
		_mc=Mcol("",3)
		_mr=Mrow("",3)
		Do Case
			Case (_sch-_mr>This.Height).And.(_scw-_mc>This.Width)
				This.Left=_mc-3
				This.Top=_mr-3
			Case (_sch-This.Height>0).And.(_scw-_mc>This.Width)
				This.Left=_mc-3
				This.Top=_mr+3-This.Height
			Case (_sch-_mr>This.Height).And.(_scw-This.Width>0)
				This.Left=_mc+3-This.Width
				This.Top=_mr-3
			Case (_sch-This.Height>0).And.(_scw-This.Width>0)
				This.Left=_mc+3-This.Width
				This.Top=_mr+3-This.Height
			Otherwise
				This.Left=_mc-This.Width/2
				This.Top=_mr-This.Height/2
		Endcase
		Set Date German
		This.odat=cDate
		Do Case
			Case Type("cDate")="D"
				This.vd=cDate
				This.vday=Iif(Empty(cDate),Day(Date()),Day(cDate))
				This.vmonth=Iif(Empty(cDate),Month(Date()),Month(cDate))
				This.vyear=Iif(Empty(cDate),Year(Date()),Year(cDate))
			Case Type("cDate")="O"
				This.vd=cDate.Value
				This.vday=Iif(Empty(cDate.Value),Day(Date()),Day(cDate.Value))
				This.vmonth=Iif(Empty(cDate.Value),Month(Date()),Month(cDate.Value))
				This.vyear=Iif(Empty(cDate.Value),Year(Date()),Year(cDate.Value))
		Endcase

		This.AddObject("olm","label")
		With This.olm
			.FontBold=.T.
			.FontSize=13
			.BorderStyle=1
			.BackColor=Rgb(255,255,255)
			.Top=2 && 2
			.Width=131*1.4
			.Alignment=2
			.Height=18*1.4
			.Visible=.T.
			.Left=(Thisform.Width-.Width)/2
		Endwith

		With This.obml
			.ToolTipText="Месяц назад"
			.Caption="<"
			.Left=34
		Endwith

		With This.obmr
			.ToolTipText="Следующий месяц"
			.Caption=">"
			.Left=Thisform.Width-64
		Endwith

		With This.obyl
			.ToolTipText="Год назад"
			.Caption="<<"
			.Left=2
		Endwith

		With This.obyr
			.ToolTipText="Следующий год"
			.Caption=">>"
			.Left=Thisform.Width-32
		Endwith

		For i=1 To 7
			This.AddObject("old"+Allt(Str(i)),"dowlabel")
			us="no=this.old"+Allt(Str(i))
			&us
			With no
				.Left=2+(i-1)*(.Width+1)
				Do Case
					Case i=1
						.Caption="Пн"
					Case i=2
						.Caption="Вт"
					Case i=3
						.Caption="Ср"
					Case i=4
						.Caption="Чт"
					Case i=5
						.Caption="Пт"
					Case i=6
						.Caption="Сб"
						.BackColor=Rgb(254,0,0)
						.ForeColor=Rgb(254,254,254)
					Case i=7
						.Caption="Вс"
						.BackColor=Rgb(254,0,0)
						.ForeColor=Rgb(254,254,254)
				Endcase
				.Visible=.T.
			Endwith
		Next
		nomb=0
		For N=1 To 6
			For i=1 To 7
				nomb=nomb+1
				This.AddObject("oldy"+Allt(Str(nomb)),"daylabel")
				us="no=this.oldy"+Allt(Str(nomb))
				&us
				With no
					.Top=68+(N-1)*(.Height+1)  && .top=57+(n-1)*(.height+1)  - это числа
					.Left=2+(i-1)*(.Width+1)
					.Visible=.T.
				Endwith
			Next
		Next
		This.Draw
	Endproc

	Procedure setday
		Local fdd,fd,i,N,us
		fdd=Ctod("01."+Allt(Str(This.vmonth))+"."+Allt(Str(This.vyear)))
		fd=Dow(fdd)
		fd=Iif(fd=1,7,fd-1)
		For i=1 To fd-1
			us="this.oldy"+Allt(Str(i))+".visible=.f."
			&us
		Next
		For i=1 To 40
			us="eobj=this.oldy"+Allt(Str(fd+i-1))
			&us
			eobj.Caption=Allt(Str(i))
			eobj.Visible=.T.
			edat=Ctod(Tran(i)+"."+Tran(This.vmonth)+"."+Tran(This.vyear))
			Do Case
				Case edat=This.vd
					eobj.BackColor=Rgb(250,150,150)
					eobj.ToolTipText='Изменяемая дата'
				Case edat=Date()
					eobj.BackColor=Rgb(150,250,150)
					eobj.ToolTipText='Сегодня'
				Otherwise
					eobj.BackColor=Rgb(250,250,250)
					eobj.ToolTipText=''
			Endcase
			If Month(fdd)!=Month(fdd+i)
				Exit
			Endif
		Next
		ezn=(fd+i-1)/7
		_colday=Iif(ezn=Int(ezn),ezn,Int(ezn)+1)
		This.Height=206-(17*(6-_colday))
		For N=fd+i To 42
			us="this.oldy"+Allt(Str(N))+".visible=.f."
			&us
		Next
	Endproc

	Procedure Draw
		This.vday=1
		This.olm.Caption=Allt(Str(This.vday))+" "+This.mns[this.vmonth]+" "+Allt(Str(This.vyear))
		Thisform.setday
	Endproc

	Procedure obml.Click
		If Thisform.vmonth=1
			Thisform.vmonth=12
			Thisform.vyear=Thisform.vyear-1
		Else
			Thisform.vmonth=Thisform.vmonth-1
		Endif
		Thisform.Draw
	Endproc

	Procedure obmr.Click
		If Thisform.vmonth=12
			Thisform.vmonth=1
			Thisform.vyear=Thisform.vyear+1
		Else
			Thisform.vmonth=Thisform.vmonth+1
		Endif
		Thisform.Draw
	Endproc

	Procedure obyl.Click
		Thisform.vyear=Thisform.vyear-1
		Thisform.Draw
	Endproc

	Procedure obyr.Click
		Thisform.vyear=Thisform.vyear+1
		Thisform.Draw
	Endproc

	Procedure Click
		This.Release
	Endproc

	Procedure tmer.Init
		This.Interval=100
	Endproc

	Procedure tmer.Timer
		Local xval,yval
		xval=Mrow(Thisform.Caption,3)
		yval=Mcol(Thisform.Caption,3)
		If xval<0.Or.yval<0
			This.Cycl=This.Cycl+1
		Else
			This.Cycl=0
		Endif
		If This.Cycl=7
			Thisform.Release
		Endif
	Endproc
Enddefine
