Parameters xdt
Local _returndate
_changeDate=createobject("cdate",xdt)
_changeDate.show(1)
Return

Define class daylabel as label
	FontSize=13
	FontBold=.t.
	Height=16*1.4
	Width=18*1.4
	Alignment=2
	BorderStyle=1
	BackColor=rgb(250,250,250)
	Caption=""
	Procedure click
		Local vdt
		vdt=ctod(this.caption+"."+allt(str(this.parent.vmonth))+"."+allt(str(this.parent.vyear)))
		If ! empty(vdt)
			Do case
				Case type("thisform.odat")="D"
					Thisform.odat=vdt
				Case type("thisform.odat")="O"
					Thisform.odat.value=vdt
					Thisform.odat.refresh
			Endcase
		Endif
		This.parent.release
	Endproc
	Procedure mousemove
		Lparameters nButton, nShift, nXCoord, nYCoord
		If ! empty(this.caption)
			If this.parent.vday!=val(this.caption)
				This.parent.vday=val(this.caption)
				This.parent.olm.caption=allt(str(this.parent.vday))+" "+this.parent.mns[this.parent.vmonth]+" "+allt(str(this.parent.vyear))
			Endif
		Endif
	Endproc
	Procedure rightclick
		Thisform.release
	Endproc
Enddefine

Define class cdcommandbutton as commandbutton
	FontBold=.t.
	FontSize=13
	Top=26 && 21
	Width=30*1.4
	Height=16*1.4
	Visible=.t.
	TabStop=.f.
Enddefine

Define class dowlabel as label
	FontSize=13
	fontbold=.f.
	height=16*1.4
	width=18*1.4
	alignment=2
	borderstyle=1
	top=47 && 37
	backcolor=rgb(0,0,254)
	forecolor=rgb(254,254,254)
Enddefine

Define class cDate as form
	odat=.f.
	Height=206 && 161 - основная форма
	Width=185  && 135 - основная форма
	BorderStyle=1
	TitleBar=0
	ShowTips=.t.
	vday=0
	vmonth=0
	vyear=0
	vd=date()
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
	Add object "obml" as "cdcommandbutton"
	Add object "obmr" as "cdcommandbutton"
	Add object "obyl" as "cdcommandbutton"
	Add object "obyr" as "cdcommandbutton"
	Add object "tmer" as "timer"
	tmer.cycl=0
	Procedure init
		Parameters  cDate
		If parameters()=0
			cDate=date()
		Endif
		_scw=_screen.width
		_sch=_screen.height
		_mc=mcol("",3)
		_mr=mrow("",3)
		Do case
			Case (_sch-_mr>this.height).and.(_scw-_mc>this.width)
				This.left=_mc-3
				This.top=_mr-3
			Case (_sch-this.height>0).and.(_scw-_mc>this.width)
				This.left=_mc-3
				This.top=_mr+3-this.height
			Case (_sch-_mr>this.height).and.(_scw-this.width>0)
				This.left=_mc+3-this.width
				This.top=_mr-3
			Case (_sch-this.height>0).and.(_scw-this.width>0)
				This.left=_mc+3-this.width
				This.top=_mr+3-this.height
			Otherwise
				This.left=_mc-this.width/2
				This.top=_mr-this.height/2
		Endcase
		Set date german
		This.odat=cDate
		Do case
			Case type("cDate")="D"
				This.vd=cDate
				This.vday=iif(empty(cDate),day(date()),day(cDate))
				This.vmonth=iif(empty(cDate),month(date()),month(cDate))
				This.vyear=iif(empty(cDate),year(date()),year(cDate))
			Case type("cDate")="O"
				This.vd=cDate.value
				This.vday=iif(empty(cDate.value),day(date()),day(cDate.value))
				This.vmonth=iif(empty(cDate.value),month(date()),month(cDate.value))
				This.vyear=iif(empty(cDate.value),year(date()),year(cDate.value))
		Endcase

		This.addobject("olm","label")
		With this.olm
			.fontbold=.t.
			.fontsize=13
			.borderstyle=1
			.backcolor=rgb(255,255,255)
			.top=2 && 2
			.width=131*1.4
			.alignment=2
			.height=18*1.4
			.visible=.t.
			.left=(thisform.width-.width)/2
		Endwith

		With this.obml
			.tooltiptext="Месяц назад"
			.caption="<"
			.left=34
		Endwith

		With this.obmr
			.tooltiptext="Следующий месяц"
			.caption=">"
			.left=thisform.width-64
		Endwith

		With this.obyl
			.tooltiptext="Год назад"
			.caption="<<"
			.left=2
		Endwith

		With this.obyr
			.tooltiptext="Следующий год"
			.caption=">>"
			.left=thisform.width-32
		Endwith

		For i=1 to 7
			This.addobject("old"+allt(str(i)),"dowlabel")
			us="no=this.old"+allt(str(i))
			&us
			With no
				.left=2+(i-1)*(.width+1)
				Do case
					Case i=1
						.caption="Пн"
					Case i=2
						.caption="Вт"
					Case i=3
						.caption="Ср"
					Case i=4
						.caption="Чт"
					Case i=5
						.caption="Пт"
					Case i=6
						.caption="Сб"
						.backcolor=rgb(254,0,0)
						.forecolor=rgb(254,254,254)
					Case i=7
						.caption="Вс"
						.backcolor=rgb(254,0,0)
						.forecolor=rgb(254,254,254)
				Endcase
				.visible=.t.
			Endwith
		Next
		nomb=0
		For n=1 to 6
			For i=1 to 7
				nomb=nomb+1
				This.addobject("oldy"+allt(str(nomb)),"daylabel")
				us="no=this.oldy"+allt(str(nomb))
				&us
				With no
					.top=68+(n-1)*(.height+1)  && .top=57+(n-1)*(.height+1)  - это числа
					.left=2+(i-1)*(.width+1)
					.visible=.t.
				Endwith
			Next
		Next
		This.draw
	Endproc

	Procedure setday
		Local fdd,fd,i,n,us
		fdd=ctod("01."+allt(str(this.vmonth))+"."+allt(str(this.vyear)))
		fd=dow(fdd)
		fd=iif(fd=1,7,fd-1)
		For i=1 to fd-1
			us="this.oldy"+allt(str(i))+".visible=.f."
			&us
		Next
		For i=1 to 40
			us="eobj=this.oldy"+allt(str(fd+i-1))
			&us
			eobj.caption=allt(str(i))
			eobj.visible=.t.
			edat=ctod(allt(str(i))+"."+allt(str(this.vmonth))+"."+allt(str(this.vyear)))
			Do case
				Case edat=this.vd
					eobj.backcolor=rgb(250,150,150)
					eobj.tooltiptext='Изменяемая дата'
				Case edat=date()
					eobj.backcolor=rgb(150,250,150)
					eobj.tooltiptext='Сегодня'
				Otherwise
					eobj.backcolor=rgb(250,250,250)
					eobj.tooltiptext=''
			Endcase
			If month(fdd)!=month(fdd+i)
				Exit
			Endif
		Next
		ezn=(fd+i-1)/7
		_colday=iif(ezn=int(ezn),ezn,int(ezn)+1)
		This.height=206-(17*(6-_colday))
		For n=fd+i to 42
			us="this.oldy"+allt(str(n))+".visible=.f."
			&us
		Next
	Endproc

	Procedure draw
		This.vday=1
		This.olm.caption=allt(str(this.vday))+" "+this.mns[this.vmonth]+" "+allt(str(this.vyear))
		Thisform.setday
	Endproc

	Procedure obml.click
		If thisform.vmonth=1
			Thisform.vmonth=12
			Thisform.vyear=thisform.vyear-1
		Else
			Thisform.vmonth=thisform.vmonth-1
		Endif
		Thisform.draw
	Endproc

	Procedure obmr.click
		If thisform.vmonth=12
			Thisform.vmonth=1
			Thisform.vyear=thisform.vyear+1
		Else
			Thisform.vmonth=thisform.vmonth+1
		Endif
		Thisform.draw
	Endproc

	Procedure obyl.click
		Thisform.vyear=thisform.vyear-1
		Thisform.draw
	Endproc

	Procedure obyr.click
		Thisform.vyear=thisform.vyear+1
		Thisform.draw
	Endproc

	Procedure click
		This.release
	Endproc

	Procedure tmer.init
		This.interval=100
	Endproc

	Procedure tmer.timer
		Local xval,yval
		xval=mrow(thisform.caption,3)
		yval=mcol(thisform.caption,3)
		If xval<0.or.yval<0
			This.cycl=this.cycl+1
		Else
			This.cycl=0
		Endif
		If this.cycl=7
			Thisform.release
		Endif
	Endproc
Enddefine
