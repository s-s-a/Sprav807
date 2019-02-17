DatDel = Dtos(Date()-Int(Val(MyValue4)))
 
Dimension aMasks[5]
aMasks[1] = pathdata+'a807*.*'
aMasks[2] = pathdata+'h807*.*'
aMasks[3] = pathdata+'acc807*.*'
aMasks[4] = pathdata+'*_ED807_full.xml'
aMasks[5] = path_zip+'*ED01OSBR.zip'

For I = 1 To Alen(aMasks)
  Adir(aFiles, aMasks[i])
  For J = 1 To Alen(aFiles, 1)
    If JustStem(aFiles[j,1]) < JustStem(Strtran(Strtran(Upper(aMasks[i]), '*.*', '*'), '*', DatDel))
      WAIT WINDOW NOWAIT Addbs(Justpath(aMasks[i]))+aFiles[j,1]
      Erase (Addbs(Justpath(aMasks[i]))+aFiles[j,1])
    Endif
  Next
Next
