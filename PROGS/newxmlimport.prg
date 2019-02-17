Lparameters ;
  lcXMLFileName as String, ; && откуда взять (входной XML-файл)
  lcBnkSeekPath as String    && куда положить BNKSEEK.DBF

lcBnkSeekPath = Iif(Empty(lcBnkSeekPath) or Not Vartype(lcBnkSeekPath)='C', PathData, Addbs(Alltrim(lcBnkSeekPath)))

ox = Createobject('xmladapter')
ox.XMLSchemaLocation = FullPath('ED807.xsd')
With ox
  If .LoadXML(lcXMLFileName, .T., .T.) And .Tables.Count > 0
    .Tables(5).ToCursor() && BICdirectoryEntry
    .Tables(2).ToCursor() && ParticipantInfo
    Select Recno() As Pid, * From participantinfo Into Cursor PartInfo NOFILTER
    Select Recno() As Id, * From bicdirectoryentry Into Cursor Bic NOFILTER
    Select * From Bic b inner Join PartInfo p On Id = Pid into cursor tmp_BNKSEEK
    Copy To (Addbs(lcBnkSeekPath) + 'BNKSEEK') type fox2x as 866
  Endif
Endwith
