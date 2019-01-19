ox = Createobject('xmladapter')
ox.XMLSchemaLocation = '20190117_ED807.xsd'
?ox.LoadXML('20190117_ED807_full.xml', .T., .T.)

If ox.Tables.Count > 0
*ssa*    ox.Tables.Item(5).ToCursor() && BICdirectoryEntry
*ssa*    ox.Tables.Item(2).ToCursor() && ParticipantInfo
  For i =1 To ox.Tables.Count
    ?ox.Tables.Item(i).Alias
    ox.Tables.Item(i).ToCursor()
  Next
Endif
select Recno() as Pid, * from participantinfo into cursor PartInfo NOFILTER 
select Recno() as id, * from bicdirectoryentry into cursor Bic NOFILTER
select * from Bic b inner join PartInfo p on id = pid