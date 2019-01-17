 IF MessageBox('Вы действительно хотите закрыть приложение?',4+32+256,'Закрытие приложения')>=7
   RETURN
 ENDIF
 
 flag_sd = 1
 
 Do CloseMutex with .F.  &&  По окончании работы приложения надо удалить объект Mutex, хотя это и не обязательно 
 _Screen.Caption = 'Microsoft Visual Foxpro'

 CLEAR EVENTS
 ON SHUTDOWN 
* CLEAR ALL 
 fr_start.Release
 CLOSE ALL
 QUIT 
