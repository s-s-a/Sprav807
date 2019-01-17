If Messagebox('Вы действительно хотите закрыть приложение?',4+32+256,'Закрытие приложения')>=7
	Return
Endif

flag_sd = 1

Do CloseMutex With .F.  &&  По окончании работы приложения надо удалить объект Mutex, хотя это и не обязательно
_Screen.Caption = 'Microsoft Visual Foxpro'

Clear Events
On Shutdown
* CLEAR ALL
fr_start.Release
Close All
Quit
