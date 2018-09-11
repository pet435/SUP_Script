'Fedotov_PV сентябрь 2012. Для службы поддержки DIRECTUM
'Сценарий запускает поиск обращений в системе ТЕХКАСНПО
'VBScript.


  'Константы или значения "по умолчанию"
  SYSCODE = "XXXXXXX"									'Код системы ТЕХКАС
  ScriptName = "Открытие справочника ПДД с фильтрацией" 'Имя сценария
  
  'Создать подключение
  DIM Login, App, Scripts, Script
  SET Login = CreateObject("SBLogon.LoginPoint")
  SET App = Login.GetApplication("SystemCode=" & SYSCODE)
  'Получить фабрику сценариев
  SET Scripts = App.ScriptFactory
  SET Script = Scripts.GetObjectByName(ScriptName)
  'Используем гашение исключения, т.к. действие "отмена" в сценариях IS-Builder возвращает исключение, которое мешает этому сценарию завершиться.
  On Error resume next
  Script.Execute
  
  ' АЛЬТЕРНАТИВНЫЙ ВАРИАНТ, более медленный
  'Получить рабочую папку приложения
  'ApplicationFolder = App.Connection.SystemInfo.CorePath 
  'Пример запуска из командной строки
  '"C:\Program Files\NPO Computer\IS-Builder 7.10.1\SBLauncher.exe" -SYS=XXXXXX -CT=Script -F="Открытие справочника ПДД с фильтрацией"
  'SET Shell = CreateObject("WScript.Shell")
  ' экранирование ковычек осуществляется символом ковычек. Запись & """" означает, что к строке добавляются одни ковычки
  'CMD = """" & ApplicationFolder & "SBLauncher.exe" & """" & "-SYS=" & SYSCODE & " -CT=Script -F=""" & ScriptName & """"
  'Shell.Run CMD
  
  'Конец сценария
  'MsgBox("End")
