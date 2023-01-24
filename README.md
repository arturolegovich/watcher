# watcher
### Скрипт ограничивающий количество запускаемых процессов на терминальном сервере (Visual Basic script)
Источник: http://forum.oszone.net/thread-64395-4.html

#### Установка скрипта
1. Скопировать файлы в папку %systemroot%\scripts (создать папку scripts, если отсутствует)  
2. Добавить в автозагрузку ОС команду:  
cscript.exe %systemroot%\scripts\watcher.vbs -f %systemroot%\scripts\watcher.cfg  

#### Настройка
В файле watcher.cfg с каждой новой строки (в столбик) пишется "<процесс.exe> <максимальное число процессов>"  

#### Логирование
Лог записывается в файл %systemroot%\scripts\watcher.log  
```
24.01.2023 14:20:41 - -- ProcessName-notepad.exe, ProcessId-4808, strOwnerProcess-artur, strDomain-pc14
24.01.2023 14:20:41 - strNameOfUser-artur
24.01.2023 14:20:46 - -- ProcessName-notepad.exe, ProcessId-6132, strOwnerProcess-artur, strDomain-pc14
24.01.2023 14:20:46 - strNameOfUser-artur
24.01.2023 14:20:46 - strNameOfUser-artur
24.01.2023 14:20:50 - -- ProcessName-notepad.exe, ProcessId-5388, strOwnerProcess-artur, strDomain-pc14
24.01.2023 14:20:50 - strNameOfUser-artur
24.01.2023 14:20:50 - strNameOfUser-artur
24.01.2023 14:20:50 - strNameOfUser-artur
24.01.2023 14:20:50 - Count of process notepad.exe is 3... Terminate last process.
```

#### Доработать
Необходимо добавить список исключений (файл-список) с имена пользователей  
```
if strNameOfUser = "Director" then
CountUserProcess = 1
end if
```
