# watcher
### Скрипт ограничивающий количество запускаемых процессов на терминальном сервере (Visual Basic script)

#### Установка скрипта
1. Скопировать файлы в папку %systemroot%\scripts (создать папку scripts, если отсутствует)  
2. Добавить в автозагрузку ОС команду:  
cscript.exe %systemroot%\scripts\watcher.vbs -f %systemroot%\scripts\watcher.cfg  

#### Настройка
В файле watcher.cfg с каждой новой строки (в столбик) пишется "<процесс.exe> <максимальное число процессов>"  

#### Логирование
Лог записывается в файл %systemroot%\scripts\watcher.log  
