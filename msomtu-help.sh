if [[ "${LANG%\.*}" == "ru_RU" ]]; then
	
	if [[ $cmd_all -eq 1 ]]; then
	printb "НАЗНАЧЕНИЕ:" 
	print-column 0 $p4 "" "msomtu - скрипт конфигурации состава компонентов Microsoft Office (MSO)." 
	echo
	
	printb "ОПИСАНИЕ:"
	print-column 0 $p4 "" "В Microsoft Office 2016 for Mac реализована изолированная архитектура компонентов (песочница), поэтому приложения хранят все необходимые ресурсы в своем собственном контейнере, и поэтому многие компоненты задублированы, что приводит к бесполезной трате дискового пространства. Этот скрипт безопасно удаляет «лишние» части следующих компонентов: языковые пакеты локализации; языковые файлы проверки правописания; файлы описаний шрифтов (.plist); задублированные системные шрифты. Скрипт также выполняет резервирование или копирование шрифтов в заданную или предопределенную папку." 
	echo
	
	printb "ПРИМЕЧАНИЯ:"
	print-column 0 $p6 "" "В скрипте реализован принцип безопасного исполнения - \"защита от дураков\". Режим исполнения по умолчанию - эмуляция. Скрипт не сможет внести изменения в систему без параметра '-run'. Параметр '-cache' не зависит от '-run'." '-'
	print-column 0 $p6 "" "Так как приложения устанавливаются с привилегиями root, для запуска скрипта (когда нужно сделать изменения) необходимо использовать sudo." '-'
	print-column 0 $p6 "" "Удаление шрифтов работает только с версией MSO 15.17 и выше, поскольку Microsoft изменила организацию шрифтов. В папке 'Fonts' обязательные шрифты, в 'DFonts' - не обязательные." '-'
	print-column 0 $p6 "" "Если удаляете шрифты, также удаляйте файлы списков шрифтов. Папка 'DFonts' и файлы списков шрифтов не обязательны - их можно удалять. Некоторые некириллические шрифты могут быть полезны - сохраните их перед удалением." '-'
	print-column 0 $p6 "" "Внимание: не удаляйте шрифты из папки 'Fonts'! Они необходимы для работы самих приложений." '-'
	print-column 0 $p6 "" "Предопределенные шрифтовые наборы не пересекаются." '-'
	print-column 0 $p6 "" "Все файловые операции регистронезависимы." '-'
	print-column 0 $p6 "" "Скрипт обрабатывает только именные параметры." '-'
	print-column 0 $p6 "" "Скрипт необходимо запускать после каждого обновления программ MSO." '-'
	print-column 0 $p6 "" "Установки по умолчанию для параметров '-lang' и '-proof': english и russian. Это зависит от системного языка и здравого смысла: для целостности MSO английский лучше оставить. Установки по умолчанию, специфичные для вашей системы, настраиваются прямо в скрипте." '-'
	echo
	fi
	
	printb "ИСПОЛЬЗОВАНИЕ:"
	print-column 0 $p4 "" "[sudo] $util [-<parameter> [<arguments>]]..."
	echo
	print-column 0 $p4 "" "[sudo] $util [-backup|-fcopy [<destination>]] [-app [<app>]] [-font [<font_pattern>]] [-ex|-x <font_pattern>] [-run]"
	echo
	print-column 0 $p4 "" "[sudo] $util [-app [\"<app_list>\"]] [-lang|-ui [\"<lang_list>\"]] [-proof|-p [\"<proof_list>\"]] [-font [<font_pattern>]] [-flist|-fl] [-ex|-x <font_pattern>] [-cache] [-report|-rep] [-verbose|-verb] [-fontset|-fs] [-all|-full] [-rev] [-help|-h|-?] [-run]"
	echo
	
	if [[ $cmd_all -eq 1 ]]; then
#	p30=43
#	local fmt21='%*s %-40s  %s %b\n'
	local mp4=-4; p12=12
	printb "СЦЕНАРИИ ИСПОЛЬЗОВАНИЯ:"
	print-padding $mp4 "- Просмотр статистики MSO" 
		print-column 0 $p12 "" "Параметр '-report'."
		echo
	print-padding $mp4 "- Оценка экономии дискового пространства (в режиме просмотра)" 
		print-column 0 $p12 "" "Параметр '-verbose' в комбинации с выбранным компонентом: font, flist, lang, proof."
		echo
	print-padding $mp4 "- Просмотр/Удаление языков локализации" 
		print-column 0 $p12 "" "Параметр '-lang'. Значение параметра - список кодов языков, разделенных пробелом. Пустое значение - удаление всех языков, кроме en и ru."
		echo
	print-padding $mp4 "- Просмотр/Удаление языковых пакетов правописания" 
		print-column 0 $p12 "" "Параметр '-proof'. Значение параметра - список названий языков, разделенных пробелом. Пустое значение - удаление всех языков, кроме english и russian."
		echo
	print-padding $mp4 "- Просмотр/Удаление файлов списков шрифтов" 
		print-column 0 $p12 "" "Параметр '-fontlist'."
		echo
	print-padding $mp4 "- Просмотр/Удаление шрифтов" 
		print-column 0 $p12 "" "Параметр '-font' с указанием списка шаблонов поиска. Шаблоном поиска может быть маска, например arial*, или имя шрифтового набора. Пустой шаблон удаляет папку 'DFonts'."
		echo
	print-padding $mp4 "- Удаление кэша шрифтов" 
		print-column 0 $p12 "" "Параметр '-cache'."
		echo
	print-padding $mp4 "- Просмотр наборов шрифтов" 
		print-column 0 $p12 "" "Параметр '-fontset'."
		echo
	print-padding $mp4 "- Поиск новых шрифтов" 
		print-column 0 $p12 "" "Параметр '-font -rev'."
		echo
	print-padding $mp4 "- Резервное копирование шрифтов" 
		print-column 0 $p12 "" "Параметр '-backup'."
		echo
	print-padding $mp4 "- Копирование шрифтов в библиотеки шрифтов" 
		print-column 0 $p12 "" "Параметр '-backup'."
		echo
	fi
	
	printb "АРГУМЕНТЫ:"
	print-column $p4 $p20 "app_list" "Список выбора приложений. w - Word, e - Excel, p - PowerPoint, o - Outlook, n - OneNote. Значение по умолчанию - 'w e p o n'." ":"
	print-column $p4 $p20 "lang_list" "Список языков локализации: ru pl de и т.д.; смотрите имена файлов с параметром '-verb'. Значение по умолчанию: 'en ru'." ":"
	print-column $p4 $p20 "proof_list" "Список языков проверки правописания: russian finnish german и т.д.; смотрите имена файлов с параметром '-verb'. Можно использовать символ подстановки '*'. Значение по умолчанию: 'english russian'." ":"
	print-column $p4 $p20 "font_pattern" "Операции со шрифтами выполняются на основе шаблонов. Шаблоны имен шрифтов: пустой - удаляется папка 'DFonts' (по умолчанию); <fontset> - имя предопределенного набора; <mask> - маска имен шрифтов: *.*, arial*, *.ttc и т.д. При использовании одиночного символа подстановки '*', заключайте его в двойные кавычки: \"*\". Предопределенные наборы: library, $fs. См. описание параметра '-fontset' и код скрипта. Набор 'library' удаляет дубликаты из системной и пользовательской библиотек шрифтов; так как сравнение выполняется по файлам (а не по семействам шрифтов), то совпадение не всегда будет полным. Также можно задавать списки наборов." ":"
	print-column $p4 $p20 "destination" "Путь к папке резервного копирования шрифтов. Значение по умолчанию: '~/Desktop/MSOFonts'. Можно применять синонимы предопределенных путей: 'syslib' - системная библиотека; 'userlib' - пользовательская библиотека." ":"
	echo
	
	printb "ПАРАМЕТРЫ:"
	print-column $p4 $p20 "-app" "Фильтр <app_list>. Список приложений, к которым применяются действия." ":"
	print-column $p4 $p20 "-lang" "Исключающий фильтр <lang_list>. Удаление языковых файлов локализации, кроме языков по умолчанию и выбора пользователя. Также см. параметр '-rev'; он инвертирует выбор пользователя, кроме умолчаний." ":"
	print-column $p4 $p20 "-proof" "Исключающий фильтр <proof_list>. Удаление языковых файлов проверки правописания, кроме языков по умолчанию и выбора пользователя. Также см. параметр '-rev'; он инвертирует выбор пользователя, кроме умолчаний." ":"
	print-column $p4 $p20 "-font" "Фильтр <font_pattern>. Удаление заданных шрифтов или папки 'DFonts'. Предопределенные наборы шрифтов: cyrdfonts, noncyr, chinese, sysfonts, symfonts. Для шрифтов параметр '-rev' отменяет выбор пользователя и изменяет функцию поиска: будет произведен поиск новых шрифтов. Это проверку имеет смысл производить после обновления программ MSO." ":"
	print-column $p4 $p20 "-backup" "Резервное копирование шрифтов в заданную или предопределенную папку. Если целевая папка не существует, она будет создана. Для копирования шрифтов можно использовать системную и пользовательскую библиотеки; см. раздел АРГУМЕНТЫ. Этот параметр отменяет все операции удаления." ":"
	print-column $p4 $p20 "-ex" "Исключающий фильтр <font_pattern> для параметра '-font'. Инвертирует выбор пользователя. В качестве 'font_pattern' можно использовать только маску." ":"
	print-column $p4 $p20 "-flist" "Ключ. Удаление файлов списков шрифтов (.plist)." ":"
	print-column $p4 $p20 "-all" "Ключ. Активация всех параметров очистки: lang, proof, font, flist, cache. Не влияет на параметр '-app'." ":"
	print-column $p4 $p20 "-cache" "Ключ. Очистка кэша шрифтов." ":"
	print-column $p4 $p20 "-verbose" "Ключ. Показывает детальную информацию по объектам выбора в режиме эмуляции." ":"
	print-column $p4 $p20 "-report" "Ключ. Показывает статистику по приложениям." ":"
	print-column $p4 $p20 "-fontset" "Ключ. Показывает предопределенные шрифтовые наборы." ":"
	print-column $p4 $p20 "-rev" "Ключ. Изменяет результат работы фильтров '-lang' и '-proof' на обратный." ":"
	print-column $p4 $p20 "-run" "Ключ. Разрешает режим изменений." ":"
	print-column $p4 $p20 "-help" "Ключ. Показывает страницу помощи. Есть два вида страницы: краткая и полная. По умолчанию (без параметров) выводится краткая страница. Для полного вида используйте параметр '-help -full'." ":"
	echo
	
	if [[ $cmd_all -eq 1 ]]; then
	printb "ПРИМЕРЫ:"
	p4=$((0-$p4)); p8=$((0-$p8))
	print-padding $p4 "Очистить все приложения с параметрами по умолчанию:"
	  print-padding $p8 "sudo $util -all -run" b
	print-padding $p4 "Показать статистику по приложениям:"
	  print-padding $p8 "$util -report" b
	print-padding $p4 "Показать установленные языковые пакеты локализации для Word и Excel:"
	  print-padding $p8 "$util -app \"w e\" -lang -verbose" b
	print-padding $p4 "Удалить заданные языковые пакеты локализации:"
	  print-padding $p8 "sudo $util -lang \"nl no de\" -rev -run" b
	print-padding $p4 "Удалить все языковые файлы правописания для Word:"
	  print-padding $p8 "sudo $util -proof -app w -run" b
	print-padding $p4 "Удалить заданные языковые файлы правописания:"
	  print-padding $p8 "sudo $util -proof \"Indonesian Isix*\" -rev -run" b
	print-padding $p4 "Найти в Word дубликаты шрифтов в библиотеках:"
	  print-padding $p8 "$util -font lib -app w -verbose" b
	print-padding $p4 "Удалить в Word дубликаты шрифтов в библиотеках:"
	  print-padding $p8 "sudo $util -font lib -app w -run" b
	print-padding $p4 "Удалить шрифты Arial и из набора 'chinese':"
	  print-padding $p8 "sudo $util -font \"chinese arial*\" -run" b
	print-padding $p4 "Найти новые шрифты в Outlook:"
	  print-padding $p8 "$util -font -rev -app o" b
	print-padding $p4 "Удалить в Word все шрифты, кроме заданных:"
	  print-padding $p8 "sudo $util -font *.* -ex \"brit* rockwell*\" -app w -run" b
	print-padding $p4 "Очистить кэш шрифтов:"
	  print-padding $p8 "sudo $util -cache" b
	print-padding $p4 "Выполнить резерное копирование выбранных шрифтов по умолчанию:"
	  print-padding $p8 "$util -backup -font \"cyrdfonts britanic*\" -run" b
	print-padding $p4 "Скопировать оригинальные русские шрифты в системную библиотеку:"
	  print-padding $p8 "sudo $util -backup syslib -font cyrdfonts -run" b
	print-padding $p4 "Показать предопределенные наборы шрифтов:"
	  print-padding $p8 "$util -fontset" b
	echo
	fi
	exit 0
fi

# template block for help in your language
#if [[ "${LANG%\.*}" == "????" ]]; then
#	if [[ $cmd_all -eq 1 ]]; then
#	printb "SYNOPSIS:"
#	echo
#	printb "DESCRIPTION:"
#	echo 
#	printb "NOTES:"
#	echo 
#	fi
#	printb "USAGE:"
#	print-column 0 $p4 "" "[sudo] $util [-<parameter> [<arguments>]]..."
#	echo
#	print-column 0 $p4 "" "[sudo] $util [-backup|-fcopy [<destination>]] [-app [<app>]] [-font [<font_pattern>]] [-ex|-x <font_pattern>] [-run]"
#	echo
#	print-column 0 $p4 "" "[sudo] $util [-app [\"<app_list>\"]] [-lang|-ui [\"<lang_list>\"]] [-proof|-p [\"<proof_list>\"]] [-font [<font_pattern>]] [-flist|-fl] [-ex|-x <font_pattern>] [-cache] [-report|-rep] [-verbose|-verb] [-fontset|-fs] [-all|-full] [-rev] [-help|-h|-?] [-run]"
#	echo 

#	if [[ $cmd_all -eq 1 ]]; then
#	local mp4=-4 p12=12
#	printb "USE CASES:"
#	echo 
#	fi
#	printb "ARGUMENTS:"
#	echo 
#	printb "PARAMETERS:"
#	echo 
#	if [[ $cmd_all -eq 1 ]]; then
#	printb "EXAMPLES:"
#	echo 
#	fi
#	exit 0
#fi

echo "No help for your language? Use parameter '-nl' for english."