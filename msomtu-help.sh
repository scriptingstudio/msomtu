# non-english help pages
if [[ "${LANG%\.*}" == "ru_RU" ]]; then
SYNOPSIS=(
	"НАЗНАЧЕНИЕ" 
	"MSOMTU - скрипт конфигурации состава компонентов Microsoft Office (MSO)." 
)
DESCRIPTION=(
	"ОПИСАНИЕ"
	"В Microsoft Office 2016 for Mac реализована изолированная архитектура компонентов (песочница), поэтому приложения хранят все необходимые ресурсы в своем собственном контейнере, и поэтому многие компоненты задублированы, что приводит к бесполезной трате дискового пространства. Этот скрипт безопасно удаляет «лишние» части следующих компонентов: языковые пакеты локализации; языковые файлы проверки правописания; файлы описаний шрифтов (.plist); задублированные системные шрифты. Скрипт также выполняет резервирование или копирование шрифтов в заданную или предопределенную папку."
)
NOTES=(
	"ПРИМЕЧАНИЯ"
	"В скрипте реализован принцип безопасного исполнения - \"защита от дураков\". Режим исполнения по умолчанию - эмуляция. Скрипт не сможет внести изменения в систему без параметра '-run'."
	
	"Так как приложения устанавливаются с привилегиями root, для запуска скрипта (когда нужно сделать изменения) необходимо использовать sudo." 
	
	"Удаление шрифтов работает только с версией MSO 15.17 и выше, поскольку Microsoft изменила организацию шрифтов. В папке 'Fonts' обязательные шрифты, в 'DFonts' - не обязательные."
	
	"Если удаляете шрифты, также удаляйте файлы списков шрифтов. Папка 'DFonts' и файлы списков шрифтов не обязательны - их можно удалять. Шрифты в папках 'DFonts' и 'Fonts' не видны другим приложениям. Некоторые некириллические шрифты могут быть полезны - сохраните их перед удалением."
	
	"Внимание: не удаляйте шрифты из папки 'Fonts'! Они необходимы для работы самих приложений."
	
	"Все файловые операции регистронезависимы." 
	"Скрипт обрабатывает только именные параметры."
	
	"Скрипт необходимо запускать после каждого обновления программ MSO."
	
	"Установки по умолчанию для параметров '-lang' и '-proof': english и russian. Это зависит от системного языка и здравого смысла: для целостности MSO английский лучше оставить. Установки по умолчанию, специфичные для вашей системы, настраиваются прямо в скрипте."
	
	"Классификация шрифтов в предопределенных наборах: кириллические, некириллические, иероглифические, символьные, системные. Предопределенные шрифтовые наборы не пересекаются."
)
USAGE=(
	"ИСПОЛЬЗОВАНИЕ"
	"[sudo] $util [-<parameter> [<arguments>]]..."
	
	"[sudo] $util [-backup|-fcopy [<destination>]] [-app [<app>]] [-font [<font_pattern>]] [-ex|-x <font_pattern>] [-run]"
	
	"[sudo] $util [-app [\"<app_list>\"]] [-lang|-ui [\"<lang_list>\"]] [-proof|-p [\"<proof_list>\"]] [-font [<font_pattern>]] [-flist|-fl] [-ex|-x <font_pattern>] [-cache] [-report|-rep] [-verbose|-verb [nl]] [-fontset|-fs] [-all|-full] [-rev] [-help|-h|-? [en]] [-run]"
)
USE_CASES=(
	"СЦЕНАРИИ ИСПОЛЬЗОВАНИЯ"
	"- Просмотр статистики MSO||Параметр '-report'."
	
	"- Оценка экономии дискового пространства (в режиме просмотра)||Параметр '-verbose' в комбинации с выбранным компонентом: font, flist, lang, proof."
	
	"- Просмотр/Удаление языков локализации||Параметр '-lang'. Значение параметра - список кодов языков, разделенных пробелом. Пустое значение - удаление всех языков, кроме en и ru."
	
	"- Просмотр/Удаление языковых пакетов правописания||Параметр '-proof'. Значение параметра - список названий языков, разделенных пробелом. Пустое значение - удаление всех языков, кроме english и russian."
	
	"- Просмотр/Удаление файлов списков шрифтов||Параметр '-fontlist'."
	
	"- Просмотр/Удаление шрифтов||Параметр '-font' с указанием списка шаблонов поиска. Шаблоном поиска может быть маска, например arial*, или имя шрифтового набора. Пустой шаблон удаляет папку 'DFonts'."
	
	"- Просмотр/Удаление дубликатов шрифтов||Параметр '-font lib' с аргументом 'lib' - библиотеки."
	"- Очитска кэша шрифтов||Параметр '-cache'."
	"- Просмотр наборов шрифтов||Параметр '-fontset'."
	"- Поиск новых шрифтов||Параметр '-font -rev'."
	"- Резервное копирование шрифтов||Параметр '-backup'."
	"- Копирование шрифтов в библиотеки шрифтов||Параметр '-backup'."
	"- Проверка новых версий||Параметр '-check'."
)
ARGUMENTS=(
	"АРГУМЕНТЫ"
	"app_list||Список выбора приложений. w - Word, e - Excel, p - PowerPoint, o - Outlook, n - OneNote. Значение по умолчанию - 'w e p o n'."
	
	"lang_list||Список языков локализации: ru pl de и т.д.; смотрите имена файлов с параметром '-verbose'. Значение по умолчанию: 'en ru'." 
	
	"proof_list||Список языков проверки правописания: russian finnish german и т.д.; смотрите имена файлов с параметром '-verbose'. Можно использовать символ подстановки '*'. Значение по умолчанию: 'english russian'." 
	
	"font_pattern||Операции со шрифтами выполняются на основе шаблонов. Шаблоны имен шрифтов: пустой - удаляется папка 'DFonts' (по умолчанию); <fontset> - имя предопределенного набора; <mask> - маска имен шрифтов: *.*, arial*, *.ttc и т.д. При использовании одиночного символа подстановки '*', заключайте его в двойные кавычки: \"*\". Предопределенные наборы: library, $fs. См. описание параметра '-fontset' и код скрипта. Набор 'library' удаляет дубликаты из системной и пользовательской библиотек шрифтов; так как сравнение выполняется по файлам (а не по семействам шрифтов), то совпадение не всегда будет полным. Также можно задавать списки наборов."
	
	"destination||Путь к папке резервного копирования шрифтов. Значение по умолчанию: '~/Desktop/MSOFonts'. Можно применять синонимы предопределенных путей: 'syslib' - системная библиотека; 'userlib' - пользовательская библиотека."
)
PARAMETERS=(
	"ПАРАМЕТРЫ"
	"-all||Ключ. Активация всех параметров очистки: lang, proof, font, flist, cache. Не влияет на параметр '-app'."

	"-app||Фильтр <app_list>. Список приложений, к которым применяются действия."	
	
	"-backup||Резервное копирование шрифтов в заданную или предопределенную папку. Если целевая папка не существует, она будет создана. Для копирования шрифтов можно использовать системную и пользовательскую библиотеки; см. раздел АРГУМЕНТЫ. Этот параметр отменяет все операции удаления."
	
	"-cache||Ключ. Очистка кэша шрифтов." 
	"-check||Ключ. Проверка новых версий на веб-сайте; открывает веб-страницу в браузере."
	
	"-ex||Исключающий фильтр <font_pattern> для параметра '-font'. Инвертирует выбор пользователя. В качестве 'font_pattern' можно использовать только маску."
	
	"-flist||Ключ. Удаление файлов списков шрифтов (.plist)." 
	
	"-font||Фильтр <font_pattern>. Удаление заданных шрифтов или папки 'DFonts'. Предопределенные наборы шрифтов: cyrdfonts, noncyr, chinese, sysfonts, symfonts. Для шрифтов параметр '-rev' отменяет выбор пользователя и изменяет функцию поиска: будет произведен поиск новых шрифтов. Это проверку имеет смысл производить после обновления программ MSO." 
	
	"-fontset||Ключ. Показывает предопределенные шрифтовые наборы." 
	
	"-help||Ключ. Показывает страницу помощи. Есть два вида страницы: краткая и полная. По умолчанию (без параметров) выводится краткая страница. Для полного вида используйте параметр '-help -full'. С аргументом 'en' выводит английскую страницу помощи."
	
	"-lang||Исключающий фильтр <lang_list>. Удаление языковых файлов локализации, кроме языков по умолчанию и выбора пользователя. Также см. параметр '-rev'; он инвертирует выбор пользователя, кроме умолчаний." 
	
	"-proof||Исключающий фильтр <proof_list>. Удаление языковых файлов проверки правописания, кроме языков по умолчанию и выбора пользователя. Также см. параметр '-rev'; он инвертирует выбор пользователя, кроме умолчаний."
	
	"-report||Ключ. Показывает статистику по приложениям."

	"-rev||Ключ. Изменяет результат работы фильтров '-lang' и '-proof' на обратный, кроме умолчаний. Для параметра '-font' изменяется тип поиска."
	"-run||Ключ. Разрешает режим изменений." 
	
	"-verbose||Ключ. Показывает детальную информацию по объектам выбора в режиме эмуляции. С аргументом 'nl' отменяет вывод списка файлов. Не зависит от '-run'."
)
EXAMPLES=(
	"ПРИМЕРЫ"
	"Очистить все приложения с параметрами по умолчанию||sudo $util -all -run"
	
	"Показать статистику по приложениям||$util -report"
	
	"Показать установленные языковые пакеты локализации для Word и Excel||$util -app \"w e\" -lang -verbose"
	
	"Удалить заданные языковые пакеты локализации||sudo $util -lang \"nl no de\" -rev -run"
	
	"Удалить все языковые файлы правописания для Word||sudo $util -proof -app w -run"
	
	"Удалить заданные языковые файлы правописания||sudo $util -proof \"Indonesian Isix*\" -rev -run"
	
	"Найти в Word дубликаты шрифтов в библиотеках||$util -font lib -app w -verbose"
	
	"Удалить в Word дубликаты шрифтов в библиотеках||sudo $util -font lib -app w -run"
	
	"Удалить шрифты Arial и из набора 'chinese'||sudo $util -font \"chinese arial*\" -run"
	
	"Найти новые шрифты в Outlook||$util -font -rev -app o"
	
	"Удалить в Word все шрифты, кроме заданных||sudo $util -font *.* -ex \"brit* rockwell*\" -app w -run"
	
	"Очистить кэш шрифтов||sudo $util -cache"
	
	"Выполнить резерное копирование выбранных шрифтов по умолчанию||$util -backup -font \"cyrdfonts britanic*\" -run"
	
	"Скопировать оригинальные русские шрифты в системную библиотеку||sudo $util -backup syslib -font cyrdfonts -run"
	
	"Показать предопределенные наборы шрифтов||$util -fontset"
)
mylang="${LANG%\.*}"
fi

### Template block for help in your language
### Copy the block below, uncomment, 
### and insert text in your language by the model
#if [[ "${LANG%\.*}" == "????" ]]; then
#	SYNOPSIS=("SYNOPSIS ???")
#	DESCRIPTION=("DESCRIPTION ???")
#	NOTES=("NOTES ???")
#	USAGE=(
#	"USAGE ???"
#	"[sudo] $util [-<parameter> [<arguments>]]..."
#	"[sudo] $util [-backup|-fcopy [<destination>]] [-app [<app>]] [-font [<font_pattern>]] [-ex|-x <font_pattern>] [-run]"
#	"[sudo] $util [-app [\"<app_list>\"]] [-lang|-ui [\"<lang_list>\"]] [-proof|-p [\"<proof_list>\"]] [-font [<font_pattern>]] [-flist|-fl] [-ex|-x <font_pattern>] [-cache] [-report|-rep] [-verbose|-verb [nl]] [-fontset|-fs] [-all|-full] [-rev] [-help|-h|-? [en]] [-run]"
#)
#	USE_CASES=("USE_CASES ???")
#	ARGUMENTS=(
#	"ARGUMENTS ???"
#	"app_list||."
#	"lang_list||."
#	"proof_list||."
#	"font_pattern||."
#	"destination||."
#)
#	PARAMETERS=(
#	"PARAMETERS ???"
#	"-all||."
#	"-app||."
#	"-backup||."
#	"-cache||."
#	"-check||."
#	"-ex||."
#	"-flist||."
#	"-font||."
#	"-fontset||."
#	"-help||."
#	"-lang||."
#	"-report||."
#	"-proof||."
#	"-rev||."
#	"-run||."
#	"-verbose||."
#)
#	EXAMPLES=(
#	"EXAMPLES ???"
#	"||$util -report"
#	"||sudo $util -all -run"
#	"||$util -app \"w e\" -lang -verbose"
#	"||sudo $util -lang \"nl no de\" -rev -run"
#	"||sudo $util -proof -app w -run"
#	"||sudo $util -proof \"Indonesian Isix*\" -rev -run"
#	"||$util -font lib -app w -verbose"
#	"||sudo $util -font lib -app w -run"
#	"||sudo $util -font \"chinese arial*\" -run"
#	"||$util -font -rev -app o"
#	"||sudo $util -font *.* -ex \"brit* rockwell*\" -app w -run"
#	"||sudo $util -cache"
#	"||$util -backup -font \"cyrdfonts britanic*\" -run"
#	"||sudo $util -backup syslib -font cyrdfonts -run"
#	"||$util -fontset"
#)
#	mylang="${LANG%\.*}"
#fi

print-topic o SYNOPSIS 0 4
print-topic o DESCRIPTION 0 4
print-topic o NOTES 0 6 '-'
print-topic   USAGE 0 4 'lh'
print-topic o USE_CASES -4 12
print-topic   ARGUMENTS 4 20 ":"
print-topic   PARAMETERS 4 20 ":"
print-topic o EXAMPLES -4 -8 ":"

[[ -z "$mylang" ]] && echo "No help for your language? Create your own, or remove this module file, or use parameter '-help en' for english."
