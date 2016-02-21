#!/bin/bash

# Welcome user!
# To get help just run this script without parameters.

# .SYNOPSIS
#	msomtu - Microsoft Office maintenance utility. Purpose - clean up MSO.
#
# .LINKS
#	Inspiration idea     : https://github.com/goodbest/OfficeThinner
#	On OS X & MSO fonts  : http://www.jklstudios.com/misc/osxfonts.html
#	Git Repo             : https://github.com/scriptingstudio/msomtu
#
# .TODO
#	- cleanup code; there are artefacts/leftovers 
#	  after migration from the previous version
#	- get more specific on duplicates and fontsets
#	- migrate to input parser G3 !!!
#	- new fonts finder: remove or not?
#	- uninstall MSO option (?)
#	- logging (???)
#	- help page: more clear text; format;
#	- rename to 'msomt' - Microsoft Office maintenance tool???
# 

defstate='skip'
cmd_clean=0
# params
run=0
cmd_help=0
cmd_all=0
cmd_app='0'
cmd_proof=$defstate
cmd_lang=$defstate
cmd_font=$defstate
cmd_fontlist=0
cmd_report=0
cmd_backup=''
cmd_cache=0
cmd_fontset=0
cmd_verb=0
cmd_reverse=0
cmd_exclude=''
cmd_log=''
cmd_uninstall=''
# END params

# Definitions
toolname="Microsoft Office 2016 Maintenance Utility"
version='2.8.38'
util="${0##*/}"
principalname='msomtu'
defapp='w e p o n'
		# THERE ARE DEPENDENCIES IN CODE!
deflang='ru' # user pref; english is added in create-filter;
defproof='russian' # user pref; english is added in create-filter

WordPATH="Microsoft Word.app" # app-path hashtable -> bash4 :-(
ExcelPATH="Microsoft Excel.app"
PowerPointPATH="Microsoft PowerPoint.app"
OutlookPATH="Microsoft Outlook.app"
OneNotePATH="Microsoft OneNote.app"

basePATH="/Applications/"
proofingPATH="/Contents/SharedSupport/"
proofingName="Proofing Tools"
fontPATH="/Contents/Resources"
backupPATH=~/"Desktop/MSOFonts/"

## Predefined fontsets. You can change them here
 # os x duplicates; in folder DFonts
 # exclusion: Microsoft no longer provides cyrillic MonotypeCorsiva, so
 # I included it here for deletion 
sysfonts=("arial*" "ariblk*" "Baskerville*" "Book Antiqua*" "ComicSans*" "Cooper*" "Gill Sans*" "GillSans*" "MonotypeCors*" "pala*" "Trebuchet*" "verdana*")
 # OS X duplicates; in folder Fonts - delete (???)
msofonts=("tahoma*" "Wingding*" "webding*")
 # custom sets; in folder DFonts
 	# chfonts - all kind of hieroglyphic/eastern fonts; chinese is idiomatic name 
chfonts=("Fangsong*" "Deng*" "gulim*" "HGR*" "kaiti*" "malgun*" "Meiryo*" "mingliu*" "MSJH*" "msyh*" "SimHei*" "simsun*" "STH*" "STX*" "STZ*" "STL*" "taile*" "YuGoth*" "yumin*")
noncyr=("Abadi*" "angsa*" "BellMT*" "Bauhaus93*" "BernardMT*" "Calisto MT*" "Braggadocio*" "Britannic*" "CalistoMT*" "ColonnaMT*" "COOPBL*" "CopperplateGothic*" "CurlzMT*" "Desdemona*" "EdwardianScriptITC*" "EngraversMT*" "Eurostile*" "FootlightMT*" "GloucesterMT*" "Goudy Old Style*" "Haettenschweiler*" "Harrington*" "ImprintMTShadow*" "KinoMT*" "Lucida Sans.*" "Lucida Sans Demibold*" "Lucida Sans Italic.*" "LucidaBright*" "LucidaBlackletter.*" "LucidaFax*" "LucidaCalligraphy*" "LucidaHandwriting*" "LucidaSansTypewrite*" "MaturaMTScriptCapitals*" "ModernNo.20*" "News Gothic MT*" "ntailu*" "Onyx*" "Perpetua*" "Rockwell*" "Stencil*" "Tw Cen*" "WideLatin*") # but some of non-cyr fonts may be useful
symfonts=("Bookshelf Symbol*" "Marlett*" "MS Reference Specialty*" "MonotypeSorts*")
cyrdfonts=("batang*" "Bookman Old Style*" "Candara*" "Century*" "Consola*" "Constan*" "Corbel*" "Franklin Gothic*" "Gabriola*" "GARA*" "Lucida Console*" "Lucida Sans Unicode*" "Mistral*" "MS Reference Sans Serif.*" "msgothic*" "Segoe Print Bold.*" "Segoe Script Bold.*")
 # custom sets; in folder Fonts
cyrfonts=("Calibri*" "Cambria*" "Century.*" "Corbel.*") # original cyr fonts
# 
dfontsets=(cyrdfonts noncyr chfonts sysfonts symfonts)
fontsets=(cyrfonts msofonts)
allfontsets=("${dfontsets[@]}" "${fontsets[@]}")
# END definitions

	
function main () {
	printb "$toolname. Version $version."

	############ Simple input named parameter parser v2 [inline method]
	local cmd='' # operation list - "clean=(flist font proof lang) report cache backup fontset"
	local script_params="$#"
	local PARGS=() INPUTPARAM=''
	while [ "$#" != 0 ]; do
		[[ "${1:0:1}" == "-" ]] && { INPUTPARAM="$1"; shift; }
		PARGS=()
		while [ "${1:0:1}" != "-" ] && [ "$#" != 0 ]; do
			PARGS+=("$1") # collect arguments of current parameter
			shift
		done
		case "$INPUTPARAM" in # check parameters
			-help|-h|-\?) cmd_help=1 ;;
			-all|-full) 
				cmd_fontlist=1; cmd_font=1; cmd_cache=1; cmd_all=1 ;;
			-flist|-fl) 
				cmd+=" clean"; cmd_fontlist=1 ;;
			-app) 
				cmd_app="$PARGS" ;;
			-font) # library (remove syslib/userlib from DFonts); pattern; fontset
				cmd+=" clean"; cmd_font="folder"
				[[ "$PARGS" ]] && cmd_font="$PARGS" ;;
			-proof|-p) 
				cmd+=" clean"; cmd_proof=""
				[[ "$PARGS" ]] && cmd_proof="$PARGS" ;;
			-lang|-ui)
				cmd+=" clean"; cmd_lang=""
				[[ "$PARGS" ]] && cmd_lang="$PARGS" ;;
			-verbose|-verb) cmd_verb=1 ;;
			-report|-rep|-info|-inf) cmd+=" report"; cmd_report=1 ;;
			-cache) cmd+=" cache"; cmd_cache=1 ;;
			-fontset|-fs) 
				cmd+=" fontset"; cmd_fontset=1; cmd_app=() ;;
			-rev) cmd_reverse=1; deflang=''; defproof='' ;; # reverse filter
			-ex|-x) [[ "$PARGS" ]] && cmd_exclude="$PARGS" ;;
			-run) run=1 ;;
			-fcopy|-backup) # -backup [<destination>]
				cmd="backup"; cmd_backup="$backupPATH"
				[[ "$PARGS" ]] && cmd_backup="$PARGS" ;;
			-check) open -a safari "http://macadmins.software" ;;
			--test) mktest; ;;
			-nl) nl=1; let "script_params--" ;;
			-log) cmd_log=1; [[ "$PARGS" ]] && cmd_log=2 ;;
			-uninstall) cmd_uninstall=1; cmd="uninstall" ;;
		esac
	done
	## Post parsing alignment
	if [[ $cmd_all -eq 1 ]]; then
		[[ "$cmd_proof" == $defstate ]] && cmd_proof=''
		[[ "$cmd_lang" == $defstate ]] && cmd_lang=''
		cmd_font=$defstate
		cmd_fontset=0
	else
		if [[ $cmd_fontset -eq 1 ]]; then
			cmd_app=(); cmd_report=0; run=0
			cmd_backup=''; cmd_fontlist=0
		fi
	fi

	if [[ -z "${cmd_lang// }" ]]; then 
		cmd_lang="$deflang"
	elif [[ "$cmd_lang" == $defstate ]]; then
		cmd_lang=''
	else
		cmd_lang=$(unique "$cmd_lang $deflang")
	fi
	if [[ -z "${cmd_proof// }" ]]; then 
		cmd_proof="$defproof"
	elif [[ "$cmd_proof" == $defstate ]]; then
		cmd_proof=''
	else
		cmd_proof=$(unique "$cmd_proof $defproof")
	fi
	if [[ -z "${cmd_font// }" ]]; then 
		cmd_font='folder'
	elif [[ "$cmd_font" == $defstate ]]; then
		cmd_font=''
	else
		cmd_font=${cmd_font/chinese/chfonts} # !!!!! confusing names
	fi
	[[ "$cmd_app" == '0' || -z "${cmd_app// }" ]] && 
		cmd_app="$defapp"
	# solo operations: backup, cache, fontset, report, help
	[[ $cmd_uninstall -eq 1 ]] && cmd="uninstall"
	if [[ -n "$cmd_backup" ]];     then cmd="backup"; fi
	if [[ "$cmd_cache" == 1 ]];    then cmd="$cmd cache"; fi
	if [[ "$cmd_fontset" == 1 ]];  then cmd="fontset"; fi
	if [[ "$cmd_report" == 1 ]];   then
		cmd="report"; #cmd_app="$defapp"
	fi
	[[ "$script_params" == 0 || "$cmd_help" -eq 1 ]] && cmd='help'
	[[ "$cmd" == '' ]] && 
		{ echo -e "\nNo valid actions or parameters defined (font lang proof flist report fontset cache backup help). Correct your command line parameters. See help.\n"; exit 3; }
	# END input parser
	
	############ Operation selector
	local c
	cmd=$(unique "$cmd")
	prepare-env	
	echo
	for c in $cmd; do
		case $c in
			report) 
				check-app 'skip'
				echo
				make-report ;;
				
			clean) 
				check-app
				echo
				local score1 score2 appdu=()
				get-diskusage score1 
				clean-application
				get-diskusage score2 
				show-appdu score1 score2 appdu 
				echo ;;
				
			backup) 
				invoke-backup ;;
				
			fontset) 
				display-fontset ;;
				
			cache) 
				clean-cache ;;
				
			uninstall) uninstall-mso ;;
				
			help)
				show-helppage ;;
		esac
	done # END operation selector
	echo
} # END main

function uninstall-mso () { # UNDER CONSTRUCTION
	if [[ $run -eq 0 ]]; then
		echo "  TRACE: uninstall Microsoft Office."
		return
	fi 
} # END uninstall MSO

function prepare-env () {
# Preparing runtime environment
	local temp=$(unique "$cmd_app") appPATH i
	cmd_app=()
	for i in $temp; do # hashtable: bash4
		[[ $i == 'w' ]] && cmd_app+=("$WordPATH")
		[[ $i == 'e' ]] && cmd_app+=("$ExcelPATH")
		[[ $i == 'p' ]] && cmd_app+=("$PowerPointPATH")
		[[ $i == 'o' ]] && cmd_app+=("$OutlookPATH")
		[[ $i == 'n' ]] && cmd_app+=("$OneNotePATH")
	done

	appInstalled=()
	temp=("$WordPATH" "$ExcelPATH" "$PowerPointPATH" "$OutlookPATH" "$OneNotePATH")
	for appPATH in "${temp[@]}"; do
		[[ -d "$basePATH$appPATH" ]] &&
			appInstalled+=("${appPATH/.app/}")
	done
	appPathArray=()
	for appPATH in "${cmd_app[@]}"; do
		[[ -d "$basePATH$appPATH" ]] &&
			appPathArray+=("$appPATH")
	done
} # END prepare

function check-app () {
	local appPATH
	if [[ ${#appPathArray} -eq 0 ]]; then
		echo "Nothing to do - no app defined or found. Bye."
		exit
	fi
	printb "Installed apps:"
	for appPATH in "${appInstalled[@]}"; do
		echo "- $appPATH"
	done
	if [[ "$1" == '' ]]; then
		echo
		printb "Apps to process:"
		for appPATH in "${appPathArray[@]}"; do
			echo "- ${appPATH/.app/}"
		done
	fi
} # END checking app list

function make-report () {
	__normalize-number () { # restr: input - kilobytes
		echo "$1" | awk '{e1=$1;} END {
			unit="K"; fmt1="%.f %s"; kilo=1024
			if (e1 > kilo) {e1/=1024; unit="M"; fmt1="%.f %s"}
			if (e1 > kilo) {e1/=1024; unit="G"; fmt1="%.1f %s"}
			printf fmt1,e1,unit
		}'
	}
	__separate-unit () { echo "${1/[A-Z]/} ${1:(-1)}"; }
	local versionPATH="/Contents/Info.plist"
	local versionKey="CFBundleShortVersionString"
	local fmt1='   %-18s : %10s  %10s\n' na='--' sfx='' appfat appfiles
	local appPATH wpath appVersion msbuild fs fc plist flist filter
	for appPATH in "${appPathArray[@]}"; do
		printb "Processing '$appPATH'"
		wpath="$basePATH$appPATH$fontPATH/DFonts"
		appVersion=$(defaults read "$basePATH$appPATH$versionPATH" $versionKey)
		msbuild=$(defaults read "$basePATH$appPATH$versionPATH" "MicrosoftBuildNumber")
		printf "$fmt1" "Version (build)" "$appVersion" "($msbuild)"
		appfat=0; appfiles=0
		##### fonts
		if [[ -d "$wpath" ]]; then
			fs=''; fs=$(du -sh -k "$wpath" | cut -f 1); let "appfat+=${fs}"
			fs=$(__normalize-number $fs)
			fc=$(ls -A "$wpath" | wc -l) #fc=$(find "$wpath" -type f | wc -l)
			let "appfiles+=${fc// }"
			printf "$fmt1" "DFonts" "${fc// }" "${fs}$sfx"
		else
			printf "$fmt1" "DFonts" "does not exist"
			cmd_font=''
		fi
		wpath="$basePATH$appPATH$fontPATH/Fonts"
		fs=''; fs=$(du -sh "$wpath" | cut -f 1)
		fc=$(ls -A "$wpath" | wc -l) 
		fs=$(__separate-unit $fs);
		printf "$fmt1" "Fonts" "${fc// }" "${fs}$sfx"
		##### fontlists
		wpath="$basePATH$appPATH$fontPATH"
					#- flist=$(find "$wpath" -type f -name font*.plist -d 1)
		flist=$(ls "$wpath/"font*.plist 2> /dev/null)
		if [[ -z "$flist" ]]; then
			printf "$fmt1" "Plists" $na
		else
			fs=$(du -sh -k "$wpath/"*.plist | awk '{ total += $1 }; END {print total}')
			let "appfat+=${fs}"; fs=$(__normalize-number $fs)
			fc=$(echo "$flist" | wc -l)
			let "appfiles+=${fc// }"
			printf "$fmt1" "Plists" "${fc// }" "${fs}$sfx"
		fi
		##### UI languages
		filter=$( create-filter "lang" )
		flist=$(find -E "$wpath" -type d -d 1 -name *.lproj ! -iregex $filter)
		if [[ -z "$flist" ]]; then
			printf "$fmt1" "Langpacks" $na
		else
			fs=$(du -sh -k "$wpath/"*.lproj | awk '{ total += $1 }; END {print total}')
			let "appfat+=${fs}"; fs=$(__normalize-number $fs)
			fc=$(echo "$flist" | wc -l); 
			let "appfiles+=${fc// }"
			printf "$fmt1" "Langpacks" "${fc// }" "${fs}$sfx"
		fi
		##### proofing tools
		wpath="$basePATH$appPATH$proofingPATH$proofingName"
		filter=$( create-filter "proof" )
		flist=$(find -E "$wpath" -type d -d 1 -name *.proofingtool ! -name Grammar.proofingtool ! -iregex $filter)
		if [[ -z "$flist" ]]; then
			printf "$fmt1" "Proofingtools" $na
		else
			fs=''; fs=$(du -sh -k "$wpath" | cut -f 1); let "appfat+=${fs}"
			fs=$(__normalize-number $fs)	
			fc=$(echo "$flist" | wc -l); let "appfiles+=${fc// }"
			printf "$fmt1" "Proofingtools" "${fc// }" "${fs}$sfx"
		fi
		##### total disk usage
		fs=$(du -sh "$basePATH$appPATH" | cut -f 1)
		fc=$(ls -AR "$basePATH$appPATH" | wc -l)
		fs=$(__separate-unit $fs); fc=$(printf "%'d" $fc)
		printf "$fmt1" '' '' '------'
		printf "$fmt1" "Total app bundle" "$fc" "${fs}$sfx"
		appfat=$(__normalize-number $appfat)
		printf "$fmt1" "Approx. thinning" "$appfiles*" "-${appfat}$sfx"

		echo
	done # app selection
	echo -e "----\n[*] Includes the default file sets which are reserved\n    by settings. The folder 'Fonts' is reserved."
	echo
} # END report

function clean-application () {
	local fmt1='   %-14s : %6s %s\n' appPATH wpath appfat fl fc
	for appPATH in "${appPathArray[@]}"; do
		appfat=0
		printb "Processing '$appPATH'"
	# - removing of font files/folder DFonts
		[[ "$cmd_font" != '' ]] && 
			clean-font "$basePATH$appPATH$fontPATH/DFonts"

	# - removing of .plist files 
		wpath="$basePATH$appPATH$fontPATH"
		if [[ $cmd_fontlist -eq 1 ]]; then
			if [[ $run -eq 1 ]]; then
				find "$wpath" -type f -name font*.plist -d 1 -exec rm -f {} \;
			else
				echo "  TRACE: remove font-list files (.plist)."
				if [[ $cmd_verb -eq 1 ]]; then
					fl=$(find "$wpath" -type f -name font*.plist -d 1 -exec basename {} \;)
					[[ -z $nl ]] && echo "$fl"; aggregator "$fl" "$wpath"
				fi
			fi
		fi # END plist (fontlist files)
	
	# - cleaning of lproj folders; keep en_GB.lproj en.lproj
		[[ "$cmd_lang" != '' ]] && clean-lang "$wpath"
	
	# - cleaning of Proofing Tools; keep English*.proofingtool Grammar.proofingtool
		[[ "$cmd_proof" != '' ]] &&
			clean-ptools "$basePATH$appPATH$proofingPATH$proofingName"
		
		appfat=$( echo "$appfat" | awk '{e1=$1;} END {
			unit="K"; fmt1="%.f%s"; kilo=1024
			if (e1 > kilo) {e1/=1024; unit="M"; fmt1="%.f%s"}
			if (e1 > kilo) {e1/=1024; unit="G"; fmt1="%.1f%s"}
			printf fmt1,e1,unit
		}' )
		appdu+=("${appfat}	$basePATH$appPATH") # tab char between
		echo
	done
} # END cleanup wrapper

function clean-font () {
	local wpath="$1" fl=() name='' f m filter fc fs
	if [[ ! -d "$wpath" ]]; then
		echo "  Folder 'DFonts' does not exist."
		return
	fi
	if [[ $cmd_reverse -eq 1 ]]; then
		find-newfont "$wpath"
		return
	fi
	
	cmd_font=$(unique "$cmd_font")
	for f in $cmd_font; do 	# expand fontsets
		[[ $f == 'userlib' || $f == 'syslib' ]] && continue
		m=$(inarray "$f" dfontsets)
		if [[ $m ]]; then
			name=$m[@]; a=("${!name}")
			fl+=("${a[@]}")
		else
			fl+=("$f")
		fi
	done # fontset selector
	filter=$(create-filter 'exclude')
	if [[ $run -eq 1 ]]; then # action mode
		for f in $cmd_font; do # folder alone
			[[ "$f" == 'folder' ]] && rm -fdr "$wpath"
			[[ "$f" == lib* ]] && remove-duplicate "$wpath"
		done
		for f in "${fl[@]}"; do # font list
			echo "Removing '$f'..."
			if [[ "$filter" != '' ]]; then
				find -E "$wpath" -type f -iname "$f" -d 1 ! -iregex "$filter" -exec rm -f {} \; 
			else
				find "$wpath" -type f -iname "$f" -d 1 -exec rm -f {} \; 
			fi
		done
		return
	fi
	
	echo "  TRACE: remove fonts/folder '$wpath'."
	[[ $cmd_verb -eq 0 ]] && return
	for f in $cmd_font; do # folder alone
		if [[ "$f" == 'folder' ]]; then
			fs=$(ls -A "$wpath"); [[ -z $nl ]] && echo "$fs"
			aggregator "$fs" "$wpath"
		fi
		[[ "$f" == lib* ]] && remove-duplicate "$wpath"
	done
	for f in "${fl[@]}"; do # font list
		if [[ "$filter" != '' ]]; then
			fs=$(find -E "$wpath" -type f -iname "$f" -d 1 ! -iregex "$filter" -exec basename {} \;)
		else
			fs=$(find "$wpath" -type f -iname "$f" -d 1 -exec basename {} \;)
		fi
		[[ -z $nl ]] && echo "$fs"; aggregator "$fs" "$wpath"
	done
} # END clean fonts

function remove-duplicate () {
	__invoke-deduplication () {
		local wpath="$3"
		printf '%s' "$4"
		local dup=$( comm -1 -2 -i <(printf '%s\n' "${1}") <(printf '%s\n' "${2}") )
		local fc=$(echo "${dup[@]}" | wc -l); [[ "$dup" == '' ]] && fc=0
		echo "[ ${fc// } ]"
		if [[ $run -eq 0 ]]; then
			echo "------------------------------"
			echo "$dup" | xargs -I{} echo "- {}" 
		else
			echo "$dup" | xargs -I{} rm -f "$wpath/{}" 
		fi
	} # ENDBLOCK deduplication
	local wpath="$1"
	local sysl=/Library/Fonts
	local userl=~/Library/Fonts

	local dfonts=$(find "$wpath" -type f -name *.* -d 1 -exec basename {} \;)
	local lib=$(find "$sysl" -type f -name *.* -d 1 -exec basename {} \;)
	local ulib=$(find "$userl" -type f -name *.* -d 1 -exec basename {} \;)
	
	printf -v obanner '\n%-30s' "SYSTEM LIBRARY <-- DFonts"
	__invoke-deduplication "$dfonts" "$lib" "$wpath" "$obanner"
	printf -v obanner '\n%-30s' "USER LIBRARY <-- DFonts"
	__invoke-deduplication "$dfonts" "$ulib" "$wpath" "$obanner"
} # END remove duplicate fonts

function find-newfont () {
	local fsfiles='' allfiles=() list=''
	local fc f i name fs newfont wpath="$1"
	[[ $run -eq 0 ]] &&
		echo -e "  TRACE: find and remove new fonts.\n"
	echo "Searching for new fonts..."
	for i in ${dfontsets[@]}; do # expand fontsets
    	name=$i[@]
		fs=("${!name}")
		for f in "${fs[@]}"; do
			list=$(find "$wpath" -type f -d 1 -iname "$f" -exec basename {} \;)
			[[ "$list" != '' ]] && fsfiles+=$'\n'"$list"
		done
	done
	fsfiles="${fsfiles:1}"
	allfiles=$(find "$wpath" -type f -d 1 -exec basename {} \;)
	fc=$(echo "${allfiles[@]}" | wc -l)
		echo "Total font files  : ${fc// }"
	fc=$(echo "${fsfiles[@]}" | wc -l)
		echo "Fonts by fontsets : ${fc// }"
	newfont=$( comm -2 -3 -i <(echo "${allfiles[@]}" | sort -u) <(echo "${fsfiles[@]}" | sort -u) )
	if [[ "${#newfont}" -eq 0 ]]; then
		echo "No new fonts found."
		return
	fi
	echo
	if [[ $run -eq 0 ]]; then
		echo -e "New fonts found:\n----------------"
		echo "$newfont"
		echo -e "\nUpdate predefined fontsets."
	else
		for f in "${newfont[@]}"; do # remove or someth else ???
			echo "Removing '$f'..."
			#find "$wpath" -type f -d 1 -name "$f" -exec rm -f {} \;
		done
	fi
} # END new fonts

function clean-lang () {
	__list-lang () { 
		local fs fc
		fs=$( find -E "$wpath" -type d -d 1 -name *.lproj ! -iregex $filter -exec basename {} \; )
		[[ -z $nl ]] && echo "$fs"; aggregator "$fs" "$wpath" 
	}
	__rm-lang () { find -E "$wpath" -type d -d 1 -name *.lproj ! -iregex $filter -exec rm -fdr {} \; ; }
	local wpath="$1"
	local filter=$( create-filter "lang" )
	if [[ cmd_reverse -eq 0 ]]; then
		if [[ $run -eq 1 ]]; then
			__rm-lang
		else # trace mode
			cmd_lang=$(unique "en $cmd_lang")
			echo "  TRACE: remove extra UI languages. Keep '$cmd_lang'."
			[[ $cmd_verb -eq 1 ]] && __list-lang
		fi
	else # reverse filter
		if [[ $run -eq 1 ]]; then
			__rm-lang
		else # trace mode
			echo "  TRACE: remove extra UI languages - '$cmd_lang'."
			[[ $cmd_verb -eq 1 ]] && __list-lang 
		fi
	fi
} # END lang

function clean-ptools () {
	local wpath="$1" fs fc
	local filter=$( create-filter "proof" )
	if [[ cmd_reverse -eq 0 ]]; then
		if [[ $run -eq 1 ]]; then
			find -E "$wpath" -type d -d 1 -name *.proofingtool ! -name Grammar.proofingtool ! -iregex $filter -exec rm -fdr {} \;
		else # trace mode
			cmd_proof=$(unique "english $cmd_proof")
			echo "  TRACE: remove extra proofing tools. Keep '$cmd_proof'."
			if [[ $cmd_verb -eq 1 ]]; then 
				fs=$( find -E "$wpath" -type d -d 1 -name *.proofingtool ! -name Grammar.proofingtool ! -iregex $filter -exec basename {} \; )
				[[ -z $nl ]] && echo "$fs"; aggregator "$fs" "$wpath" 
			fi
		fi
	else # reverse filter
		if [[ $run -eq 1 ]]; then
			find -E "$wpath" -type d -d 1 -name *.proofingtool -iregex $filter -exec rm -fdr {} \;
		else # trace mode
			echo "  TRACE: remove extra proofing tools - '$cmd_proof'."
			if [[ $cmd_verb -eq 1 ]]; then 
				fs=$( find -E "$wpath" -type d -d 1 -name *.proofingtool -iregex $filter -exec basename {} \; )
				[[ -z $nl ]] && echo "$fs"; aggregator "$fs" "$wpath"
			fi
		fi
	fi
} # END proofingtools

function display-fontset () {
	local p4=4 p20=20
	local fs1=$(joina , "${sysfonts[@]}")
	local fs2=$(joina , "${chfonts[@]}")
	local fs3=$(joina , "${noncyr[@]}")
	local fs4=$(joina , "${cyrdfonts[@]}")
	local fs5=$(joina , "${cyrfonts[@]}")
	local fs6=$(joina , "${symfonts[@]}")
	printb "Predefined fontsets:"
	print-column $p4 $p20 "sysfonts" "OS duplicated fonts (in DFonts):" "-"
	print-column $p4 $p20 "" "${fs1//,/, }"
	echo
	print-column $p4 $p20 "chinese" "All kind of hieroglyphic/eastern fonts (in DFonts):" "-"
	print-column $p4 $p20 "" "${fs2//,/, }"
	echo
	print-column $p4 $p20 "noncyr" "Non-cyrillic fonts (in DFonts):" "-"
	print-column $p4 $p20 "" "${fs3//,/, })"
	echo
	print-column $p4 $p20 "cyrdfonts" "Cyrillic original fonts (in DFonts; do not include 'sysfonts'):" "-"
	print-column $p4 $p20 "" "${fs4//,/, }"
	echo
	print-column $p4 $p20 "cyrfonts" "Cyrillic original fonts (in Fonts):" "-"
	print-column $p4 $p20 "" "${fs5//,/, }"
	echo
	print-column $p4 $p20 "symfonts" "Symbolic fonts (in DFonts):" "-"
	print-column $p4 $p20 "" "${fs6//,/, }"
	echo
} # END fontsets

function invoke-backup () { # for fonts only
	if [[ -z $cmd_backup ]]; then
		printf '%s\n\n' "ERROR: Invalid set of backup arguments specified."	
		return -1
	fi
	
	printb "Backup fonts of '$appPathArray'"
	local bsrc="$basePATH$appPathArray$fontPATH" bset=()
	local bdest="${cmd_backup}"
	local ffolder="DFonts" f i m a name fl fc=0 c
	if [[ "$bdest" == 'syslib' ]]; then
		bdest="/Libraries/Fonts/"
	elif [[ "$bdest" == 'userlib' ]]; then
		bdest=~/Libraries/Fonts/
	fi
	if [[ ! -d "$bdest" ]]; then
		if [[ $run -eq 1 ]]; then
			mkdir -p "$bdest"
			if [[ $? -ne 0 ]]; then 
				echo "ERROR: failed to create directory '$bdest'."
				bdest=''
			fi
		fi
	fi
	[[ "$bdest" == '' ]] && return 
	
	[[ "$cmd_font" == '' ]] && cmd_font="*.*"
	for f in $(unique "${cmd_font}"); do # expand array
		[[ $f == userlib || $f == syslib ]] && continue
		[[ $f == folder ]] && f="*.*"
		[[ "$f" == 'cyrfonts' || "$f" == 'msofonts' ]] && ffolder='Fonts'
		m=$(inarray "$f" allfontsets)
		if [[ $m ]]; then
			name=$m[@]; a=("${!name}")
			bset+=("${a[@]}")
		else
			bset+=("$f")
		fi
	done # fontset selector
	bsrc+="/$ffolder"
	if [[ $run -eq 1 ]]; then
		for i in "${bset[@]}"; do
			fl=$(find "$bsrc" -type f -iname "$i" -d 1)
			[[ "${fl// }" == '' ]] && continue
			echo "$fl" | xargs -I{} cp -f "{}" "$bdest"/
			error=$?
			c=$(echo "$fl" | wc -l); [[ "$fl" == '' ]] && c=0;
			let "fc+=$c"
		done
		echo "-----------"
		echo "Total files : ${fc// }."
		echo "Destination : '$bdest'."
		if [[ $error -ne 0 ]]; then
			echo "Error ($error) in copying."
		else
			echo "Done. Files copied : $fc. If 0 copied check fontsets (see help)."
		fi
	else # trace mode
		echo "TRACE: source folder : '$bsrc'."
		for i in "${bset[@]}"; do
			fl=$(find "$bsrc" -type f -iname "$i" -d 1 -exec basename {} \;)
			[[ "${fl// }" == '' ]] && continue
			c=$(echo "$fl" | wc -l); [[ "$fl" == "" ]] && c=0;
			let "fc+=$c"
			[[ $cmd_verb -eq 1 ]] && echo "$fl"
		done
		echo "-----------"
		echo "Total files : ${fc// }."
		echo "Destination : '$bdest'."
		[[ ! -d "$bdest" ]] &&
			echo "TRACE: cannot access '$bdest'."
	fi
} # END backup

function clean-cache () {
	echo
	atsutil databases -remove
} # END cache

function show-helppage () {
	local p3=3 p4=4 p6=6 p8=8 p20=20
	local fs="${dfontsets[@]/chfonts/chinese}"; fs=${fs// /, }

	if [[ "${LANG%\.*}" != "en_US" && "${LANG%\.*}" != "en_GB" && -z $nl ]]; then
		local extless="${0%.*}"
		local helpfile="${extless%-*}-help.sh"  # clean dev stage suffix
		[[ -f "$helpfile" ]] && { . "$helpfile"; exit 0; }
	fi

	if [[ $cmd_all -eq 1 ]]; then
	printb "SYNOPSIS:"
	print-column 0 $p4 "" "MSOMTU is Microsoft Office maintenance utility."
	echo
	
	printb "DESCRIPTION:"
	print-column 0 $p4 "" "Microsoft Office 2016 for Mac uses an isolated resource architecture (sandboxing), so the MSO apps duplicate all of the components in its own application container that's waisting gigabytes of your disk space. This script safely removes (thinning) extra parts of the folowing components: UI languages; proofing tools; fontlist files (.plist); OS duplicated font files. It also can backup/copy font files to predefined and user defined destinations." 
	echo

	printb "NOTES:"
	print-column 0 $p6 "" "Safe scripting technique - 'Foolproof' or 'Harmless Run'. The default running mode is view. You cannot change or harm your system without switch '-run'. Parameter '-cache' does not depend on '-run'." '-'
	print-column 0 $p6 "" "As MSO is installed with root on /Applications directory you have to run this script with sudo to make changes." '-'
	print-column 0 $p6 "" "As application font structure has been changed since MSO version 15.17 font deletion only works with 15.17 or later. Microsoft separated font sets for some reasons. Essential fonts to the MSO apps are in the 'Fonts' folder within each app. The rest are in the 'DFonts' folder." '-'
	print-column 0 $p6 "" "If you remove fonts, remove font lists as well. The 'DFonts' folder and font lists are safe to remove. Neither third party app can see MSO fonts installed to the 'DFonts' folder. Some of the fonts you may find useful, save them before deletion." '-'
	print-column 0 $p6 "" "Caution: do not remove fonts from the 'Fonts' folder! These are minimum needed for the MSO applications to work." '-'
	print-column 0 $p6 "" "File operations are case insensitive." '-'
	print-column 0 $p6 "" "Script only accepts named parameters." '-'
	print-column 0 $p6 "" "Apply thinning after every MSO update." '-'
	print-column 0 $p6 "" "Default settings for the '-lang' and '-proof' parameters: english and russian. It depends on your system locale and common sense: for MSO integrity it is better to leave english. You can change any default settings in code for your needs." '-'
	print-column 0 $p6 "" "Font classification spicifics in predefined fontsets (in descending): cyrillic, non-cyrillic, hieroglyphic, symbolic, system. Fontsets do not intersect." '-'
	echo
	fi

	printb "USAGE:"
	print-column 0 $p4 "" "[sudo] $util [-<parameter> [<arguments>]]..."
	echo
	print-column 0 $p4 "" "[sudo] $util [-backup|-fcopy [<destination>]] [-app [<app>]] [-font [<font_pattern>]] [-ex|-x <font_pattern>] [-run]"
	echo
	print-column 0 $p4 "" "[sudo] $util [-app [\"<app_list>\"]] [-lang|-ui [\"<lang_list>\"]] [-proof|-p [\"<proof_list>\"]] [-font [<font_pattern>]] [-flist|-fl] [-ex|-x <font_pattern>] [-cache] [-report|-rep] [-verbose|-verb] [-fontset|-fs] [-all|-full] [-rev] [-help|-h|-?] [-run]"
	echo 

	if [[ $cmd_all -eq 1 ]]; then
	local mp4=-4 p12=12
	printb "USE CASES:"
	print-padding $mp4 "Solo actions: 'backup', 'cache', 'fontset', 'report', 'help'."
	echo
	print-padding $mp4 "- Getting MSO info" 
		print-column 0 $p12 "" "Parameter '-report'."
	print-padding $mp4 "- Getting assessment of thinning (view mode)" 
		print-column 0 $p12 "" "Parameter '-verbose' along with resource selector: font, flist, lang, proof."
	print-padding $mp4 "- Listing/Removing UI languages" 
		print-column 0 $p12 "" "Parameter '-lang'."
	print-padding $mp4 "- Listing/Removing proofingtools" 
		print-column 0 $p12 "" "Parameter '-proof'."
	print-padding $mp4 "- Listing/Removing fonts" 
		print-column 0 $p12 "" "Parameter '-font'."
	print-padding $mp4 "- Listing/Removing font list files" 
		print-column 0 $p12 "" "Parameter '-fontlist'."
	print-padding $mp4 "- Listing fontsets" 
		print-column 0 $p12 "" "Parameter '-fontset'."
	print-padding $mp4 "- Removing font cache" 
		print-column 0 $p12 "" "Parameter '-cache'."
	print-padding $mp4 "- Finding new fonts" 
		print-column 0 $p12 "" "Parameter '-font -rev'."
	print-padding $mp4 "- Backing up fonts" 
		print-column 0 $p12 "" "Parameter '-backup'."
	print-padding $mp4 "- Copying fonts to font libraries" 
		print-column 0 $p12 "" "Parameter '-backup'."
	echo
	fi
	
	printb "ARGUMENTS:"
	print-column $p4 $p20 "app_list" "App list. w - Word, e - Excel, p - PowerPoint, o - Outlook, n - OneNote. Default: 'w e p o n'." ":"
	print-column $p4 $p20 "lang_list" "Langauge list: ru pl de etc, see filenames with parameter '-verbose'. Default: 'en ru', see NOTES." ":"
	print-column $p4 $p20 "proof_list" "Proofingtools list: russian finnish german etc, see filenames with parameter '-verbose'. Wildcard '*' is available. Default: 'english russian', see NOTES." ":"
	print-column $p4 $p20 "font_pattern" "Font operations are based on patterns. Font patterns: empty - removes the 'DFonts' folder (default); <fontset> - removes fonts of predefined fontset; <mask> - removes selection: *.*, arial*, *.ttc etc. If you use single '*' enclose it in quotation marks: \"*\". Predefined fontsets: library, $fs. See parameter '-fontset' and details in code. Fontset 'library' removes duplicates of system and user libraries; it may not exactly match fonts because based on file-by-file (unlike font family) comparison (DFonts against libraries). You can use list of fontsets as well." ":"
	print-column $p4 $p20 "destination" "Backup destination folderpath for fonts. Default value is '~/Desktop/MSOFonts'. You can use predefined destinations as well: 'syslib' - system library; 'userlib' - user library." ":"
	echo

	printb "PARAMETERS:"
	print-column $p4 $p20 "-app" "Filter <app_list>. Selects application to process." ":"
	print-column $p4 $p20 "-lang" "Exclusive filter <lang_list>. Removes UI languages except defaults and user list. See also parameter '-rev'; it reverses user selection except defaults." ":"
	print-column $p4 $p20 "-proof" "Exclusive filter <proof_list>. Removes proofing tools except defaults and user list. See also parameter '-rev'; it reverses user selection except defaults." ":"
	print-column $p4 $p20 "-font" "Filter <font_pattern>. Removes selected fonts or the 'DFonts' folder. Available fontsets: cyrdfonts, noncyr, chinese, sysfonts. Parameter '-rev' ignores user selection and alternates search function: new fonts are going to be discovered. It is useful to check new fonts up after new update." ":"
	print-column $p4 $p20 "-backup" "Backs up fonts to user defined destination. If destination folder does not exist it will be created. You can use system and user libraries as destination, see ARGUMENTS. Backup alternates all deletions to backup." ":"
	print-column $p4 $p20 "-ex" "Exclusive filter <font_pattern>. Excludes font selection with parameter '-font'. Only mask can be used as 'font_pattern'." ":"
	print-column $p4 $p20 "-flist" "Switch. Removes fontlist (.plist) files." ":"
	print-column $p4 $p20 "-all" "Switch. Activates all cleaning options: lang, proof, font, flist, cache. It does not affect a parameter '-app'." ":"
	print-column $p4 $p20 "-cache" "Switch. Cleans up font cache." ":"
	print-column $p4 $p20 "-verbose" "Switch. Shows objects to be removed in view mode." ":"
	print-column $p4 $p20 "-report" "Switch. Shows statistics on objects." ":"
	print-column $p4 $p20 "-fontset" "Switch. Shows predefined fontsets." ":"
	print-column $p4 $p20 "-rev" "Switch. Reverses effect of the 'lang' and 'proof' filters. For parameter '-font' it is to search for the new fonts." ":"
	print-column $p4 $p20 "-run" "Switch. The default mode is view (test). Activates operations execution." ":"
	print-column $p4 $p20 "-help" "Switch. Shows the help page. There are two kinds of help page: short and full. The default is short one (no paramaters). To get the full page use parameters '-help -full'." ":"
	print-column $p4 $p20 "-help" "Special switch. With parameter '-help' forces english help. With parameter '-verbose' skips file listing; the same as without '-verbose' but displays report table." ":"
	echo
	
	if [[ $cmd_all -eq 1 ]]; then
	printb "EXAMPLES:"
	p4=$((0-$p4)); p8=$((0-$p8))
	print-padding $p4 "Get app statistics:"
	  print-padding $p8 "$util -report" b
	print-padding $p4 "Thin all apps with all parameters:"
	  print-padding $p8 "sudo $util -all -run" b
	print-padding $p4 "Show app ('w e' for Word and Excel) language files installed:"
	  print-padding $p8 "$util -app \"w e\" -lang -verbose" b
	print-padding $p4 "Remove a number of languages:"
	  print-padding $p8 "sudo $util -lang \"nl no de\" -rev -run" b
	print-padding $p4 "Remove all proofing tools except defaults for Word:"
	  print-padding $p8 "sudo $util -proof -app w -run" b
	print-padding $p4 "Remove a number of proofing tools:"
	  print-padding $p8 "sudo $util -proof \"Indonesian Isix*\" -rev -run" b
	print-padding $p4 "Show duplicates of library fonts for Word:"
	  print-padding $p8 "$util -font lib -app w -verbose" b
	print-padding $p4 "Remove duplicated fonts in libraries for Word:"
	  print-padding $p8 "sudo $util -font lib -app w -run" b
	print-padding $p4 "Remove 'chinese' and Arial fonts:"
	  print-padding $p8 "sudo $util -font \"chinese arial*\" -run" b
	print-padding $p4 "Show new fonts for Outlook:"
	  print-padding $p8 "$util -font -rev -app o" b
	print-padding $p4 "Exclude a few useful fonts from deletion for Word:"
	  print-padding $p8 "sudo $util -font *.* -ex \"brit* rockwell*\" -app w -run" b
	print-padding $p4 "Clean font cache:"
	  print-padding $p8 "sudo $util -cache" b
	print-padding $p4 "Backup fonts to default destination:"
	  print-padding $p8 "$util -backup -font \"cyrdfonts britanic*\" -run" b
	print-padding $p4 "Copy original cyrillic fonts to system library:"
	  print-padding $p8 "sudo $util -backup syslib -font cyrdfonts -run" b
	print-padding $p4 "Show predefined fontsets:"
	  print-padding $p8 "$util -fontset" b
	echo
	fi
	
	echo $helpfile
	exit 0
} # END help page

function write-log () { echo; } # under construction
function show-appdu () {
	[[ $run -eq 0 && $cmd_verb -eq 0 ]] && return
	local s1 s2 a1 a2 as1 as2 du1=$1[@] du2=$2[@] dif='-' difpc
	[[ $run -eq 0 ]] && du2=$3[@]
	du1=("${!du1}"); du2=("${!du2}")
	[[ -z "${du1[@]}" && -z "${du2[@]}" ]] && return
	local fmt1='%-12s %8s %8s %14s\n'
	printb "After-effect of MSO thinning report"
	printf "$fmt1" 'Application' 'Before' 'After' 'Thin effect'
	printf "$fmt1" '-----------' '------' '-----' '-----------'
	for s1 in "${du1[@]}"; do
		a1="${s1#*/}"
		as1=${s1%%/*}; as1=$(echo -n $as1)
		for s2 in "${du2[@]}"; do
			if [[ "$a1" == "${s2#*/}" ]]; then
				as2=${s2%%/*}; as2=$(echo -n $as2)
				a1="${a1##*/}"; a1=${a1/.app/}
				if [[ $run -eq 0 ]]; then 
					dif=$as2
					as2=$(math-diffexpr $as1 $as2 1); 
					difpc="${as2#*|}"; as2="${as2%|*}"
				else
					dif=$(math-diffexpr $as1 $as2); 
					difpc="${dif#*|}"; dif="${dif%|*}"
				fi
				[[ ${#dif} -eq 2 && ${dif:0:1} == '0' ]] && dif='n/a'
				###[[ ${#dif} -eq 2 && ${dif:0:1} == '0' ]] && dif='<1'${dif:(-1)}
				#if [[ $dif != 'n/a' && $difpc ]]; then
					printf -v difpc ": %3s" "$difpc"
					dif="$dif $difpc"
				#fi
				printf "$fmt1" "${a1/Microsoft /}" "$as1" "$as2" "$dif"
			fi
		done
	done
} # END app disk usage score report
function get-diskusage () {
	[[ $run -eq 0 && $cmd_verb -eq 0 ]] && return
	local app s1 score=()
	for app in "${appPathArray[@]}"; do
		s1=$( du -sh "$basePATH$app" )
		score+=("$s1")
	done
	eval $1='("${score[@]}")'
} # END MSO DU
function aggregator () {
# $1 - '\n' delimited file list; $2 - filepath;	
	local fc=$( echo "${1}" | 
		xargs -I{} du -sh -k "$2"/"{}" | 
		awk '{ total += $1 }; END {print total}' )
	let "appfat+=fc" # appfat - prev declared var
} # END assessment data aggregator

function create-filter () {
	local l='' pre='' sfx='' search='' i
	if [[ "$1" == "lang" ]]; then
		l="$cmd_lang"
		pre=".+/(en|en_GB"
		sfx=")\.lproj"
	elif [[ "$1" == "proof" ]]; then
		l="$cmd_proof"
		pre=".+/(English"
		sfx=").+\.proofingtool"
	elif [[ "$1" == "exclude" && $cmd_exclude ]]; then
		local patterns=($cmd_exclude)
		for i in "${patterns[@]}"; do
			i=${i/\*/.*}
			search+="|$i"
		done
		[[ $search ]] && search="${search:1}"
		if [[ "$search" ]]; then
			printf '%s%s%s' ".+/(" $search ")\..+"
		else
			echo ''
		fi
		return
	fi
	[[ $cmd_reverse -eq 1 ]] && pre=".+/("
	for i in $l; do
		search+="|$i"
	done
	[[ $cmd_reverse -eq 1 && ${search:0:1} == '|' ]] && search="${search:1}"
	printf '%s%s%s' "$pre" "$ss" "$sfx"
} # END search filter

##########  Common routines library
function joina { local oldifs="$IFS"; IFS="$1"; shift; echo "$*"; IFS="$oldifs"; }
function printb () { echo -e "\033[1m$1\033[0m"; }
function printu () { echo -e "\033[4m$1\033[0m"; }
function print-padding () {
	[[ ! -n "$2" || -z "$1" ]] && return
	local str=$2 pad=''
	[[ $1 -ne 0 ]] && printf -v pad "%*s" $1 ' '
	case "$3" in
		b) if [[ $1 -gt 0 ]]; then echo -e "\033[1m$str\033[0m$pad"; 
			else echo -e "$pad\033[1m$str\033[0m"; fi ;;
		u) if [[ $1 -gt 0 ]]; then echo -e "\033[4m$str\033[0m$pad"; 
			else echo -e "$pad\033[4m$str\033[0m"; fi ;;
		*) if [[ $1 -gt 0 ]]; then echo -e "$str$pad"; 
			else echo -e "$pad$str"; fi ;;
	esac
} # END print-padding
function print-column () {
# Prints 1 or 2 columns of text. UNIX sucks: printf is not unicode-aware!
# $1 - column1 padding
# $2 - column2 padding
# $3 - column1 text
# $4 - column2 text
# $5 - divider char
	local pad1=$1 pad2=$2
	local str1="$3" str2="$4" div=${5:-' '} 
	local w=$(tput cols); let "w--"; [[ $w -gt 120 ]] && w=100
	local rightpad=4 gap=3 # predefined margins
	local s1 s2
	
	[[ $gap -lt 3 ]] && div=''
	[[ $pad2 -lt $pad1 ]] && pad2=$pad1
	[[ $pad1 -eq 0 ]] && let "pad2--"
	
	local wcol1=$(($pad2-$pad1-$gap)) 
	local oldifs="$IFS"; IFS=$'\n'
	[[ -n "$str1" ]] &&
		local col1=( $( echo "$str1" | fmt -w $wcol1 ) )
	[[ -n "$str2" ]] &&
		local col2=( $( echo "$str2" | fmt -w $(($w-$pad2-$rightpad)) ) )
	IFS="$oldifs"
	[[ ${#col1} -eq 0 && ${#col2} -eq 0 ]] && return
	
	wcol1=$((0-$wcol1))
	[[ ${#col1[@]} -gt ${#col2[@]} ]] &&
		count=${#col1[@]} || count=${#col2[@]}; 
	for (( i=0; i < $count; i++ )); do
		s1="${col1[$i]}"
		s2="${col2[$i]}"
		[[ ${#s1} -gt $((0-$wcol1)) ]] && s1=${s1:0:$((0-$wcol1-1))}
		printf "%${pad1}s%${wcol1}s%${gap}s%s\n" ' ' "$s1" "$div " "$s2"
		div=''
	done
} # END print-column
function unique () { # rm dup words in string; 'echo' removes awk tail
	[[ "${1// }" == '' ]] && { echo; return; }
	echo $( awk 'BEGIN{RS=ORS=" "}!a[$0]++' <<<"$1 " )
}
function inarray () { # test item ($1) in array ($2 - passing by name!)
    local s=$1
    local name=$2[@] # var name, NOT value!
    local a=("${!name}")
	comm -1 -2 -i <(printf '%s\n' "${s[@]}" | sort -u) <(printf '%s\n' "${a[@]}" | sort -u)
}
function math-diffexpr () { # simple unit calculator
	local dif difpc exp='{
		e1=toupper($1); e2=toupper($2); child=$3}; END {
		u1=substr(e1,length(e1),1); u2=substr(e2,length(e2),1)
		x=match(u1,"[KMG]"); if (!x) u1="K" # type cast
		x=match(u2,"[KMG]"); if (!x) u2=u1; 
		kilo=1024; unit="K"
		m1=e1; sub(/[A-Z]/,"",m1); m2=e2; sub(/[A-Z]/,"",m2);
		m1 = 0 + m1; m2 = 0 + m2
		if (!m1 && !m2 || m1 == 0) {printf "%s|", "0"; exit 1}

		if (u1 == "G") m1*=(1024*1024); else
			if (u1 == "M") m1*=1024
		if (u2 == "G") m2*=(1024*1024); else
			if (u2 == "M") m2*=1024
		if (m2 > m1) {r=m1; m1=m2; m2=r} # swap or exit ???
		r = m1-m2; pc = (child) ? m2/m1 : r/m1; pc*=100
		if (r > kilo) {r/=1024; unit="M"}; if (r > kilo) {r/=1024; unit="G"}

		fmt1 = (unit == "G") ? "%.1f" : "%.f"
		if (r%1 >= 0.95) fmt1="%.f"; if (r == 0) unit=""	
		fmtpc = "%.f%%"	
		if (pc < 1) {pc="\\<1"; fmtpc="%s%%"} else
			if (pc > 99.5) {pc="abs"; fmtpc="%s"}
		printf "dif="fmt1"%s; difpc="fmtpc, r, unit, pc
	}'
	eval $(echo "${@}" | awk "$exp"); echo "$dif|$difpc"
} # END difference of byte expressions

function mktest () {
# development environment; I use it for testing; test data manager is in another script
# remove this function (and input parsing option '--test') before public distribution
	WordPATH="Microsoft Word"
	ExcelPATH="Microsoft Excel"
	PowerPointPATH="Microsoft PowerPoint"
	OutlookPATH="Microsoft Outlook"
	OneNotePATH="Microsoft OneNote"
	basePATH=~/"Desktop/msotest/"
	[[ ! -d "$basePATH" ]] && exit 1
}

############# Your show begins here
main "${@}"
