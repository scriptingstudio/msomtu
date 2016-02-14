#!/bin/bash

# Welcome user!
# To get help just run this script without parameters.

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
version='2.8.21'
util="${0##*/}"
principalname='msomtu'
defapp='w e p o n'
deflang='ru' # user pref; english is added in create-filter
defproof='russian' # user pref; english is added in create-filter

WordPATH="Microsoft Word.app" # hashtable: bash4
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
 # I included it for deletion 
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
	local PARGS=(); local INPUTPARAM=''
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
			-report|-rep) cmd+=" report"; cmd_report=1 ;;
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
	# backup? backup first; exit
	# cache? cache last
	# fontset? fs first; exit
	# report? rep first; exit
	[[ $cmd_uninstall -eq 1 ]] && cmd="uninstall"
	if [[ -n "$cmd_backup" ]];     then cmd="backup"; fi
	if [[ "$cmd_cache" == 1 ]];    then cmd="$cmd cache"; fi
	if [[ "$cmd_fontset" == 1 ]];  then cmd="fontset"; fi
	if [[ "$cmd_report" == 1 ]];   then
		cmd_app="$defapp"
		cmd="report"
	fi
	[[ "$script_params" == 0 || "$cmd_help" -eq 1 || "$cmd" == '' ]] && cmd='help'
	# END input parser
	
	############ Operation selector
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
				display-initialDU
				echo
				clean-application
				echo
				display-finalDU ;;
				
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
	local temp=$(unique "$cmd_app")
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
	__normalize-number () {
		unit='KB'
		[[ $fs -gt 1000 ]] && { let "fs/=1024"; unit='MB'; }
		[[ $fs -gt 1000 ]] && { let "fs/=1024"; unit='GB'; }
	}
	__separate-unit () { fs=${fs/G/ G}; fs=${fs/M/ M}; }
	local versionPATH="/Contents/Info.plist"
	local versionKey="CFBundleShortVersionString"
	local fmt1='   %-18s : %8s  %10s\n'; local na='--'
	for appPATH in "${appPathArray[@]}"; do
		printb "Processing '$appPATH'"
		local wpath="$basePATH$appPATH$fontPATH/DFonts"
		local appVersion=$(defaults read "$basePATH$appPATH$versionPATH" $versionKey)
		local msbuild=$(defaults read "$basePATH$appPATH$versionPATH" "MicrosoftBuildNumber")
		printf "$fmt1" "Version (build)" "$appVersion" "($msbuild)"
		##### fonts
		if [[ -d "$wpath" ]]; then
			fs=''; fs=$(du -sh "$wpath" | cut -f 1)
			fc=$(ls -A "$wpath" | wc -l) #fc=$(find "$wpath" -type f | wc -l)
			__separate-unit
			printf "$fmt1" "DFonts" "${fc// }" "${fs}B"
		else
			printf "$fmt1" "DFonts" "does not exist"
			cmd_font=''
		fi
		wpath="$basePATH$appPATH$fontPATH/Fonts"
		fs=''; fs=$(du -sh "$wpath" | cut -f 1)
		fc=$(ls -A "$wpath" | wc -l) 
		__separate-unit
		printf "$fmt1" "Fonts" "${fc// }" "${fs}B"
		##### fontlists
		wpath="$basePATH$appPATH$fontPATH"
		#local plist=$(find "$wpath" -type f -name font*.plist -d 1)
		local plist=("$wpath/"font*.plist)
		if [[ -z "$plist" ]]; then
			printf "$fmt1" "Plists" $na
		else
			fs=$(du -sh -k "$wpath/"*.plist | awk '{ total += $1 }; END {print total}')
			__normalize-number
			fc=$(echo "$flist" | wc -l)
			printf "$fmt1" "Plists" "${fc// }" "$fs $unit"
		fi
		##### UI languages
		local filter=$( create-filter "lang" )
		local flist=$(find -E "$wpath" -type d -d 1 -name *.lproj ! -iregex $filter)
		if [[ -z "$flist" ]]; then
			printf "$fmt1" "Langpacks" $na
		else
			fs=$(du -sh -k "$wpath/"*.lproj | awk '{ total += $1 }; END {print total}')
			__normalize-number
			fc=$(echo "$flist" | wc -l)
			printf "$fmt1" "Langpacks" "${fc// }" "$fs $unit"
		fi
		##### proofing tools
		wpath="$basePATH$appPATH$proofingPATH$proofingName"
		filter=$( create-filter "proof" )
		flist=$(find -E "$wpath" -type d -d 1 -name *.proofingtool ! -name Grammar.proofingtool ! -iregex $filter)
		if [[ -z "$flist" ]]; then
			printf "$fmt1" "Proofingtools" $na
		else
			fs=''; fs=$(du -sh "$wpath" | cut -f 1)
			fc=$(echo "$flist" | wc -l)
			__separate-unit
			printf "$fmt1" "Proofingtools" "${fc// }" "${fs}B"
		fi
		##### total disk usage
		fs=$(du -sh "$basePATH$appPATH" | cut -f 1)
		fc=$(ls -AR "$basePATH$appPATH" | wc -l)
		__separate-unit; fc=$(printf "%'d" $fc)
		printf "$fmt1" "Total app bundle" "$fc" "${fs}B"

		echo
	done # app selection
	##echo "[*] Total includes all of the files in the application container."
	echo
} # END report

function clean-application () {
	local fmt1='   %-14s : %6s %s\n'
	for appPATH in "${appPathArray[@]}"; do
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
				[[ $cmd_verb -eq 1 ]] &&
					find "$wpath" -type f -name font*.plist -d 1 -exec basename {} \;
			fi
		fi # END plist (fontlist files)
	
	# - cleaning of lproj folders; keep en_GB.lproj en.lproj
		[[ "$cmd_lang" != '' ]] && clean-lang "$wpath"
	
	# - cleaning of Proofing Tools; keep English*.proofingtool Grammar.proofingtool
		[[ "$cmd_proof" != '' ]] &&
			clean-ptools "$basePATH$appPATH$proofingPATH$proofingName"

		echo
	done
} # END cleanup wrapper

function clean-font () {
	local wpath="$1"
	if [[ ! -d "$wpath" ]]; then
		echo "  Folder 'DFonts' does not exist."
		return
	fi
	if [[ $cmd_reverse -eq 1 ]]; then
		find-newfont "$wpath"
		return
	fi
	
	cmd_font=$(unique "$cmd_font")
	local fl=(); local name=''
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
	local filter=$(create-filter 'exclude')
	if [[ $run -eq 1 ]]; then # action mode
		for f in $cmd_font; do
			[[ "$f" == 'folder' ]] && rm -fdr "$wpath"
			[[ "$f" == lib* ]] && remove-duplicate "$wpath"
		done
		for f in "${fl[@]}"; do 
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
	for f in $cmd_font; do
		[[ "$f" == 'folder' ]] && ls -A "$wpath"
		[[ "$f" == lib* ]] && remove-duplicate "$wpath"
	done
	for f in "${fl[@]}"; do
		if [[ "$filter" != '' ]]; then
			find -E "$wpath" -type f -iname "$f" -d 1 ! -iregex "$filter" -exec basename {} \; 
		else
			find "$wpath" -type f -iname "$f" -d 1 -exec basename {} \;
		fi
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
	local fsfiles=''; local allfiles=(); local list=''
	echo "Searching for new fonts..."
	for i in ${dfontsets[@]}; do
    	local name=$i[@]
		local fs=("${!name}")
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
	local newfont=$( comm -2 -3 -i <(echo "${allfiles[@]}" | sort -u) <(echo "${fsfiles[@]}" | sort -u) )
	if [[ "${#newfont}" -eq 0 ]]; then
		echo "No new fonts found."
		return
	fi
	if [[ $run -eq 0 ]]; then
		echo "  TRACE: remove new fonts."
		echo "$newfont"
	else
		for f in "${newfont[@]}"; do # ???
			echo "Removing '$f'..."
			#find "$wpath" -type f -d 1 -name "$f" -exec rm -f {} \;
		done
	fi
} # END new fonts

function clean-lang () {
	__list-lang () { find -E "$wpath" -type d -d 1 -name *.lproj ! -iregex $filter -exec basename {} \; ; }
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
	local wpath="$1"
	local filter=$( create-filter "proof" )
	if [[ cmd_reverse -eq 0 ]]; then
		if [[ $run -eq 1 ]]; then
			find -E "$wpath" -type d -d 1 -name *.proofingtool ! -name Grammar.proofingtool ! -iregex $filter -exec rm -fdr {} \;
		else # trace mode
			cmd_proof=$(unique "english $cmd_proof")
			echo "  TRACE: remove extra proofing tools. Keep '$cmd_proof'."
			[[ $cmd_verb -eq 1 ]] &&
				find -E "$wpath" -type d -d 1 -name *.proofingtool ! -name Grammar.proofingtool ! -iregex $filter -exec basename {} \;
		fi
	else # reverse filter
		if [[ $run -eq 1 ]]; then
			find -E "$wpath" -type d -d 1 -name *.proofingtool -iregex $filter -exec rm -fdr {} \;
		else # trace mode
			echo "  TRACE: remove extra proofing tools - '$cmd_proof'."
			[[ $cmd_verb -eq 1 ]] &&
				find -E "$wpath" -type d -d 1 -name *.proofingtool -iregex $filter -exec basename {} \;
		fi
	fi
} # END proofingtools

function display-fontset () {
	local fmt2='%*s %-13s  %s %b\n'
	p4=4; p20=20;
	local fs1=$(joina , "${sysfonts[@]}")
	local fs2=$(joina , "${chfonts[@]}")
	local fs3=$(joina , "${noncyr[@]}")
	local fs4=$(joina , "${cyrdfonts[@]}")
	local fs5=$(joina , "${cyrfonts[@]}")
	local fs6=$(joina , "${symfonts[@]}")
	printb "Predefined fontsets:"
	print-table $p4 $p20 "$fmt2" "sysfonts" "OS duplicated fonts (in DFonts):" "-"
	print-table $p4 $p20 "$fmt2" "" "${fs1//,/, }"
	echo
	print-table $p4 $p20 "$fmt2" "chinese" "All kind of hieroglyphic/eastern fonts (in DFonts):" "-"
	print-table $p4 $p20 "$fmt2" "" "${fs2//,/, }"
	echo
	print-table $p4 $p20 "$fmt2" "noncyr" "Non-cyrillic fonts (in DFonts):" "-"
	print-table $p4 $p20 "$fmt2" "" "${fs3//,/, })"
	echo
	print-table $p4 $p20 "$fmt2" "cyrdfonts" "Cyrillic original fonts (in DFonts; do not include 'sysfonts'):" "-"
	print-table $p4 $p20 "$fmt2" "" "${fs4//,/, }"
	echo
	print-table $p4 $p20 "$fmt2" "cyrfonts" "Cyrillic original fonts (in Fonts):" "-"
	print-table $p4 $p20 "$fmt2" "" "${fs5//,/, }"
	echo
	print-table $p4 $p20 "$fmt2" "symfonts" "Symbolic fonts (in DFonts):" "-"
	print-table $p4 $p20 "$fmt2" "" "${fs6//,/, }"
	echo
} # END fontsets

function invoke-backup () { # for fonts only
	if [[ -z $cmd_backup ]]; then
		printf '%s\n\n' "ERROR: Invalid set of backup arguments specified."	
		return -1
	fi
	
	printb "Backup fonts of '$appPathArray'"
	local bsrc="$basePATH$appPathArray$fontPATH"; local bset=()
	local bdest="${cmd_backup}"
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
	local ffolder="DFonts"
	for f in $(unique "${cmd_font}"); do # expand array
		[[ $f == userlib || $f == syslib ]] && continue
		[[ $f == folder ]] && f="*.*"
		[[ "$f" == 'cyrfonts' || "$f" == 'msofonts' ]] && ffolder='Fonts'
		m=$(inarray "$f" allfontsets)
		if [[ $m ]]; then
			local name=$m[@]; a=("${!name}")
			bset+=("${a[@]}")
		else
			bset+=("$f")
		fi
	done # fontset selector
	bsrc+="/$ffolder"; local fc=0
	if [[ $run -eq 1 ]]; then
		for i in "${bset[@]}"; do
			local fl=$(find "$bsrc" -type f -iname "$i" -d 1)
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
			local fl=$(find "$bsrc" -type f -iname "$i" -d 1 -exec basename {} \;)
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
	local p3=3; local p4=4; local p6=6; local p8=8; local p20=20
	local fs="${dfontsets[@]/chfonts/chinese}"; fs=${fs// /, }

	if [[ "${LANG%\.*}" != "en_US" && "${LANG%\.*}" != "en_GB" && -z $nl ]]; then
		local un="${util%-*}"
		local helpfile="${0%/*}/$un-help.sh" #"${0%/*}/${util/.sh/-ru.sh}"
		[[ -f "$helpfile" ]] && { . "$helpfile"; exit 0; }
	fi

	printb "SYNOPSIS:"
	print-column 0 $p4 "" "MSOMTU is Microsoft Office maintenance utility."
	echo
	
	printb "DESCRIPTION:"
	print-column 0 $p4 "" "Microsoft Office 2016 for Mac uses an isolated resource architecture (sandboxing), so apps duplicate all of the components in its own application container that's waisting gigabytes of disk space. This script safely removes (thinning) extra parts of the folowing components: UI languages; proofing tools; fontlist files (.plist); OS duplicated font files. It also can backup/copy font files to predefined or user defined destinations." 
	echo

	printb "NOTES:"
	print-column 0 $p6 "" "Safe scripting technique - 'Foolproof' or 'Harmless Run'. Default running mode is view. You cannot change or harm your system without switch '-run'. Parameter '-cache' does not depend on '-run'." '-'
	print-column 0 $p6 "" "As MSO is installed with root on /Applications directory you have to run this script with sudo to make changes." '-'
	print-column 0 $p6 "" "File operations are case insensitive." '-'
	print-column 0 $p6 "" "As application font structure has been changed since MSO version 15.17 font deletion only works with 15.17 or later." '-'
	print-column 0 $p6 "" "If you remove fonts, remove font lists as well. 'DFonts' folder and font lists are safe to remove. Some of the fonts you may find useful, save them before deletion." '-'
	print-column 0 $p6 "" "Caution: do not remove fonts from 'Fonts' folder! These are minimum needed for MSO applications to work." '-'
	print-column 0 $p6 "" "Predefined fontsets do not intersect." '-'
	print-column 0 $p6 "" "Script only accepts named parameters." '-'
	print-column 0 $p6 "" "Apply thinning after every MSO update." '-'
	print-column 0 $p6 "" "You can change default settings in code for your needs." '-'
	echo

	printb "USAGE:"
	print-column 0 $p4 "" "[sudo] $util [-<parameter> [<arguments>]]..."
	echo
	print-column 0 $p4 "" "[sudo] $util [-backup|-fcopy [<destination>]] [-app [<app>]] [-font [<font_pattern>]] [-ex|-x <font_pattern>] [-run]"
	echo
	print-column 0 $p4 "" "[sudo] $util [-app [\"<app_list>\"]] [-lang|-ui [\"<lang_list>\"]] [-proof|-p [\"<proof_list>\"]] [-font [<font_pattern>]] [-flist|-fl] [-ex|-x <font_pattern>] [-cache] [-report|-rep] [-verbose|-verb] [-fontset|-fs] [-all|-full] [-rev] [-help|-h|-?] [-run]"
	echo 

	local mp4=-4; local p12=12
	printb "USE CASES:"
	print-padding $mp4 "- Geting MSO info" 
		print-column 0 $p12 "" "Parameter '-report'."
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
	
	printb "ARGUMENTS:"
	print-column $p4 $p20 "app_list" "App list. w - Word, e - Excel, p - PowerPoint, o - Outlook, n - OneNote. Default: 'w e p o n'." ":"
	print-column $p4 $p20 "lang_list" "Langauge list: ru pl de etc, see filenames with parameter '-verb'. Default: 'en ru'." ":"
	print-column $p4 $p20 "proof_list" "Proofingtools list: russian finnish german etc, see filenames with parameter '-verb'. Wildcard '*' is available. Default: 'english russian'." ":"
	print-column $p4 $p20 "font_pattern" "Font operations are based on patterns. Font patterns: empty - removes folder 'DFonts' (default); <fontset> - removes fonts of predefined fontset; <mask> - removes selection: *.*, arial*, *.ttc etc. If you use single '*' enclose it in quotation marks: \"*\". Predefined fontsets: library, $fs. See parameter '-fontset' and details in code. Fontset 'library' removes duplicates of system and user libraries; it may not exactly match fonts because based on file-by-file (unlike font family) comparison (DFonts against libraries). You can use list of fontsets." ":"
	print-column $p4 $p20 "destination" "Backup destination folderpath for fonts. Default value is '~/Desktop/MSOFonts'. You can use predefined destinations: 'syslib' - system library; 'userlib' - user library." ":"
	echo

	printb "PARAMETERS:"
	print-column $p4 $p20 "-app" "Filter <app_list>. Selects application to process." ":"
	print-column $p4 $p20 "-lang" "Exclusive filter <lang_list>. Removes UI languages except defaults and user list. See also parameter '-rev'; it reverses user selection except defaults." ":"
	print-column $p4 $p20 "-proof" "Exclusive filter <proof_list>. Removes proofing tools except defaults and user list. See also parameter '-rev'; it reverses user selection except defaults." ":"
	print-column $p4 $p20 "-font" "Filter <font_pattern>. Removes selected fonts or 'DFonts' folder. Available fontsets: cyrdfonts, noncyr, chinese, sysfonts. Parameter '-rev' ignores user selection and alternates search function: new fonts are going to be discovered. It is useful to check new fonts up after new update." ":"
	print-column $p4 $p20 "-backup" "Backs up fonts to user defined destination. If destination folder does not exist it will be created. You can use system and user libraries as destination, see ARGUMENTS. Backup alternates all deletions to backup." ":"
	print-column $p4 $p20 "-ex" "Exclusive filter <font_pattern>. Excludes font selection with parameter '-font'. Only mask can be used as 'font_pattern'." ":"
	print-column $p4 $p20 "-flist" "Switch. Removes fontlist (.plist) files." ":"
	print-column $p4 $p20 "-all" "Switch. Activates all cleaning options: lang, proof, font, flist, cache. It does not affect a parameter '-app'." ":"
	print-column $p4 $p20 "-cache" "Switch. Cleans font cache." ":"
	print-column $p4 $p20 "-verbose" "Switch. Shows objects to be removed in view mode." ":"
	print-column $p4 $p20 "-report" "Switch. Shows statistics on objects." ":"
	print-column $p4 $p20 "-fontset" "Switch. Shows predefined fontsets." ":"
	print-column $p4 $p20 "-rev" "Switch. Reverses effect of 'lang' and 'proof' filters." ":"
	print-column $p4 $p20 "-help" "Switch. Shows this screen. (Optional)" ":"
	print-column $p4 $p20 "-run" "Switch. Default mode is view (test). Activates operations execution." ":"
	echo
	
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
	  print-padding $p8 "sudo $util -font *.* -ex \"brit* rockwell*\" -app w" b
	print-padding $p4 "Clean font cache:"
	  print-padding $p8 "sudo $util -cache" b
	print-padding $p4 "Backup fonts to default destination:"
	  print-padding $p8 "$util -backup -font \"cyrdfonts britanic*\" -run" b
	print-padding $p4 "Copy original cyrillic fonts to system library:"
	  print-padding $p8 "sudo $util -backup syslib -font cyrdfonts -run" b
	print-padding $p4 "Show predefined fontsets:"
	  print-padding $p8 "$util -fontset" b
	echo
	
	exit 0
} # END help page

function write-log () { echo; } # under construction
function display-initialDU () {
	echo "Before cleaning MSO is taking:"
	diskUsage "${appPathArray[@]}"
} # END print du

function display-finalDU () {
	if [[ $run -eq 1 ]]; then
		echo "After cleaning MSO is taking:"
		diskUsage "${appPathArray[@]}"
		printf '\n%s\n' "Office thinning complete."
	fi
} # END print du

########## Common routines
function create-filter () {
	local l=''; local pre=''; local sfx=''; local search=''
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

function joina { local oldifs="$IFS"; IFS="$1"; shift; echo "$*"; IFS="$oldifs"; }
function printb () { echo -e "\033[1m$1\033[0m"; }
function printu () { echo -e "\033[4m$1\033[0m"; }
function print-padding () {
	[[ ! -n "$2" || -z "$1" ]] && return
	local str=$2; local pad=''
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
	local pad1=$1; local pad2=$2; 
	local str1="$3"; local str2="$4"; local div=${5:-' '} 
	local w=$(tput cols); let "w--"; [[ $w -gt 120 ]] && w=100
	local rightpad=4; local gap=3 # predefined margins
	
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
	if [[ ${#col1[@]} -gt ${#col2[@]} ]]; then 
		count=${#col1[@]}; else count=${#col2[@]}; 
	fi 
	for (( i=0; i < $count; i++ )); do
		local s1="${col1[$i]}"
		local s2="${col2[$i]}"
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
    local a=("${!name}") # expand to value - magic!
	comm -1 -2 -i <(printf '%s\n' "${s[@]}" | sort -u) <(printf '%s\n' "${a[@]}" | sort -u)
}

function diskUsage () {
	for appPATH in "${@}"; do
		du -sh "$basePATH$appPATH"
	done
} # END MSO DU

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
