#!/bin/bash

# Welcome user!
# To get help just run this script without parameters.

##< 
# .SYNOPSIS
#	msomtu - Microsoft Office maintenance utility. Purpose - clean up MSO.
#
# .LINKS
#	Inspiration idea     : https://github.com/goodbest/OfficeThinner
#	On OS X & MSO fonts  : http://www.jklstudios.com/misc/osxfonts.html
#	Git Repo             : https://github.com/scriptingstudio/msomtu
##>

# command line param block
cmd_app=''			# command to select applications
cmd_font=''			# command to view/remove fonts
cmd_lang=''			# command to view/remove UI langs
cmd_proof=''		# command to view/remove proofingtools
cmd_fontlist=''		# command to view/remove font list files
cmd_report=''		# command to view MSO info
cmd_backup=''		# command to copy/backup fonts
cmd_cache=''		# command to clean font cache
cmd_fontset=''		# command to view predefined fontsets
cmd_checkupdate=''	# command to check new version
cmd_help=''			# command to select help pages
cmd_all=''			# switch to force all commands
cmd_verb=''			# switch to activate verbose mode
cmd_inverse=''		# inverting search filter
cmd_exclude=''		# exclusive font search filter
cmd_run=''			# switch to unlock commands
# END params

# Definitions: defaults and constants
toolname="Microsoft Office 2016 Maintenance Utility"
version='2.9.8'
util="${0##*/}"; util="${util%%-*}.sh"; util="${util%%.sh*}.sh"
helpfile="${0%.*}"; helpfile="${helpfile%%-*}-help.sh"
defapp='w e p o n'
deflang='ru'       # user pref; + english never deleted
defproof='Russian' # user pref; + english never deleted
backupPATH=~/Desktop/MSOFonts/ # user pref
basePATH="/Applications/"
fontPATH="/Contents/Resources"
prooftoolPATH="/Contents/SharedSupport/Proofing Tools"
WordPATH="Microsoft Word.app" # declare -A msoapps -> bash4
ExcelPATH="Microsoft Excel.app"
PowerPointPATH="Microsoft PowerPoint.app"
OutlookPATH="Microsoft Outlook.app"
OneNotePATH="Microsoft OneNote.app"

## Predefined fontsets definition block. You can change them here.
## in bash4 it would be a hashtable
 # os x duplicates; in  DFonts folder.
 # exclusion: Microsoft no longer provides cyrillic MonotypeCorsiva, so
 # I included it here for deletion (Ive got it cyrillic in the lib).
sysfonts=("arial*" "ariblk*" "Baskerville*" "Book Antiqua*" "ComicSans*" "Cooper*" "Gill Sans*" "GillSans*" "pala*" "Trebuchet*" "verdana*" "MonotypeCors*")
 # OS X duplicates; in Fonts folder
msoessential=("tahoma*" "Wingding*" "webding*")
 # chfonts - all kinds of hieroglyphic/eastern fonts (in DFonts folder)
chfonts=("Fangsong*" "Deng*" "gulim*" "HGR*" "kaiti*" "malgun*" "Meiryo*" "mingliu*" "MSJH*" "msyh*" "SimHei*" "simsun*" "STH*" "STX*" "STZ*" "STL*" "taile*" "YuGoth*" "yumin*")
noncyr=("Abadi*" "angsa*" "BellMT*" "Bauhaus93*" "BernardMT*" "Calisto MT*" "Braggadocio*" "Britannic*" "CalistoMT*" "ColonnaMT*" "COOPBL*" "CopperplateGothic*" "CurlzMT*" "Desdemona*" "EdwardianScriptITC*" "EngraversMT*" "Eurostile*" "FootlightMT*" "GloucesterMT*" "Goudy Old Style*" "Haettenschweiler*" "Harrington*" "ImprintMTShadow*" "KinoMT*" "Lucida Sans.*" "Lucida Sans Demibold*" "Lucida Sans Italic.*" "LucidaBright*" "LucidaBlackletter.*" "LucidaFax*" "LucidaCalligraphy*" "LucidaHandwriting*" "LucidaSansTypewrite*" "MaturaMTScriptCapitals*" "ModernNo.20*" "News Gothic MT*" "ntailu*" "Onyx*" "Perpetua*" "Rockwell*" "Stencil*" "Tw Cen*" "WideLatin*") # but some of non-cyr fonts may be useful
symfonts=("Bookshelf Symbol*" "Marlett*" "MS Reference Specialty*" "MonotypeSorts*")
cyrdfonts=("batang*" "Bookman Old Style*" "Candara*" "Century*" "Consola*" "Constan*" "Corbel*" "Franklin Gothic*" "Gabriola*" "GARA*" "Lucida Console*" "Lucida Sans Unicode*" "Mistral*" "MS Reference Sans Serif.*" "msgothic*" "Segoe Print Bold.*" "Segoe Script Bold.*")
 # original cyr fonts; in Fonts folder
cyrfonts=("Calibri*" "Cambria*" "Century.*" "Corbel.*")
 # fontset blocks
dfontsets=(cyrdfonts noncyr chfonts sysfonts symfonts)
fontsets=(cyrfonts msoessential)
allfontsets=("${dfontsets[@]}" "${fontsets[@]}")
fsdescriptor=( # for display-fontset function
	"sysfonts|-|OS X duplicated fonts (in DFonts)"
	"chfonts|chinese|All kind of hieroglyphic/eastern fonts (in DFonts)" # 2 replmnt in code
	"noncyr|-|Non-cyrillic fonts (in DFonts)"
	"cyrdfonts|-|Cyrillic original fonts (in DFonts; do not include 'sysfonts')"
	"cyrfonts|-|Cyrillic original fonts (in Fonts)"
	"symfonts|-|Symbolic fonts (in DFonts)"
)
# END Definitions region

	
function main () {
	printb "$toolname. Version $version."
	export starttime=$( perl -MTime::HiRes -e 'print int(1000 * Time::HiRes::gettimeofday)' )
	
	############ Simple input parameter parser v2.2 [inline method]
	local script_params="$#" prefix='cmd_' paramorder='' unknown=''
	local PARGS=() INPUTPARAM='' param cmd='' c
	while [ "$#" != 0 ]; do
		[[ "${1:0:1}" == "-" ]] && { INPUTPARAM="$1"; shift; }
		PARGS=()
		while [ "${1:0:1}" != "-" ] && [ "$#" != 0 ]; do
			PARGS+=("$1") # collect arguments of the current parameter
			shift
		done
		param=''
		case "$INPUTPARAM" in # translate parameters
			-report|-rep|-info|-inf) param="report" ;;
			-app)             param="app" ;;
			-font)            param="font" ;;
			-proof|-p)        param="proof" ;;
			-lang|-ui)        param="lang" ;;
			-flist|-fl)       param="fontlist" ;;
			-cache|-fc)       param="cache" ;;
			-fontset|-fs)     param="fontset" ;;
			-fcopy|-backup)   param="backup" ;;
			-verbose|-verb)   param="verb" ;;
			-inv|-rev)        param="inverse" ;;
			-ex|-x)           param="exclude" ;;
			-run|-ok)         param="run" ;;
			-help|-h|-\?)     param="help" ;;
			-all|-full)       param="all" ;;
			-check)           param="checkupdate" ;;
			*) [[ $INPUTPARAM ]] && unknown+=" $INPUTPARAM" ;;
		esac
		if [[ "$param" ]]; then
			paramorder+=" $param" # actual parameters
			[[ "${#PARGS[@]}" > 0 ]] && # output: paramvar = value of its arguments
				eval $prefix$param='("${PARGS[@]}")' || 
				eval $prefix$param=true
		fi
	done # END input parser
	[[ "$unknown" ]] && { echo "ERROR: unknown parameter(s) '${unknown:1}' specified. See help."; exit 1; }

	# default value parser
	[[ "$cmd_run" && "$cmd_run" != true ]] && cmd_run=''
	[[ "$cmd_exclude" == true ]] && cmd_exclude=''
	[[ "$cmd_inverse" != true ]] && cmd_inverse='' || { deflang=''; defproof=''; }
	[[ "$cmd_backup" == true ]] && cmd_backup="$backupPATH"
	[[ "$cmd_app" == true || -z "$cmd_app" ]] && cmd_app="$defapp"
	if [[ $cmd_all == true ]]; then
		cmd_proof=${cmd_proof:=true}
		cmd_lang=${cmd_lang:=true}
		cmd_font=${cmd_font:=true}; 
		cmd_fontlist=true
	elif [[ "$cmd_all" ]]; then
		cmd_all=''
	fi
	if [[ "$cmd_lang" == true ]]; then 
		cmd_lang="$deflang"
	elif [[ "$cmd_lang" ]]; then
		cmd_lang=$(unique "$deflang $cmd_lang") # default langs come here
	fi
	if [[ "$cmd_proof" == true ]]; then 
		cmd_proof="$defproof"
	elif [[ "$cmd_proof" ]]; then
		cmd_proof=$(unique "$defproof $cmd_proof") # default langs come here
	fi	
	if [[ "$cmd_font" == true ]]; then 
		cmd_font='folder'
	elif [[ "$cmd_font" ]]; then
		cmd_font=${cmd_font//chinese/chfonts} # name/disp name
		cmd_font=$(unique "$cmd_font")
	fi
	
	# group operation list: "clean=(flist font proof lang) check all"
	[[ "$cmd_lang" || "$cmd_proof" || "$cmd_font" ||
		"$cmd_fontlist" == true || "$cmd_all" == true ]] && cmd+=" clean"
	[[ "$cmd_checkupdate" == true ]] && cmd+=" checkupdate"
	
	# solo operation list filter: backup, fontset, report, cache, help
	param='backup fontset report cache help'
	for i in $paramorder; do for s in $param; do
		[[ $s == $i ]] && m=$i # search for the last actual param
	done done
	[[ $m ]] && [[ $m == "${paramorder##* }" || -z "$cmd" ]] && cmd=$m # priority
	[[ $script_params -eq 0 ]] && cmd="help"
	[[ -z "$cmd" ]] && 
		{ echo "No valid command specified. Correct your command line parameters. See help."; exit 3; }
	
	############ Operation selector
	cmd=$(unique "$cmd")
	create-AppPath	
	echo
	for c in $cmd; do
		case $c in
			report) 
				show-appbanner 'skip2' 
				echo
				get-msoinfo ;;
				
			clean) 
				show-appbanner 
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
				
			checkupdate) 
				echo "See in your default web-browser."
				open -a safari "http://macadmins.software" ;;
				
			help)
				show-helppage ;;
		esac
	done # END operation selector
	echo
	[[ $cmd_verb || $cmd_run ]] && {
		runtime=$( perl -MTime::HiRes -e '$s=int(1000 * Time::HiRes::gettimeofday); $s-=$ENV{starttime}; $s-=80; if ($s > 1000) {$s/=1000} else {$s="0.$s"}; print $s' );
		echo "Runtime: $runtime sec"; }
} # END main

function create-AppPath () {
	local temp=$(unique "$cmd_app") appPATH i
	cmd_app=()
	for i in $temp; do # hashtable: bash4
		[[ $i == 'w' ]] && cmd_app+=("$WordPATH")
		[[ $i == 'e' ]] && cmd_app+=("$ExcelPATH")
		[[ $i == 'p' ]] && cmd_app+=("$PowerPointPATH")
		[[ $i == 'o' ]] && cmd_app+=("$OutlookPATH")
		[[ $i == 'n' ]] && cmd_app+=("$OneNotePATH")
	done
	[[ ${#cmd_app[@]} -eq 0 ]] && 
		{ echo "ERROR: Invalid app ID ['${temp[@]}'] specified."; exit 1; }

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

function show-appbanner () {
	local appPATH
	printb "Installed apps:"
	for appPATH in "${appInstalled[@]}"; do
		echo "- $appPATH"
	done
	[[ "$1" ]] && return
		echo
		printb "Apps to process:"
		for appPATH in "${appPathArray[@]}"; do
			echo "- ${appPATH/.app/}"
		done
} # END script header

function get-msoinfo () {
	__separate-unit () { echo "${1/[A-Z]/} ${1:(-1)}"; }
	local versionPATH="/Contents/Info.plist"
	local versionKey="CFBundleShortVersionString"
	local fmt1='   %-18s : %10s  %10s\n' na='--' sfx='' appfat appfiles
	local appPATH wpath appVersion msbuild fs fc plist flist filter
	
	for appPATH in "${appPathArray[@]}"; do
		printb "Application: ${appPATH/.app/}"
		wpath="$basePATH$appPATH$fontPATH/DFonts"
		appVersion=$(defaults read "$basePATH$appPATH$versionPATH" $versionKey)
		msbuild=$(defaults read "$basePATH$appPATH$versionPATH" "MicrosoftBuildNumber")
		printf "$fmt1" "Version (build)" "$appVersion" "($msbuild)"
		appfat=0; appfiles=0
		##### fonts
		if [[ -d "$wpath" ]]; then
			fs=$(du -sh -k "$wpath" | cut -f 1); let "appfat+=${fs}"
			fs=$(normalize-number $fs)
			fc=$(ls -A "$wpath" | wc -l) #fc=$(find "$wpath" -type f | wc -l)
			let "appfiles+=${fc// }"
			printf "$fmt1" "DFonts" "${fc// }" "${fs}$sfx"
		else
			printf "$fmt1" "DFonts" "does not exist"
			cmd_font=''
		fi
		wpath="$basePATH$appPATH$fontPATH/Fonts"
		fs=$(du -sh "$wpath" | cut -f 1)
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
			let "appfat+=${fs}"; fs=$(normalize-number $fs)
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
			let "appfat+=${fs}"; fs=$(normalize-number $fs)
			fc=$(echo "$flist" | wc -l); 
			let "appfiles+=${fc// }"
			printf "$fmt1" "Langpacks" "${fc// }" "${fs}$sfx"
		fi
		##### proofing tools
		wpath="$basePATH$appPATH$prooftoolPATH"
		filter=$( create-filter "proof" )
		flist=$(find -E "$wpath" -type d -d 1 -name *.proofingtool ! -name Grammar.proofingtool ! -iregex $filter)
		if [[ -z "$flist" ]]; then
			printf "$fmt1" "Proofingtools" $na
		else
			fs=''; fs=$(du -sh -k "$wpath" | cut -f 1); let "appfat+=${fs}"
			fs=$(normalize-number $fs)	
			fc=$(echo "$flist" | wc -l); let "appfiles+=${fc// }"
			printf "$fmt1" "Proofingtools" "${fc// }" "${fs}$sfx"
		fi
		##### total disk usage
		fs=$(du -sh "$basePATH$appPATH" | cut -f 1)
		fc=$(ls -AR "$basePATH$appPATH" | wc -l)
		fs=$(__separate-unit $fs); fc=$(printf "%'d" $fc)
		printf "$fmt1" '' '' '------'
		printf "$fmt1" "Total app bundle" "$fc" "${fs}$sfx"
		appfat=$(normalize-number $appfat)
		printf "$fmt1" "Approx. thinning" "$appfiles*" "-${appfat}$sfx"

		echo
	done # app selection
	echo -e "----\n[*] Includes the default file sets which are reserved\n    by settings. The folder 'Fonts' is reserved."
	echo
} # END MSO report

function clean-application () {
	__list-langrc () { # for clean-lang, clean-ptools
		local fs name="$1" wpath="$2" filter="$3"
		if [[ -z $cmd_inverse ]]; then
			[[ "$name" == 'lproj' ]] &&
			fs=$( find -E "$wpath" -type d -d 1 -name "*.$name" ! -iregex $filter -exec basename {} \; ) ||
			fs=$(find -E "$wpath" -type d -d 1 -name "*.$name" ! -name Grammar.proofingtool ! -iregex $filter -exec basename {} \;) 
		else
			fs=$( find -E "$wpath" -type d -d 1 -name "*.$name" -iregex $filter -exec basename {} \; )
		fi
		[[ $cmd_verb != 'nl' && "$fs" ]] && echo "$fs"; 
		[[ "$fs" ]] && update-counter "$fs" "$wpath" 
	} # END listing language resource
	local fmt1='   %-14s : %6s %s\n' appPATH wpath appfat fl
	for appPATH in "${appPathArray[@]}"; do
		appfat=0
		printb "Processing '${appPATH/.app/}'"
	# - removing of font files/folder DFonts
		[[ "$cmd_font" ]] && 
			clean-font "$basePATH$appPATH$fontPATH/DFonts"

	# - removing of .plist files 
		wpath="$basePATH$appPATH$fontPATH"
		if [[ $cmd_fontlist ]]; then
			if [[ $cmd_run ]]; then
				echo "Remove font-list files (.plist)"
				find "$wpath" -type f -name font*.plist -d 1 -exec rm -f {} \;
			else
				echo "TRACE: remove font-list files (.plist)."
				if [[ $cmd_verb ]]; then
					fl=$(find "$wpath" -type f -name font*.plist -d 1 -exec basename {} \;)
					[[ $cmd_verb != 'nl' ]] && echo "$fl"; update-counter "$fl" "$wpath"
				fi
			fi
		fi # END fontlist files
	
	# - cleaning of lproj folders
		[[ "$cmd_lang" ]] && clean-lang "$wpath";
	
	# - cleaning of proofing tools
		[[ "$cmd_proof" ]] &&
			clean-ptools "$basePATH$appPATH$prooftoolPATH";
		
		appfat=$(normalize-number "$appfat" ns)
		appdu+=("${appfat}	$basePATH$appPATH") # tab char between
		echo
	done
} # END cleanup wrapper

function clean-font () {
	local wpath="$1" fl=() name='' f m a filter fc fs
	if [[ ! -d "$wpath" ]]; then
		echo "  Folder 'DFonts' does not exist."
		return
	fi
	if [[ $cmd_inverse ]]; then
		find-newfont "$wpath"
		return
	fi
	
	for f in $cmd_font; do 	# expand fontsets
		[[ $f == 'folder' || $f == 'userlib' || $f == 'syslib' ]] && continue
		m=$(inarray "$f" dfontsets)
		if [[ $m ]]; then
			name=$m[@]; a=("${!name}")
			fl+=("${a[@]}")
		else
			fl+=("$f")
		fi
	done # fontset selector
	filter=$(create-filter 'exclude')
	if [[ $cmd_run ]]; then # action mode
		echo "Remove fonts"
		if [[ "$cmd_font" == 'folder' ]]; then # folder alone
			if [[ "$filter" == '' ]]; then
				rm -fdr "$wpath"
			else
				find -E "$wpath" -type f -d 1 ! -iregex "$filter" -exec rm -f {} \;
			fi
		elif [[ "$cmd_font" == lib* ]]; then 
			remove-duplicate "$wpath"
		fi
		for f in "${fl[@]}"; do # font lists
			echo " Removing '$f'..."
			if [[ "$filter" != '' ]]; then
				find -E "$wpath" -type f -iname "$f" -d 1 ! -iregex "$filter" -exec rm -f {} \;
			else
				find "$wpath" -type f -iname "$f" -d 1 -exec rm -f {} \; 
			fi
		done
		return
	fi
	
	echo "TRACE: remove fonts in '$wpath'."
	[[ -z $cmd_verb ]] && return
	if [[ "$cmd_font" == 'folder' ]]; then # folder alone
		if [[ "$filter" == '' ]]; then
			fs=$(ls -A "$wpath")
		else
			fs=$(find -E "$wpath" -type f -d 1 ! -iregex "$filter" -exec basename {} \;)
		fi
		if [[ "$fs" ]]; then
			[[ $cmd_verb != 'nl' ]] && echo "$fs"
			update-counter "$fs" "$wpath"; 
		fi
	elif [[ "$cmd_font" == lib* ]]; then 
		remove-duplicate "$wpath"
	fi
	for f in "${fl[@]}"; do # font list
		if [[ "$filter" != '' ]]; then
			fs=$(find -E "$wpath" -type f -iname "$f" -d 1 ! -iregex "$filter" -exec basename {} \;)
		else
			fs=$(find "$wpath" -type f -iname "$f" -d 1 -exec basename {} \;)
		fi
		[[ "$fs" ]] &&
		{ [[ $cmd_verb != 'nl' ]] && echo "$fs"; update-counter "$fs" "$wpath"; }
	done
} # END clean fonts

function remove-duplicate () { # fonts
	__invoke-deduplication () {
		local wpath="$3"
		printf '%s' "$4"
		local dup=$( comm -1 -2 -i <(printf '%s\n' "${1}") <(printf '%s\n' "${2}") )
		local fc=$(echo "${dup[@]}" | wc -l); [[ "$dup" == '' ]] && fc=0
		echo "[ ${fc// } ]"
		if [[ ! $cmd_run ]]; then
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
	[[ ! $cmd_run ]] &&
		echo -e "TRACE: find and remove new fonts.\n"
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
	if [[ ! $cmd_run ]]; then
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
	local wpath="$1" filter=$( create-filter "lang" ) cmdrm msg
	if [[ -z $cmd_inverse ]]; then
		cmdrm='find -E "$wpath" -type d -d 1 -name *.lproj ! -iregex $filter -exec rm -fdr {} \;'
		cmd_lang=$(unique "en $cmd_lang") # english never deleted
		msg="Remove optional UI languages. Keep '$cmd_lang'."
	else # inverse filter
		cmdrm='find -E "$wpath" -type d -d 1 -name *.lproj -iregex $filter -exec rm -fdr {} \;'
		msg="Remove optional UI languages - '$cmd_lang'."
	fi
	if [[ $cmd_run ]]; then
		echo "$msg"
		eval '$cmdrm'
	else
		echo "TRACE: $msg"
		[[ $cmd_verb ]] && __list-langrc 'lproj' "$wpath" "$filter"
	fi
} # END lang

function clean-ptools () {
	local wpath="$1" filter=$( create-filter "proof" ) cmdrm msg
	if [[ -z $cmd_inverse ]]; then
		cmdrm='find -E "$wpath" -type d -d 1 -name *.proofingtool ! -name Grammar.proofingtool ! -iregex $filter -exec rm -fdr {} \;'
		cmd_proof=$(unique "English $cmd_proof") # english never deleted
		msg="Remove optional proofing tools. Keep '$cmd_proof'."
	else # inverse filter
		cmdrm='find -E "$wpath" -type d -d 1 -name *.proofingtool -iregex $filter -exec rm -fdr {} \;'
		msg="Remove optional proofing tools - '$cmd_proof'."
	fi
	if [[ $cmd_run ]]; then
		echo "$msg"
		eval '$cmdrm'
	else
		echo "TRACE: $msg";
		[[ $cmd_verb ]] && __list-langrc 'proofingtool' "$wpath" "$filter";
	fi
} # END proofingtools

function display-fontset () {
	local fset name array fn f disp desc
	printb "Predefined fontsets:"
	for f in "${fsdescriptor[@]}"; do
		fn="${f%%|*}"; disp="${f%|*}"; disp="${disp#*|}"; desc="${f##*|}"
		[[ "$disp" == '-' ]] && disp="$fn"
		name=${fn}[@]; array=("${!name}")
		fset=$(joina ',' "${array[@]}")
		print-row 4 20 "$disp" "$desc:" "-"
		print-row 4 20 "" "${fset//,/, }"
		echo
	done
} # END fontsets

function invoke-backup () { # for fonts only
	if [[ -z $cmd_backup ]]; then
		printf '%s\n\n' "ERROR: Invalid set of backup arguments specified."	
		return -1
	fi
	
	local bsrc="$basePATH$appPathArray$fontPATH" bset=()
	local bdest="${cmd_backup}"
	local ffolder="DFonts" f i m a name fl fc=0 c
	printb "Backup fonts of '$appPathArray'"
	if [[ "$bdest" == 'syslib' ]]; then
		bdest="/Libraries/Fonts/"
	elif [[ "$bdest" == 'userlib' ]]; then
		bdest=~/Libraries/Fonts/
	fi
	if [[ ! -d "$bdest" ]]; then
		if [[ $cmd_run ]]; then
			mkdir -p "$bdest"
			if [[ $? -ne 0 ]]; then 
				echo "ERROR: failed to create directory '$bdest'."
				bdest=''
			fi
		fi
	fi
	[[ -z "$bdest" ]] && return 
	
	[[ -z "$cmd_font" ]] && cmd_font="*.*"
	for f in $(unique "${cmd_font}"); do # expand array
		[[ $f == 'userlib' || $f == 'syslib' ]] && continue
		[[ $f == 'folder' ]] && f="*.*"
		[[ "$f" == 'cyrfonts' || "$f" == 'msoessential' ]] && ffolder='Fonts'
		m=$(inarray "$f" allfontsets)
		if [[ $m ]]; then
			name=$m[@]; a=("${!name}")
			bset+=("${a[@]}")
		else
			bset+=("$f")
		fi
	done # fontset selector
	bsrc+="/$ffolder"
	if [[ $cmd_run ]]; then
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
			[[ $cmd_verb ]] && echo "$fl"
		done
		echo "-----------"
		echo "Total files : ${fc// }."
		echo "Destination : '$bdest'."
		[[ ! -d "$bdest" ]] &&
			echo "TRACE: cannot access '$bdest'."
	fi
} # END backup

function show-helppage () {
	print-topic () {
	# $1 - opt-flag/data; $2 - pad1; $3 - pad2; $4 - lineheight/delim
		[[ $1 == 'full' ]] && shift || return
		local name=$1[@]; local array=("${!name}")
		local pad1=$2 pad2=$3 item h2 val delim lh
		[[ "$4" == 'lh' ]] && { lh=$4; shift; }
		delim=$4
	
		[[ ${#array[@]} -lt 2 ]] && return
		[[ "${array[0]}" ]] && echo -e "\033[1m"${array[0]}:"\033[0m"
		for item in "${array[@]:1}"; do
			[[ -z $item ]] && continue
			h2="${item%%||*}"; val="${item#*||}"
			[[ "$h2" == "$val" ]] && h2=''
			if [[ $pad1 -gt -1 && $pad2 -gt -1 ]]; then # type 1
				print-row $pad1 $pad2 "$h2" "$val" $delim
			elif [[ $pad1 -lt 0 && $pad2 -gt -1 ]]; then # type 2
				print-padding $pad1 "$h2$delim"
				print-row 0 $pad2 "" "$val"
			else # type 3
				print-padding $pad1 "$h2$delim"
				print-padding $pad2 "$val" b
			fi
			[[ "$lh" ]] && echo
		done
		[[ -z "$lh" ]] && echo
	} # END print topic
	local fs="${dfontsets[@]/chfonts/chinese}"; fs=${fs// /, } # name/disp name
	local opt=$(inarray 'full' cmd_help)

	if [[ "${LANG%\.*}" != en_* && -z $(inarray 'en' cmd_help) ]]; then
		[[ -f "$helpfile" ]] && { . "$helpfile"; return; }
	fi

	SYNOPSIS=(
		"SYNOPSIS"
		"MSOMTU is Microsoft Office Maintenance Utility. The solid fitness solution for your Office."
	)
	DESCRIPTION=(
		"DESCRIPTION"
		"Microsoft Office 2016 for Mac uses an isolated resource architecture (sandboxing), so the MSO apps duplicate all of the components in its own application container that's waisting gigabytes of your disk space. This script safely removes (thins) extra parts of the following components: UI languages; proofing tools; fontlist files (font*.plist); OS X duplicated font files. It also can backup/copy font files to predefined and user defined destinations."
	)
	NOTES=(
		"NOTES"
		"Safe Scripting technique - \"Foolproof\" or \"Harmless Run\". The default running mode is view. You can think of it as \"what-if\" mode. The script cannot make changes or harm your system without parameter '-run'."
		
		"As MSO is installed with root on '/Applications' directory you have to run this script with sudo to make changes."
	
		"As application font structure has been changed since MSO version 15.17 font deletion only works with 15.17 or later. Microsoft separated font sets for some reasons. Essential fonts to the MSO apps are in the 'Fonts' folder within each app. The rest are in the 'DFonts' folder."
	
		"If you remove fonts, remove font lists (font*.plist) as well; see PARAMETERS. The 'DFonts' folder and font lists are safe to remove. No third party app can see MSO fonts installed to the 'DFonts' and 'Fonts' folders. Some of the fonts you may find useful, save them before deletion."
	
		"Caution: do not remove fonts from the 'Fonts' folder! These are minimum needed for the MSO applications to work."
	
		"File operations are case insensitive."
	
		"The script only accepts named parameters. In case of parameter duplicates the latter wins."
	
		"Apply thinning after every MSO update."
	
		"Default settings for the '-lang' and '-proof' parameters: english and russian. It depends on your system locale and common sense: for MSO integrity it is better to leave english. You can change any default settings in code for your needs. Default languages are reserved from deletion."
	
		"Font classification spicifics in predefined fontsets: cyrillic, non-cyrillic, hieroglyphic, symbolic, system. Fontsets do not intersect."
	)
	USAGE=(
		"USASE"
		"[sudo] $util [-<parameter> [<arguments>]]..."
	
		"[sudo] $util -backup [<destination>] [-app [<app>]] [-font [<font_pattern>]] [-ex|-x <font_pattern>] [-run|-ok]"
	
		"[sudo] $util [-app [\"<app_list>\"]] [-lang|-ui [\"<lang_list>\"]] [-proof|-p [\"<proof_list>\"]] [-font [<font_pattern>]] [-flist|-fl] [-ex|-x <font_pattern>] [-cache|-fc] [-report|-rep|-info] [-verbose|-verb [nl]] [-fontset|-fs] [-all|-full] [-inv] [-help|-h|-? [en] [full]] [-run|-ok]"
	)
	USE_CASES=(
		"USE CASES"
		"Solo actions: 'report', 'fontset', 'backup', 'cache', 'help'.||"
		"- Getting MSO info||Parameter '-report'."
		"- Getting assessment of thinning (view mode)||Parameter '-verbose' along with resource selector: font, flist, lang, proof."
		"- Listing/Removing UI languages||Parameter '-lang'."
		"- Listing/Removing proofingtools||Parameter '-proof'."
		"- Listing/Removing fonts||Parameter '-font'."
		"- Listing/Removing font duplicates||Parameter '-font lib'."
		"- Listing/Removing font list files||Parameter '-fontlist'."
		"- Listing fontsets||Parameter '-fontset'."
		"- Clearing font cache||Parameter '-cache'."
		"- Finding new fonts||Parameter '-font -inv'."
		"- Backing up fonts||Parameter '-backup'."
		"- Copying fonts to font libraries||Parameter '-backup'."
		"- Checking for new versions||Parameter '-check'."
	)
	ARGUMENTS=(
		"ARGUMENTS"
		"app_list||App list. w - Word, e - Excel, p - PowerPoint, o - Outlook, n - OneNote. Default: 'w e p o n'."
	
		"lang_list||Langauge list: ru pl de etc, see filenames with parameter '-verbose'. Default: 'en ru', see NOTES."
	
		"proof_list||Proofingtools list: russian finnish german etc, see filenames with parameter '-verbose'. Wildcard '*' is available. Default: 'english russian', see NOTES."
	
		"font_pattern||Font operations are based on patterns. Font patterns: empty - removes the 'DFonts' folder (default); <fontset> - removes fonts of predefined fontset; <mask> - removes selection: *.*, arial*, *.ttc etc. If you use single '*' enclose it in quotation marks: \"*\". Predefined fontsets: library, $fs. See parameter '-fontset' and details in code. Fontset 'library' removes duplicates of system and user libraries; it may not exactly match fonts because based on file-by-file (unlike font family) comparison (DFonts against libraries). You can use list of fontsets as well."
	
		"destination||Backup destination folderpath for fonts. Default value is '~/Desktop/MSOFonts'. You can use predefined destinations as well: 'syslib' - system library; 'userlib' - user library."
	)
	PARAMETERS=(
		"PARAMETERS"
		"-all||Switch. Activates all cleaning options: lang, proof, font, flist, cache. It does not affect a parameter '-app'."
	
		"-app||Filter <app_list>. Selects application to process."
	
		"-backup||Backs up fonts to user defined destination. If destination folder does not exist it will be created. You can use system and user libraries as destination, see ARGUMENTS. Backup alternates all deletions to backup."
	
		"-cache||Switch. Cleans up font cache."
		"-check||Switch. Checks for new versions; opens the web-page in browser."
	
		"-ex||Exclusive filter <font_pattern>. Excludes font selection with parameter '-font'. Only mask can be used as 'font_pattern'."
	
		"-flist||Switch. Removes fontlist (font*.plist) files. Fontlists are like cache. When you remove unneeded fonts you can also have to clear all non existent fonts from its lists. Since discovering fonts through all lists is difficult remove all of the .plist files. They all have to do with the fixed font lists you see in Office."
	
		"-font||Filter <font_pattern>. Removes selected fonts or the 'DFonts' folder. Available fontsets: cyrdfonts, noncyr, chinese, sysfonts. Parameter '-inv' ignores user selection and alternates search function: new fonts are going to be discovered. It is useful to check new fonts up after new update. Argument 'library' alters searching in libraries for duplicates."

		"-fontset||Switch. Shows predefined fontsets."
	
		"-help||Switch. Shows the help page. Text language depends on your locale. There are two kinds of help page: short and full. The default is short one (no parameters). To get the full page use parameter '-help full'. Special argument 'en' forces english help page."
	
		"-inv||Switch. Inverts effect of the 'lang' and 'proof' filters, but defaults are reserved. For parameter '-font' it is to search for the new fonts."
	
		"-lang||Exclusive filter <lang_list>. Removes UI languages except defaults and user list. See also parameter '-inv'; it inverts user selection except defaults."
	
		"-report||Switch. Shows statistics on objects."
	
		"-proof||Exclusive filter <proof_list>. Removes proofing tools except defaults and user list. See also parameter '-inv'; it inverts user selection except defaults."
	
		"-run||Switch. The default mode is view (test). Activates operations execution."
	
		"-verbose||Switch. View mode: shows objects to be removed. With special argument 'nl' skips file listing. It does not depend on '-run'."
	)
	EXAMPLES=(
		"EXAMPLES"
		"Get app statistics||$util -report"
		"Thin all apps with all parameters||sudo $util -all -run"
		"Show app ('w e' for Word and Excel) language files installed||$util -app \"w e\" -lang -verbose"
		"Remove a number of languages||sudo $util -lang \"nl no de\" -inv -run"
		"Remove all proofing tools except defaults for Word||sudo $util -proof -app w -run"
		"Remove a number of proofing tools||sudo $util -proof \"Indonesian Isix*\" -inv -run"
		"Show duplicates of library fonts for Word||$util -font lib -app w -verbose"
		"Remove duplicated fonts in libraries for Word||sudo $util -font lib -app w -run"
		"Remove 'chinese' and Arial fonts||sudo $util -font \"chinese arial*\" -run"
		"Show new fonts for Outlook||$util -font -inv -app o"
		"Exclude a few useful fonts from deletion for Word||sudo $util -font *.* -ex \"brit* rockwell*\" -app w -run"
		"Clean font cache||sudo $util -cache"
		"Backup fonts to default destination||$util -backup -font \"cyrdfonts britanic*\" -run"
		"Copy original cyrillic fonts to system library||sudo $util -backup syslib -font cyrdfonts -run"
		"Show predefined fontsets||$util -fontset"
	)
	LINKS=(
		"RELATED LINKS"
		"- Inspiration idea of 'thinning'||https://github.com/goodbest/OfficeThinner"
		"- On OS X & MSO fonts||http://www.jklstudios.com/misc/osxfonts.html"
		"- The project Github repo||https://github.com/scriptingstudio/msomtu"
	)
	
	print-topic $opt   SYNOPSIS     0 4
	print-topic $opt   DESCRIPTION  0 4
	print-topic $opt   NOTES        0 6 '-'
	print-topic full   USAGE        0 4 'lh'
	print-topic $opt   USE_CASES   -4 12 ':'
	print-topic full   ARGUMENTS    4 20 ':'
	print-topic full   PARAMETERS   4 20 ':'
	print-topic $opt   EXAMPLES    -4 -8 ':'
	print-topic $opt   LINKS       -4 12 ':'
} # END help page

function clean-cache () { echo; atsutil databases -remove; }

function show-appdu () { # final/assessment report
	[[ ! $cmd_run && -z $cmd_verb ]] && return
	local s1 s2 a1 a2 as1 as2 du1=$1[@] du2=$2[@] dif='-' difpc
	[[ ! $cmd_run ]] && du2=$3[@]
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
				if [[ ! $cmd_run ]]; then 
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
	[[ ! $cmd_run && -z $cmd_verb ]] && return
	local app s1 score=()
	for app in "${appPathArray[@]}"; do
		s1=$( du -sh "$basePATH$app" )
		score+=("$s1")
	done
	eval $1='("${score[@]}")'
} # END MSO DU
function update-counter () {
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
		pre=".+/(en|en_GB" # english never deleted
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
		[[ ${search:0:1} == '|' ]] && search="${search:1}"
		[[ "$search" ]] &&
			printf '%s%s%s' ".+/(" $search ")\..+" ||
			echo ''
		return
	fi
	[[ $cmd_inverse ]] && pre=".+/("
	for i in $l; do search+="|$i"; done
	[[ $cmd_inverse && ${search:0:1} == '|' ]] && search="${search:1}"
	printf '%s%s%s' "$pre" "$search" "$sfx"
} # END search filter

##########  Common-purpose routines library
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
function print-row () { # simple table formatter
# Prints 1 or 2 columns of text. UNIX sucks: printf is not unicode-aware!
# $1 - column1 padding
# $2 - column2 padding
# $3 - column1 text
# $4 - column2 text
# $5 - divider char
	local pad1=$1 pad2=$2 s1 s2
	local str1="$3" str2="$4" div=${5:-' '} 
	local w=$(tput cols); let "w--"; [[ $w -gt 120 ]] && w=100
	local rightpad=4 gap=3 # predefined margins
	
	[[ $gap -lt 3 ]] && div=''
	[[ $pad2 -lt $pad1 ]] && pad2=$pad1
	[[ $pad1 -eq 0 ]] && ((pad2--))
	
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
} # END print-row
function joina { local oldifs="$IFS"; IFS="$1"; shift; echo "$*"; IFS="$oldifs"; }
function unique () { # rm dup words in string; 'echo' removes awk tail
	[[ "${1// }" == '' ]] && { echo; return; }
	echo $( awk 'BEGIN{RS=ORS=" "}!a[$0]++' <<<"$1 " )
}
function instr () { awk -v a="$2" -v b="$1" 'BEGIN{print index(a,b)}'; }
function inarray () { # test item ($1) in array ($2 - passing by name!)
    local s=$1 name=$2[@]; local a=("${!name}")
	comm -1 -2 -i <(printf '%s\n' "${s[@]}" | sort -u) <(printf '%s\n' "${a[@]}" | sort -u)
}
function normalize-number () { 
	echo "$1" "$2" | awk '{e1=$1; v2=$2} END {
		unit="K"; fmt1="%.f"; kilo=1024
		if (e1 > kilo) {e1/=1024; unit="M"; fmt1="%.f"}
		if (e1 > kilo) {e1/=1024; unit="G"; fmt1="%.1f"}
		fmt1 = (v2 == "ns") ? fmt1"%s" : fmt1" %s"
		printf fmt1,e1,unit
	}'
}
function math-diffexpr () { # simple unit calculator; couldnot find better
	local dif difpc exp='{
		e1=toupper($1); e2=toupper($2); child=$3}; END {
		u1=substr(e1,length(e1),1); u2=substr(e2,length(e2),1)
		x=match(u1,"[KMG]"); if (!x) u1="K" # type cast
		x=match(u2,"[KMG]"); if (!x) u2=u1; 
		kilo=1024; unit="K"
		m1=e1; sub(/[A-Z]/,"",m1); m2=e2; sub(/[A-Z]/,"",m2);
		m1 = 0 + m1; m2 = 0 + m2
		if (!m1 && !m2 || m1 == 0) {printf "%s|", "0"; exit 1}
#		if (m1 > kilo && u1 != "G") {m1/=1024; u1 = (u1 == "K") ? "M":"G"}
#		if (m2 > kilo && u2 != "G") {m2/=1024; u2 = (u2 == "K") ? "M":"G"}
		if (m1 < 1 && u1 != "K") {m1*=1024; u1 = (u1 == "G") ? "M":"K"}
		if (m2 < 1 && u2 != "K") {m2*=1024; u2 = (u2 == "G") ? "M":"K"}

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

############# Your show begins here
main "${@}"
