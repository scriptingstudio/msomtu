# MSOMTU  
***Msomtu*** is Microsoft Office 2016 for Mac maintenance utility.

## Description
Microsoft Office 2016 for Mac uses an isolated resource architecture (sandboxing), so apps duplicate all of the components in its own application container that's waisting gigabytes of disk space. This script safely removes (thinning) extra parts of the folowing components: UI languages; proofing tools; fontlist files (.plist); OS duplicated font files. It also can backup/copy font files to predefined or user defined destinations.

## Notes
* *Safe scripting* technique — "Foolproof" or "Harmless Run". Default running mode is view. You cannot change or harm your system without switch '-run'. Parameter '-cache' does not depend on '-run'.
* As MSO is installed with root on /Applications directory you have to run this script with *sudo* to make changes.
* As application font structure has been changed since MSO version 15.17 font deletion only works with 15.17 or later.
* File operations are case insensitive.
* If you remove fonts, remove font lists as well. 'DFonts' folder and font lists are safe to remove. Some of the fonts you may find useful, save them before deletion.
* **Caution**: do not remove fonts from 'Fonts' folder! These are minimum needed for MSO applications to work.
* Apply thinning after every MSO update.
* You can change default settings in code for your needs.

## Usage

~~~sh
$ [sudo] msomtu.sh [-<parameter> [<arguments>]]...

$ [sudo] msomtu.sh [-backup|-fcopy [<destination>]] [-app [<app>]] [-font [<font_pattern>]] [-ex|-x <font_pattern>] [-run]

$ [sudo] msomtu.sh [-app ["<app_list>"]] [-lang|-ui ["<lang_list>"]] [-proof|-p ["<proof_list>"]] [-font [<font_pattern>]] [-flist|-fl] [-ex|-x <font_pattern>] [-cache] [-report|-rep] [-verbose|-verb] [-fontset|-fs] [-all|-full] [-rev] [-help|-h|-?] [-run]
~~~

### Use Cases
<table>
<thead style="text-align:left;"><tr>
<th>Action</th> <th>Parameter</th>
</tr></thead>
<tr>
<td>Getting MSO info</td><td><code>-report</code></td>
<tr>
<td>Listing/Removing UI languages</td><td><code>-lang</code></td>
<tr>
<td>Listing/Removing proofingtools</td><td><code>-proof</code></td>
</tr>
<tr>
<td>Listing/Removing fonts</td><td><code>-font</code></td>
</tr>
<tr>
<td>Listing/Removing font list files</td><td><code>-flist</code></td>
</tr>
<tr>
<td>Listing fontsets</td><td><code>-fontset</code></td>
</tr>
<tr>
<td>Removing font cache</td><td><code>-cache</code></td>
</tr>
<tr>
<td>Finding new fonts</td><td><code>-font -rev</code></td>
</tr>
<tr>
<td>Backing up fonts</td><td><code>-backup</code></td>
</tr>
<tr>
<td>Copying fonts to font libraries</td><td><code>-backup syslib -font cyrdfonts</code></td>
</tr>
</table>

## Arguments
<table>
<tr><td><code>app_list</code></td><td>App list. w - Word, e - Excel, p - PowerPoint, o - Outlook, n - OneNote. Default: 'w e p o n'.</td></tr>
<tr><td><code>lang_list</code></td><td>Langauge list: ru pl de etc, see filenames with parameter '-verb'. Default: 'en ru'.</td></tr>
<tr><td><code>proof_list</code></td><td>Proofingtools list: russian finnish german etc, see filenames with parameter '-verb'. Wildcard '*' is available. Default: 'english russian'.</td></tr>
<tr><td><code>font_pattern</code></td><td>Font operations are based on patterns. Font patterns: empty - removes folder 'DFonts' (default); <i>fontset</i> - removes fonts of predefined fontset; <i>mask</i> - removes selection: <i>*.*, arial*, *.ttc</i> etc. If you use single '*' enclose it in quotation marks: "*". Predefined fontsets: <i>library, cyrdfonts, noncyr, chfonts, sysfonts, symfonts</i>. See parameter '-fontset' and details in code. Fontset 'library' removes duplicates of system and user libraries; it may not exactly match fonts because based on file-by-file (unlike font family) comparison (DFonts against libraries). You can use list of fontsets.</td></tr>
<tr><td><code>destination</code></td><td>Backup destination folderpath for fonts. Default value is '~/Desktop/MSOFonts'. You can use predefined destinations: <i>syslib</i> - system library; <i>userlib</i> - user library.</td></tr>
</table>
## Parameters
<table>
<tr><td><code>-app</code></td> <td>Filter <i>app_list</i>. Selects application to process.</td></tr>
<tr><td><code>-lang</code></td> <td>Exclusive filter <i>ang_list</i>. Removes UI languages except defaults and user list. See also parameter '-rev'; it reverses user selection except defaults.</td></tr>
<tr><td><code>-proof</code></td> <td>Exclusive filter <i>proof_list</i>. Removes proofing tools except defaults and user list. See also parameter '-rev'; it reverses user selection except defaults.</td></tr>
<tr><td><code>-font</code></td> <td>Filter <i>font_pattern</i>. Removes selected fonts or 'DFonts' folder. Available fontsets: <i>cyrdfonts, noncyr, chinese, sysfonts</i>. Parameter '-rev' ignores user selection and alternates search function: new fonts are going to be discovered. It is useful to check new fonts up after new update.</td></tr>
<tr><td><code>-backup</code></td> <td>Backs up fonts to user defined destination. If destination folder does not exist it will be created. You can use system and user libraries as destination, see ARGUMENTS. Backup alternates all deletions to backup.</td></tr>
<tr><td><code>-ex</code></td> <td>Exclusive filter <i>font_pattern</i>. Excludes font selection with parameter '-font'. Only mask can be used as <i>font_pattern</i>.</td></tr>
<tr><td><code>-flist</code></td> <td>Switch. Removes fontlist (.plist) files.</td></tr>
<tr><td><code>-all</code></td> <td>Switch. Activates all cleaning options: <i>lang, proof, font, flist, cache</i>. It does not affect a parameter '-app'.</td></tr>
<tr><td><code>-cache</code></td> <td>Switch. Cleans font cache.</td></tr>
<tr><td nowrap><code>-verbose</code></td> <td>Switch. Shows objects to be removed in view mode.</td></tr>
<tr><td><code>-report</code></td> <td>Switch. Shows statistics on objects.</td></tr>
<tr><td nowrap><code>-fontset</code></td> <td>Switch. Shows predefined fontsets.</td></tr>
<tr><td><code>-rev</code></td> <td>Switch. Reverses effect of 'lang' and 'proof' filters.</td></tr>
<tr><td><code>-help</code></td> <td>Switch. Shows help page. (Optional)</td></tr>
<tr><td><code>-run</code></td> <td>Switch. Default mode is view (test). Activates operations execution.</td></tr>
</table>
## Examples
Get app statistics:

~~~sh
$ msomtu.sh -report
~~~

Thin all apps with all parameters:

~~~sh
$ sudo msomtu.sh -all -run
~~~

Show app ('w e' for Word and Excel) language files installed:

~~~sh
$ msomtu.sh -app "w e" -lang -verbose
~~~

Remove a number of languages:

~~~sh
$ sudo msomtu.sh -lang "nl no de" -rev -run 
~~~

Remove all proofing tools except defaults for Word:

~~~sh
$ sudo msomtu.sh -proof -app w -run 
~~~

Remove a number of proofing tools:

~~~sh
$ sudo msomtu.sh -proof "Indonesian Isix*" -rev -run 
~~~

Show duplicates of library fonts for Word:

~~~sh
$ msomtu.sh -font lib -app w -verbose 
~~~

Remove duplicated fonts in libraries for Word:

~~~sh
$ sudo msomtu.sh -font lib -app w -run 
~~~

Remove 'chinese' and Arial fonts:

~~~sh
$ sudo msomtu.sh -font "chinese arial*" -run 
~~~

Show new fonts for Outlook:

~~~sh
$ msomtu.sh -font -rev -app o 
~~~

Exclude a few useful fonts from deletion for Word:

~~~sh
$ sudo msomtu.sh -font *.* -ex "brit* rockwell*" -app w 
~~~

Clean font cache:

~~~sh
$ sudo msomtu.sh -cache
~~~

Backup fonts to default destination:

~~~sh
$ msomtu.sh -backup -font "cyrdfonts britanic*" -run
~~~

Copy original cyrillic fonts to system library:

~~~sh
$ sudo msomtu.sh -backup syslib -font cyrdfonts -run 
~~~

Show predefined fontsets:

~~~sh
$ msomtu.sh -fontset
~~~

## Links
* Inspiration idea: [https://github.com/goodbest/OfficeThinner]()
* More on OS X & MSO fonts: [http://www.jklstudios.com/misc/osxfonts.html]()

