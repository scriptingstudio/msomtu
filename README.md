# MSOMTU  
***Msomtu*** is Microsoft Office for Mac Maintenance Utility.

## Description
Microsoft Office 2016 for Mac uses an isolated resource architecture (sandboxing), so the MSO apps duplicate all of the components in their own application container waisting gigabytes of your disk space and making MSO unmanageable. This script safely removes (thinning) extra parts of the folowing components: 

* UI languages; 
* Proofing tools; 
* OS X duplicated font files; 
* Fontlist files.

It also can backup/copy font files to predefined and user defined destinations.

## Requirements
* Microsoft Office 2016 for Mac 15.17 or later.

## Features
* ***Safe scripting*** technique — "Foolproof" or "Harmless Run". The default running mode is view. You cannot change or harm your system without switch `-run`. Parameter `-cache` does not depend on `-run`.
* Proactive **assessment**. You can evaluate your disk space taken by the MSO app components before thinning.
* **Duplicate fonts finder**. You can find out conflicting and extra app fonts against the font libraries.
* **New fonts finder**. You can find out new (standard sets considered) fonts added in each app after new MSO update.
* **Predefined fontsets** are custom classes of fonts. Font classification specifics (in descending): cyrillic, non-cyrillic, hieroglyphic of any kind, symbolic, system. Fontsets do not intersect. You can modify fontsets for your needs.
* Backup. You can backup your fonts before deletion.
* Copy or move fonts to the libraries.
* Flexible parameter filters.
* Multilanguage help. Currently english and russian. Non-english help pages are in separate module. You can extend help with your language by the template in the module.

## Notes
* As MSO is installed with root on the `/Applications` directory you have to run this script with `sudo` to make changes.
* As application font structure has been changed since MSO version 15.17 font deletion only works with 15.17 or later. Microsoft separated font sets for some reasons. Essential fonts to the MSO apps are in the `Fonts` folder within each app. The rest are in the `DFonts` folder.
* If you remove fonts, remove font lists as well. The `DFonts` folder and font lists are safe to remove. No third party app can see MSO fonts installed to the `DFonts` folder. Some of the fonts you may find useful, save them before deletion.
* **Caution**: do not remove fonts from the `Fonts` folder! These are minimum needed for the MSO applications to work.
* File operations are case insensitive.
* Apply thinning after every MSO update.
* Default settings for the `-lang` and `-proof` parameters: *english* and *russian*. It depends on your system locale and common sense: for MSO integrity it is better to leave english. You can change any default settings in code for your needs.

## Usage

```sh
$ [sudo] msomtu.sh [-<parameter> [<arguments>]]...

$ [sudo] msomtu.sh [-backup|-fcopy [<destination>]] [-app [<app>]] [-font [<font_pattern>]] [-ex|-x <font_pattern>] [-run]

$ [sudo] msomtu.sh [-app ["<app_list>"]] [-lang|-ui ["<lang_list>"]] [-proof|-p ["<proof_list>"]] [-font [<font_pattern>]] [-flist|-fl] [-ex|-x <font_pattern>] [-cache] [-report|-rep] [-verbose|-verb [nl]] [-fontset|-fs] [-all|-full] [-rev] [-help|-h|-? [en]] [-run]
```

#### Use Cases
**Solo actions:** `-report`, `-fontset`, `-backup`, `-cache`, `-help`.

| Action  | Parameter |
|:--------|:----------|
| Getting MSO info | `-report` |
| Getting proactive assessment of thinning (view mode) <br/><br/> | Fonts: `-font -verbose nl` <br/>UI langs: `-lang -verbose nl` <br/>Proofingtools: `-proof -verbose nl` <br/>Fontlists: `-flist -verbose nl` |
| Listing/Removing UI languages | `-lang` |
| Listing/Removing proofingtools | `-proof` |
| Listing/Removing fonts | `-font` |
| Listing/Removing font list files | `-flist` |
| Listing fontsets | `-fontset` |
| Removing font cache | `-cache` |
| Finding new fonts | `-font -rev` |
| Backing up fonts | `-backup -font` |
| Copying fonts to font libraries | `-backup -font` |
| Getting help | Short page — no parameters <br/>Full page — `-help -full` <br/>Force english page — `-help en` |

## Arguments
<table>
<tr><td valign="top"><code>app_list</code></td><td>Application list. <b>w</b> — Word, <b>e</b> — Excel, <b>p</b> — PowerPoint, <b>o</b> — Outlook, <b>n</b> — OneNote. Default: <code>w e p o n</code>.</td></tr>
<tr><td valign="top"><code>lang_list</code></td><td>Langauge list: <code>ru pl de</code> etc; see filenames with parameter <code>-verbose</code>. Default: <code>en ru</code>, see NOTES.</td></tr>
<tr><td valign="top"><code>proof_list</code></td><td>Proofingtools list: <code>russian finnish german</code> etc; see filenames with parameter <code>-verbose</code>. Wildcard <code>*</code> is available to use. Default: <code>english russian</code>, see NOTES.</td></tr>
<tr><td valign="top"><code>font_pattern</code></td><td>Font operations are based on patterns. Font patterns: empty — removes the <code>DFonts</code> folder (default); <i>fontset</i> — removes fonts of predefined fontset; <i>mask</i> — removes selection: <i>*.*, arial*, *.ttc</i> etc. If you use single <code>*</code> enclose it in quotation marks: <code>"*"</code>. Predefined fontsets: <code>library</code>, <code>cyrdfonts</code>, <code>noncyr</code>, <code>chinese</code>, <code>sysfonts</code>, <code>symfonts</code>. See parameter <code>-fontset</code> and details in code. Fontset <code>library</code> removes duplicates from system and user libraries; it may not exactly match fonts because based on file-by-file (unlike font family) comparison (<code>DFonts</code> against libraries). You can use list of fontsets as well.</td></tr>
<tr><td valign="top"><code>destination</code></td><td>Backup destination folderpath for fonts. Default value is <code>~/Desktop/MSOFonts</code>. You can use predefined destinations as well: <code>syslib</code> — system library; <code>userlib</code> — user library.</td></tr>
</table>

## Parameters
<table>
<tr><td valign="top"><code>-app</code></td> <td>Filter <code>app_list</code>. Selects application to process.</td></tr>
<tr><td valign="top"><code>-lang</code></td> <td>Exclusive filter <code>lang_list</code>. Removes UI languages except defaults and user list. See also parameter <code>-rev</code>; it reverses user selection except defaults.</td></tr>
<tr><td valign="top"><code>-proof</code></td> <td>Exclusive filter <code>proof_list</code>. Removes proofing tools except defaults and user list. See also parameter <code>-rev</code>; it reverses user selection except defaults.</td></tr>
<tr><td valign="top"><code>-font</code></td> <td>Filter <code>font_pattern</code>. Removes selected fonts or the <code>DFonts</code> folder. Available fontsets: <code>cyrdfonts</code>, <code>noncyr</code>, <code>chinese</code>, <code>sysfonts</code>. Parameter <code>-rev</code> ignores user selection and alternates search function: new fonts are going to be discovered. It is useful to check new fonts up after new update.</td></tr>
<tr><td valign="top"><code>-backup</code></td> <td>Backs up fonts to user defined destination. If destination folder does not exist it will be created. You can use system and user libraries as destination, see ARGUMENTS. Backup alternates all deletions to backup.</td></tr>
<tr><td valign="top"><code>-ex</code></td> <td>Exclusive filter <code>font_pattern</code>. Excludes font selection with parameter <code>-font</code>. Only mask can be used as <i>font_pattern</i>.</td></tr>
<tr><td valign="top"><code>-flist</code></td> <td>Switch. Removes fontlist (.plist) files.</td></tr>
<tr><td valign="top"><code>-all</code></td> <td>Switch. Activates all cleaning options: <code>lang</code>, <code>proof</code>, <code>font</code>, <code>flist</code>, <code>cache</code>. It does not affect the parameter <code>-app</code>.</td></tr>
<tr><td valign="top"><code>-cache</code></td> <td>Switch. Cleans up font cache.</td></tr>
<tr><td nowrap valign="top"><code>-verbose</code></td> <td>Switch. Shows objects to be removed, in view mode. With special argument <code>nl</code> skips file listing.</td></tr>
<tr><td valign="top"><code>-report</code></td> <td>Switch. Shows statistics on objects.</td></tr>
<tr><td nowrap valign="top"><code>-fontset</code></td> <td>Switch. Shows predefined fontsets.</td></tr>
<tr><td valign="top"><code>-rev</code></td> <td>Switch. Reverses effect of the <code>-lang</code> and <code>-proof</code> filters. For parameter <code>-font</code> it is to search for the new fonts.</td></tr>
<tr><td valign="top"><code>-run</code></td> <td>Switch. The default mode is view (test). Activates operations execution.</td></tr>
<tr><td valign="top"><code>-help</code></td> <td>Switch. Shows the help page. There are two kinds of help page: short and full. The default is short one (no paramaters). To get the full page use parameters <code>-help -full</code>. Special argument <code>en</code> forces english help page.</td></tr>
</table>

## Examples
Get app statistics:

```sh
$ msomtu.sh -report
```

Thin all apps with all parameters:

```sh
$ sudo msomtu.sh -all -run
```

Show which app ('w e' for Word and Excel) language files installed:

```sh
$ msomtu.sh -app "w e" -lang -verbose
```

Remove a number of languages:

```sh
$ sudo msomtu.sh -lang "nl no de" -rev -run 
```

Remove all proofing tools except defaults for Word:

```sh
$ sudo msomtu.sh -proof -app w -run 
```

Remove a number of proofing tools:

```sh
$ sudo msomtu.sh -proof "Indonesian Isix*" -rev -run 
```

Show duplicates of the library fonts for Word:

```sh
$ msomtu.sh -font lib -app w -verbose 
```

Remove duplicated fonts in the libraries for Word:

```sh
$ sudo msomtu.sh -font lib -app w -run 
```

Remove `chinese` and Arial fonts:

```sh
$ sudo msomtu.sh -font "chinese arial*" -run 
```

Show new fonts for Outlook:

```sh
$ msomtu.sh -font -rev -app o 
```

Exclude a few useful fonts from deletion for Word:

```sh
$ sudo msomtu.sh -font *.* -ex "brit* rockwell*" -app w -run 
```

Clean font cache:

```sh
$ sudo msomtu.sh -cache
```

Backup fonts to default destination:

```sh
$ msomtu.sh -backup -font "cyrdfonts britanic*" -run
```

Copy original cyrillic fonts to system library:

```sh
$ sudo msomtu.sh -backup syslib -font cyrdfonts -run 
```

Show predefined fontsets:

```sh
$ msomtu.sh -fontset
```

## Links
* Inspiration idea of "thinning": [OfficeThinner](https://github.com/goodbest/OfficeThinner)
* More on OS X & MSO fonts: [Font Management in OS X, by Kurt Lang](http://www.jklstudios.com/misc/osxfonts.html)
