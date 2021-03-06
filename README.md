# MSOMTU
***Msomtu*** is Microsoft Office for Mac Maintenance Utility—the integrated fitness solution for your Office.

## Table of contents
* [The initial idea] (#the-initial-idea)
* [Description] (#description)
* [Requirements] (#requirements)
* [Features] (#features)
* [Notes] (#notes)
* [Usage] (#usage)
* [Arguments] (#arguments)
* [Parameters] (#parameters)
* [Examples] (#examples)
* [Customization] (#customization)
* [Links] (#links)

## The initial idea
**Font management** for Microsoft Office. MSO has completely isolated and partly duplicated fonts. Why? 

## Description
Microsoft Office 2016 for Mac uses an isolated resource architecture (sandboxing), so the MSO apps duplicate all of the components in their own application container that's waisting gigabytes of your disk space and making MSO unmanageable—it has stopped being integrated. This script safely removes (thins) **optional parts** of the following components: 

* UI languages.
* Proofing tools.
* OS X duplicated fonts.
* Fontlist files (font*.plist).

## Requirements
* Microsoft Office 2016 for Mac 15.17 or later to work with fonts.

## Features
* ***Safe Scripting*** technique—“Foolproof” or “Harmless Run”. The default running mode is view. You can think of it as **what-if** mode. The script cannot make changes or harm your system without parameter `-run`. It unlocks commands.
* **Analytical tools** and **proactive assessment**. You can explore the MSO app resources: fonts, languages of localization and proofing tools, disk size taken by the resources. You can evaluate your disk space taken by the MSO app components before thinning.
* **Duplicate fonts finder**. You can find out conflicting and extra app fonts against the font libraries.
* **New fonts finder**. You can find out new (standard sets considered) fonts added in each app after new MSO update.
* **Predefined fontsets** are custom classes of fonts for easy manipulation with fonts. Font classification specifics: cyrillic, non-cyrillic, hieroglyphic of any kind, symbolic, system. Fontsets do not intersect. You can modify fontsets for your needs.
* **Font backup**. You can backup your fonts to predefined and user defined destinations as well before deletion.
* Copy/move fonts to the font libraries.
* Flexible search filters: exclusion, inversion, macros, masks, lists.
* Multi-language help. Currently english and russian. Non-english help pages are in the separate module. You can extend help with your language by the template in the module.
* MSO new version checker.

## Notes
* As MSO is installed with root on the `/Applications` directory you will be asked for an administrative account's password to make changes.
* As application font structure has been changed since MSO version 15.17 font deletion only works with 15.17 or later. Microsoft separated font sets for some reasons. Essential fonts to the MSO apps are in the `Fonts` folder within each app. The rest are in the `DFonts` folder.
* If you remove fonts, remove font lists (font*.plist) as well; see PARAMETERS. The `DFonts` folder and font lists are safe to remove. No third party app can see MSO fonts installed to the `DFonts` and `Fonts` folders. Some of the fonts you may find useful, save them before deletion.
* **Caution**: do not remove fonts from the `Fonts` folder! These are minimum needed for the MSO applications to work.
* File operations are case insensitive.
* Apply thinning after every MSO update.
* Default settings for the `-lang` and `-proof` parameters: *english* and *russian*. It depends on your system locale and common sense: for MSO integrity it is better to leave english. You can change any default settings in code for your needs. Default languages are reserved from deletion.

## Usage

```sh
# Syntax schema
$ msomtu.sh [-<parameter> [<arguments>]]...

$ msomtu.sh -backup [<destination>] [-app [<app>]] [-font [<font_pattern>]] [-ex|-x <font_pattern>] [-run|-ok]

$ msomtu.sh [-app ["<app_list>"]] [-lang|-ui ["<lang_list>"]] [-proof|-p ["<proof_list>"]] [-font [<font_pattern>]] [-flist|-fl] [-ex|-x <font_pattern>] [-cache|-fc [u|user]] [-report|-rep|-info] [-verbose|-verb [nl]] [-fontset|-fs] [-all|-full] [-inv] [-help|-h|-? [en] [full]] [-run|-ok]
```

#### Use Cases
**Solo actions:** `-report`, `-fontset`, `-backup`, `-help`. <br/>If two or more actions specified the latter wins.

**Group actions:** `-font`, `-lang`, `-proof`, `-flist`, `-cache`, `-check`, `-all`.

| Action  | Parameter |
|:--------|:----------|
| Getting MSO info | `-report` |
| Getting proactive assessment <br/>of thinning (view mode) <br/><br/> | Fonts: `-font -verbose nl` <br/>UI langs: `-lang -verbose nl` <br/>Proofingtools: `-proof -verbose nl` <br/>Fontlists: `-flist -verbose nl` |
| Listing/Removing UI languages    | `-lang` |
| Listing/Removing proofingtools   | `-proof` |
| Listing/Removing fonts | `-font` |
| Listing/Removing font duplicates | `-font lib` |
| Listing/Removing font list files | `-flist` |
| Listing fontsets     | `-fontset` |
| Clearing font cache  | `-cache` |
| Finding new fonts    | `-font -inv` |
| Backing up fonts     | `-backup -font` |
| Copying fonts to font libraries | `-backup -font` |
| Checking for new versions       | `-check` |
| Getting help | Short page: no parameters <br/>Full page: `-help full` <br/>Force english page: `-help en` |

## Arguments
<table>
<tr><td valign="top"><code>app_list</code></td><td>Application list. <b>w</b>—Word, <b>e</b>—Excel, <b>p</b>—PowerPoint, <b>o</b>—Outlook, <b>n</b>—OneNote. Default: <code>w e p o n</code>.</td></tr>
<tr><td valign="top"><code>lang_list</code></td><td>Langauge list: <code>ru pl de</code> etc; see filenames with parameter <code>-verbose</code>. Default: <code>en ru</code>, see NOTES.</td></tr>
<tr><td valign="top"><code>proof_list</code></td><td>Proofingtools list: <code>russian finnish german</code> etc; see filenames with parameter <code>-verbose</code>. Wildcard <code>*</code> is available to use. Default: <code>english russian</code>, see NOTES.</td></tr>
<tr><td valign="top"><code>font_pattern</code></td><td>Font operations are based on patterns. Font patterns: empty—removes the <code>DFonts</code> folder (default); <i>fontset</i>—removes fonts of predefined fontset; <i>mask</i>—removes selection: <code>*.*</code>, <code>arial*</code>, <code>*.ttc</code> etc. If you use single <code>*</code> enclose it in quotation marks: <code>"*"</code>. Predefined fontsets: <code>library</code>, <code>cyrdfonts</code>, <code>noncyr</code>, <code>glyph</code>, <code>sysfonts</code>, <code>symfonts</code>. See parameter <code>-fontset</code> and details in code. Fontset <code>library</code> removes duplicates from system and user libraries; it may not exactly match fonts because based on file-by-file (unlike font family) comparison (<code>DFonts</code> against libraries). You can use list of fontsets as well.</td></tr>
<tr><td valign="top"><code>destination</code></td><td>Backup destination folderpath for fonts. Default value is <code>~/Desktop/MSOFonts</code>. You can use predefined destinations as well: <code>syslib</code>—system library; <code>userlib</code>—user library.</td></tr>
</table>

## Parameters
**Note:** In case of parameter duplicates the latter wins.

<table>
<tr><td valign="top"><code>-all</code></td> <td>Switch. Activates all cleaning options: <code>lang</code>, <code>proof</code>, <code>font</code>, <code>flist</code>, <code>cache</code>. It does not affect the parameter <code>-app</code>.</td></tr>
<tr><td valign="top"><code>-app</code></td> <td>Filter <code>app_list</code>. Selects application to process.</td></tr>
<tr><td valign="top"><code>-backup</code></td> <td>Backs up fonts to user defined destination. If destination folder does not exist it will be created. You can use system and user libraries as destination, see ARGUMENTS. Backup command alternates all deletions to backup.</td></tr>
<tr><td valign="top"><code>-cache</code></td> <td>Switch. Cleans up font cache (the system and the current user). The argument <code>user</code> indicates to clean cache for the current user only. It does not depend on <code>-run</code>.</td></tr>
<tr><td valign="top"><code>-check</code></td><td>Switch. Checks for new versions; opens the web-page in browser.</td></tr>
<tr><td valign="top"><code>-ex</code></td> <td>Exclusive filter <code>font_pattern</code>. Excludes font selection with parameter <code>-font</code>. Only mask can be used as <code>font_pattern</code>.</td></tr>
<tr><td valign="top"><code>-flist</code></td> <td>Switch. Removes fontlist (font*.plist) files. Fontlists are like cache. When you remove unneeded fonts you can also have to clear all non existent fonts from its lists. Since discovering fonts through all lists is difficult, remove all of the .plist files. They all have to do with the fixed font lists you see in Office.</td></tr>
<tr><td valign="top"><code>-font</code></td> <td>Filter <code>font_pattern</code>. Removes selected fonts or the <code>DFonts</code> folder. Available fontsets: <code>cyrdfonts</code>, <code>noncyr</code>, <code>glyph</code>, <code>sysfonts</code>. Parameter <code>-inv</code> ignores user selection and alternates search function: new fonts are going to be discovered. It is useful to check new fonts up after new update. Argument <code>library</code> alters searching in libraries for duplicates. </td></tr>
<tr><td nowrap valign="top"><code>-fontset</code></td> <td>Switch. Shows predefined fontsets.</td></tr>
<tr><td valign="top"><code>-help</code></td> <td>Switch. Shows the help page. Text language depends on your locale. There are two kinds of help page: short and full. The default is short one (no parameters). To get the full page use parameter <code>-help full</code>. Special argument <code>en</code> forces english help page.</td></tr>
<tr><td valign="top"><code>-inv</code></td> <td>Switch. Inverts effect of the <code>-lang</code> and <code>-proof</code> filters, but defaults are reserved. For parameter <code>-font</code> it is to search for the new fonts.</td></tr>
<tr><td valign="top"><code>-lang</code></td> <td>Exclusive filter <code>lang_list</code>. Removes UI languages except defaults and user list. See also parameter <code>-inv</code>; it inverts user selection except defaults.</td></tr>
<tr><td valign="top"><code>-proof</code></td> <td>Exclusive filter <code>proof_list</code>. Removes proofing tools except defaults and user list. See also parameter <code>-inv</code>; it inverts user selection except defaults.</td></tr>
<tr><td valign="top"><code>-report</code></td> <td>Switch. Shows statistics on objects.</td></tr>
<tr><td valign="top"><code>-run</code></td> <td>Switch. The default mode is view (test). Activates operations execution.</td></tr>
<tr><td nowrap valign="top"><code>-verbose</code></td> <td>Switch. Shows objects to be removed or searched, and assessment report. With special argument <code>nl</code> skips file listing. This parameter only works in view mode.</td></tr>
</table>

## Examples
Get app statistics:

```sh
$ msomtu.sh -report
```

Thin all apps with all parameters:

```sh
$ msomtu.sh -all -run
```

Show which app ('w e' for Word and Excel) language files installed:

```sh
$ msomtu.sh -app "w e" -lang -verbose
```

Remove a number of languages:

```sh
$ msomtu.sh -lang "nl no de" -inv -run 
```

Remove all proofing tools except defaults for Word:

```sh
$ msomtu.sh -proof -app w -run 
```

Remove a number of proofing tools:

```sh
$ msomtu.sh -proof "Indonesian Isix*" -inv -run 
```

Show duplicates of the library fonts for Word:

```sh
$ msomtu.sh -font lib -app w -verbose 
```

Remove duplicated fonts in the libraries for Word:

```sh
$ msomtu.sh -font lib -app w -run 
```

Remove `glyph` and Arial fonts:

```sh
$ msomtu.sh -font "glyph arial*" -run 
```

Show new fonts for Outlook:

```sh
$ msomtu.sh -font -inv -app o 
```

Exclude a few useful fonts from deletion for Word:

```sh
$ msomtu.sh -font *.* -ex "brit* rockwell*" -app w -run 
```

Clean font cache:

```sh
# system and user cache
$ msomtu.sh -cache
# current user cache
$ msomtu.sh -cache u
```

Backup fonts to default destination:

```sh
$ msomtu.sh -backup -font "cyrdfonts britanic*" -run
```

Copy original cyrillic fonts to system library:

```sh
$ msomtu.sh -backup syslib -font cyrdfonts -run 
```

Show predefined fontsets:

```sh
$ msomtu.sh -fontset
```

## Customization
You can easily modify the code for your needs:

* Languages to reserve from deletion.
* Font sets to work with fonts.
* Extend help with your language.

## Links
* Inspirational idea of “thinning”: [ @goodbest ](https://github.com/goodbest/OfficeThinner)
* More on OS X & MSO fonts: [Font Management in OS X, by Kurt Lang](http://www.jklstudios.com/misc/osxfonts.html)
