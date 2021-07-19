# ![logo](https://www.gstatic.com/images/branding/product/2x/apps_script_48dp.png)<br> Google Admin Chromebook  Device Management Script

<link href="https://fonts.googleapis.com/icon?family=Material+Icons"
      rel="stylesheet">
<link href="https://fonts.googleapis.com/icon?family=Material+Icons"
      rel="stylesheet">

<!-- 
https://wordpress.com/support/markdown-quick-reference/

https://marketplace.visualstudio.com/items?itemName=bierner.markdown-preview-github-styles

https://gist.github.com/rxaviers/7360908

https://stackoverflow.com/questions/58737436/how-to-create-a-good-looking-notification-or-warning-box-in-github-flavoured-mar -->

| :exclamation: USE AT OWN RISK :exclamation: |
|---------------------------------------------|


Table of contents
=================

<!--ts-->
   * [Installation](#installation)
      * [Clasp](#clasp)
      * [Copy](#copy)
   * [Usage](#usage)
   * [Resources](#resources)
     * [Andrew Stillman](#andrew-stillman)
     * [mhawksey](#mhawksey)
     * [Adam L](#adam-l)
   * [Troubleshooting](#troubleshooting)
     * [PowerShell](#powershell)
<!--te-->

Installation
============

Clasp
-----

1. Install Node.js from [here](https://nodejs.org/en/). The version must be at least 4.7.4.

2. Follow google developers directions to install clasp from Node.js [here](https://developers.google.com/apps-script/guides/clasp#requirements).

3. Verify you have Google Apps Script API enabled [here](https://script.google.com/home/usersettings).

4. Download latest release [here](https://github.com/myxxmikeyxx/Google-Chromebook-Admin-Script/releases/latest) and extract it.

5. Open terminal/cmd/power shell **inside project folder**. (In Visual Studio Code do Ctrl+Shift+`)

6. Do the following commands: 
   * Create package.json (Doesn't really matter what you make it, just is needed for clasp login step):
   <pre>
   npm init
   </pre>
   * Clasp login:
   <pre>
   clasp login
   </pre>
   * Create Google Sheet with script attached:
   <pre>
   clasp create cbManagement
   </pre>

7. Select **sheets** for where you want the script to install to.

8. Do the following commands: 
   <pre>
   clasp push
   clasp open
   </pre> This will push it to google script, then open the script.

9. Click the <span class="material-icons">info</span> in the upper left. 

10. Now click the sheet. This will take you to the sheet the script is attached to. From here rename it to whatever you want.


Copy
----
1. Download latest release [here](https://github.com/myxxmikeyxx/Google-Chromebook-Admin-Script/releases/latest) and extract it.

2. Open a new google sheets document and name it.

3. In the tool bar click ```Tools > Script Editor```.

4. In ```Apps Script``` rename the script at the top.

5. Copy and paste all of code.js into the code file, replacing everything.
6. In the Files location click the + and chose script.
7. Name one "columnRelation" & another "rowsData".
8. Copy the script into the respective ones.

Usage
=====

This script is for getting active chrome devices. Sorting them, adding and changing fields, and moving OU's. 


Resources
=========
I would like to thank everyone below for the great scripts that I was able to use and tweak to make this project happen.

[Andrew Stillman](https://www.linkedin.com/in/astillman)
---------------
Script & Spreadsheet [here](http://chromebookedu.blogspot.com/2014/02/a-new-script-for-chromebook-admins-via.html).
This script was what I used as a base to make mine. I had to change a lot of stuff and later just went in my own direction, but the main get devices and update devices parts are from his script with minor tweaks.

[mhawksey](https://github.com/mhawksey)
----------
The orginal script can be found [here](https://gist.github.com/mhawksey/51a1501493787bc5b7f1). I ended up tweaking it a bit to make it so the users and not an array of recent users. I also made it only show "ACTIVE" devices. (Google admin keeps deprovisoned device info for a while after they have been deprovisioned, and the get devices would show all devices without this change).

[Adam L.](https://stackoverflow.com/users/1373663/adaml)
-------
I used his script for letter to column & column to letter. The post I found it on is [here](https://stackoverflow.com/questions/21229180/convert-column-index-into-corresponding-column-letter). This script was not changed, plus it was very useful.


Troubleshooting
================

PowerShell
----------
If you are getting "Execution Policy" error in PowerShell try [here](https://tecadmin.net/powershell-running-scripts-is-disabled-system/) for help.


