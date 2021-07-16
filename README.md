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
   * [Dependency](#dependency)
   * [Docker](#docker)
     * [Local](#local)
<!--te-->

Installation
============

Clasp
-----

1. Install Node.js

2. Follow google developers directions to install clasp from Node.js [here](https://developers.google.com/apps-script/guides/clasp#requirements).

3. Follow clasp login directions.

4. Verify you have Google Apps Script API enabled [here](https://script.google.com/home/usersettings).

5. Download latest release [here](/releases/latest) and extract it.

6. Open terminal/cmd/power shell in project folder.

7. Do the following command: 
   <pre>
   clasp create cbManagement
   </pre>

8. Select sheets for where you want the script to install to.

9. Do the following command: 
   <pre>
   clasp push
   clasp open
   </pre> This will open the script.

10. Click the <span class="material-icons">info</span> in the upper left. 

11. Now click the sheet. This will take you to the sheet the script is attached to. From here rename it to whatever you want.


Copy
----
1. Download latest release [here](/myxxmikeyxx/Google-Chromebook-Admin-Script/releases/latest/download/) and extract it.

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