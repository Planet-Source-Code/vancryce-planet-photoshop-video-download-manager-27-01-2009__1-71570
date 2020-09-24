Planet Photoshop Video Download Manager
---------------------------------------


27.01.2009
----------
I abandoned the WizMedia WebBrowser Control.
There were Runtime Errors associate to webbrowser that i could't solve.
I change to Microsoft Internet Controls instead.
The rest is basic the same, with little changes of the design.



OLD INFORMATION
----------------
Some Explanation
----------------
I like VB, but i like Design and Photoshop also.
Planet Photoshop (www.planet photoshop) have good videos and tutorials.
Planet Source Code have lots information about vb.
I combine both here.

I like to catch the tutorial videos from Planet Photoshop to see off-line, but it's not an easy thing to do.
First i must open the tutorial, read the source code of the HTML page, find the *.flv file and use this link to catch the video: www.planetphotoshop.com/videos/....flv.
There's a lot of videos there.
I thought in a way to cath those videos faster. It's here that enter Visual Basic 6 and the Planet Source Code.
I use controls and source code from PSC users, because i'm not a programmer.
I just like to do things my way and have some fun with photoshop and vb.
Sorry for using all your work, and Thanks!!


File Usage
----------

MODULES:
- AmodParsers.bas (to show frmSplash and frmAbout);
- basBrowse (to open Browse files when clicking botBrowse);
- basClipboard (to manage and Hook Clipboard);
- basGeral (mix functions);
- basIcon (used to make your application icon with better quality in Windows Systems (Alt+Tab, Systray, Taskbar, etc..);
- basItemExist (function to determine if file or folder exist);
- basProgPath (a substitute of the App.Path);
- basResource (used to manage the PPVDM.res);

CLASSES:
- Ac32bppDIB.cls, AcGDIPlus.cls, AcGIFparser.cls, COparser.cls, AcPNGparser.cls, AcPNGwriter.cls (to show frmSplash and frmAbout);
- VicsDL.cls (to use progressbar during download files)

REFERENCES:
- Olelib.tlb (to use progressbar during download files)

UPX
---
-Upx.exe, Upx.bat (use this tool compress the *.exe file - very good)

-----------------------------------------------------------------------------------------------------------------

My Skills in VB are very bad.
- I couldn't find a way to load the frmSplash and frmAbout image from the Resource file. I must use the image frmSplash.png to that;
- I didn't found a way to show controls on frmSplash and frmAbout forms when using LaVolpe code. It's an amazing effect for VB, but only to show images;

Please, implement this and other things as you like and show me the result.