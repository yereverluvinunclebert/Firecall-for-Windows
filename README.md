# Firecall for Windows

Firecall for Windows, written in VB6. A WoW64 utility for Reactos, 
Windows XP, Win7, 8 and 10+.

Fire Call Win is a simple chat tool that allows two users, typically developers, 
to communicate person-to-person using encrypted transfer via  shared files and 
folders. Files can be transferred and exchanged, chat and data swapped on almost 
real-time basis. This version uses the basic Dropbox service to share files and 
folders but it can also use any shared file or folder resource over a network of 
some sort as long as it appears as a locally accessible resource to Windows. In 
addition to chat, Emojis, voice recordings and user files can be shared too.

![FCWWider01](https://github.com/yereverluvinunclebert/Firecall-for-Windows/assets/2788342/fa09d58f-04bb-4a82-bbfb-4a30928c47b9)

This program is free and FOSS. The basic file and folder sharing services are 
free, this program uses Dropbox out of the box but Google Drive and Microsofts 
Onedrive all provide basic file and folder sharing facilities for free.

Firecall for Windows is Beta-grade software, under development, not yet 
ready to use on a production system - use at your own risk.

This version was developed on Windows 7 using 32 bit VisualBasic 6 as a FOSS 
project creating a WoW64 program for the desktop. 

It is open source to allow easy configuration, bug-fixing, enhancement and 
community contribution towards free-and-useful VB6 utilities that can be created
by anyone. The first step was the creation of this template program to form the 
basis for the conversion of other desktop utilities or widgets. A future step 
is conversion to RADBasic/TwinBasic for future-proofing and 64bit-ness. 

This utility is one of a set of steampunk and dieselpunk desktop widgets. That 
you can find here on Deviantart: https://www.deviantart.com/yereverluvinuncleber/gallery

![document-unknown](https://github.com/yereverluvinunclebert/Firecall-for-Windows/assets/2788342/178e5248-ea23-454e-a1be-bb2ba8b9f7a1)

I do hope you enjoy using this utility and others. Your own software 
enhancements and contributions will be gratefully received if you choose to 
contribute.
 
 Credits :   
 
	LA Volpe (VB Forums) for his transparent picture handling.  
	Shuja Ali (codeguru.com) for his settings.ini code.  
	Registry reading code from ALLAPI.COM.  
	Rxbagain on codeguru for his Open File common dialog code without dependent OCX.  
	Krool on the VBForums for his impressive common control replacements, slider and textboxW.  
	si_the_geek for his special folder code.  
	theTrick for his sound recording and saving to a WAV file.  
	Elroy for his kind help with subclassing and balloon tooltips and all his other kindness.  
	Wqweto for his innovative email injection work and help  .

	Thats all as far as I know. There may be others but it is not my intention to hide their efforts.

 Built using: VB6, MZ-TOOLS 3.0, CodeHelp Core IDE Extender Framework 2.2 & Rubberduck 2.4.1
 
	MZ-TOOLS https://www.mztools.com/  
	vBAdvance  
	CodeHelp http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=62468&lngWId=1  
	Rubberduck http://rubberduckvba.com/  
	Registry code ALLAPI.COM  
	La Volpe superb VB6 non-native image types  http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=67466&lngWId=1  
	Open File common dialog code without dependent OCX - http://forums.codeguru.com/member.php?92278-rxbagain  
	Open font dialog code without dependent OCX - unknown URL  
	Krools superb replacement Controls http://www.vbforums.com/showthread.php?698563-CommonControls-%28Replacement-of-the-MS-common-controls%29  
	Chris Fannin (AbbydonKrafts) Copying a folder  http://vbcity.com/forums/t/129391.aspx  
	Austin Hickl fnGetDateInUniversalFormat  http://computer-programming-forum.com/66-vb-controls/6dff1bae05df0a6e.htm  
	Ellis Dee VB6 quicksort https://www.vbforums.com/showthread.php?473677-VB6-Sorting-algorithms-(sort-array-sorting-arrays)  
	KayJay fnIsGoodURL that utilises the isValidURL API  https://www.vbforums.com/showthread.php?231061-Validate-URL&p=1361958&viewfull=1#post1361958  
	JCIs resize image https://www.vbforums.com/member.php?40893-jcis  
	qvb6 vb6 date from epoch  https://www.vbforums.com/member.php?291519-qvb6  
	Elroys superb balloon Tooltips.  
	Wqwetos superb TLS/STARTTLS code to enable email from VB6 using STARTTLS.  
	theTricks superb sound code allowing recording of high quality sound.  
	Keith Lacelle for the alternative FSO code to read a value from an INI file when GetPrivateProfileString fails.  
	 https://gist.github.com/Grimthorr/d17810f34cd361769ed0  
	Olaf Schmidt and his Date to Epoch code  

 Tested on :
 
	ReactOS 0.4.14 32bit on virtualBox  
	Windows 7 Professional 32bit on Intel  
	Windows 7 Ultimate 64bit on Intel  
	Windows 7 Professional 64bit on Intel  
	Windows XP SP3 32bit on Intel  
	Windows 10 Home 64bit on Intel  
	Windows 10 Home 64bit on AMD  

 Dependencies:
 
Krools replacement for the Microsoft Windows Common Controls found in mscomctl.ocx (slider) is replicated
by the addition of one dedicated OCX file that is shipped with this package - CCRSlider.ocx

This OCX will reside in the program folder. The program reference to this OCX is 
contained within the supplied resource file Panzer Earth Gauge.RES.
It is compiled into the binary. 

	requires a FireCallWin folder in C:\Users\<user>\AppData\Roaming\ eg: C:\Users\<user>\AppData\Roaming\FireCallWin  
	requires a Recordings folder in C:\Users\<user>\AppData\Roaming\FireCallWin eg: C:\Users\<user>\AppData\Roaming\FireCallWin\Recordings  
	requires a settings.ini file to exist in C:\Users\<user>\AppData\Roaming\FireCallWin  
	requires CCRSlider.ocx to exist in the program folder  
	requires an archive folder in app.path  
	requires a backup folder in app.path  
	
Project References:  

	VisualBasic for Applications  
	VisualBasic Runtime Objects and Procedures  
	VisualBasic Objects and Procedures  
	OLE Automation - drag and drop  
	Microsoft ActiveX Data Objects 2.8 Library msador28.tlb as shipped with Windows XP +
	Microsoft CDO for windows 2000 library component cdosys.dll

LICENCE AGREEMENTS:

Copyright 2023 Dean Beedell

In addition to the GNU General Public Licence please be aware that you may use
any of my own imagery in your own creations but commercially only with my
permission. In all other non-commercial cases I require a credit to the
original artist using my name or one of my pseudonyms and a link to my site.
With regard to the commercial use of incorporated images, permission and a
licence would need to be obtained from the original owner and creator, ie. me.

Program Notes:

The VB6 non native images (PNGs &c) are displayed using Lavolpes transparent DIB image code,
except for the .ico files which use his earlier StdPictureEx class.
Lavolpes later ico code caused many strange visual artifacts and complete failures to show .ico files.
especially when other image types were displayed on screen simultaneously.

The sound is recorded using theTricks sound code. It previously used MCISendString to record but Cortana on Win10+
hijacks the sound device so it does not work on those oses.

We have two comboboxes to store the audio input devices. The main combobox on the main form is used on form
startup, reason this is done this way is because the enumeration must be done on form_load for the recording
button to operate in HQ mode. Although we normally store the config. data in the prefs form, if we read that
construct on startup it will try to load the whole prefs form and the prefs program variables are not ready
for that to occur. Basically, we cannot have the combobox on another form and instead we keep the two in synch.

![firecallDesktopPrefs01](https://github.com/yereverluvinunclebert/Firecall-for-Windows/assets/2788342/c6039ce2-efd2-438c-ae0c-5b4994b96e94)

The email is achieved using a tool from Microsoft called CDO, Firecall uses this to make the email point-to-point
connection. Microsoft have failed to update CDO for a while so STARTTLS is not supported by default. In order to
make STARTTLS function we have a proxy on port 10025 that takes any STARTTLS connection and manually injects the
STARTTLS command into the stream just at the right time, for correct negotiation of a secure connection. The proxy
forwards on the connection to the users chosen port. This is the only way to make CDO negotiate a STARTTLS
connection.
