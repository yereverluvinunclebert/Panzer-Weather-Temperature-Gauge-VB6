# Panzer-Weather-Gauges-VB6

A FOSS Weather Gauges VB6 WoW64 Widget for Reactos, XP, Win7, 8 and 10/11+.

My current VB6/RC6 PSD program being finished now, 12% complete, awaiting a lot of work to get the thing functioning, generatging sunrise/sunset times, four additional gauges, two sub-forms, addition of smooth rotation to the Weather pointers, testing on laptops, testing on Windows XP and Win7 32bit and some multi-monitor checking, completion of the CHM help file and the creation of the setup.exe. This Panzer widget is based upon the Yahoo widget of the same visual design and very similar operation.

Why VB6? Well, with a 64 bit, modern-language improvement upgrade on the way with 100% compatible TwinBasic coupled with support for transparent PNGs via RC/Cairo, VB6 code has an amazing future.

![vb6-logo-350](https://github.com/yereverluvinunclebert/Panzer-RAM-Gauge-VB6/assets/2788342/2f60380d-29f5-4737-8392-e7d747c61f25)

I created as a variation of the previous gauges I had previously created for the World of Tanks and War Thunder
communities. The Panzer Weather Gauge widget is an attractive dieselpunk VB6 widget for your desktop.
Functional and gorgeous at the same time. The graphics are my own, I took original inspiration from a clock face by Italo Fortana combining it with an aircraft gauge surround. It is all my code with some help from the chaps at VBForums (credits given).

![panzerWeather650](https://github.com/yereverluvinunclebert/Panzer-Weather-Widget/assets/2788342/73c299e9-b8f6-422d-9a95-5f70a16183e3)

The Panzer Weather Gauge VB6 is a useful utility displaying the Weather your chosen locality in a dieselpunk fashion on your desktop. This Widget is a moveable widget that you can move anywhere around the desktop as you require. The Weather data is extracted via an XML request to aviation.gov. The program extracts the temperature abnd barometric data from that XML and displays it via analogue pointers on several desktop gauges.

Hope the code is useful to anyone else building system metric utilities using VB6/VBS/VBA.

![weather-icon-01](https://github.com/yereverluvinunclebert/Panzer-Weather-Widget/assets/2788342/ff953574-718b-47d1-84af-b425771a7db1)

![background](https://github.com/yereverluvinunclebert/Panzer-Weather-Widget/assets/2788342/07b0c7b4-a4e9-4b6c-89d3-fde55b0b735b)

This widget can be increased in size, animation speed can be changed,
opacity/transparency may be set as to the users discretion. The widget can
also be made to hide for a pre-determined period.

![panzer-temperature-icon](https://github.com/yereverluvinunclebert/Panzer-Weather-Widget/assets/2788342/f4b3b246-c895-4458-ab29-57a5f84e0d26)

Right clicking will bring up a menu of options. Double-clicking on the widget will cause a personalised Windows application to
fire up. The first time you run it there will be no assigned function and so it
will state as such and then pop up the preferences so that you can enter the
command of your choice. The widget takes command line-style commands for
windows. Mouse hover over the widget and press CTRL+mousewheel up/down to resize. It works well on Windows XP
to Windows 11.

![panzer-weather-help](https://github.com/yereverluvinunclebert/Panzer-Weather-Widget/assets/2788342/23582667-4a0d-4719-b6a0-b1a1407ccf7f)

The Panzer Battery Gauge VB6 gauge is Beta-grade software, under development, not yet
ready to use on a production system - use at your own risk.

This version was developed on Windows 7 using 32 bit VisualBasic 6 as a FOSS
project creating a WoW64 widget for the desktop.

![Licence002](https://github.com/yereverluvinunclebert/Panzer-RAM-Gauge-VB6/assets/2788342/09dd88fd-0bff-4115-8fda-9b3e6b6852f5)

It is open source to allow easy configuration, bug-fixing, enhancement and
community contribution towards free-and-useful VB6 utilities that can be created
by anyone. The first step was the creation of this template program to form the
basis for the conversion of other desktop utilities or widgets. A future step
is new VB6 widgets with more functionality and then hopefully, conversion of
each to RADBasic/TwinBasic for future-proofing and 64bit-ness.

![menu01](https://github.com/yereverluvinunclebert/Panzer-RAM-Gauge-VB6/assets/2788342/ee727437-e6e4-4b91-8c0d-90e7e43352b4)

This utility is one of a set of steampunk and dieselpunk widgets. That you can
find here on Deviantart: https://www.deviantart.com/yereverluvinuncleber/gallery

I do hope you enjoy using this utility and others. Your own software
enhancements and contributions will be gratefully received if you choose to
contribute.

![panzer-weather-gauges](https://github.com/yereverluvinunclebert/Panzer-Weather-Widget/assets/2788342/9d1fb5ee-0e4e-467a-a337-36fd2fa9bc64)

BUILD: The program runs without any Microsoft plugins.



Built using: VB6, MZ-TOOLS 3.0, VBAdvance, CodeHelp Core IDE Extender
Framework 2.2 & Rubberduck 2.4.1, RichClient 6

Links:

    https://www.vbrichclient.com/#/en/About/
    MZ-TOOLS https://www.mztools.com/
    CodeHelp http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=62468&lngWId=1
    Rubberduck http://rubberduckvba.com/
    Rocketdock https://punklabs.com/
    Registry code ALLAPI.COM
    La Volpe http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=67466&lngWId=1
    PrivateExtractIcons code http://www.activevb.de/rubriken/
    Persistent debug code http://www.vbforums.com/member.php?234143-Elroy
    Open File common dialog code without dependent OCX - http://forums.codeguru.com/member.php?92278-rxbagain
    VBAdvance

Tested on :

    ReactOS 0.4.14 32bit on virtualBox
    Windows 7 Professional 32bit on Intel
    Windows 7 Ultimate 64bit on Intel
    Windows 7 Professional 64bit on Intel
    Windows XP SP3 32bit on Intel
    Windows 10 Home 64bit on Intel
    Windows 10 Home 64bit on AMD
    Windows 11 64bit on Intel

CREDITS:

I have really tried to maintain the credits as the project has progressed. If I
have made a mistake and left someone out then do forgive me. I will make amends
if anyone points out my mistake in leaving someone out.

MicroSoft in the 90s - MS built good, lean and useful tools in the late 90s and
early 2000s. Thanks for VB6.

Olaf Schmidt - This tool was built using the RichClient RC5 Cairo wrapper for
VB6. Specifically the components using transparency and reading images directly
from PSD. Thanks for the massive effort Olaf in creating Cairo counterparts for
all VB6 native controls and giving us access to advanced features on controls
such as transparency.

Shuja Ali @ codeguru for his settings.ini code.

ALLAPI.COM For the registry reading code.

Rxbagain on codeguru for his Open File common dialog code without a dependent
OCX - http://forums.codeguru.com/member.php?92278-rxbagain

si_the_geek for his special folder code

Elroy on VB forums for the balloon tooltips

Harry Whitfield for his quality testing, brain stimulation and being an
unwitting source of inspiration.

![panzer-clipboard-help](https://github.com/yereverluvinunclebert/Panzer-Weather-Widget/assets/2788342/91ef78ab-c7c6-4836-a9b1-d01e78014116)

Dependencies:

o A windows-alike o/s such as Windows XP, 7-11 or Apple Mac OSX 11.

o Microsoft VB6 IDE installed with its runtime components. The program runs
without any additional Microsoft OCX components, just the basic controls that
ship with VB6.

![vb6-logo-350](https://github.com/yereverluvinunclebert/Panzer-RAM-Gauge-VB6/assets/2788342/2479af5a-82bf-42ae-bdb1-28c22160f93c)

- Uses the latest version of the RC6 Cairo framework from Olaf Schmidt.

During development the RC6 components need to be registered. These scripts are
used to register. Run each by double-clicking on them.

    RegisterRC6inPlace.vbs
    RegisterRC6WidgetsInPlace.vbs

During runtime on the users system, the RC6 components are dynamically
referenced using modRC6regfree.bas which is compiled into the binary.

Requires a PzRAM Gauge folder in C:\Users\<user>\AppData\Roaming\
eg: C:\Users\<user>\AppData\Roaming\PzRAM Gauge
Requires a settings.ini file to exist in C:\Users\<user>\AppData\Roaming\PzRAM Gauge
The above will be created automatically by the compiled program when run for the
first time.

Uses just one OCX control extracted from Krools mega pack (slider). This is part
of Krools replacement for the whole of Microsoft Windows Common Controls found
in mscomctl.ocx. The slider control OCX file is shipped with this package.

- CCRSlider.ocx

This OCX will reside in the program folder. The program reference to this OCX is
contained within the supplied resource file Panzer RAM Gauge Gauge.RES. It is
compiled into the binary.

- OLEGuids.tlb

This is a type library that defines types, object interfaces, and more specific
API definitions needed for COM interop / marshalling. It is only used at design
time (IDE). This is a Krool-modified version of the original .tlb from the
vbaccelerator website. The .tlb is compiled into the executable.
For the compiled .exe this is NOT a dependency, only during design time.

From the command line, copy the tlb to a central location (system32 or wow64
folder) and register it.

COPY OLEGUIDS.TLB %SystemRoot%\System32\
 REGTLIB %SystemRoot%\System32\OLEGUIDS.TLB

In the VB6 IDE - project - references - browse - select the OLEGuids.tlb

![prefs-about](https://github.com/yereverluvinunclebert/Panzer-Battery-Gauge-VB6/assets/2788342/302e59dd-c767-4b38-b748-b3119f3a6d15)

- SETUP.EXE - The program is currently distributed using setup2go, a very useful
  and comprehensive installer program that builds a .exe installer. Youll have to
  find a copy of setup2go on the web as it is now abandonware. Contact me
  directly for a copy. The file "install PzRAM Gauge 0.1.0.s2g" is the configuration
  file for setup2go. When you build it will report any errors in the build.

- HELP.CHM - the program documentation is built using the NVU HTML editor and
  compiled using the Microsoft supplied CHM builder tools (HTMLHelp Workshop) and
  the HTM2CHM tool from Yaroslav Kirillov. Both are abandonware but still do
  the job admirably. The HTML files exist alongside the compiled CHM file in the
  HELP folder.

Project References:

    VisualBasic for Applications
    VisualBasic Runtime Objects and Procedures
    VisualBasic Objects and Procedures
    OLE Automation
    vbRichClient6 - RC6Widgets (RC6Widgets.DLL) - RC6 (RC6.DLL)


![weather](https://github.com/yereverluvinunclebert/Panzer-Weather-Widget/assets/2788342/4ab9945b-f460-43c9-b2d2-92a0643c50d0)    

LICENCE AGREEMENTS:

Copyright © 2023 Dean Beedell

In addition to the GNU General Public Licence please be aware that you may use
any of my own imagery in your own creations but commercially only with my
permission. In all other non-commercial cases I require a credit to the
original artist using my name or one of my pseudonyms and a link to my site.
With regard to the commercial use of incorporated images, permission and a
licence would need to be obtained from the original owner and creator, ie. me.


![about](https://github.com/yereverluvinunclebert/Panzer-Weather-Gauges-VB6/assets/2788342/d36355ff-0289-4145-b9bf-641aad58041e)

![Panzer-CPU-Gauge-onDesktop](https://github.com/yereverluvinunclebert/Panzer-RAM-Gauge-VB6/assets/2788342/6dc97b14-9954-4f8c-a775-5657b2aeec85)
