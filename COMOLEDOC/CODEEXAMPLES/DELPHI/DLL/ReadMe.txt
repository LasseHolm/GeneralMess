When Automation starts up, it looks in its install dir (the same dir as 
the file PCsELcad.exe is).

This directory, is where your DLL-file should be placed.
Also place a MIF-file here. The MIF-file tells Automation
about the DLL-filename, and what text to display.

Edit the Mif-file so it points to the DLL, and correct text as needed.


Help for the MIF-file:
[Modul] 
	The text in "Title=" is the text displayed in "Files->Modules"
	
	The text in "LC_Title=" will overrule the previous caption if languagecode is found
	(LC = Language code. Ex. DK=Denmark, RU=Russian, etc...)


[PLUGIN]
	The text in "MenuX=" tells where the call should be availeble.
	EX = "Menu1=Tools" will make it availeble through the tools-menu.
	
	The caption to the menu will be fetch From "NameX=".
	The Text from "LC_NameX=" will overrule the previous caption if languagecode is found
	(LC = Language code. Ex. DK=Denmark, RU=Russian, etc...)
	 
	The Text from "CMDX=" is the command to run. Here is where you DLL-filename should be placed.
	
	The text from "IndexX=" is at what location in the selected menu, the called should be placed.


Notice:
A call can be made availeble to call from more than one specific location in Automation.

Ex.
[PLUGIN]
Menu1=Tools
Name1=1TestDll
DK_Name1=2 Test af DLL
RU_Name1=3 Test of DLL
CN_Name1=4 Test of DLL
SE_Name1=5 Test of DLL
CZ_Name1=6 Test of DLL
Cmd1=TestDLL.dll,0,0,0
Index1=40

Menu2=Popup
Name2=TestDlla
DK_Name2=Test af DLL b
RU_Name2=Test of DLL c
CN_Name2=Test of DLL d
SE_Name2=Test of DLL e
CZ_Name2=Test of DLL f
Cmd2=TestDLL.dll,0,0,0
Index2=10

Makes the call availeble from the Toolsmenu and a pop-up-menu.
This will be descriped more in one of the next versions of this documentation.
Please notice, that for each place a call should be made availeble, the number
behind the parametertext, should increase.
So for the first items in [PLUGIN] the values are 1.
For the next items in [PLUGIN] the values are 2.
etc... (if more items)
