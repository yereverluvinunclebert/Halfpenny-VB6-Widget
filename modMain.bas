Attribute VB_Name = "modMain"
'---------------------------------------------------------------------------------------
' Module    : modMain
' Author    : beededea
' Date      : 27/04/2023
' Purpose   : Main module containing sub main
'---------------------------------------------------------------------------------------

Option Explicit

'------------------------------------------------------ STARTS
' for SetWindowPos z-ordering
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const HWND_TOP As Long = 0 ' for SetWindowPos z-ordering
Private Const HWND_TOPMOST As Long = -1
Private Const HWND_BOTTOM As Long = 1
Private Const SWP_NOMOVE  As Long = &H2
Private Const SWP_NOSIZE  As Long = &H1
Private Const OnTopFlags  As Long = SWP_NOMOVE Or SWP_NOSIZE
'------------------------------------------------------ ENDS


Public fMain As New cfMain
Public revealWidgetTimerCount As Integer

Public thingummyWidget As cwThingummy
Public aboutWidget As cwAbout




'--------------------------------------------------------------------------------------------------------------
' BUILD:
'
'
' Credits : Standing on the shoulders of the following giants:
'
'           Olaf Schmidt for his Rich Client 5 Cairo wrapper.
'           Shuja Ali (codeguru.com) for his settings.ini code.
'           Registry reading code from ALLAPI.COM.
'           Rxbagain on codeguru for his Open File common dialog code without dependent OCX
'           Krool on the VBForums for his impressive common control replacements
'           si_the_geek for his special folder code
'           Elroy for the balloon tooltips
'           Rod Stephens vb-helper.com for ResizeControls
'
' Tools:    Built on a 3.3ghz Dell Latitude E6410 running Windows 7 Ultimate 64bit using VB6 SP6, VbAdvance, MZ-TOOLS 3.0,
'           CodeHelp Core IDE Extender Framework 2.2 & Rubberduck 2.4.1
'
'           MZ-TOOLS https://www.mztools.com/
'           CodeHelp http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=62468&lngWId=1
'           Rubberduck http://rubberduckvba.com/
'           Registry code ALLAPI.COM
'           VbAdvance
'           La Volpe  http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=67466&lngWId=1
'           Open File common dialog code without dependent OCX - http://forums.codeguru.com/member.php?92278-rxbagain
'           Open font dialog code without dependent OCX - unknown URL
'           Krool's replacement Controls http://www.vbforums.com/showthread.php?698563-CommonControls-%28Replacement-of-the-MS-common-controls%29
'
'   Tested on :
'           ReactOS 0.4.14 32bit on virtualBox
'           Windows 7 gblPlofessional 32bit on Intel - wip
'           Windows 7 Ultimate 64bit on Intel
'           Windows 7 gblPlofessional 64bit on Intel - done
'           Windows XP SP3 32bit on Intel - wip
'           Windows 10 Home 64bit on Intel - done
'           Windows 10 Home 64bit on AMD
'           Windows 11 64bit on Intel - done
'
' Dependencies:
'
'           Requires a thingummy folder in C:\Users\<user>\AppData\Roaming\ eg: C:\Users\<user>\AppData\Roaming\halfpenny
'           Requires a settings.ini file to exist in C:\Users\<user>\AppData\Roaming\halfpenny
'           The above will be created automatically by the compiled program when run for the first time.
'
'           Uses just one OCX control extracted from Krool's mega pack (slider). This is part of Krool's replacement for the
'           whole of Microsoft Windows Common Controls found in mscomctl.ocx. The slider control OCX file is shipped with
'           this package.
'
'           * CCRSlider.ocx
'
'           This OCX will reside in the program folder. The program reference to this OCX is contained within the
'           supplied resource file halfpenny.RES. It is compiled into the binary.
'
'           * OLEGuids.tlb
'
'           This is a type library that defines types, object interfaces, and more specific API definitions
'           needed for COM interop / marshalling. It is only used at design time (IDE). This is a Krool-modified
'           version of the original .tlb from the vbaccelerator website. The .tlb is compiled into the executable.
'           For the compiled .exe this is NOT a dependency, only during design time.
'
'           From the command line, copy the tlb to a central location (system32 or wow64 folder) and register it.
'
'           COPY OLEGUIDS.TLB %SystemRoot%\System32\
'           REGTLIB %SystemRoot%\System32\OLEGUIDS.TLB
'
'           In the VB6 IDE - project - references - browse - select the OLEGuids.tlb
'
'           * Uses the RC5 Cairo wrapper from Olaf Schmidt.
'
'           During development the RC5 components need to be registered. These scripts are used to register. Run each by
'           double-clicking on them.
'
'           RegisterRC5inPlace.vbs
'           RegisterVBWidgetsInPlace.vbs
'
'           During runtime on the user's system, the RC5 components are dynamically referenced using modRC5regfree.bas
'           which is compiled into the binary.
'
'           * SETUP.EXE - The program is currently distributed using setup2go, a very useful and comprehensive installer program
'           that builds a .exe installer. You'll have to find a copy of setup2go on the web as it is now abandonware. Contact me
'           directly for a copy. The file "install halfpenny 0.1.0.s2g" is the configuration file for setup2go. When you build it will
'           report any errors in the build.
'
'           * HELP.CHM - the program documentation is built using the NVU HTML editor and compiled using the Microsoft supplied
'           CHM builder tools (HTMLHelp Workshop) and the HTM2CHM tool from Yaroslav Kirillov. Both are abandonware but still do
'           the job admirably. The HTML files exist alongside the compiled CHM file in the HELP folder.
'
' gblPloject References:
'
'           VisualBasic for Applications
'           VisualBasic Runtime Objects and Procedures
'           VisualBasic Objects and Procedures
'           OLE Automation
'           vbRichClient5
'
' Summary:
'
'           The program is quite simple but forms the structure for other similar programs yet to come. These
'           will be funcxtional replicas of the graphical Steam/Dieselpunk javascript widgets I previously built using the
'           Yahoo widget engine.
'
'           This program is a mix of native VB6 forms and controls and 3rd party additions.
'           The superb RC5 Cairo wrapper from Olaf Schmidt is used in a very limited manner.
'           RC5's transparency capability is used for the main halfpenny and the about image only. I haven't used
'           Olaf's other Cairo controls to build forms as I need a graphical IDE to operate. Only testing RC5
'           at the moment, there should be no problems upgrading to RC6.
'
'           The helpForm is a standard VB6 window without a titlebar nor controls, displaying a fullsize image.
'           Standard VB6 forms are used for the preference and licence windows.
'           The standard VB6 timers are located on an invisible standard VB6 form - frmTimer.
'           The standard VB6 menu is located on it's own invisible VB6 form - menuForm.
'           The frmTimer invisible form is also the container for the large 128x128px overall project icon.

'           The utility itself has some configuration details that it stores in the settings.ini file
'           within the user appdata roaming folder.
'
'           The font selection and file/folder dialogs are generated using Win32 APIs rather than the
'           common dialog OCX which dispensed with another MS OCX.
'
'           As stated above, I have used Krool's amazing control replacement project. The specific code for
'           just one of the controls (slider) has been incorporated rather than all 32 of
'           Krool's complete package.
'
'           The preference form is resizable - which allows it to run on high DPI systems. In my mind, this is a poor
'           man's method of handling high DPI, 4K screens. I find the creation of DPI aware VB6 programs with a working
'           side by side configuration using manifests difficult with VB6 on modern systems. Instead the controls are
'           resized dynamically when the form is dragged. The images are reloaded with higher res. versions on 1500 twip
'           intervals. I could have done with GDI+ using multiple embedded icons but it is OK for the moment. All in all,
'           it's a bit sh1t but it works well enough, so it'll do...
'
'           There is one useful .BAT file - unhide.bat which will reveal a halfpenny mistakenly 'hidden' for an
'           extended period of time from the right click menu. This will allow you to open the prefs and unset the hidden
'           configuration option.
'
'
'    LICENCE AGREEMENTS:
'
'    Copyright � 2023 Dean Beedell
'
'    Using this program implies you have accepted the licence. The GPL licence applies to the code
'    this software Is provided 'as-is', without any express or implied warranty. In no event will the
'    author be held liable for any damages arising from the use of this software. Permission is granted to
'    anyone to use this software for any purpose, including commercial applications, and to alter it and
'    redistribute it freely, subject to the following restrictions:
'
'    1. The origin of this software must not be misrepresented; you must not claim that you wrote the original software. If you use this software in a product, an acknowledgment in the product documentation is required.
'    2. Altered source versions must be plainly marked as such, and must not be misrepresented as being the original software.
'    3. This notice may not be removed or altered from any source distribution.
'
'    This program is free software; you can redistribute it and/or modify it under the terms of the
'    GNU General Public Licence as published by the Free Software Foundation; either version 2 of the
'    License, or (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without
'    even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
'    General Public Licence for more details.
'
'    You should have received a copy of the GNU General Public Licence along with this program; if not,
'    write to the Free Software Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301
'    USA
'
'    If you use this software in any way whatsoever then that implies acceptance of the licence. If you
'    do not wish to comply with the licence terms then please remove the download, binary and source code
'    from your systems immediately.
'
'    FYI - I like CALLing subroutines, it may be old-fashioned but its what I do.
'
'--------------------------------------------------------------------------------------------------------------




'---------------------------------------------------------------------------------------
' Procedure : Main
' Author    : beededea
' Date      : 27/04/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Main()
    On Error GoTo Main_Error
    
    Call mainRoutine(False)

   On Error GoTo 0
   Exit Sub

Main_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Main of Module modMain"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : main_routine
' Author    : beededea
' Date      : 27/06/2023
' Purpose   : called by sub main() to allow this routine to be called elsewhere,
'             a reload for example.
'---------------------------------------------------------------------------------------
'
Public Sub mainRoutine(ByVal restart As Boolean)
    Dim extractCommand As String: extractCommand = vbNullString

   On Error GoTo main_routine_Error

    extractCommand = Command$ ' capture any parameter passed
    
    ' initialise global vars
    Call initialiseGlobalVars
    
    'add Resources to the global ImageList
    Call addImagesToImageList
    
    ' check the Windows version
    classicThemeCapable = fTestClassicThemeCapable
    
    ' get this tool's entry in the trinkets settings file and assign the app.path
    Call getTrinketsFile
    
    ' get the location of this tool's settings file (appdata)
    Call getToolSettingsFile
    
    ' read the dock settings from the new configuration file
    Call readSettingsFile(softwareThingummy, gblPlSettingsFile)
        
    ' check first usage and display licence screen
    Call checkLicenceState

    ' initialise and create the main forms on the current display
    Call createFormOnCurrentDisplay
    
    ' place the form at the saved location
    Call makeVisibleFormElements
    
    ' resolve VB6 sizing width bug
    Call determineScreenDimensions
    
    ' run the functions that are also called at reload time.
    Call adjustMainControls ' this needs to be here after the initialisation of the Cairo forms and widgets
    
    ' check the selected monitor properties to determine form placement
    'Call monitorProperties(frmHidden) - might use RC5 for this?
    
    ' move/hide onto/from the main screen
    Call mainScreen
        
    ' if the program is run in unhide mode, write the settings and exit
    Call handleUnhideMode(extractCommand)
    
    ' if a first time run shows prefs
    If gblPlFirstTimeRun = "true" Then     'parse the command line
        Call makeProgramPreferencesAvailable
    End If
    
    ' check for first time running
    Call checkFirstTime

    ' configure any global timers here
    Call configureTimers
    
    'load the preferences form but don't yet show it, speeds up access to the prefs via the menu
    Load halfpennyPrefs

    ' RC message pump will auto-exit when Cairo Forms = 0 so we run it only when 0, this prevents message interruption
    ' when running twice on reload.
    If Cairo.WidgetForms.Count = 0 Then Cairo.WidgetForms.EnterMessageLoop

   On Error GoTo 0
   Exit Sub

main_routine_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure main_routine of Module modMain"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : checkFirstTime
' Author    : beededea
' Date      : 12/05/2023
' Purpose   : check for first time running
'---------------------------------------------------------------------------------------
'
Private Sub checkFirstTime()

   On Error GoTo checkFirstTime_Error

    If gblPlFirstTimeRun = "true" Then
        gblPlFirstTimeRun = "false"
        sPutINISetting softwareThingummy, "firstTimeRun", gblPlFirstTimeRun, gblPlSettingsFile
    End If

   On Error GoTo 0
   Exit Sub

checkFirstTime_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure checkFirstTime of Module modMain"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : initialiseGlobalVars
' Author    : beededea
' Date      : 12/05/2023
' Purpose   : initialise global vars
'---------------------------------------------------------------------------------------
'
Private Sub initialiseGlobalVars()
      
    On Error GoTo initialiseGlobalVars_Error

    thisThingummy = "\Halfpenny"
    softwareThingummy = "Software" & thisThingummy

    ' general
    gblPlStartup = vbNullString
    gblPlGaugeFunctions = vbNullString
    'PrStampSelection = vbNullString
    
    ' config
    gblPlEnableTooltips = vbNullString
    gblPlEnableBalloonTooltips = vbNullString
    gblPlShowTaskbar = vbNullString
    
    gblPlGaugeSize = vbNullString
    gblPlScrollWheelDirection = vbNullString
    
    ' position
    gblPlAspectHidden = vbNullString
    gblPlWidgetPosition = vbNullString
    gblPlWidgetLandscape = vbNullString
    gblPlWidgetPortrait = vbNullString
    gblPlLandscapeFormHoffset = vbNullString
    gblPlLandscapeFormVoffset = vbNullString
    gblPlPortraitHoffset = vbNullString
    gblPlPortraitYoffset = vbNullString
    gblPlvLocationPercPrefValue = vbNullString
    gblPlhLocationPercPrefValue = vbNullString
    
    ' sounds
    gblPlEnableSounds = vbNullString
    
    ' development
    gblPlDebug = vbNullString
    gblPlDblClickCommand = vbNullString
    gblPlOpenFile = vbNullString
    gblPlDefaultEditor = vbNullString
         
    ' font
    gblPlPrefsFont = vbNullString
    gblPlPrefsFontSize = vbNullString
    gblPlPrefsFontItalics = vbNullString
    gblPlPrefsFontColour = vbNullString
    
    ' window
    gblPlWindowLevel = vbNullString
    gblPlPreventDragging = vbNullString
    gblPlOpacity = vbNullString
    gblPlWidgetHidden = vbNullString
    gblPlHidingTime = vbNullString
    gblPlIgnoreMouse = vbNullString
    gblPlFirstTimeRun = vbNullString
    
    ' general storage variables  initialised

    gblTrinketsDir = vbNullString
    gblTrinketsFile = vbNullString

    gblPlSettingsDir = vbNullString
    gblPlSettingsFile = vbNullString
    gblPlMaximiseFormX = vbNullString
    gblPlMaximiseFormY = vbNullString
    gblPlLastSelectedTab = vbNullString
    gblPlSkinTheme = vbNullString
    
    ' general variables declared
    toolSettingsFile = vbNullString
    classicThemeCapable = False
    storeThemeColour = 0
    windowsVer = vbNullString
    
    ' vars to obtain correct screen width (to correct VB6 bug) STARTS
    screenTwipsPerPixelX = 0
    screenTwipsPerPixelY = 0
    screenWidthTwips = 0
    screenHeightTwips = 0
    screenHeightPixels = 0
    screenWidthPixels = 0
    oldScreenHeightPixels = 0
    oldScreenWidthPixels = 0
    
    ' key presses
    CTRL_1 = False
    SHIFT_1 = False
    
    ' other globals
    debugflg = 0
    minutesToHide = 0
    aspectRatio = vbNullString
    revealWidgetTimerCount = 0
    oldPrSettingsModificationTime = #1/1/2000 12:00:00 PM#

   On Error GoTo 0
   Exit Sub

initialiseGlobalVars_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure initialiseGlobalVars of Module modMain"
    
End Sub

        
'---------------------------------------------------------------------------------------
' Procedure : addImagesToImageList
' Author    : beededea
' Date      : 27/04/2023
' Purpose   : add Resources to the global ImageList
'---------------------------------------------------------------------------------------
'
Private Sub addImagesToImageList()
    Dim useloop As Integer: useloop = 0
    
    On Error GoTo addImagesToImageList_Error

    Cairo.ImageList.AddImage "about", App.Path & "\Resources\images\about.png"
    
    'add Resources to the global ImageList
    Cairo.ImageList.AddImage "halfpenny", App.Path & "\Resources\images\Halfpenny.png"

    ' prefs icons

    Cairo.ImageList.AddImage "about-icon-dark", App.Path & "\Resources\images\about-icon-dark-1010.jpg"
    Cairo.ImageList.AddImage "about-icon-light", App.Path & "\Resources\images\about-icon-light-1010.jpg"
    Cairo.ImageList.AddImage "config-icon-dark", App.Path & "\Resources\images\config-icon-dark-1010.jpg"
    Cairo.ImageList.AddImage "config-icon-light", App.Path & "\Resources\images\config-icon-light-1010.jpg"
    Cairo.ImageList.AddImage "development-icon-light", App.Path & "\Resources\images\development-icon-light-1010.jpg"
    Cairo.ImageList.AddImage "development-icon-dark", App.Path & "\Resources\images\development-icon-dark-1010.jpg"
    Cairo.ImageList.AddImage "general-icon-dark", App.Path & "\Resources\images\general-icon-dark-1010.jpg"
    Cairo.ImageList.AddImage "general-icon-light", App.Path & "\Resources\images\general-icon-light-1010.jpg"
    Cairo.ImageList.AddImage "sounds-icon-light", App.Path & "\Resources\images\sounds-icon-light-1010.jpg"
    Cairo.ImageList.AddImage "sounds-icon-dark", App.Path & "\Resources\images\sounds-icon-dark-1010.jpg"
    Cairo.ImageList.AddImage "windows-icon-light", App.Path & "\Resources\images\windows-icon-light-1010.jpg"
    Cairo.ImageList.AddImage "windows-icon-dark", App.Path & "\Resources\images\windows-icon-dark-1010.jpg"
    Cairo.ImageList.AddImage "font-icon-dark", App.Path & "\Resources\images\font-icon-dark-1010.jpg"
    Cairo.ImageList.AddImage "font-icon-light", App.Path & "\Resources\images\font-icon-light-1010.jpg"
    Cairo.ImageList.AddImage "position-icon-light", App.Path & "\Resources\images\position-icon-light-1010.jpg"
    Cairo.ImageList.AddImage "position-icon-dark", App.Path & "\Resources\images\position-icon-dark-1010.jpg"

    Cairo.ImageList.AddImage "general-icon-dark-clicked", App.Path & "\Resources\images\general-icon-dark-600-clicked.jpg"
    Cairo.ImageList.AddImage "config-icon-dark-clicked", App.Path & "\Resources\images\config-icon-dark-600-clicked.jpg"
    Cairo.ImageList.AddImage "font-icon-dark-clicked", App.Path & "\Resources\images\font-icon-dark-600-clicked.jpg"
    Cairo.ImageList.AddImage "sounds-icon-dark-clicked", App.Path & "\Resources\images\sounds-icon-dark-600-clicked.jpg"
    Cairo.ImageList.AddImage "position-icon-dark-clicked", App.Path & "\Resources\images\position-icon-dark-600-clicked.jpg"
    Cairo.ImageList.AddImage "development-icon-dark-clicked", App.Path & "\Resources\images\development-icon-dark-600-clicked.jpg"
    Cairo.ImageList.AddImage "windows-icon-dark-clicked", App.Path & "\Resources\images\windows-icon-dark-600-clicked.jpg"
    Cairo.ImageList.AddImage "about-icon-dark-clicked", App.Path & "\Resources\images\about-icon-dark-600-clicked.jpg"

    Cairo.ImageList.AddImage "general-icon-light-clicked", App.Path & "\Resources\images\general-icon-light-600-clicked.jpg"
    Cairo.ImageList.AddImage "config-icon-light-clicked", App.Path & "\Resources\images\config-icon-light-600-clicked.jpg"
    Cairo.ImageList.AddImage "font-icon-light-clicked", App.Path & "\Resources\images\font-icon-light-600-clicked.jpg"
    Cairo.ImageList.AddImage "sounds-icon-light-clicked", App.Path & "\Resources\images\sounds-icon-light-600-clicked.jpg"
    Cairo.ImageList.AddImage "position-icon-light-clicked", App.Path & "\Resources\images\position-icon-light-600-clicked.jpg"
    Cairo.ImageList.AddImage "development-icon-light-clicked", App.Path & "\Resources\images\development-icon-light-600-clicked.jpg"
    Cairo.ImageList.AddImage "windows-icon-light-clicked", App.Path & "\Resources\images\windows-icon-light-600-clicked.jpg"
    Cairo.ImageList.AddImage "about-icon-light-clicked", App.Path & "\Resources\images\about-icon-light-600-clicked.jpg"
   
   On Error GoTo 0
   Exit Sub

addImagesToImageList_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure addImagesToImageList of Module modMain"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : adjustMainControls
' Author    : beededea
' Date      : 27/04/2023
' Purpose   : called at runtime and on restart, sets the characteristics of the halfpenny and menus
'---------------------------------------------------------------------------------------
'
Public Sub adjustMainControls()
   
   On Error GoTo adjustMainControls_Error

    ' validate the inputs of any data from the input settings file
    Call validateInputs
    
    thingummyWidget.Zoom = Val(gblPlGaugeSize) / 100
    thingummyWidget.ZoomDirection = gblPlScrollWheelDirection
    
    If thingummyWidget.Hidden = False Then
        thingummyWidget.opacity = Val(gblPlOpacity) / 100
        thingummyWidget.Widget.Refresh
    End If
    
    If gblPlGaugeFunctions = "1" Then
        menuForm.mnuSwitchOff.Checked = False
        menuForm.mnuTurnFunctionsOn.Checked = True
    Else
        menuForm.mnuSwitchOff.Checked = True
        menuForm.mnuTurnFunctionsOn.Checked = False
    End If
    
    If gblPlDefaultEditor <> vbNullString And gblPlDebug = "1" Then
        menuForm.mnuEditWidget.Caption = "Edit Widget using " & gblPlDefaultEditor
        menuForm.mnuEditWidget.Visible = True
    Else
        menuForm.mnuEditWidget.Visible = False
    End If
        
    If gblPlPreventDragging = "0" Then
        menuForm.mnuLockWidget.Checked = False
        thingummyWidget.Locked = False
    Else
        menuForm.mnuLockWidget.Checked = True
        thingummyWidget.Locked = True ' this is just here for continuity's sake, it is also set at the time the control is selected
    End If
    
    If gblPlShowTaskbar = "0" Then
        fMain.thingummyForm.ShowInTaskbar = False
    Else
        fMain.thingummyForm.ShowInTaskbar = True
    End If
                 
    ' set the z-ordering of the window
    Call setWindowZordering
    
    ' set the tooltips on the main screen
    Call setMainTooltips
    
    ' set the hiding time for the hiding timer, can't read the minutes from comboxbox as the prefs isn't yet open
    Call setHidingTime
    
    'Call setVisibleStampType

    If minutesToHide > 0 Then menuForm.mnuHideWidget.Caption = "Hide Widget for " & minutesToHide & " min."

   On Error GoTo 0
   Exit Sub

adjustMainControls_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure adjustMainControls of Module modMain"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : setWindowZordering
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : set the z-ordering of the window
'---------------------------------------------------------------------------------------
'
Public Sub setWindowZordering()

   On Error GoTo setWindowZordering_Error

    If Val(gblPlWindowLevel) = 0 Then
        Call SetWindowPos(fMain.thingummyForm.hwnd, HWND_BOTTOM, 0&, 0&, 0&, 0&, OnTopFlags)
    ElseIf Val(gblPlWindowLevel) = 1 Then
        Call SetWindowPos(fMain.thingummyForm.hwnd, HWND_TOP, 0&, 0&, 0&, 0&, OnTopFlags)
    ElseIf Val(gblPlWindowLevel) = 2 Then
        Call SetWindowPos(fMain.thingummyForm.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, OnTopFlags)
    End If

   On Error GoTo 0
   Exit Sub

setWindowZordering_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setWindowZordering of Module modMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : readSettingsFile
' Author    : beededea
' Date      : 12/05/2020
' Purpose   : read the application's setting file and assign values to public vars
'---------------------------------------------------------------------------------------
'
Public Sub readSettingsFile(ByVal location As String, ByVal gblPlSettingsFile As String)
    On Error GoTo readSettingsFile_Error

    If fFExists(gblPlSettingsFile) Then
        
        ' general
        gblPlStartup = fGetINISetting(location, "startup", gblPlSettingsFile)
        gblPlGaugeFunctions = fGetINISetting(location, "gaugeFunctions", gblPlSettingsFile)
        gblHalfpennySelection = fGetINISetting(location, "halfpennySelection", gblPlSettingsFile)
        
        ' configuration
        gblPlEnableTooltips = fGetINISetting(location, "enableTooltips", gblPlSettingsFile)
        gblPlEnableBalloonTooltips = fGetINISetting(location, "enableBalloonTooltips", gblPlSettingsFile)
        gblPlShowTaskbar = fGetINISetting(location, "showTaskbar", gblPlSettingsFile)
        
        gblPlGaugeSize = fGetINISetting(location, "gaugeSize", gblPlSettingsFile)
        gblPlScrollWheelDirection = fGetINISetting(location, "scrollWheelDirection", gblPlSettingsFile)
        'gblPlWidgetSkew = fGetINISetting(location, "widgetSkew", gblPlSettingsFile)
        
        ' position
        gblPlAspectHidden = fGetINISetting(location, "aspectHidden", gblPlSettingsFile)
        gblPlWidgetPosition = fGetINISetting(location, "widgetPosition", gblPlSettingsFile)
        gblPlWidgetLandscape = fGetINISetting(location, "widgetLandscape", gblPlSettingsFile)
        gblPlWidgetPortrait = fGetINISetting(location, "widgetPortrait", gblPlSettingsFile)
        gblPlLandscapeFormHoffset = fGetINISetting(location, "landscapeHoffset", gblPlSettingsFile)
        gblPlLandscapeFormVoffset = fGetINISetting(location, "landscapeYoffset", gblPlSettingsFile)
        gblPlPortraitHoffset = fGetINISetting(location, "portraitHoffset", gblPlSettingsFile)
        gblPlPortraitYoffset = fGetINISetting(location, "portraitYoffset", gblPlSettingsFile)
        gblPlvLocationPercPrefValue = fGetINISetting(location, "vLocationPercPrefValue", gblPlSettingsFile)
        gblPlhLocationPercPrefValue = fGetINISetting(location, "hLocationPercPrefValue", gblPlSettingsFile)

        ' font
        gblPlPrefsFont = fGetINISetting(location, "prefsFont", gblPlSettingsFile)
        gblPlPrefsFontSize = fGetINISetting(location, "prefsFontSize", gblPlSettingsFile)
        gblPlPrefsFontItalics = fGetINISetting(location, "prefsFontItalics", gblPlSettingsFile)
        gblPlPrefsFontColour = fGetINISetting(location, "prefsFontColour", gblPlSettingsFile)
        
        ' sound
        gblPlEnableSounds = fGetINISetting(location, "enableSounds", gblPlSettingsFile)
        
        ' development
        gblPlDebug = fGetINISetting(location, "debug", gblPlSettingsFile)
        gblPlDblClickCommand = fGetINISetting(location, "dblClickCommand", gblPlSettingsFile)
        gblPlOpenFile = fGetINISetting(location, "openFile", gblPlSettingsFile)
        gblPlDefaultEditor = fGetINISetting(location, "defaultEditor", gblPlSettingsFile)
        
        ' other
        gblPlMaximiseFormX = fGetINISetting(softwareThingummy, "maximiseFormX", gblPlSettingsFile)
        gblPlMaximiseFormY = fGetINISetting(softwareThingummy, "maximiseFormY", gblPlSettingsFile)
        gblPlLastSelectedTab = fGetINISetting(location, "lastSelectedTab", gblPlSettingsFile)
        gblPlSkinTheme = fGetINISetting(location, "skinTheme", gblPlSettingsFile)
        
        ' window
        gblPlWindowLevel = fGetINISetting(location, "windowLevel", gblPlSettingsFile)
        gblPlPreventDragging = fGetINISetting(location, "preventDragging", gblPlSettingsFile)
        gblPlOpacity = fGetINISetting(location, "opacity", gblPlSettingsFile)
        
        ' we do not want the widget to hide at startup
        'gblPlWidgetHidden = fGetINISetting(location, "widgetHidden", gblPlSettingsFile)
        gblPlWidgetHidden = "0"
        
        gblPlHidingTime = fGetINISetting(location, "hidingTime", gblPlSettingsFile)
        gblPlIgnoreMouse = fGetINISetting(location, "ignoreMouse", gblPlSettingsFile)
         
        gblPlFirstTimeRun = fGetINISetting(location, "firstTimeRun", gblPlSettingsFile)
        
    End If

   On Error GoTo 0
   Exit Sub

readSettingsFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readSettingsFile of Module common2"

End Sub


    
'---------------------------------------------------------------------------------------
' Procedure : validateInputs
' Author    : beededea
' Date      : 17/06/2020
' Purpose   : validate the relevant entries from the settings.ini file in user appdata
'---------------------------------------------------------------------------------------
'
Public Sub validateInputs()
    
   On Error GoTo validateInputs_Error
            
        ' general
        If gblPlGaugeFunctions = vbNullString Then gblPlGaugeFunctions = "1" ' always operate
        If gblHalfpennySelection = vbNullString Then gblHalfpennySelection = "0"
        If gblPlStartup = vbNullString Then gblPlStartup = "1"
        
        ' Config
        If gblPlEnableTooltips = vbNullString Then gblPlEnableTooltips = "1"
        If gblPlEnableBalloonTooltips = vbNullString Then gblPlEnableBalloonTooltips = "1"
        If gblPlShowTaskbar = vbNullString Then gblPlShowTaskbar = "0"
        
        If gblPlGaugeSize = vbNullString Then gblPlGaugeSize = "25"
        If gblPlScrollWheelDirection = vbNullString Then gblPlScrollWheelDirection = "1"
               
        ' fonts
        If gblPlPrefsFont = vbNullString Then gblPlPrefsFont = "times new roman" 'prefsFont", gblPlSettingsFile)
        If gblPlPrefsFontSize = vbNullString Then gblPlPrefsFontSize = "8" 'prefsFontSize", gblPlSettingsFile)
        If gblPlPrefsFontItalics = vbNullString Then gblPlPrefsFontItalics = "false"
        If gblPlPrefsFontColour = vbNullString Then gblPlPrefsFontColour = "0"

        ' sounds
        If gblPlEnableSounds = vbNullString Then gblPlEnableSounds = "1"

        ' position
        If gblPlAspectHidden = vbNullString Then gblPlAspectHidden = "0"
        If gblPlWidgetPosition = vbNullString Then gblPlWidgetPosition = "0"
        If gblPlWidgetLandscape = vbNullString Then gblPlWidgetLandscape = "0"
        If gblPlWidgetPortrait = vbNullString Then gblPlWidgetPortrait = "0"
        If gblPlLandscapeFormHoffset = vbNullString Then gblPlLandscapeFormHoffset = vbNullString
        If gblPlLandscapeFormVoffset = vbNullString Then gblPlLandscapeFormVoffset = vbNullString
        If gblPlPortraitHoffset = vbNullString Then gblPlPortraitHoffset = vbNullString
        If gblPlPortraitYoffset = vbNullString Then gblPlPortraitYoffset = vbNullString
        If gblPlvLocationPercPrefValue = vbNullString Then gblPlvLocationPercPrefValue = vbNullString
        If gblPlhLocationPercPrefValue = vbNullString Then gblPlhLocationPercPrefValue = vbNullString
                
        ' development
        If gblPlDebug = vbNullString Then gblPlDebug = "0"
        If gblPlDblClickCommand = vbNullString Then gblPlDblClickCommand = vbNullString
        If gblPlOpenFile = vbNullString Then gblPlOpenFile = vbNullString
        If gblPlDefaultEditor = vbNullString Then gblPlDefaultEditor = vbNullString
        If gblPlPreventDragging = vbNullString Then gblPlPreventDragging = "0"
        
        ' window
        If gblPlWindowLevel = vbNullString Then gblPlWindowLevel = "1" 'WindowLevel", gblPlSettingsFile)
        If gblPlOpacity = vbNullString Then gblPlOpacity = "100"
        If gblPlWidgetHidden = vbNullString Then gblPlWidgetHidden = "0"
        If gblPlHidingTime = vbNullString Then gblPlHidingTime = "0"
        If gblPlIgnoreMouse = vbNullString Then gblPlIgnoreMouse = "0"
        
        ' other
        If gblPlFirstTimeRun = vbNullString Then gblPlFirstTimeRun = "true"
        If gblPlLastSelectedTab = vbNullString Then gblPlLastSelectedTab = "general"
        If gblPlSkinTheme = vbNullString Then gblPlSkinTheme = "dark"
        
   On Error GoTo 0
   Exit Sub

validateInputs_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure validateInputs of form modMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : getTrinketsFile
' Author    : beededea
' Date      : 17/10/2019
' Purpose   : get this tool's entry in the trinkets settings file and assign the app.path
'---------------------------------------------------------------------------------------
'
Private Sub getTrinketsFile()
    On Error GoTo getTrinketsFile_Error

    Dim iFileNo As Integer: iFileNo = 0

    gblTrinketsDir = fSpecialFolder(feUserAppData) & "\trinkets" ' just for this user alone
    gblTrinketsFile = gblTrinketsDir & "\" & App.EXEName & ".ini"

    'if the folder does not exist then create the folder
    If Not fDirExists(gblTrinketsDir) Then
        MkDir gblTrinketsDir
    End If

    'if the settings.ini does not exist then create the file by copying
    If Not fFExists(gblTrinketsFile) Then

        iFileNo = FreeFile
        'open the file for writing
        Open gblTrinketsFile For Output As #iFileNo
        Write #iFileNo, App.Path & "\" & App.EXEName & ".exe"
        Write #iFileNo,
        Close #iFileNo
    End If

   On Error GoTo 0
   Exit Sub

getTrinketsFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getTrinketsFile of Form modMain"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : getToolSettingsFile
' Author    : beededea
' Date      : 17/10/2019
' Purpose   : get this tool's settings file and assign to a global var
'---------------------------------------------------------------------------------------
'
Private Sub getToolSettingsFile()
    On Error GoTo getToolSettingsFile_Error
    ''If debugflg = 1  Then Debug.Print "%getToolSettingsFile"
    
    Dim iFileNo As Integer: iFileNo = 0
    
    gblPlSettingsDir = fSpecialFolder(feUserAppData) & thisThingummy ' just for this user alone
    gblPlSettingsFile = gblPlSettingsDir & "\settings.ini"
        
    'if the folder does not exist then create the folder
    If Not fDirExists(gblPlSettingsDir) Then
        MkDir gblPlSettingsDir
    End If

    'if the settings.ini does not exist then create the file by copying
    If Not fFExists(gblPlSettingsFile) Then

        iFileNo = FreeFile
        'open the file for writing
        Open gblPlSettingsFile For Output As #iFileNo
        Close #iFileNo
    End If
    
   On Error GoTo 0
   Exit Sub

getToolSettingsFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getToolSettingsFile of Form modMain"

End Sub



'
'---------------------------------------------------------------------------------------
' Procedure : configureTimers
' Author    : beededea
' Date      : 07/05/2023
' Purpose   : configure any global timers here
'---------------------------------------------------------------------------------------
'
Private Sub configureTimers()

    On Error GoTo configureTimers_Error
    
    oldPrSettingsModificationTime = FileDateTime(gblPlSettingsFile)

    frmTimer.rotationTimer.Enabled = True
    frmTimer.settingsTimer.Enabled = True

    On Error GoTo 0
    Exit Sub

configureTimers_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure configureTimers of Module modMain"
            Resume Next
          End If
    End With
 
End Sub

'
'---------------------------------------------------------------------------------------
' Procedure : setHidingTime
' Author    : beededea
' Date      : 07/05/2023
' Purpose   : set the hiding time for the hiding timer, can't read the minutes from comboxbox as the prefs isn't yet open
'---------------------------------------------------------------------------------------
'
Private Sub setHidingTime()
    
    On Error GoTo setHidingTime_Error

    If gblPlHidingTime = "0" Then minutesToHide = 1
    If gblPlHidingTime = "1" Then minutesToHide = 5
    If gblPlHidingTime = "2" Then minutesToHide = 10
    If gblPlHidingTime = "3" Then minutesToHide = 20
    If gblPlHidingTime = "4" Then minutesToHide = 30
    If gblPlHidingTime = "5" Then minutesToHide = 60

    On Error GoTo 0
    Exit Sub

setHidingTime_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setHidingTime of Module modMain"
            Resume Next
          End If
    End With

End Sub



'---------------------------------------------------------------------------------------
' Procedure : createFormOnCurrentDisplay
' Author    : beededea
' Date      : 07/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub createFormOnCurrentDisplay()
    On Error GoTo createFormOnCurrentDisplay_Error

    With New_c.Displays(1) 'get the current Display
      fMain.InitAndShowAsFreeForm .WorkLeft, .WorkTop, 1000, 1000, "halfpenny Widget"
    End With

    On Error GoTo 0
    Exit Sub

createFormOnCurrentDisplay_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure createFormOnCurrentDisplay of Module modMain"
            Resume Next
          End If
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : handleUnhideMode
' Author    : beededea
' Date      : 13/05/2023
' Purpose   : when run in 'unhide' mode it writes the settings file then exits, the other
'             running but hidden process will unhide itself by timer.
'---------------------------------------------------------------------------------------
'
Private Sub handleUnhideMode(ByVal thisUnhideMode As String)
    
    On Error GoTo handleUnhideMode_Error

    If thisUnhideMode = "unhide" Then     'parse the command line
        gblPlUnhide = "true"
        sPutINISetting softwareThingummy, "unhide", gblPlUnhide, gblPlSettingsFile
        Call thingummyForm_Unload
        End
    End If

    On Error GoTo 0
    Exit Sub

handleUnhideMode_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure handleUnhideMode of Module modMain"
            Resume Next
          End If
    End With
End Sub
