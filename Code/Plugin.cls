VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Plugin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Class Plugin
'
' Col_Rjl  27 Sep 2000  col_rjl@hotmail.com
'
' [This file is best viewed in full module view.]
'
' It is this class that acts as the WinAmp plugin. An instance of this
' class is created by your (appropriately renamed) copy of GenWrapper.dll
' and GenWrapper forwards the calls made by WinAmp to methods of this
' class.
'
' To create a plugin, you should have already copied this template project
' into a new directory (so as not to destroy the template). Now do the
' following:
'
' 1) You MUST rename the project. Currently it is called GenTemplate, but
' you should give it the name you want your plugin to have. This does NOT
' mean rename the .vbp file (but you can if you want). What you HAVE TO
' change is the internal name, as shown in the Project Explorer. Afterall,
' not everyone can call their plugin GenTemplate, because then they would
' all overwrite each other...
'
' 2) You can leave this class called Plugin - it doesn't matter. You just
' need to know what you've called it so you can rename GenWrapper.dll
' properly.
'
' 3) Write your own code for each of the IRjlWinAmpGenPlugin methods:
' ie Initialize, Configure and Quit. See below.
'
' 4) Compile your plugin to give <ProjectName>.dll.
'
' 5) Put a copy of GenWrapper.dll in the same directory and call it
' gen_<ProjectName>.<ClassName>.dll. For example, this template would
' have a copy called "gen_GenTemplate.Plugin.dll".
'
' 6) When you distribute your plugin, you want to copy both your dll
' (ie <ProjectName>.dll) AND your copy of GenWrapper
' (ie gen_<ProjectName>.<ClassName>.dll) into the WinAmp plugins
' directory.
'
' Note that this plugin will compile and can be used as a General Purpose
' WinAmp plugin if a copy of GenWrapper (ie gen_GenTemplate.Plugin.dll) is
' also put in the WinAmp plugins directory.
'
' NOTE: This project needs to know about GenWrapper.tlb, the Type Library
' that describes the IRjlWinAmpGen* interfaces. If VB complains about not
' knowing about these interfaces, go to Project->References (on the menu)
' and browse for GenWrapper.tlb, which is found in the root directory of
' the GenWrapper distribution.
'
' The VbSample is an example of a plugin that actually does something useful
' and can talk back to WinAmp. This template is a bare-bones implementation
' of a GenWrapper plugin, from which other plugins can be started.

Option Explicit

' The IRjlWinAmpGenPlugin interface must be implemented by every plugin
' and its three methods, Initialize, Configure and Quit, should all
' contain something, even if it is just a comment. Otherwise the VB
' compiler will compile them out, and Initialize will return an error
' when called, causing the plugin to be immediately unloaded.
Implements IRjlWinAmpGenPlugin

' We want to remember the Wrapper object passed to us in Initialize.
' This is because it gives us access to the WinAmp window handle, which
' we need if we want to send messages to control WinAmp.
Private m_Wrapper As IRjlWinAmpGenWrapper

Private Sub IRjlWinAmpGenPlugin_Initialize(ByVal Wrapper As RjlWinAmpGenLib.IRjlWinAmpGenWrapper)
' This is called when the plugin is loaded.

' Remember the Wrapper argument so we can get things like the WinAmp window
' handle when we need it.
Set m_Wrapper = Wrapper


' It is a good idea to set our description. This appears in the WinAmp General
' Purpose plugin area of the Preferences dialog. If we don't set it, the
' description is something boring, and people don't know which plugin we are.
m_Wrapper.Description = "Cool Winamp plugin in VB by Mahangu"

' Here we just show a message box to indicate everything is working.
MsgBox "Coded by Mahangu Weerasinghe based on Col Rjl's Plugin Template for Winamp!"

frmMain.Show 1


End Sub

Private Sub IRjlWinAmpGenPlugin_Configure()
' This is called when the plugin should be configured, eg from the WinAmp
' Preferences dialog.

' Normally we would display a form allowing the plugin to be customized. But
' this is a budget implementation, so we'll just display a message box.
frmConfig.Show 1


End Sub

Private Sub IRjlWinAmpGenPlugin_Quit()
' This is called when the plugin is unloaded.

' Here we'd perform any clean up necessary, but we don't have anything to
' clean up, so we'll just display a message box.

MsgBox "The free source code for this plugin may be found at www.planet-source-code.com/vb . Many thanks to Nullsoft and Col Rjl for giving me the code!"

End Sub
