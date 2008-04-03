# Microsoft Developer Studio Project File - Name="MFCMapi" - Package Owner=<4>
# Microsoft Developer Studio Generated Build File, Format Version 6.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Application" 0x0101

CFG=MFCMapi - Win32 Debug Unicode
!MESSAGE This is not a valid makefile. To build this project using NMAKE,
!MESSAGE use the Export Makefile command and run
!MESSAGE 
!MESSAGE NMAKE /f "MFCMapi.mak".
!MESSAGE 
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "MFCMapi.mak" CFG="MFCMapi - Win32 Debug Unicode"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "MFCMapi - Win32 Release" (based on "Win32 (x86) Application")
!MESSAGE "MFCMapi - Win32 Debug" (based on "Win32 (x86) Application")
!MESSAGE "MFCMapi - Win32 Debug Unicode" (based on "Win32 (x86) Application")
!MESSAGE "MFCMapi - Win32 Release Unicode" (based on "Win32 (x86) Application")
!MESSAGE 

# Begin Project
# PROP AllowPerConfigDependencies 0
# PROP Scc_ProjName "//depot/mfcmapi"
# PROP Scc_LocalPath "."
CPP=cl.exe
MTL=midl.exe
RSC=rc.exe

!IF  "$(CFG)" == "MFCMapi - Win32 Release"

# PROP BASE Use_MFC 6
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "Release"
# PROP BASE Intermediate_Dir "Release"
# PROP BASE Target_Dir ""
# PROP Use_MFC 6
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "Release"
# PROP Intermediate_Dir "Release"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MD /W4 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_AFXDLL" /Yu"stdafx.h" /FD /c
# ADD CPP /nologo /G6 /MD /W4 /WX /GX /Zi /Od /Ob1 /I ".\include" /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_AFXDLL" /D "_MBCS" /FR /YX"stdafx.h" /FD /GF /c
# ADD BASE MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x409 /d "NDEBUG" /d "_AFXDLL"
# ADD RSC /l 0x409 /d "NDEBUG" /d "_AFXDLL"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 /nologo /subsystem:windows /machine:I386
# ADD LINK32 version.lib /nologo /subsystem:windows /debug /machine:I386 /nodefaultlib:"LIBC" /libpath:".\lib" /release
# SUBTRACT LINK32 /pdb:none
# Begin Special Build Tool
WkspDir=.
TargetDir=.\Release
TargetPath=.\Release\MFCMapi.exe
TargetName=MFCMapi
SOURCE="$(InputPath)"
PostBuild_Desc=Copy Release build
PostBuild_Cmds=copy "$(TargetPath)" "$(WkspDir)"	copy "$(TargetDir)\$(TargetName).pdb" "$(WkspDir)"
# End Special Build Tool

!ELSEIF  "$(CFG)" == "MFCMapi - Win32 Debug"

# PROP BASE Use_MFC 6
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir "Debug"
# PROP BASE Intermediate_Dir "Debug"
# PROP BASE Target_Dir ""
# PROP Use_MFC 6
# PROP Use_Debug_Libraries 1
# PROP Output_Dir "Debug"
# PROP Intermediate_Dir "Debug"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MDd /W4 /Gm /GX /ZI /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_AFXDLL" /Yu"stdafx.h" /FD /GZ /c
# ADD CPP /nologo /G6 /MDd /W4 /WX /Gm /GX /Zi /Od /I ".\include" /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_AFXDLL" /D "_MBCS" /FR /YX"stdafx.h" /FD /GF /GZ /c
# ADD BASE MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x409 /d "_DEBUG" /d "_AFXDLL"
# ADD RSC /l 0x409 /d "_DEBUG" /d "_AFXDLL"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 /nologo /subsystem:windows /debug /machine:I386 /pdbtype:con
# ADD LINK32 version.lib /nologo /subsystem:windows /incremental:no /debug /machine:I386 /nodefaultlib:"LIBC" /libpath:".\lib" /release
# SUBTRACT LINK32 /pdb:none

!ELSEIF  "$(CFG)" == "MFCMapi - Win32 Debug Unicode"

# PROP BASE Use_MFC 6
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir "Debug_Unicode"
# PROP BASE Intermediate_Dir "Debug_Unicode"
# PROP BASE Ignore_Export_Lib 0
# PROP BASE Target_Dir ""
# PROP Use_MFC 6
# PROP Use_Debug_Libraries 1
# PROP Output_Dir "Debug_Unicode"
# PROP Intermediate_Dir "Debug_Unicode"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MDd /W4 /Gm /GX /Zi /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_AFXDLL" /D "_MBCS" /D "UNICODE" /Fr /YX"stdafx.h" /FD /GZ /c
# ADD CPP /nologo /G6 /MDd /W4 /WX /Gm /GX /Zi /Od /I ".\include" /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_AFXDLL" /D "_UNICODE" /D "UNICODE" /FR /YX"stdafx.h" /FD /GF /GZ /c
# ADD BASE MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x409 /d "_DEBUG" /d "_AFXDLL"
# ADD RSC /l 0x409 /d "_DEBUG" /d "_AFXDLL"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 mapi32.lib EdkMapi.Lib MbLogon.Lib EdkDebug.Lib EdkGUID.Lib EdkUtils.Lib /nologo /subsystem:windows /debug /machine:I386 /nodefaultlib:"LIBC" /pdbtype:con
# SUBTRACT BASE LINK32 /pdb:none
# ADD LINK32 version.lib /nologo /entry:"wWinMainCRTStartup" /subsystem:windows /incremental:no /debug /machine:I386 /nodefaultlib:"mfc42d.lib" /nodefaultlib:"mfcs42d.lib" /nodefaultlib:"mfco42d.lib" /nodefaultlib:"LIBC" /libpath:".\lib" /release
# SUBTRACT LINK32 /pdb:none

!ELSEIF  "$(CFG)" == "MFCMapi - Win32 Release Unicode"

# PROP BASE Use_MFC 6
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "Release_Unicode"
# PROP BASE Intermediate_Dir "Release_Unicode"
# PROP BASE Ignore_Export_Lib 0
# PROP BASE Target_Dir ""
# PROP Use_MFC 6
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "Release_Unicode"
# PROP Intermediate_Dir "Release_Unicode"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MD /W4 /GX /Zi /Od /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_AFXDLL" /D "_MBCS" /FR /YX"stdafx.h" /FD /c
# ADD CPP /nologo /G6 /MD /W4 /WX /GX /Zi /Od /I ".\include" /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_AFXDLL" /D "UNICODE" /D "_UNICODE" /FR /YX"stdafx.h" /FD /GF /c
# ADD BASE MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x409 /d "NDEBUG" /d "_AFXDLL"
# ADD RSC /l 0x409 /d "NDEBUG" /d "_AFXDLL"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 mapi32.lib MbLogon.Lib EdkMapi.Lib EdkDebug.Lib EdkGUID.Lib EdkUtils.Lib /nologo /subsystem:windows /debug /machine:I386 /nodefaultlib:"LIBC" /pdbtype:con
# ADD LINK32 version.lib /nologo /entry:"wWinMainCRTStartup" /subsystem:windows /debug /machine:I386 /nodefaultlib:"mfc42d.lib" /nodefaultlib:"mfcs42d.lib" /nodefaultlib:"mfco42d.lib" /nodefaultlib:"LIBC" /libpath:".\lib" /release
# SUBTRACT LINK32 /pdb:none

!ENDIF 

# Begin Target

# Name "MFCMapi - Win32 Release"
# Name "MFCMapi - Win32 Debug"
# Name "MFCMapi - Win32 Debug Unicode"
# Name "MFCMapi - Win32 Release Unicode"
# Begin Group "Source Files"

# PROP Default_Filter "cpp;c;cxx;rc;def;r;odl;idl;hpj;bat"
# Begin Source File

SOURCE=.\ABContDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\AbDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\AboutDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\AclDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\AddIns.cpp
# End Source File
# Begin Source File

SOURCE=.\AdviseSink.cpp
# End Source File
# Begin Source File

SOURCE=.\AttachmentsDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\BaseDialog.cpp
# End Source File
# Begin Source File

SOURCE=.\ContentsTableDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\ContentsTableListCtrl.cpp
# End Source File
# Begin Source File

SOURCE=.\DumpStore.cpp
# End Source File
# Begin Source File

SOURCE=.\Editor.cpp
# End Source File
# Begin Source File

SOURCE=.\Error.cpp
# End Source File
# Begin Source File

SOURCE=.\FakeSplitter.cpp
# End Source File
# Begin Source File

SOURCE=.\File.cpp
# End Source File
# Begin Source File

SOURCE=.\FileDialogEx.cpp
# End Source File
# Begin Source File

SOURCE=.\FolderDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\FormContainerDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\Guids.cpp
# End Source File
# Begin Source File

SOURCE=.\HexEditor.cpp
# End Source File
# Begin Source File

SOURCE=.\HierarchyTableDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\HierarchyTableTreeCtrl.cpp
# End Source File
# Begin Source File

SOURCE=.\ImportProcs.cpp
# End Source File
# Begin Source File

SOURCE=.\InterpretProp.cpp
# End Source File
# Begin Source File

SOURCE=.\InterpretProp2.cpp
# End Source File
# Begin Source File

SOURCE=.\MailboxTableDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\MainDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\MAPIABFunctions.cpp
# End Source File
# Begin Source File

SOURCE=.\MAPIFormFunctions.cpp
# End Source File
# Begin Source File

SOURCE=.\MAPIFunctions.cpp
# End Source File
# Begin Source File

SOURCE=.\MAPIMime.cpp
# End Source File
# Begin Source File

SOURCE=.\MapiObjects.cpp
# End Source File
# Begin Source File

SOURCE=.\MapiProcessor.cpp
# End Source File
# Begin Source File

SOURCE=.\MAPIProfileFunctions.cpp
# End Source File
# Begin Source File

SOURCE=.\MAPIProgress.cpp
# End Source File
# Begin Source File

SOURCE=.\MAPIStoreFunctions.cpp
# End Source File
# Begin Source File

SOURCE=.\MFCMapi.rc
# End Source File
# Begin Source File

SOURCE=.\MFCOutput.cpp
# End Source File
# Begin Source File

SOURCE=.\MFCUtilityFunctions.cpp
# End Source File
# Begin Source File

SOURCE=.\MsgServiceTableDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\MsgStoreDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\MyMAPIFormViewer.cpp
# End Source File
# Begin Source File

SOURCE=.\MySecInfo.cpp
# End Source File
# Begin Source File

SOURCE=.\MyWinApp.cpp
# End Source File
# Begin Source File

SOURCE=.\Output.cpp
# End Source File
# Begin Source File

SOURCE=.\ParentWnd.cpp
# End Source File
# Begin Source File

SOURCE=.\ProfileListDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\PropertyEditor.cpp
# End Source File
# Begin Source File

SOURCE=.\PropertyTagEditor.cpp
# End Source File
# Begin Source File

SOURCE=.\PropTagArray.cpp
# End Source File
# Begin Source File

SOURCE=.\ProviderTableDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\PublicFolderTableDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\RecipientsDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\Registry.cpp
# End Source File
# Begin Source File

SOURCE=.\RestrictEditor.cpp
# End Source File
# Begin Source File

SOURCE=.\RulesDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\SingleMAPIPropListCtrl.cpp
# End Source File
# Begin Source File

SOURCE=.\SortHeader.cpp
# End Source File
# Begin Source File

SOURCE=.\SortListCtrl.cpp
# End Source File
# Begin Source File

SOURCE=.\StreamEditor.cpp
# End Source File
# Begin Source File

SOURCE=.\TagArrayEditor.cpp
# End Source File
# End Group
# Begin Group "Header Files"

# PROP Default_Filter "h;hpp;hxx;hm;inl"
# Begin Source File

SOURCE=.\ABContDlg.h
# End Source File
# Begin Source File

SOURCE=.\ABDlg.h
# End Source File
# Begin Source File

SOURCE=.\AboutDlg.h
# End Source File
# Begin Source File

SOURCE=.\AclDlg.h
# End Source File
# Begin Source File

SOURCE=.\AddIns.h
# End Source File
# Begin Source File

SOURCE=.\AdviseSink.h
# End Source File
# Begin Source File

SOURCE=.\AttachmentsDlg.h
# End Source File
# Begin Source File

SOURCE=.\BaseDialog.h
# End Source File
# Begin Source File

SOURCE=.\bldver.h
# End Source File
# Begin Source File

SOURCE=.\ColumnTags.h
# End Source File
# Begin Source File

SOURCE=.\ContentsTableDlg.h
# End Source File
# Begin Source File

SOURCE=.\ContentsTableListCtrl.h
# End Source File
# Begin Source File

SOURCE=.\DumpStore.h
# End Source File
# Begin Source File

SOURCE=.\Editor.h
# End Source File
# Begin Source File

SOURCE=.\Enums.h
# End Source File
# Begin Source File

SOURCE=.\Error.h
# End Source File
# Begin Source File

SOURCE=.\ExtraPropTags.h
# End Source File
# Begin Source File

SOURCE=.\FakeSplitter.h
# End Source File
# Begin Source File

SOURCE=.\File.h
# End Source File
# Begin Source File

SOURCE=.\FileDialogEx.h
# End Source File
# Begin Source File

SOURCE=.\FolderDlg.h
# End Source File
# Begin Source File

SOURCE=.\FormContainerDlg.h
# End Source File
# Begin Source File

SOURCE=.\Guids.h
# End Source File
# Begin Source File

SOURCE=.\HexEditor.h
# End Source File
# Begin Source File

SOURCE=.\HierarchyTableDlg.h
# End Source File
# Begin Source File

SOURCE=.\HierarchyTableTreeCtrl.h
# End Source File
# Begin Source File

SOURCE=.\ImportProcs.h
# End Source File
# Begin Source File

SOURCE=.\InterpretProp.h
# End Source File
# Begin Source File

SOURCE=.\InterpretProp2.h
# End Source File
# Begin Source File

SOURCE=.\MailboxTableDlg.h
# End Source File
# Begin Source File

SOURCE=.\MainDlg.h
# End Source File
# Begin Source File

SOURCE=.\MAPIABFunctions.h
# End Source File
# Begin Source File

SOURCE=.\MAPIFormFunctions.h
# End Source File
# Begin Source File

SOURCE=.\MAPIFunctions.h
# End Source File
# Begin Source File

SOURCE=.\MAPIMime.h
# End Source File
# Begin Source File

SOURCE=.\MapiObjects.h
# End Source File
# Begin Source File

SOURCE=.\MapiProcessor.h
# End Source File
# Begin Source File

SOURCE=.\MAPIProfileFunctions.h
# End Source File
# Begin Source File

SOURCE=.\MAPIProgress.h
# End Source File
# Begin Source File

SOURCE=.\MAPIStoreFunctions.h
# End Source File
# Begin Source File

SOURCE=.\MFCOutput.h
# End Source File
# Begin Source File

SOURCE=.\MFCUtilityFunctions.h
# End Source File
# Begin Source File

SOURCE=.\MsgServiceTableDlg.h
# End Source File
# Begin Source File

SOURCE=.\MsgStoreDlg.h
# End Source File
# Begin Source File

SOURCE=.\MyMAPIFormViewer.h
# End Source File
# Begin Source File

SOURCE=.\MySecInfo.h
# End Source File
# Begin Source File

SOURCE=.\MyWinApp.h
# End Source File
# Begin Source File

SOURCE=.\Output.h
# End Source File
# Begin Source File

SOURCE=.\ParentWnd.h
# End Source File
# Begin Source File

SOURCE=.\ProfileListDlg.h
# End Source File
# Begin Source File

SOURCE=.\PropertyEditor.h
# End Source File
# Begin Source File

SOURCE=.\PropertyTagEditor.h
# End Source File
# Begin Source File

SOURCE=.\PropTagArray.h
# End Source File
# Begin Source File

SOURCE=.\ProviderTableDlg.h
# End Source File
# Begin Source File

SOURCE=.\PublicFolderTableDlg.h
# End Source File
# Begin Source File

SOURCE=.\RecipientsDlg.h
# End Source File
# Begin Source File

SOURCE=.\Registry.h
# End Source File
# Begin Source File

SOURCE=.\Resource.h
# End Source File
# Begin Source File

SOURCE=.\RestrictEditor.h
# End Source File
# Begin Source File

SOURCE=.\RulesDlg.h
# End Source File
# Begin Source File

SOURCE=.\SingleMAPIPropListCtrl.h
# End Source File
# Begin Source File

SOURCE=.\SortHeader.h
# End Source File
# Begin Source File

SOURCE=.\SortListCtrl.h
# End Source File
# Begin Source File

SOURCE=.\StdAfx.h
# End Source File
# Begin Source File

SOURCE=.\StreamEditor.h
# End Source File
# Begin Source File

SOURCE=.\TagArrayEditor.h
# End Source File
# End Group
# Begin Group "Resource Files"

# PROP Default_Filter "ico;cur;bmp;dlg;rc2;rct;bin;rgs;gif;jpg;jpeg;jpe"
# Begin Source File

SOURCE=.\Res\down.bmp
# End Source File
# Begin Source File

SOURCE=.\MFCMapi.ico
# End Source File
# Begin Source File

SOURCE=.\res\MFCMapi.rc2
# End Source File
# Begin Source File

SOURCE=.\Res\up.bmp
# End Source File
# End Group
# Begin Source File

SOURCE=.\mfcmapi.exe.manifest
# End Source File
# End Target
# End Project
