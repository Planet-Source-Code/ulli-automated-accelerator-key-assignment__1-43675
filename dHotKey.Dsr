VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} dHotkey 
   ClientHeight    =   11385
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   12600
   _ExtentX        =   22225
   _ExtentY        =   20082
   _Version        =   393216
   Description     =   "VB Hotkey Generator"
   DisplayName     =   "Ulli's VB HotKey Generator"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "dHotkey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'© 2003     UMGEDV GmbH  (umgedv@aol.com)
'
'Author     UMG (Ulli K. Muehlenweg)
'
'Title      VB6 Hotkey Add-In
'
'Purpose    This Add-In analyzes all (or selected) controls in a form and
'           generates the necessary accelerator hotkeys, ie those keys to
'           be pressed together with the Alt-key to access a control.
'
'           Compile the DLL into your VB directory and then use the Add-Ins
'           Manager to load the Hotkey Add-In into VB.
'
'**********************************************************************************
'Development History
'**********************************************************************************
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'04Mar2003 Version 1.1.5 - UMG
'
'  Killed some un-used variables and renamed some others
'  Prevent form close while busy
'  Fixed tab order and initial focus in fHotkey
'  Optimized permutation function (in particular the "sort")
'  Altered termination alert
'  Added cell backcolor to highlight unresolved ambiguities
'  Added cell alignment
'  Adjusted column width without scrollbar
'  Some code cosmetics
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'02Mar2003 Version 1.1.2 - UMG
'
'  Got about 35% run time improvement by some optimizing:
'    Exit Examination Loop early
'    Store Options in local variables
'    Transform to lower instead of case-insensitive InStr
'    Moved all arrays into a structure and used With whenever possible to eliminate
'    frequent indexing
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'01Mar2003 Version 1.0.12 - UMG
'
'  Prototype
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
DefLng A-Z 'we're 32 bit

'control classes that will not be processed
'you may modify this as you find necessary
Private Const ExcludedClassName1 As String = "Frame"
Private Const ExcludedClassName2 As String = ""

Private Const MenuName          As String = "Add-Ins" 'you may need to localize "Add-Ins"

Private Const SplashTime        As Long = 555
Private VBInstance              As VBIDE.VBE
Private CommandBarMenu          As CommandBar
Private MenuItem                As CommandBarControl
Private WithEvents MenuEvents   As CommandBarEvents
Attribute MenuEvents.VB_VarHelpID = -1

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)

  Dim ClipboardText As String
  Dim i             As Long

    Set VBInstance = Application
    If ConnectMode = ext_cm_External Then
        AssignHotkeys
      Else 'NOT CONNECTMODE...
        On Error Resume Next
            Set CommandBarMenu = VBInstance.CommandBars(MenuName)
        On Error GoTo 0
        If CommandBarMenu Is Nothing Then
            MsgBox "Hotkey was loaded but could not be connected to the " & MenuName & " menu.", vbCritical
          Else 'NOT COMMANDBARMENU...
            fSplash.Show
            DoEvents
            With CommandBarMenu
                Set MenuItem = .Controls.Add(msoControlButton)
                i = .Controls.Count - 1
                If .Controls(i).BeginGroup And Not .Controls(i - 1).BeginGroup Then
                    'menu separator required
                    MenuItem.BeginGroup = True
                End If
            End With 'COMMANDBARMENU
            'set menu caption
            With App
                MenuItem.Caption = "&" & .ProductName & " V" & .Major & "." & .Minor & "." & .Revision & "..."
            End With 'APP
            With Clipboard
                ClipboardText = .GetText
                'set menu picture
                .SetData fSplash.picMenu.Image
                MenuItem.PasteFace
                .Clear
                .SetText ClipboardText
            End With 'CLIPBOARD
            'set event handler
            Set MenuEvents = VBInstance.Events.CommandBarEvents(MenuItem)
            'done connecting
            Sleep SplashTime
            Unload fSplash
        End If
    End If

End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)

    On Error Resume Next
        MenuItem.Delete
    On Error GoTo 0

End Sub

Public Sub AssignHotkeys()

  Dim Cntls         As Object 'we have to use a neutral variable with late bind because either VBControls or SelectedVBControls is assigned to it
  Dim Cntl          As VBControl
  Dim Count         As Long
  Dim Capt          As String
  Dim Hdr           As String
  Dim Name          As String
  Dim AllControls   As Boolean

    If Not VBInstance.SelectedVBComponent Is Nothing Then
        With VBInstance.SelectedVBComponent
            Name = " [" & .Name & "]"
            If Len(Name) = 3 Then 'just a space and the two brackets - ie no name
                Name = " [unknown]"
            End If
            Hdr = AppDetails & Name
            If .Type = vbext_ct_VBForm Or _
               .Type = vbext_ct_VBMDIForm Or _
               .Type = vbext_ct_DocObject Or _
               .Type = vbext_ct_UserControl Or _
               .Type = vbext_ct_PropPage Then
                With .Designer
                    AllControls = (.SelectedVBControls.Count = 0)
                    If AllControls Then
                        Set Cntls = .VBControls
                      Else 'ALLCONTROLS = FALSE/0
                        Set Cntls = .SelectedVBControls
                    End If
                End With '.DESIGNER
                If Cntls.Count Then
                    Load fHotkey
                    Count = 0
                    On Error Resume Next
                        For Each Cntl In Cntls 'cycle through all controls extracting name and caption
                            With Cntl
                                If .ClassName <> ExcludedClassName1 And .ClassName <> ExcludedClassName2 Then
                                    Err.Clear
                                    Capt = .Properties("Caption")
                                    If Err = 0 And Len(Trim$(Capt)) Then 'send name and caption to fHotkey
                                        Inc Count
                                        fHotkey.LetCapt(Capt) = .Properties("Name")
                                    End If
                                End If
                            End With 'CNTL
                        Next Cntl
                    On Error GoTo 0
                    If Count Then 'some where sent
                        With fHotkey
                            .Caption = Hdr
                            .FillTheGrid False
                            .Show vbModal
                            If .Tag = "0" Then 'Apply was clicked
                                Count = 0
                                On Error Resume Next
                                    For Each Cntl In Cntls 'cycle through all controls
                                        With Cntl
                                            If .ClassName <> ExcludedClassName1 And .ClassName <> ExcludedClassName2 Then
                                                Err.Clear
                                                Capt = .Properties("Caption")
                                                If Err = 0 Then 'has a caption property
                                                    If Len(Trim$(Capt)) Then 'caption is not empty
                                                        Inc Count
                                                        .Properties("Caption") = fHotkey.GetCapt(Count) 'update caption from fHotkey item
                                                    End If
                                                End If
                                            End If
                                        End With 'CNTL
                                    Next Cntl
                                On Error GoTo 0
                            End If
                        End With 'FHOTKEY
                        Unload fHotkey
                      Else 'COUNT = FALSE/0
                        If AllControls Then
                            MsgBox "There are no controls with a caption in" & Name & ".", vbExclamation, Hdr
                          Else 'ALLCONTROLS = FALSE/0
                            If Cntls.Count = 1 Then
                                MsgBox "The control [" & Cntls(0).Properties("Name") & "] has no caption.", vbExclamation, Hdr
                              Else 'NOT CNTLS.COUNT...
                                MsgBox "There are no controls with a caption among those you have selected.", vbExclamation, Hdr
                            End If
                        End If
                    End If
                  Else 'CNTLS.COUNT = FALSE/0
                    MsgBox "There are no controls in" & Name & ".", vbExclamation, Hdr
                End If
              Else 'NOT .TYPE...
                MsgBox "The currently selected component" & Name & " is not a form.", vbExclamation, Hdr
            End If
        End With 'VBINSTANCE.SELECTEDVBCOMPONENT
      Else 'NOT NOT...
        MsgBox "Cannot find any selected component - you must select a form first.", vbExclamation, App.ProductName & " [none]"
    End If

End Sub

Private Sub MenuEvents_Click(ByVal CommandBarControl As Object, Handled As Boolean, CancelDefault As Boolean)

    AssignHotkeys

End Sub

':) Ulli's VB Code Formatter V2.16.11 (2003-Mrz-04 23:24) 61 + 160 = 221 Lines
