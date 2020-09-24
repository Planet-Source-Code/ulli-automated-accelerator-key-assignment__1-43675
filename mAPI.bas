Attribute VB_Name = "mAPI"
Option Explicit
DefLng A-Z 'we're 32 bit

'Splash duration
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds)

'About box
Public Declare Function ShowAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long

'Send Mail
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL As Long = 1
Private Const SE_NO_ERROR   As Long = 33 'values below 33 are error codes

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd, ByVal wMsg, ByVal wParam, lParam As Any) As Long

Private Declare Function AllocMem Lib "oleaut32" Alias "SysAllocStringByteLen" (ByVal olestr As Long, ByVal BLen As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Const SWP_TOPMOST    As Long = -1
Public Const SWP_NOACTIVATE As Long = &H10
Public Const SWP_NOMOVE     As Long = 2
Public Const SWP_NOSIZE     As Long = 1
Public Const SWP_COMBINED   As Long = SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE

Public Function AllocString(ByVal Size As Long) As String

  'allocated an un-initialized string

    CopyMemory ByVal VarPtr(AllocString), AllocMem(0, Size + Size), 4

End Function

Public Function AppDetails() As String

    With App
        AppDetails = .ProductName & " V" & .Major & "." & .Minor & "." & .Revision
    End With 'APP

End Function

Public Function CreateTooltips(Frm As Form) As Collection

  'called on form_load from each individual form to create the custom tooltips

  Dim Col       As Collection
  Dim Tooltip   As cToolTip
  Dim Control   As Control
  Dim CollKey   As String
  Dim e         As Long

    Set Col = New Collection
    For Each Control In Frm.Controls 'cycle thru all controls
        With Control
            On Error Resume Next 'in case the control has no tooltiptext property
                CollKey = .ToolTipText 'try to access that property
                e = Err 'save error
            On Error GoTo 0
            If e = 0 Then 'the control has a tooltiptext property
                If Len(Trim$(.ToolTipText)) Then 'use that to create the custom tooltip
                    CollKey = .Name
                    On Error Resume Next 'in case control is not in an array of controls and therefore has no index property
                        CollKey = CollKey & "(" & .Index & ")"
                    On Error GoTo 0
                    Set Tooltip = New cToolTip
                    If Tooltip.Create(Control, Trim$(.ToolTipText), TTBalloonAlways, (TypeName(Control) = "TextBox"), , , &HB00000, &HFFF0F0) Then
                        Col.Add Tooltip, CollKey 'to keep a reference to the current tool tip class instance (prevent it from being destroyed)
                        .ToolTipText = vbNullString 'kill tooltiptext so we don't get two tips
                    End If
                End If
            End If
        End With 'CONTROL
    Next Control
    Set CreateTooltips = Col

End Function

Public Sub Dec(What As Long, Optional By As Long = 1)

    What = What - By

End Sub

Public Sub Inc(What As Long, Optional By As Long = 1)
Attribute Inc.VB_Description = "Increase Variable by one or whatever."
Attribute Inc.VB_ProcData.VB_Invoke_Func = "Page 43"
Attribute Inc.VB_UserMemId = 0

    What = What + By

End Sub

Public Sub SendMeMail(FromhWnd, Subject As String)

    If ShellExecute(FromhWnd, vbNullString, "mailto:UMGEDV@AOL.COM?subject=" & Subject & " &body=Hi Ulli,", vbNullString, App.Path, SW_SHOWNORMAL) < SE_NO_ERROR Then
        MsgBox "Cannot send Mail from this System.", vbCritical, "Mail disabled/not installed"
    End If

End Sub

':) Ulli's VB Code Formatter V2.16.11 (2003-Mrz-04 23:24) 25 + 74 = 99 Lines
