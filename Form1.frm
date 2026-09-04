VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wait for Shelled Application to Terminate"
   ClientHeight    =   2370
   ClientLeft      =   5355
   ClientTop       =   2415
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   4890
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtApp 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   4455
   End
   Begin VB.CommandButton cmdShell 
      Caption         =   "&Start App"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   $"Form1.frx":0000
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the path to the app you want to start and wait for (a good example is calc.exe)"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const SYNCHRONIZE = &H100000
Const INFINITE = &HFFFF    'Wait forever
Const WAIT_OBJECT_0 = 0    'The state of the specified object is signaled.
Const WAIT_TIMEOUT = &H102 'The time-out interval elapsed, and the
                                  'object’s state is nonsignaled.

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Sub cmdQuit_Click()
Unload Me
End Sub

Private Sub Command1_Click()

End Sub


Private Sub cmdShell_Click()
Dim lPid As Long
Dim lHnd As Long
Dim lRet As Long

' The WaitForSingleObject function returns when one of the following occurs:
'   -   The specified object is in the signaled state.
'   -   The time-out interval elapses.
'
' The dwMilliseconds parameter specifies the time-out interval, in milliseconds.
' The function returns if the interval elapses, even if the object’s state is
' nonsignaled. If dwMilliseconds is zero, the function tests the object’s state
' and returns immediately. If dwMilliseconds is INFINITE, the function’s time-out
' interval never elapses.
'
' This example waits an INFINITE amount of time for the process to end. As a
' result this process will be frozen until the shelled process terminates. The
' down side is that if the shelled process hangs, so will this one.
'
' A better approach is to wait a specific amount of time.  Once the time-out
' interval expires, test the return value. If it is WAIT_TIMEOUT, the process
' is still not signaled.  Then you can either wait again or continue with your
' processing.
'
' DOS Applications:
'    Waiting for a DOS application is tricky because the DOS window never goes
'    away when the application is done.  To get around this, prefix the app that
'    you are shelling to with "command.com /c".
'
'    For example: lPid = Shell("command.com /c " & txtApp.Text, vbNormalFocus)
'
If Trim$(txtApp) = "" Then Exit Sub

lPid = Shell(txtApp.Text, vbNormalFocus)
If lPid <> 0 Then
    lHnd = OpenProcess(SYNCHRONIZE, 0, lPid)       'Get a handle to the shelled process.
    If lHnd <> 0 Then                              'If successful, wait for the
        lRet = WaitForSingleObject(lHnd, INFINITE) ' application to end.
        CloseHandle (lHnd)                         'Close the handle.
    End If
    MsgBox "Just terminated.", vbInformation, "Shelled Application"
End If
End Sub


