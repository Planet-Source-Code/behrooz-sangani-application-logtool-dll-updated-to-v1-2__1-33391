VERSION 5.00
Begin VB.Form frmSample 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sample LogTool Project"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   4650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Log a predefined messages"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   1800
      Width           =   3375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Show Log Tracker"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   2160
      Width           =   3375
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Show Debug Window"
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   3240
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Log An Error Message"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Log Warning"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   3375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Log Information"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   1320
      Width           =   3375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "About"
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   2280
      Width           =   3375
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show MsgBox after log"
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   2880
      Width           =   3375
   End
End
Attribute VB_Name = "frmSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================================
'  Sample Form
'  Sample project for LogTool DLL
'=========================================================================================
'  Created By: Behrooz Sangani <bs20014@yahoo.com>
'  Published Date: 27/06/2002
'  WebSite: http://www.geocities.com/bs20014
'  Legal Copyright: Behrooz Sangani © 27/06/2002
'=========================================================================================

Dim Log As New LogMeBaby
Dim PreMsg As New PredefinedMessages

Private Sub Check1_Click()
    'You may generate a message box after logging
    Log.ShowMsgAfterLog = CBool(Check1.Value)
End Sub

Private Sub Check2_Click()
    'You can set debug window back and fore colors and place it where you want
    Log.ShowDebugWindow CBool(Check2.Value), Me.Left, Me.Top + Me.Height, Me.Width, 2000, vbBlue, vbYellow
End Sub

Private Sub Command1_Click()
    On Error GoTo error
    Dim a As Integer
    'let's generate a stupid error ! :)
    a = "as"

    Exit Sub
error:
    Log.SaveToLog Err.Description, ErrorLog
End Sub

Private Sub Command2_Click()

    'Generate a sample warning
    a = 105

    If a > 100 Then
        Log.SaveToLog "Number exceeds 100, there may be some unexpected errors!", WarningLog
    End If

End Sub

Private Sub Command3_Click()
    Log.SaveToLog "Command3 was just clicked!"
End Sub

Private Sub Command4_Click()
    'There are loads of windows predefined messages collected
    'in LogTool DLL as enums. Just select the desired message
    Log.SaveToLog PreMsg.Winnet32Err(ERROR_DEVICE_ALREADY_REMEMBERED), ErrorLog
End Sub

Private Sub Command6_Click()
    Log.About
    'Comments can be added to log file
    Log.SaveToLog "This is a test comment line saved after showing about message", CommentLine
End Sub

Private Sub Command5_Click()
    'NOT YET IMPLEMEMTED
    '  Log.ShowTracker
End Sub

Private Sub Form_Load()
    'Set the default values of the log file
    Log.AppTitle = App.Title
    Log.LogFile = App.Path & "\sample.log"
    'OverWriteLog on each form_load will create a new log.
    'If you wish to add info to the old log remove it
    Log.SaveToLog "App was started", , OverWriteLog
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Log.SaveToLog "App was unloaded"
End Sub

