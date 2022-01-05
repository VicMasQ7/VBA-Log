VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check_LogToFile 
      Caption         =   "Log to File"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test  Logger"
      Height          =   855
      Left            =   1560
      TabIndex        =   0
      Top             =   960
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check_LogToFile_Click()
    If Check_LogToFile = 1 Then 'Checked
        Logger.LogCallback = "LogFile"
    Else 'UnChecked
        Logger.LogCallback = Empty
    End If

End Sub

' Attach alternative logging function(s)
Public Sub LogFile(Level As Long, Message As String, From As String)
  ' ... Impletetion
  Debug.Print "... Here i log in File...."
  
End Sub

Public Sub LogWorkbook(Level As Long, Message As String, From As String)
  ' ... Impletetion
    Debug.Print "... Here i log in Workbook...."
End Sub

Private Sub Command1_Click()
    Logger.LogDebug "Howdy!"
    ' -> does nothing because logging is disabled by default

    Logger.LogEnabled = True
    ' -> Log all levels (Trace, Debug, Info, Warn, Error)

    Logger.LogThreshold = 3
    ' -> Log levels >= 3 (Info, Warn, and Error)

    Logger.LogTrace "Start of logging"
    Logger.LogDebug "Logging has started"
    Logger.LogInfo "Logged with VBA-Log"
    Logger.LogWarn "Watch out!", "ModuleName.SubName"
    Logger.LogError "Something went wrong", "ClassName.FunctionName", Err.Number

End Sub
