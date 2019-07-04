VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "boot.ini Error Master"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4455
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "Form1.frx":0442
      Left            =   3360
      List            =   "Form1.frx":046A
      TabIndex        =   4
      Text            =   "Drive"
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "More Program"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   4215
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "Form1.frx":04AA
      Left            =   120
      List            =   "Form1.frx":04B4
      Sorted          =   -1  'True
      TabIndex        =   1
      Text            =   "Select for OS"
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ⓒ2012 Lee Dong-gun (Sotaneum)"
      Height          =   180
      Left            =   960
      TabIndex        =   6
      Top             =   2280
      Width           =   2880
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Where is setup for Windows OS?"
      Height          =   180
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   2760
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

a = Combo2.Text & "boot.ini"
If Combo2.Text = "Drive" Then
MsgBox "Erorr :: Where is setup for Windows OS?", vbExclamation, "Erorr"

Else

    If Combo1.Text = "Microsoft Windows XP Home Edition" Then
    
        Open a For Output As #1
        Print #1, "[boot loader]"
        Print #1, "timeout=5"
        Print #1, "default=multi(0)disk(0)rdisk(0)partition(1)\WINDOWS"
        Print #1, "[operating systems]"
        Print #1, "multi(0)disk(0)rdisk(0)partition(1)\WINDOWS=" & """" & "Microsoft Windows XP Home Edition" & """" & " /noexecute=optin /fastdetect"
        Close #1
        MsgBox "Success. Reboot computer. (Need)", vbInformation, "boot.ini File"
    End If
    
End If

If Combo2.Text = "Drive" Then
MsgBox "Erorr :: Where is setup for Windows OS?", vbExclamation, "Erorr"

Else

    If Combo1.Text = "Microsoft Windows XP Professional" Then
        Open a For Output As #2
        Print #2, "[boot loader]"
        Print #2, "timeout=5"
        Print #2, "default=multi(0)disk(0)rdisk(0)partition(1)\WINDOWS"
        Print #2, "[operating systems]"
        Print #2, "multi(0)disk(0)rdisk(0)partition(1)\WINDOWS=" & """" & "Microsoft Windows XP Professional" & """" & " /noexecute=optin /fastdetect"
        Close #2
        MsgBox "Success. Reboot computer. (Need)", vbInformation, "boot.ini File"
    End If
    
End If


End Sub

Private Sub Command2_Click()
Shell ("explorer.exe http://duration.digimoon.net/")

End Sub

Private Sub Command3_Click()
End

End Sub

