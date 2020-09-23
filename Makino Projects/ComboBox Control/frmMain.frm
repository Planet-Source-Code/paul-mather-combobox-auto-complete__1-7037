VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ComboBox Control"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   5880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLoadHistory 
      Caption         =   "Load History"
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdSaveHistory 
      Caption         =   "Save History"
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdAddHistory 
      Caption         =   "Add to History"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdDropCombo 
      Caption         =   "Drop Combo"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   960
      Width           =   2655
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAddHistory_Click()
    Call ComboAddToHistory(Combo1)
End Sub
Private Sub cmdSaveHistory_Click()
    Call ComboSaveHistory(Combo1)
End Sub
Private Sub Combo1_Change()
    Call ComboAutoComplete(Combo1)
End Sub
Private Sub cmdDropCombo_Click()
    Call ComboDropdown(Combo1)
End Sub
Private Sub cmdLoadHistory_Click()
    Call ComboLoadHistory(Combo1)
End Sub

Private Sub Form_Load()
    Combo1.AddItem "This is a test"
    Combo1.AddItem "Again testing"
    Combo1.AddItem "Apples"
    Combo1.AddItem "Again testing push"
    Combo1.AddItem "Again testing pop"
    Combo1.AddItem "Hello There"
    Combo1.AddItem "This is a string that is longer than the available space"
End Sub

