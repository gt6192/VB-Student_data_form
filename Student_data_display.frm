VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Student_data_display 
   Caption         =   "Form2"
   ClientHeight    =   6825
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   12135
   LinkTopic       =   "Form2"
   ScaleHeight     =   6825
   ScaleWidth      =   12135
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\Admin\Documents\VB-Databases\Students_data_list.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Studentsdata"
      Top             =   6240
      Visible         =   0   'False
      Width           =   3375
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Student_data_display.frx":0000
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   10398
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Students Data List"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   18.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   240
      Width           =   4455
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Quit 
         Caption         =   "Quit"
      End
   End
End
Attribute VB_Name = "Student_data_display"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Quit_Click()
Unload Me
End Sub
