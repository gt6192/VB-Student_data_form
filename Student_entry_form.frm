VERSION 5.00
Begin VB.Form Student_entry_form 
   Caption         =   "Form1"
   ClientHeight    =   5175
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   9975
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Data Form"
      Height          =   4095
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   9495
      Begin VB.CommandButton Command2 
         Caption         =   "Add New"
         Height          =   615
         Left            =   4800
         TabIndex        =   17
         Top             =   2400
         Width           =   4455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Submit Data"
         Height          =   615
         Left            =   4800
         TabIndex        =   16
         Top             =   3240
         Width           =   4455
      End
      Begin VB.TextBox Text7 
         DataField       =   "Fees"
         DataSource      =   "Data1"
         Height          =   405
         Left            =   240
         TabIndex        =   11
         Text            =   "40000"
         Top             =   3480
         Width           =   4215
      End
      Begin VB.TextBox Text6 
         DataField       =   "Subject3"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   2520
         Width           =   4215
      End
      Begin VB.TextBox Text4 
         DataField       =   "Subject2"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   4800
         TabIndex        =   9
         Top             =   1680
         Width           =   4455
      End
      Begin VB.TextBox Text5 
         DataField       =   "Subject1"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1680
         Width           =   4215
      End
      Begin VB.TextBox Text3 
         DataField       =   "Lastname"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   5760
         TabIndex        =   4
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox Text2 
         DataField       =   "Firstname"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         DataField       =   "id"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Fees"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Subject 3"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Subject 2"
         Height          =   255
         Left            =   4800
         TabIndex        =   13
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Subject 1"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Last Name"
         Height          =   255
         Left            =   5760
         TabIndex        =   7
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "First Name"
         Height          =   255
         Left            =   1680
         TabIndex        =   6
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "ID"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   255
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\Admin\Documents\VB-Databases\Students_data_list.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Studentsdata"
      Top             =   5520
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "Students Data Entry"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   18.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   240
      Width           =   4575
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Quit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu Options 
      Caption         =   "Options"
      Begin VB.Menu Students_Data_List 
         Caption         =   "Students Data List"
      End
   End
   Begin VB.Menu View 
      Caption         =   "View"
      Begin VB.Menu Mode 
         Caption         =   "Mode"
         Begin VB.Menu Light_Mode 
            Caption         =   "Light Mode"
         End
         Begin VB.Menu Dark_Mode 
            Caption         =   "Dark Mode"
         End
      End
   End
End
Attribute VB_Name = "Student_entry_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.Update
End Sub

Private Sub Command2_Click()
Data1.Recordset.AddNew
End Sub

Private Sub Dark_Mode_Click()
Label1.ForeColor = &HFFFFFF
Label2.ForeColor = &HFFFFFF
Label3.ForeColor = &HFFFFFF
Label4.ForeColor = &HFFFFFF
Label5.ForeColor = &HFFFFFF
Label6.ForeColor = &HFFFFFF
Label7.ForeColor = &HFFFFFF
Label8.ForeColor = &HFFFFFF

Label1.BackColor = &H0&
Label2.BackColor = &H0&
Label3.BackColor = &H0&
Label4.BackColor = &H0&
Label5.BackColor = &H0&
Label6.BackColor = &H0&
Label7.BackColor = &H0&
Label8.BackColor = &H0&

Student_entry_form.BackColor = &H0&

Frame1.BackColor = &H0&
Frame1.ForeColor = &HFFFFFF

End Sub

Private Sub Light_Mode_Click()
Label1.ForeColor = &H0&
Label2.ForeColor = &H0&
Label3.ForeColor = &H0&
Label4.ForeColor = &H0&
Label5.ForeColor = &H0&
Label6.ForeColor = &H0&
Label7.ForeColor = &H0&
Label8.ForeColor = &H0&

Label1.BackColor = &H8000000F
Label2.BackColor = &H8000000F
Label3.BackColor = &H8000000F
Label4.BackColor = &H8000000F
Label5.BackColor = &H8000000F
Label6.BackColor = &H8000000F
Label7.BackColor = &H8000000F
Label8.BackColor = &H8000000F

Student_entry_form.BackColor = &H8000000F

Frame1.BackColor = &H8000000F
Frame1.ForeColor = &H0&
End Sub

Private Sub Quit_Click()
Unload Me
Unload Student_data_display
End Sub

Private Sub Students_Data_List_Click()
Student_data_display.Show
End Sub

