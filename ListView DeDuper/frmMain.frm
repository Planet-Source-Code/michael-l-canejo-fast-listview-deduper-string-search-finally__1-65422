VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "ListView API DeDuper"
   ClientHeight    =   4035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   ScaleHeight     =   4035
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Add Test Items"
      Height          =   390
      Left            =   1740
      TabIndex        =   3
      Top             =   3495
      Width           =   1665
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove ListView Dupes via Collection"
      Height          =   315
      Left            =   1005
      TabIndex        =   2
      Top             =   3045
      Width           =   3105
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Remove ListView Dupes via API"
      Height          =   315
      Left            =   1005
      TabIndex        =   1
      Top             =   2670
      Width           =   3105
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2355
      Left            =   195
      TabIndex        =   0
      Top             =   180
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   4154
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "method2"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   4230
      TabIndex        =   7
      Top             =   3090
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "method1"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   4230
      TabIndex        =   6
      Top             =   2745
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Slower"
      Height          =   195
      Left            =   435
      TabIndex        =   5
      Top             =   2745
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Faster"
      Height          =   195
      Left            =   480
      TabIndex        =   4
      Top             =   3090
      Width           =   435
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
 Dim lngCount As Long
 lngCount = DeDupe_ListView_API(ListView1)
 
 Caption = "List Count: " & ListView1.ListItems.Count
 MsgBox lngCount & " dupes removed via API method!", vbInformation, "ListView DeDuper"
End Sub

Private Sub Command2_Click()
 Dim lngCount As Long
 lngCount = DeDupe_ListView_Col(ListView1)
 
 Caption = "List Count: " & ListView1.ListItems.Count
 MsgBox lngCount & " dupes removed via Collection method!", vbInformation, "ListView DeDuper"
End Sub

Private Sub Command3_Click()
 Dim itmAdd As ListItem
 Dim x As Long
 For x = 1 To 400
  'Add data to the ListView control
  Set itmAdd = ListView1.ListItems.Add(Text:="Joe")
  itmAdd.SubItems(1) = "05/07/97"

  Set itmAdd = ListView1.ListItems.Add(Text:="Fred")
  itmAdd.SubItems(1) = "05/17/97"

  Set itmAdd = ListView1.ListItems.Add(Text:="Anne")
  itmAdd.SubItems(1) = "04/01/97"
 Next
 
 Caption = "List Count: " & ListView1.ListItems.Count
 
End Sub

Private Sub Form_Load()

 Dim clmAdd As ColumnHeader
 

 'Add two Column Headers to the ListView control
 Set clmAdd = ListView1.ColumnHeaders.Add(Text:="Name")
 Set clmAdd = ListView1.ColumnHeaders.Add(Text:="Date")

 'Set the view property of the Listview control to Report view
 ListView1.View = lvwReport

 Command3_Click
 
 
End Sub


