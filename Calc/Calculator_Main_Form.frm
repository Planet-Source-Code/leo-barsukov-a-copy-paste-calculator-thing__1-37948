VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form Calc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator"
   ClientHeight    =   4980
   ClientLeft      =   6210
   ClientTop       =   4965
   ClientWidth     =   4530
   Icon            =   "Calculator_Main_Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Calculator_Main_Form.frx":1042
   ScaleHeight     =   4980
   ScaleWidth      =   4530
   Begin MSScriptControlCtl.ScriptControl Script 
      Left            =   1440
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   -1  'True
   End
   Begin VB.TextBox Scripto 
      Height          =   285
      Left            =   -720
      MultiLine       =   -1  'True
      TabIndex        =   30
      Top             =   3240
      Width           =   150
   End
   Begin VB.CommandButton Num 
      Appearance      =   0  'Flat
      Caption         =   ")"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   16
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Num 
      Appearance      =   0  'Flat
      Caption         =   "Sin"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   20
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton Num 
      Appearance      =   0  'Flat
      Caption         =   "Pi"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   22
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton Num 
      Appearance      =   0  'Flat
      Caption         =   "SqR"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   21
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton Num 
      Appearance      =   0  'Flat
      Caption         =   "Cos"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   19
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   -2400
      TabIndex        =   24
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton Num 
      Appearance      =   0  'Flat
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4320
      Width           =   495
   End
   Begin VB.CommandButton Num 
      Appearance      =   0  'Flat
      Caption         =   "]"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   18
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Num 
      Appearance      =   0  'Flat
      Caption         =   "["
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   17
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Num 
      Appearance      =   0  'Flat
      Caption         =   "("
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   15
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Clr 
      Appearance      =   0  'Flat
      Caption         =   "CE"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      Picture         =   "Calculator_Main_Form.frx":4ED84
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton Num 
      Appearance      =   0  'Flat
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   14
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton Num 
      Appearance      =   0  'Flat
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   13
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton Num 
      Appearance      =   0  'Flat
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   12
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton Num 
      Appearance      =   0  'Flat
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   11
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton Num 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton Num 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton Num 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   2160
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton Num 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton Num 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton Num 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton Num 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton Num 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton Num 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton Num 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Solve 
      Appearance      =   0  'Flat
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3960
      Picture         =   "Calculator_Main_Form.frx":53206
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox Code 
      Height          =   1815
      Left            =   -6840
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Calculator_Main_Form.frx":5AA88
      Top             =   2760
      Width           =   6135
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2520
      Top             =   4320
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   2520
      Top             =   4320
   End
   Begin VB.Timer Syntax 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2520
      Top             =   4320
   End
   Begin VB.Timer Finn 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   2520
      Top             =   4320
   End
   Begin VB.Timer tAns 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   2520
      Top             =   4320
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"Calculator_Main_Form.frx":5AB05
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1049
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4095
      Left            =   120
      TabIndex        =   29
      Top             =   720
      Width           =   1935
   End
End
Attribute VB_Name = "Calc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  This Calculator was made by Leo Barsukov.
'  I got this idea of making a Copy-Paste Calculator
'  So that you Wouldn't have to be worried of
'  Typing Every Problem in "WinCalc"
'  This Calculator is a Simple thing,
'  not as advanced as the Windows calc, but
'  Very, Very Useful.
'  You have the right to Modify it in ANY WAY
'  For Your self, and if you make a good Update,
'  First, E-Mail it to me, then stick it on
'  Planet-Source-Code.com
'                   If You Like This Project, Please Vote.
'
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'
'
'
'
'    cccccc         a       l           cccccc
'   c      c       a a      l          c      c
'   c             a   a     l          c
'   c            aaaaaaa    l          c
'   c      c    a       a   l      l   c      c
'    cccccc     a       a   llllllll    cccccc
'         __________
'        /_______  /\
'       //______/ / /
'      /__   __  / /
'     //_/__/_/ / /
'    /   /_/   / /
'   /_________/ /
'   \_________\/
'
'
'
'
'
'
'
'
'Declare All Variables
Dim f, fs, Answ() As String, Prob, Nums
Private Sub Syntax_Timer()
Text3.Text = ""
Text2.ZOrder 0
Text2.SetFocus
Text2.SelStart = 0
Text2.SelLength = Len(Text2)
Syntax.Enabled = False
End Sub

Private Sub tAns_Timer()
tAns.Enabled = False
End Sub

Private Sub Text3_GotFocus()
Text2.SelStart = Len(Text2.Text)
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
Text2.ZOrder 0
Text2.SetFocus
End Sub

Private Sub Clr_Click()
On Error Resume Next
Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.DeleteFile("Answer.tmp")
Text2.Text = ""
Text3.Text = ""
Text2.ZOrder 0
Text2.SetFocus
End Sub

Private Sub Solve_Click()
    If Text2.Text = "" Then Exit Sub
           Answ = Split(Text2.Text & " ")
       Text2.Text = ""
    For j = 0 To UBound(Answ())
        Text2.Text = Text2.Text & Answ(j)
    Next
    Scripto = ""
        Scripto = Scripto & "Prob=" & Text2 & Chr(13) & Chr(10) & Code
On Error GoTo Syntax:
        Script.AddCode (Scripto.Text)
    'Change Field
    Text3.ZOrder 0
    Text3.SetFocus
    'Pass the Errors



Text3 = ""

On Error Resume Next
For i = 0 To 2
Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.opentextfile("c:\Answer.tmp")
Text3 = f.readline
'Close opened file
f.Close
Next
Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.DeleteFile("C:\Answer.tmp")

Exit Sub
Syntax:
Text3.ZOrder 0
Text3.Text = "Syntax Error"
Syntax.Enabled = True
End Sub

Private Sub Form_Load()
    'Saving .vbs File with the problem
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.CreateTextFile("c:\Answer.vbs")
        'Writing Code to .vbs File
        f.WriteLine "Dim f, fs, Prob"
        f.WriteLine "On Error Resume Next"
        f.WriteLine "Prob=" & Text2.Text 'Adding the problem to "Code Format"
        f.WriteLine Code.Text 'Adding the "Save .tmp" code
    'Save Finished File
    f.Close
    
    For i = 0 To 9
        Num(i).Caption = i
    Next
    For i = 0 To 22
        Num(i).Picture = Solve.Picture
    Next
    Nums = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", "+", "-", "*", "/", "(", ")", "[", "]", "Cos(", "Sin(", "", "*3.14")
End Sub
Private Sub Num_Click(Index As Integer)
Text2.SetFocus
Text2.ZOrder 0
Text2.SelText = Nums(Index)
If Index = 21 Then
MsgBox "Under Construction", vbExclamation, "Sorry"
End If
End Sub

