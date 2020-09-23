VERSION 5.00
Begin VB.Form frmCreateControl 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dynamic Creation Control"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5955
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Designer"
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   5775
      Begin VB.CommandButton cmdApply 
         Caption         =   "Apply"
         Height          =   375
         Left            =   3840
         TabIndex        =   6
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtQuantity 
         Height          =   285
         Left            =   5040
         TabIndex        =   3
         Text            =   "5"
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox cmbObjectType 
         Height          =   315
         ItemData        =   "AddControls.frx":0000
         Left            =   1440
         List            =   "AddControls.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Quantity"
         Height          =   255
         Left            =   4080
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Object 
         Caption         =   "Control Type"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Result"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   5775
   End
End
Attribute VB_Name = "frmCreateControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Date :   January 09 2002
' Author:  Omar Vivas
'          MCP, MCSD, MCDBA, MCSE
' Version: 1.0
' Subject: Dynamic Creation Control
' Note:    It´s only to show the way to create control dynamic
'          this code not is total optimize
'          maybe I do better

Dim sTokens As String

Private Sub cmdApply_Click()
    sTokens = ""
    For i = 1 To txtQuantity.Text
        sTokens = sTokens & "," & Chr$(Asc("A") + i - 1)
    Next
    sTokens = Mid(sTokens, 2)

    RemoveControls
    
    CreateControls 600, 180, sTokens
End Sub

Private Sub Form_Load()
    cmbObjectType.AddItem "VB.TextBox"
    cmbObjectType.AddItem "VB.CheckBox"
    cmbObjectType.AddItem "VB.CommandButton"
    cmbObjectType.AddItem "VB.ListBox"
    cmbObjectType.ListIndex = 0
    cmdApply_Click
End Sub

'   Create Control
'   This routine is reusable to do that on your program
'   Useful, when you need create control and don´t know how many
'   element you have.

Sub CreateControls(ByVal TopIni As Integer, ByVal LeftIni As Integer, ByVal sTokens As String)
Dim i As Integer
Dim Left As Integer
Dim Top As Integer
Dim Width As Integer
Dim Height As Integer
Dim MaxXLinea As Integer
Dim Sep As Integer
Dim Objeto

    Set Objeto = frmCreateControl.Frame1
    TipoObjecto = cmbObjectType.Text
    Width = 1000
    Height = 200
    Sep = 100
    Item = Parsear(sTokens, ",")
    While Item <> ""
        
        Set chkArray = Controls.Add(TipoObjecto, "chk" & Item, frmCreateControl.Frame1)
        
        MaxXLinea = Int((Objeto.Width - LeftIni) / Width)
        
        iConv = i Mod MaxXLinea
        
        Left = ((Width + Sep) * iConv) + LeftIni
        Top = ((Height + Sep) * (i \ MaxXLinea)) + TopIni

        chkArray.Move Left, Top, Width, Height
        chkArray.Visible = True
        Item = Parsear(sTokens, ",")
        i = i + 1
    Wend
End Sub

' Separe the tokens
Function Parsear(ByRef sTokens As String, ByVal Sep As String) As String
Dim Pos As Integer
Dim lTamCad As Integer

    lTamCad = Len(Sep)
    Pos = InStr(sTokens, Sep)
    If Pos = 0 Then
        Pos = Len(sTokens) + 1
    End If
    Parsear = Mid(sTokens, 1, Pos - lTamCad)
    sTokens = Mid(sTokens, Pos + lTamCad)
End Function

' Remove the Control created dynamic
Sub RemoveControls()
    If Controls.Count > 7 Then
        For i = Controls.Count - 1 To 7 Step -1
            Controls.Remove i
        Next
    End If
End Sub
