VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Error Trapping Generator"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "About"
      Height          =   285
      Left            =   45
      TabIndex        =   12
      Top             =   7560
      Width           =   4650
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Line Counter"
      Height          =   285
      Left            =   45
      TabIndex        =   13
      Top             =   7290
      Width           =   4650
   End
   Begin VB.Frame Frame2 
      Height          =   4335
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4695
      Begin VB.CommandButton Command1 
         Caption         =   "Generate Error Trapping"
         Height          =   315
         Left            =   2010
         TabIndex        =   9
         Top             =   225
         Width           =   2595
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   90
         TabIndex        =   8
         Top             =   225
         Width           =   1875
      End
      Begin VB.DirListBox Dir1 
         Height          =   2790
         Left            =   90
         TabIndex        =   7
         Top             =   645
         Width           =   1875
      End
      Begin VB.FileListBox File1 
         Height          =   2820
         Left            =   2010
         MultiSelect     =   2  'Extended
         Pattern         =   "*.cls;*.frm;*.bas"
         TabIndex        =   6
         Top             =   645
         Width           =   2595
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2025
         TabIndex        =   5
         Text            =   "Markus 360"
         Top             =   3540
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2025
         TabIndex        =   4
         Text            =   "MK360E"
         Top             =   3900
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Projet"
         Height          =   195
         Left            =   1530
         TabIndex        =   11
         Top             =   3585
         Width           =   405
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Code Error"
         Height          =   195
         Left            =   1215
         TabIndex        =   10
         Top             =   3945
         Width           =   750
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   45
      TabIndex        =   0
      Top             =   4365
      Width           =   4650
      Begin VB.CommandButton Command2 
         Caption         =   "List Store Procs SQL"
         Height          =   375
         Left            =   90
         TabIndex        =   2
         Top             =   180
         Width           =   4470
      End
      Begin VB.ListBox List1 
         Height          =   2205
         ItemData        =   "ErrorGenerator.frx":0000
         Left            =   90
         List            =   "ErrorGenerator.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   585
         Width           =   4470
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'AUTHOR : Robert Turcotte
'Description : Error Trapping generator
'Tasks : Genererate Error Tapping where needed on .Frm .Cls .Bas
'        Counts lines in your program (Only what you typed)
'        Generate a SQL stored Procedure list (SQL Only)
'Special Features : Can select more than one file and it will generate/count/list
'                   stored procedures from selected files.
Dim FonctNom As String

Private Sub Command1_Click()
   Dim FonctNom As String
   Dim lastline As String
   Dim tableau(1000, 1000) As Boolean
   Dim nbProcs(1000) As Long
   Dim Trouver As Boolean
   Dim CurrentProc As Long
   Dim TypeFonction As Boolean
   
   For j = 0 To File1.ListCount - 1
      CurrentProc = 0
      If File1.Selected(j) Then
         Open Dir1.Path & "\" & File1.List(j) For Input As #1
         
         Do While Not EOF(1)   ' Loop until end of file.
            Line Input #1, textline   ' Read line into variable.
            
            If Left(textline, 12) = "End Function" Then
                CurrentProc = CurrentProc + 1
            End If
            
            If Left(Trim(textline), 8) = "On Error" Then
               tableau(File1.ListIndex, CurrentProc) = True
            End If
         Loop
         Close #1
         nbProcs(File1.ListIndex) = CurrentProc - 1
      End If
   Next j
   
   For j = 0 To File1.ListCount - 1
      CurrentProc = 0
      If File1.Selected(j) Then
         
         'SET ATTRIBUTES AS NORMAL (DEACTIVATES HIDDEN, READ-ONLY, SYSTEM ATTRIBUTES)
         SetAttr Dir1.Path & "\" & File1.List(j), vbNormal
         Open Dir1.Path & "\" & File1.List(j) For Input As #1
         Open Dir1.Path & "\" & Left(File1.List(j), Len(File1.List(j)) - 3) & "temp" For Output As #2

         lastline = ""
         Do While Not EOF(1)   ' Loop until end of file.
            Line Input #1, textline   ' Read line into variable.

            'DETECTS IF THERES A 'END SUB' || 'END FUNCTION' AND WRITE (IF APPLICABLE) THE ERROR
            'TRAPPING PROCEDURE. FINISH THE SUB || FUNCTION AND GOES ON WITH THE NEXT METHOD
            '(CURRENTPROC++)
            If Left(textline, 12) = "End Function" Or Left(textline, 7) = "End Sub" Then
               If tableau(File1.ListIndex, CurrentProc) = False Then
                  Print #2, ""
                  If TypeFonction = True Then
                     Print #2, "   Exit Function"
                  Else
                     Print #2, "   Exit Sub"
                  End If
                  
                  Print #2, ""
                  Print #2, FonctNom & "_Erreur:"
                  Print #2, ""
                  Print #2, "   MsgBox ""Projet.............: "" & """ & Text1.Text & """ & Chr(13) & _"
                  Print #2, "          ""Classe/Module/Form.: "" & """ & Left(File1.List(j), Len(File1.List(j)) - 4) & """ & Chr(13) & _"
                  Print #2, "          ""Fonction/Sub.......: "" & """ & FonctNom & """ & Chr(13) & _"
                  Print #2, "          ""Erreur.............: "" & """ & Text2.Text & """ & Err.Number & Chr(13) & _"
                  Print #2, "          ""Description........: "" & Err.Description, vbCritical, ""Erreur" & Text2.Text & """"
                  Print #2, ""
               CurrentProc = CurrentProc + 1
               End If
            End If

            'WRITE THE COPY OF THE READED LINE FROM THE ORIGINAL FILE
            Print #2, textline

            'VERIFICATION FONCTION || SUB && PUBLIC || PRIVATE
            If Left(Trim(textline), 16) = "Public Function " Then
               FonctNom = ""
               For i = 17 To Len(textline)
                  If Mid(textline, i, 1) <> "(" Then
                     FonctNom = FonctNom + Mid(textline, i, 1)
                  Else
                     Exit For
                  End If
               Next i
               If tableau(File1.ListIndex, CurrentProc) = False Then
                  Print #2, ""
                  Print #2, "   On Error GoTo " & FonctNom + "_Erreur"
                  Print #2, ""
               End If
               TypeFonction = True
            End If
            If Left(Trim(textline), 17) = "Private Function " Then
               FonctNom = ""
               For i = 18 To Len(textline)
                  If Mid(textline, i, 1) <> "(" Then
                     FonctNom = FonctNom + Mid(textline, i, 1)
                  Else
                     Exit For
                  End If
               Next i
               If tableau(File1.ListIndex, CurrentProc) = False Then
                  Print #2, ""
                  Print #2, "   On Error GoTo " & FonctNom + "_Erreur"
                  Print #2, ""
               End If
               TypeFonction = True
            End If
            If Left(Trim(textline), 12) = "Private Sub " Then
               FonctNom = ""
               For i = 13 To Len(textline)
                  If Mid(textline, i, 1) <> "(" Then
                     FonctNom = FonctNom + Mid(textline, i, 1)
                  Else
                     Exit For
                  End If
               Next i
               If tableau(File1.ListIndex, CurrentProc) = False Then
                  Print #2, ""
                  Print #2, "   On Error GoTo " & FonctNom + "_Erreur"
                  Print #2, ""
               End If
               TypeFonction = False
            End If
            If Left(Trim(textline), 11) = "Public Sub " Then
               FonctNom = ""
               For i = 12 To Len(textline)
                  If Mid(textline, i, 1) <> "(" Then
                     FonctNom = FonctNom + Mid(textline, i, 1)
                  Else
                     Exit For
                  End If
               Next i
               If tableau(File1.ListIndex, CurrentProc) = False Then
                  Print #2, ""
                  Print #2, "   On Error GoTo " & FonctNom + "_Erreur"
                  Print #2, ""
               End If
               TypeFonction = False
            End If
         Loop

         Close #1
         Close #2

         'DELETE THE ORIGINAL FILE
         SetAttr Dir1.Path & "\" & File1.List(j), vbNormal
         Kill (Dir1.Path & "\" & File1.List(j))
         Dim OldName, NewName
         'PREPARE TO RENAME THE FILES
         OldName = Dir1.Path & "\" & Left(File1.List(j), Len(File1.List(j)) - 3) & "temp"
         NewName = Dir1.Path & "\" & File1.List(j)
         'RENAMING THE BATCH FILE AS THE ORIGINAL
         Name OldName As NewName
      End If
   Next j

   MsgBox ("Done")
End Sub

Private Sub Command2_Click()

   Dim Trouver As Boolean
   
   List1.Clear
   
   'LIST ALL STORED PROCEDURES (SQL ONLY)
   For j = 0 To File1.ListCount - 1
      If File1.Selected(j) Then
         Open Dir1.Path & "\" & File1.List(j) For Input As #1
         
         Do While Not EOF(1)   ' Loop until end of file.
            Trouver = False
            Line Input #1, textline   ' Read line into variable.
            For i = 1 To Len(textline) - 20
               If Mid(textline, i, 15) = ".CommandText = " And Trouver = False Then
                  storeproc = Mid(textline, i + 16, Len(textline) - i - 16)
                  Trouver = True
               End If
            Next i
            If Trouver = True Then
               List1.AddItem (storeproc)
            End If
         Loop
         Close #1
      End If
   Next j


   Open Dir1.Path & "\StoreProcs.txt" For Output As #2
      For i = 0 To List1.ListCount - 1
         Print #2, List1.List(i)
      Next i
   Close #2

End Sub

Private Sub Command3_Click()
   'ABOUT
   Form2.Show vbModal
End Sub

Private Sub Command4_Click()
   
   Dim blnCount As Boolean
   Dim ligne As Long
   
   ligne = 0
   For j = 0 To File1.ListCount - 1
      blnCount = False
      If File1.Selected(j) Then
         Open Dir1.Path & "\" & File1.List(j) For Input As #1

         Do While Not EOF(1)   ' Loop until end of file.
            Line Input #1, textline   ' Read line into variable.

            'VERIFICATION FONCTION || SUB && PUBLIC || PRIVATE
            If Left(textline, 16) = "Public Function " Or Left(textline, 17) = "Private Function " Or Left(textline, 12) = "Private Sub " Or Left(textline, 11) = "Public Sub " Then
               blnCount = True
               ligne = ligne + 1
            End If
                        
            If blnCount = True Then
               ligne = ligne + 1
            End If
                        
            If Left(textline, 12) = "End Function" Or Left(textline, 7) = "End Sub" Then
               blnCount = False
               ligne = ligne + 1
            End If


         Loop

         Close #1
      End If
   Next j

   MsgBox ("Number of lines : " & ligne)
End Sub

Private Sub Dir1_Change()
   
   'WHEN CHANGING DIRECTORY
   File1.Path = Dir1.Path

End Sub

Private Sub Drive1_Change()
   
   'WHEN CHANGING DRIVE
   Dir1.Path = Drive1.Drive

End Sub

