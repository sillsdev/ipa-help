VERSION 5.00
Begin VB.Form frmUsngHelp 
   Caption         =   "Using IPA-Help"
   ClientHeight    =   5505
   ClientLeft      =   4065
   ClientTop       =   1710
   ClientWidth     =   7725
   FontTransparent =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "USNGHELP.frx":0000
   LinkTopic       =   "frmUsngHelp"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5505
   ScaleWidth      =   7725
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4905
      Left            =   15
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   45
      Width           =   7665
   End
End
Attribute VB_Name = "frmUsngHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************
'* frmUsngHelp version info:
'*  See basGlobals (Globals.bas).
'**************************************************

Private Sub Form_Load()
  
  Call CenterForm(Me)
  
  '************************************************
  '* Build a string containing the help file text.
  '* Assign that string to the Text property of
  '* the textbox, Text1.
  '************************************************
  Text1a = "MOUSE ACTIONS:" + vbCrLf _
    + "The mouse can be used to navigate, and to play accompanying sounds:" + vbCrLf _
    + vbCrLf _
    + " Actions                            Result" + vbCrLf _
    + " -------------------------------------------------" + vbCrLf _
    + " Click                                 Vocalic" + vbCrLf _
    + " Double Click                     Intervocalic" + vbCrLf
  Text1b = "KEYBOARD SHORTCUTS:" + vbCrLf _
    + "Some key sequences can be used as shortcuts:" + vbCrLf _
    + vbCrLf
  Text1b1 = " Key Sequence                  Result" + vbCrLf _
    + " -------------------------------------------------------" + vbCrLf _
    + " Left Arrow, Up Arrow         Move backward on chart" + vbCrLf _
    + " Right Arrow, Down Arrow  Move forward on chart" + vbCrLf _
    + " Enter                                     Play vocalic sound" + vbCrLf _
    + " Shift + Enter                         Play intervocalic sound" + vbCrLf
  Text1c = "BUTTONS:" + vbCrLf _
    + "The following buttons are available for playing back various items:" + vbCrLf _
    + vbCrLf _
    + " Context                            Available Buttons" + vbCrLf _
    + " ----------------------------------------------------------" + vbCrLf _
    + " Vowel                              Vocalic" + vbCrLf _
    + " Consonant                       Vocalic, Intervocalic" + vbCrLf _
    + " Diacritic: on vowel           Vocalic" + vbCrLf _
    + " Diacritic: on consonant    Vocalic, Intervocalic" + vbCrLf _
    + " Suprasegmental:              Vocalic" + vbCrLf
  Text1d = "DISABLED PLAY BUTTONS:" + vbCrLf _
    + "Some characters on the IPA chart only have a vocalic sound " _
    + "associated with them. When one of these characters is selected, " _
    + "the intervocalic button in the tool bar will appear disabled." + vbCrLf
  Text1e = "DISABLED CHARACTERS:" + vbCrLf _
    + "Three characters on the IPA chart do not have sounds recorded for " _
    + "them yet. These characters appear disabled (dark gray) and can" _
    + "not be selected." + vbCrLf
  Text1f = "PATHS:" + vbCrLf _
    + "The path for the sound files can be set by " _
    + "choosing the Path command from the File " _
    + "menu." + vbCrLf
  Text1g = "COMPARISON:" + vbCrLf _
    + "To compare your pronunciation of an IPA sound with IPA Help's pronunciation," + vbCrLf _
    + " *  Press the Record button (Red Circle)." + vbCrLf _
    + " *  Speak the sound." + vbCrLf _
    + " *  Press the Stop button (Black Square)." + vbCrLf _
    + " *  Press the Play button (Black Triangle) to review the sound you " _
    + "just recorded (Optional)." + vbCrLf _
    + " * Press the Compare button (Black Triangle followed by Speaker Symbol) " _
    + "to compare your sound with IPA Help's sound." + vbCrLf
  Text1h = "TESTING:" + vbCrLf _
    + "To take a computer-generated test of the IPA symbols, " + vbCrLf _
    + " *  Press the 'Test' button or select 'Start' from the Test menu." + vbCrLf _
    + " *  A dialog box will then be displayed." + vbCrLf _
    + " *  Select the desired test group and layout." + vbCrLf _
    + "      'Retest' will be available after a regular test, and " _
    + "will include items missed in the previous test." + vbCrLf _
    + "      'Brief' layout shows the IPA symbols only in a more compact layout." + vbCrLf _
    + "      'IPA Chart' layout will display the symbols on the chart "
  Text1i = "during the test. " + vbCrLf _
    + " *  Press the 'Start' button when ready to begin test." + vbCrLf _
    + " *  There will be a two-second pause, followed by the first sound." + vbCrLf _
    + " *  You will have 5 seconds to click on the symbol that matches " _
    + "the sound." + vbCrLf _
    + " *  If your selection was correct, a smile will appear." + vbCrLf _
    + "    If not, a serious face will appear." + vbCrLf _
    + "    If you do not make a selection within 5 seconds " _
    + "it will be counted as incorrect." + vbCrLf
  Text1j = " *  Press the 'Stop' button at any time to end the test."
  Text1 = Text1a + vbCrLf + Text1b _
    + Text1b1 + vbCrLf + Text1c + vbCrLf + Text1d _
    + vbCrLf + Text1e + vbCrLf + Text1f + vbCrLf _
    + Text1g + vbCrLf + Text1h + Text1i + Text1j
End Sub

Private Sub Form_Resize()
  
  On Error Resume Next
  
  With Text1
    .Move .Left, .Top, ScaleWidth - 30, ScaleHeight - 60
  End With

End Sub

