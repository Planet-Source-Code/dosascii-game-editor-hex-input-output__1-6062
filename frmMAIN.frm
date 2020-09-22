VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmMAIN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FUBAR Editor"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   7020
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2119
      Left            =   240
      TabIndex        =   5
      Top             =   468
      Width           =   5161
      Begin VB.PictureBox Picture1 
         Height          =   735
         Left            =   3000
         ScaleHeight     =   675
         ScaleWidth      =   1995
         TabIndex        =   16
         Top             =   240
         Width           =   2055
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "Put a little piccy here maybe."
            Height          =   495
            Left            =   120
            TabIndex        =   17
            Top             =   120
            Width           =   1695
         End
      End
      Begin VB.Label Label7 
         Caption         =   "http://www.frtrk.quick.com.au"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1800
         Width           =   4335
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   4560
         Picture         =   "frmMAIN.frx":0000
         Top             =   1080
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label8 
         Caption         =   "Put email address and URL on this ""TAB""."
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1200
         Width           =   4335
      End
      Begin VB.Label Label2 
         Caption         =   "flame@baz.com.au"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1560
         Width           =   4215
      End
      Begin VB.Line Line6 
         BorderColor     =   &H80000005&
         BorderStyle     =   6  'Inside Solid
         X1              =   117
         X2              =   5031
         Y1              =   1053
         Y2              =   1053
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         X1              =   117
         X2              =   5031
         Y1              =   1053
         Y2              =   1053
      End
      Begin VB.Label Label1 
         Caption         =   "Put Game Title here....etc"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   7
         Top             =   705
         Width           =   2745
      End
      Begin VB.Label lblTITLE 
         Caption         =   "FUBAR Editor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2700
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2119
      Left            =   234
      TabIndex        =   12
      Top             =   468
      Width           =   5161
      Begin VB.CommandButton cmd4 
         Caption         =   "&Help"
         Height          =   375
         Left            =   1920
         TabIndex        =   22
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Text            =   "frmMAIN.frx":0442
         Top             =   720
         Width           =   4935
      End
      Begin VB.Label lblAV 
         Caption         =   "App Version"
         Height          =   247
         Left            =   117
         TabIndex        =   14
         Top             =   468
         Width           =   4927
      End
      Begin VB.Label lblAT 
         Caption         =   "App title"
         Height          =   247
         Left            =   117
         TabIndex        =   13
         Top             =   234
         Width           =   4927
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2119
      Left            =   234
      TabIndex        =   8
      Top             =   480
      Width           =   5161
      Begin ComCtl2.UpDown UpDown1 
         Height          =   312
         Left            =   2452
         TabIndex        =   27
         Top             =   1080
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   327681
         BuddyControl    =   "txtCASH"
         BuddyDispid     =   196629
         OrigLeft        =   2400
         OrigTop         =   1080
         OrigRight       =   2640
         OrigBottom      =   1455
         Max             =   32767
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtNAME 
         Height          =   312
         Left            =   936
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   600
         Width           =   1755
      End
      Begin VB.CommandButton cmd6 
         Caption         =   "&Restore"
         Height          =   375
         Left            =   3720
         TabIndex        =   24
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton cmd5 
         Caption         =   "Ma&x Credits"
         Height          =   375
         Left            =   3720
         TabIndex        =   23
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtCASH 
         Height          =   312
         Left            =   936
         MaxLength       =   5
         TabIndex        =   9
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label Label9 
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Maybe put some notes here...."
         Height          =   364
         Left            =   117
         TabIndex        =   15
         Top             =   1638
         Width           =   4927
      End
      Begin VB.Line Line8 
         BorderColor     =   &H80000005&
         BorderStyle     =   6  'Inside Solid
         X1              =   117
         X2              =   5031
         Y1              =   1521
         Y2              =   1521
      End
      Begin VB.Line Line7 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         X1              =   117
         X2              =   5031
         Y1              =   1521
         Y2              =   1521
      End
      Begin VB.Label Label4 
         Caption         =   "Credits:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   11
         Top             =   1100
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Enter in the amount of credits you want in the game."
         Height          =   247
         Left            =   117
         TabIndex        =   10
         Top             =   234
         Width           =   4927
      End
   End
   Begin VB.Timer tmrEG 
      Interval        =   1
      Left            =   5520
      Top             =   2160
   End
   Begin VB.Timer tmrVAL 
      Interval        =   1
      Left            =   6000
      Top             =   2160
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "&Quit"
      Height          =   364
      Left            =   5640
      TabIndex        =   4
      Tag             =   "Close the Edior. This will not save changes."
      Top             =   1750
      Width           =   1300
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "&Save"
      Height          =   364
      Left            =   5640
      TabIndex        =   3
      Tag             =   "Save your changes to the games."
      Top             =   1175
      Width           =   1300
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "&Open"
      Height          =   364
      Left            =   5640
      TabIndex        =   2
      Tag             =   "Open and locate the Saved Game file, ready for editing."
      Top             =   585
      Width           =   1300
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6480
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.TabStrip ts1 
      Height          =   2587
      Left            =   117
      TabIndex        =   1
      Top             =   117
      Width           =   5395
      _ExtentX        =   9499
      _ExtentY        =   4551
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Main"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Game"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&About"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sb1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   0
      Top             =   2865
      Width           =   7020
      _ExtentX        =   12383
      _ExtentY        =   423
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3466
            Text            =   "File Size: (No File Opened)"
            TextSave        =   "File Size: (No File Opened)"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8837
         EndProperty
      EndProperty
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000005&
      BorderStyle     =   6  'Inside Solid
      X1              =   5733
      X2              =   6786
      Y1              =   1638
      Y2              =   1638
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000003&
      BorderWidth     =   2
      X1              =   5733
      X2              =   6786
      Y1              =   1638
      Y2              =   1638
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      BorderStyle     =   6  'Inside Solid
      X1              =   6786
      X2              =   5733
      Y1              =   1053
      Y2              =   1053
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      BorderWidth     =   2
      X1              =   5733
      X2              =   6786
      Y1              =   1053
      Y2              =   1053
   End
End
Attribute VB_Name = "frmMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'IT MIGHT EASYIER TO VIEW THIS TEXT AT A DISPLAY HIGHER RES THEN 800X600
'
'Created by DosAscii. dosascii@hotmail.com
'Ok this is properly the best VB code I've written in
'along time, so I hope you enjoy it.
'If you use my code in your apps, please credit
'me.
'
'
'DosAscii dosascii@hotmail.com
'
'
'IMPORTANT!!!!
'When you release your VB code to the public domain,
'remember to do the following.
'1: Delete all comments and WhiteSpace (unnesscary/All Tabs and spaces).
'2: Use small named varibles. No longer then 8 Chars.
'3: Use image boxs when ever possible
'4: Don't use too many Global vars....
'5: Use functions.
'This will make EXE a little smaller...
'
'The Status bar at the bottom of the app is for the ToolTips,
'the text for each control is kept in the TAG of that control.
'The TAG is then displayed in the sb1 (Statusbar).
'
'The scorllable text under "About" is a text box.
'With Multiline true, Vertical Scroll Bars true, Locked
'true. When in design time; Use CTRL+ENTER to type on a newline.
'
'USES:
'MS COMMON DIALOG CONTROL 6.0
'MS WINDOWS COMMON CONTROLS 6.0
'MS WINDOWS COMMON CONTROLS 2.5.0(SP2)
'
'OK for this code we will operate on the file I included
'(FILE.BAZ), I suggest you look at the file with a hex editor.
'Then run the app, and see the changes....We'll pretend this
'file contains a games saved info, pretend that we are
'modifying the amount of cash in the game....
'======================================================================
Public OpFLG As Boolean 'Ok this will be our flag to see if a file is opened.
'Ths Boolean will only hold a True or False state.


Private Sub cmd1_Click()
'OPEN CLICK.
'Some varibles...
Dim PlayerName As String
Dim FileSize As String
Dim Contents As String 'Contains the contents of the Cash txtbox.
Dim CV As Integer
Dim Result As String 'Final String to Hex Conversion.
Dim DECVal 'Decimal value of Hex.
Dim DUMMY
Dim Lgh 'Lenght
Dim Lft 'Left
Dim Rht 'Right

'This will call the CommonDialog control, ---(cd1)
'to display and Open Dialog.
'Lets trap the errors..
On Error GoTo ErrorOpen
'Close any preivously opened files
Close #1

'Lets start...
cd1.InitDir = "C:\" 'Lets just set the inital "starting"
'directoy to C:\ DRIVE.
cd1.Filter = "Cool Game File(*.BAZ)|*.BAZ" ' OK here
'we show the user what they can open.
cd1.FilterIndex = 1 'Set the order of the above filter.
'Lots of flags....just to make it a tighter on the user.
cd1.Flags = cdlOFNFileMustExist Or cdlOFNPathMustExist Or cdlOFNHideReadOnly Or cdlOFNExtensionDifferent Or cdlOFNNoReadOnlyReturn
'Tell cd1 to display an "Open Dialog".
cd1.ShowOpen
'Ok lets get and set the filename to our app.
FileName = cd1.FileName
'Open the filename for the access we need. We'll use
'binary to directly read and write to the file.
Open (FileName) For Binary Access Read Write As #1
OpFLG = True

'Just a little trick to make your app look more
'pro.
'Get the file lenght.
FileSize = LOF(1)
EOF (1)
'Now display it in the Status bar.
sb1.Panels(1).Text = "File Size: " & FileSize & " Bytes" & " "
'Ok lets get the player's name, and display it in the text box.
'Yeah..it's all pretty simple here.
Seek #1, 3
PlayerName = Input(10, 1)
txtNAME.Text = PlayerName

'Now lets get the data we need and display it the
'text box "txtCASH".
'First of, lets set the position in the file, from,
'were we'er going to start read..
'NOTE!!!
'VB file access is indexed starting from 1.
'Stupid VB.
Seek #1, 69 'Decimal position in the BAZ file.
'That number was a pure coincidence!

'Now we get the input.
txtCASH.Text = Input(2, 1) 'The 2 comes from how many bytes
'we want, and the 1 is the for the file(#1).

'Ok lets convert the inputted string to hex and
'then Decimal.
'Set the var to contain the contents of
'the text box.
Contents = txtCASH.Text
'Perform the conversion.
For CV = 1 To Len(Contents)
If Len(Hex(Asc(Mid(Contents, CV, 1)))) > 1 Then
Result = Result & Hex(Asc(Mid(Contents, CV, 1)))
End If
Next
'Before we convert the Hex, we need to swap
'the "endian" eg:
'10000 = 2710h
'we want 1027h.
'NOTE!
'Some files will deliberatly store there data in swap(little endian) order or in
'big endian, or maybe both. The Intel way is the little-endian way.
Lgh = Len(Result) 'Get the lenght.
    If Lgh = 4 Then
        Lft = Left(Result, 2) 'Get the left most of the var
        Rht = Right(Result, 2) 'Get the right most of the var
        Result = Rht & Lft 'Set the two togther, note the
        'order in which there put togther.
    End If
            
'Now we convert the hex into Decimal.
'txtCASH.Text = Result
DUMMY = Result
'Convert.
DECVal = CInt("&H" & DUMMY)
'Display conversion.
txtCASH.Text = DECVal
'Set scroll bar value.
UpDown1.Value = DECVal

'Here's what happenes if there is a error.
'It'l just simply drop.
'This is manily here for when, the user
'hits cancel on the Open dialog.
ErrorOpen:
    Exit Sub 'This will abort the cd1 procedure.
End Sub

Private Sub cmd3_Click()
'The form UnLoad code.
'Very simple
Close #1 'Close the file.
Unload Me 'Unload app.
End Sub

Private Sub cmd4_Click()
'This will display a little Display box.
MsgBox ("Put ""HELP"" link here")
End Sub

Private Sub cmd5_Click()
'Pretty simple here.
txtCASH.Text = 32767
End Sub

Private Sub cmd6_Click()
'Ohh boy this is getting hard...
'Restore to factory deflaut. (maybe put that in the
'tooptiptext).
txtCASH.Text = 10000
End Sub

Private Sub Form_Load()
'This code will simple change the text of
'the labels in the "About Tab".
'This will show App Title.
lblAT.Caption = App.Title
'This will show the App Version.
lblAV.Caption = "Version: " & App.Major & "." & App.Minor & "." & App.Revision
'Bring the main frame to the front.
'I put that there, cos I got sick of having to rearrange the
'frames at design time, instead I could just leave everything the
'way it was.
Frame1.ZOrder 0
End Sub

Private Sub Image1_Click()
'Here's what happens when the user clicks on the
'image.
'Display a form.
frmBOO.Show vbModal, Me
End Sub

'OK lets put a little sortof a "Easter Egg" in here
'For simplistic reason, we'll use a seprate timer.
Private Sub tmrEG_Timer()
Dim Check
Check = txtCASH.Text
If Check = "3546" Then 'When the text box has this value
'in it, the Easter Egg it true.
Image1.Visible = True
End If
End Sub

Private Sub tmrVAL_Timer()
'This timer is used to stop the user from
'typing a too higher number into the text box.
If Val(txtCASH.Text) > 32767 Then txtCASH.Text = "32767"
'Now because VB is Dirty, they lets us
'only have one type of a 2 byte int, and this
'is: -32768 to 32767 not 0 to 65535
'Now in the "game" you can really have 0xFFFF or
'65535.....
End Sub

Private Sub ts1_Click()
'Code for the TabStrip.
'A simple IF..THEN Statement,
'to bring the frames to the front.
'ZOrder is used to bring the frame to the front.
'First Tab Button.
If ts1.SelectedItem.Index = 1 Then
    Frame1.ZOrder 0 'Bring the frame to the front.
End If
'Second Tab Button.
If ts1.SelectedItem.Index = 2 Then
    Frame2.ZOrder 0
End If
'Third Tab Button.
If ts1.SelectedItem.Index = 3 Then
    Frame3.ZOrder 0
End If
End Sub

Private Sub txtCASH_Keypress(KeyAscii As Integer)
'OK this code will only allow the user to enter
'in numeric data only, plus accepting the Backspace
'key, for what it is.
If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
    If KeyAscii <> vbKeyBack Then
        KeyAscii = 0   ' Cancel the character.
        Beep           ' Sound error signal.
    End If
End If
End Sub

Private Sub cmd2_Click()
'Saving.
'Check to see if a file is open.
If OpFLG = False Then
'if False  then exit.
'Simple Ehh.
Exit Sub
End If
'Else is assumed.

Dim DUMMY
Dim HexVal  As String * 4 'Converted Decimal to Hex.
Dim HexWRT  As Integer 'Converted Hex string to integer
Dim Lghz    'Lenght
Dim Lftz    'Left
Dim Rhtz    'Right
DUMMY = txtCASH.Text 'Get the user input.
'Convert the Decimal.
HexVal = Hex(DUMMY)
'Now we have to swap the data again.
Lghz = Len(HexVal) 'Get the lenght.
    'Perform the swaps and evenings.
    If Lghz = 4 Then
        Lftz = Left(HexVal, 2) 'Get the left most of the var.
        Rhtz = Right(HexVal, 2) 'Get the right most of the var.
        HexVal = Lftz & Rhtz 'Swap. And add to make an even num.
    End If

    If Lghz = 3 Then
        Lftz = Left(HexVal, 2) 'Get the left most of the var.
        Rhtz = Right(HexVal, 1) 'Get the right most of the var.
        HexVal = Rhtz & "0" & Lftz 'Swap. And add to make an even num.
    End If
    'Yeah I think you know pretty well what happens here.
    If Lghz = 2 Then
        Lftz = Left(HexVal, 2)
        HexVal = "00" & Lftz
    End If
    'Owwwww this one's a real hardy....
    If Lghz = 1 Then
        Lftz = Left(HexVal, 2)
        HexVal = "000" & Lftz
    End If
    
'Convert the var into a integer
HexWRT = CInt("&H" & HexVal)
Seek #1, 69
Put #1, , HexWRT
End Sub

'Now we get started on the tooptips.
'We will only do three here, just for the command
'buttons, along the side of the form, plus the form it self.
Private Sub cmd1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Open.
sb1.Panels(2).Text = cmd1.Tag
End Sub

Private Sub cmd2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Save.
sb1.Panels(2).Text = cmd2.Tag
End Sub

Private Sub cmd3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Exit.
sb1.Panels(2).Text = cmd3.Tag
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'The Form
'This is here, cos if the user moves the cursor to fast,
'over the form, then the status bar will still have its
'previous contents....whick looks a little stupid.
sb1.Panels(2).Text = ""
End Sub

'======================================================================
'YET ANOTHER NOTE!!!!
'If you want to increas the maximum amount of cash
'written to a file, you could use a long, and set the
'position 2 bytes before the offset you would write
'to if you were using integer. EG:

'Ok say this is the file:
'OFFSET     DATA
'00000000   F7A7 3F37 8086 48FA
'           0102 0304 0506 0708
'           positions...
'Ok and you what to write over 8086, but you what to write FFFF,
'well if you use a Long, then you would have write at 03
'and the varible you would write would look this.
'var = 3F37<converted decimal>
'I hope this helps....
'whistle if you need my H-E-L-P
'....well, you know how it goes.
'
'
'Remember, DELETE all "whitespace".....This crap I mean.
'
'
'
'The END!
'
'
'Oh yeah, if you make something out of this, let me now, or send it
'to me.
'
'If someone what's to make a series of editors and wants help,
'drop me a line, I started work on a few along time ago.....but
'I never finsihed them...
'They were:
'OMF 2097 saved game editor.
'Lords of The Realm II.
'Constructor.
'I was going to have a really big one on Betrayal at Krondor......
'
'Ha! Have fun and make something really cool.
'
'
'UPDATE:
'Added an UpDown Control, and set it to txtCASH.
'Slect the UpDown Control and click "Custom" in the Propertise
'panel, and have a play with it.
