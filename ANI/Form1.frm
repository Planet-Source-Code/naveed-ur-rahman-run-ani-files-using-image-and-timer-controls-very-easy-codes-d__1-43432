VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Naveed Animation Files (*.ani) Display"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   7785
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "File Informations:"
      Height          =   5250
      Left            =   150
      TabIndex        =   15
      Top             =   1080
      Width           =   4965
      Begin MSComctlLib.ListView ListView2 
         Height          =   1230
         Left            =   165
         TabIndex        =   16
         Top             =   1260
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   2170
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView vlstFrameList 
         Height          =   2205
         Left            =   180
         TabIndex        =   17
         Top             =   2880
         Width           =   4680
         _ExtentX        =   8255
         _ExtentY        =   3889
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Size:"
         Height          =   195
         Left            =   405
         TabIndex        =   23
         Top             =   345
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Frames:"
         Height          =   195
         Left            =   195
         TabIndex        =   22
         Top             =   660
         Width           =   555
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Frames List: (Select items for still preview)"
         Height          =   195
         Left            =   180
         TabIndex        =   21
         Top             =   2640
         Width           =   2910
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Main Blocks:"
         Height          =   195
         Left            =   180
         TabIndex        =   20
         Top             =   1035
         Width           =   915
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Label9"
         Height          =   195
         Left            =   930
         TabIndex        =   19
         Top             =   345
         Width           =   480
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Label10"
         Height          =   195
         Left            =   930
         TabIndex        =   18
         Top             =   660
         Width           =   570
      End
   End
   Begin VB.CommandButton btnRun 
      Cancel          =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   7200
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Tag             =   "&Play"
      Top             =   3090
      UseMaskColor    =   -1  'True
      Width           =   405
   End
   Begin VB.Frame Frame3 
      Caption         =   "Custom Settings:"
      Height          =   1275
      Left            =   5190
      TabIndex        =   3
      Top             =   3750
      Width           =   2415
      Begin VB.CheckBox chkSuperFastSpeed 
         Caption         =   "Super Fast Speed."
         Height          =   285
         Left            =   180
         TabIndex        =   6
         Top             =   870
         Width           =   2130
      End
      Begin VB.CheckBox chkEnhanceColor 
         Caption         =   "Use &Enhance Color."
         Height          =   285
         Left            =   195
         TabIndex        =   4
         Top             =   420
         Value           =   1  'Checked
         Width           =   2130
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Preview:"
      Height          =   1920
      Left            =   5190
      TabIndex        =   2
      Top             =   1080
      Width           =   2415
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Left            =   285
         Top             =   375
      End
      Begin MSComctlLib.ImageList imglstImageHolder 
         Left            =   825
         Top             =   285
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Label3"
         Height          =   225
         Left            =   75
         TabIndex        =   5
         Top             =   225
         Width           =   2175
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   975
         Picture         =   "Form1.frx":038A
         Top             =   720
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load"
      Default         =   -1  'True
      Height          =   510
      Left            =   6585
      TabIndex        =   1
      Top             =   135
      Width           =   1020
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1815
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Text            =   "Windows 98 busy.ani"
      Top             =   240
      Width           =   4635
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Animation &File (*.ani):"
      Height          =   195
      Left            =   210
      TabIndex        =   24
      Top             =   285
      Width           =   1470
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Copyrights(c) 2002-2003 Naveed's Software. All Rights Are Reserved."
      Height          =   195
      Left            =   135
      TabIndex        =   14
      Top             =   6480
      Width           =   4935
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "- Naveed's Software"
      Height          =   195
      Left            =   6075
      TabIndex        =   13
      Top             =   6090
      Width           =   1440
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Dedicated to all VB Programmers"
      Height          =   195
      Left            =   5190
      TabIndex        =   12
      Top             =   5835
      Width           =   2325
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Date: 11 Feb 2003"
      Height          =   195
      Left            =   5190
      TabIndex        =   11
      Top             =   5595
      Width           =   1335
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "By: Naveed ur Rahman"
      Height          =   195
      Left            =   5190
      TabIndex        =   10
      Top             =   5325
      Width           =   1665
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "neenojee@hotmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   5235
      MouseIcon       =   "Form1.frx":2055
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   3210
      Width           =   1665
   End
   Begin VB.Image imgPlay 
      Height          =   240
      Left            =   7350
      Picture         =   "Form1.frx":21A7
      Top             =   3510
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgStop 
      Height          =   240
      Left            =   7095
      Picture         =   "Form1.frx":2531
      Top             =   3510
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "You may use Drag-and-Drop file browsing."
      Height          =   195
      Left            =   1815
      TabIndex        =   7
      Top             =   660
      Width           =   2970
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



'Naveed Animation Files (*.ani) Display:

'   Simple codes to run an Animation File (*.ani)
'   in a Picture/Image box using Timer control.

'   I know it was better to give you people,
'   a CONTROL instead of these complicated codes,
'   but believe me you can convert these codes
'   into an ActiveX Control very easily. This
'   may be a practice for YOU :D

'   I recommend you to read Animation file header
'   before understanding the codes. (see Related Documents)


'---------------------------------------------

'Declaring File System Object
Dim FSYS As New FileSystemObject

'Declaring a LONG variable CurrentFrame
'CurrentFrame holds the number of Current Frame
'of animation and is executed by Timer1()
Dim CurrentFrame As Long

'If you click chkEnhanceColor then...
Private Sub chkEnhanceColor_Click()
MsgBox "Please Load/Reload the file to apply this setting.", vbInformation, Me.Caption
End Sub

'If you click chkSuperFastSpeed then...
Private Sub chkSuperFastSpeed_Click()
MsgBox "Please Load/Reload the file to apply this setting.", vbInformation, Me.Caption
End Sub

'The main procedure i.e. Animation File Load procedure
Private Sub cmdLoad_Click()

'--------------------
'Checking File Existance
If FSYS.FileExists(Text1.Text) = False Then
MsgBox "File not found.", vbCritical, Me.Caption
Exit Sub
End If
'--------------------


Dim ReadFile As String      'Holds the given animation file name
'Assigning Value to ReadFile
ReadFile = Text1.Text


Dim TempFile As String      'Holds a temporary file name
'Assigning Value to TempFile
TempFile = FSYS.BuildPath(FSYS.GetSpecialFolder(TemporaryFolder).Path, App.hInstance)


'The following thre variables are used in this procedure
Dim s As Long, i As Long, j As Long, z As Long

'A very important String used in this procedure
Dim FreeString As String

'--------------------
'Some very important variables used to hold headers/blocks (as String)
Dim HDR As String            'RIFF Block
Dim RATEHDR As String        'RATE Block
Dim LISTHDR As String        'LIST Block
Dim FRAMEHDR As String       'Frame Block(s)
Dim FRAMESTRING As String    'It is the body of frame. Actually the
                             'whole icon hold by the frame.
Dim rHDR As String           'A rough header (for unknown types blocks)
'--------------------


'--------------------
'Some other very necessary variables:
Dim bSize As Long            'Animation Size
Dim tFrames As Long          'Total Number Of Frames
Dim BlockSize As Long        'RIFF Block Size
Dim InfoGap As Long          'Gap (of bytes/characters) when RATE block
                             'is not present and/or is replaced by some
                             'other rough/unknow block.
Dim DelayBytesSize As Long   'RATE block size
Dim IconListingSize As Long  'LIST block size
Dim IconListStart  As Long   'Holds the position (@#nnn) from where the
                             'icon list begins.
Dim IconBinaryStart  As Long 'Holds the position (@#nnn) from where the
                             'icon's binary begins.
Dim FrameSize  As Long       'Holds the size of frame (in bytes)
Dim DelayOfFrame As Long     'The delay of a pirticular frame
Dim rSize As Long            'Rough block size (in bytes)
Dim rStart  As Long          'The position from where the rough
                             'block begins.
'--------------------

Dim ImageAdd As ListImage    'Image List Item (When we add frames/icons in an ImageList)
Dim LItem  As ListItem       'List Item or List Vew Item

vlstFrameList.ListItems.Clear    'Clearing Listview control for the NEW filling
imglstImageHolder.ListImages.Clear      'Clearing ImageList control for the NEW filling

'--------------------
'Binary Processing Starts Here:

BlockSize = 20 'Partial block size to justify file's integrity (type of Animation file)

'Reading header of animation file
HDR = GetString(ReadFile, 1, BlockSize)


'Checking IF THE FILE IS AN ANIMATION FILE ???
If Left(HDR, 4) <> "RIFF" And Mid(HDR, 9, 4) <> "ACON" Then
    
    MsgBox "Either file is not ani ANI formatted file or Format is updated version of ANI files." & vbCrLf & "Can't read this file.", vbInformation, Me.Caption
    Exit Sub
    
End If

If Mid(HDR, 13, 4) = "LIST" Then
    
    InfoGap = GetIntMultiSize(Asc(Mid(HDR, 17, 1)), Asc(Mid(HDR, 18, 1)), Asc(Mid(HDR, 19, 1)), Asc(Mid(HDR, 20, 1)))
    BlockSize = 16 + InfoGap + 4 + 44
    
Else
    
    BlockSize = 12 + 44
    InfoGap = -8
    
End If

'Full RIFF Block Reading:
HDR = GetString(ReadFile, 1, BlockSize)

'--------------------
'RIFF Block:
ListView2.ListItems(1).SubItems(1) = "1"
ListView2.ListItems(1).SubItems(2) = Len(HDR)
'--------------------

'--------------------
'Animation Size:
bSize = GetIntMultiSize(Asc(Mid(HDR, 5, 1)), Asc(Mid(HDR, 6, 1)), Asc(Mid(HDR, 7, 1)), Asc(Mid(HDR, 8, 1)))
Label9.Caption = bSize & " bytes."
'--------------------

'--------------------
'Total Number Of Frames:
tFrames = GetIntMultiSize(Asc(Mid(HDR, InfoGap + 20 + 12 + 1, 1)), Asc(Mid(HDR, InfoGap + 20 + 12 + 2, 1)), Asc(Mid(HDR, InfoGap + 20 + 12 + 3, 1)), Asc(Mid(HDR, InfoGap + 20 + 12 + 4, 1)))
Label10.Caption = tFrames
'--------------------

FreeString = GetString(ReadFile, BlockSize + 1, 8)

'Check is RATE block is present:
If Left(FreeString, 4) <> "rate" Then
    
    'Rate Block Not Exist:
    DelayBytesSize = -8
    RATEHDR = "|"   ' | Character Means No Block (For Codes)
    
    '--------------------
    'RATE Block:
    ListView2.ListItems(2).SubItems(1) = "Not Exist"
    ListView2.ListItems(2).SubItems(2) = "Not Exist"
    '--------------------
    
    GoTo Next_After_Rate
    
End If

'--------------------
DelayBytesSize = GetIntMultiSize(Asc(Mid(FreeString, 5, 1)), Asc(Mid(FreeString, 6, 1)), Asc(Mid(FreeString, 7, 1)), Asc(Mid(FreeString, 8, 1)))
RATEHDR = GetString(ReadFile, BlockSize + 1, 5 + DelayBytesSize + 3)
'--------------------
'RATE Block:
ListView2.ListItems(2).SubItems(1) = BlockSize + 1
ListView2.ListItems(2).SubItems(2) = Len(RATEHDR)
'--------------------

Next_After_Rate:

'--------------------
'Neglecting Rough Headers:
Get_Rough_Blocks:
    rStart = BlockSize + (5 + DelayBytesSize + 3) + 1
    rHDR = GetString(ReadFile, rStart, 12)
    rSize = GetIntMultiSize(Asc(Mid(rHDR, 5, 1)), Asc(Mid(rHDR, 6, 1)), Asc(Mid(rHDR, 7, 1)), Asc(Mid(rHDR, 8, 1)))
    
    If Left(rHDR, 4) <> "LIST" Then 'Rough Header
        
        DelayBytesSize = DelayBytesSize + rSize + 8
        GoTo Get_Rough_Blocks   'Get more rough blocks if present
        
    End If
'--------------------


'--------------------
'LIST Block Reading Begins Here:
IconListStart = BlockSize + (5 + DelayBytesSize + 3) + 1
IconBinaryStart = IconListStart + 12
LISTHDR = GetString(ReadFile, IconListStart, 12)
IconListingSize = GetIntMultiSize(Asc(Mid(LISTHDR, 5, 1)), Asc(Mid(LISTHDR, 6, 1)), Asc(Mid(LISTHDR, 7, 1)), Asc(Mid(LISTHDR, 8, 1)))

'--------------------
'LIST Block:
ListView2.ListItems(3).SubItems(1) = IconListStart
ListView2.ListItems(3).SubItems(2) = IconListingSize
'--------------------

'------------------------------------------------
'Frames List: Saving Frames Into Image List
    
    'You may use some other method to save icons
    'I found it easy, I am using it :D
        
'--------------------
'Asiging initial values to the variables:
i = 0
j = 1
s = 0
'--------------------

'The loop variable "z" holds the number of current frame (which is under process)

For z = 1 To tFrames    'From  to the Total Number of frames

'Frame# Column of vlstFrameList
Set LItem = vlstFrameList.ListItems.Add(, , "# " & z)

'Delay Column of vlstFrameList
If RATEHDR = "|" Then
    
    'No Rate Block Then
    DelayOfFrame = 10
    
Else
    
    'Rate Block Present
    DelayOfFrame = GetIntMultiSize(Asc(Mid(RATEHDR, 8 + s + 1, 1)), Asc(Mid(RATEHDR, 8 + s + 2, 1)), Asc(Mid(RATEHDR, 8 + s + 3, 1)), Asc(Mid(RATEHDR, 8 + s + 4, 1)))
    
End If

'Delay Column of vlstFrameList
LItem.SubItems(1) = DelayOfFrame

'Obtaining Frame Block (Pirticular)
FRAMEHDR = GetString(ReadFile, IconBinaryStart + i, 22 + 4 + 4)

'Frame Binary Size or Size Column of vlstFrameList
FrameSize = GetIntMultiSize(Asc(Mid(FRAMEHDR, 4 + 1, 1)), Asc(Mid(FRAMEHDR, 5 + 1, 1)), Asc(Mid(FRAMEHDR, 6 + 1, 1)), Asc(Mid(FRAMEHDR, 7 + 1, 1)))
LItem.SubItems(2) = FrameSize


'Icon Height and Width Column of vlstFrameList
LItem.SubItems(3) = Asc(Mid(FRAMEHDR, 15, 1)) & " x " & Asc(Mid(FRAMEHDR, 16, 1))

   
    'Frame String Hold Whole The Icon !!!
    FRAMESTRING = GetString(ReadFile, IconBinaryStart + i + 8, FrameSize)
    
        
    'What I am going to do ?
    'Simply, I am writing FRAMESTRING
    '(carrying whole icon) to a temporary file
    'and then loading it into ImageList control
    'assigning delay information (of the frame) to
    'the TAG property of ImageList's Image.
    'Also, I am giving a suitable key name to each
    'image (e.g. "#1#", "#2#"...) for creating
    'an easy animating sequence.
        
    'Again if you don't like this method,
    'you may chage it :D
    
    'Writing An Icon File (Temporary File)
    PutString TempFile, 1, FRAMESTRING

    'Use Enhance Color
    If chkEnhanceColor.Value = 1 Then
    PutString TempFile, 3, Chr$(1)  'Coloring Image (By default it is necessary)
    End If
    
    'Adding/Loading To ImageList
    Set ImageAdd = imglstImageHolder.ListImages.Add(, , LoadPicture(TempFile))
    
    'Saving Frame Delay Period In Image's (ListImage Control) Tag Property
    
    If chkSuperFastSpeed.Value = 1 Then
    
    'Super Fast Speed
    ImageAdd.Tag = 1 / 10

    Else
    
    'Natural Speed (By default)
    ImageAdd.Tag = DelayOfFrame

    End If
    
    'Giving A Suitable Key (For creating easy animating sequence)
    ImageAdd.Key = "#" & ImageAdd.Index & "#"
    
    'Deleting Temporary File
    FSYS.DeleteFile TempFile

'Necessary Increments:
'--------------------
j = j + 1
s = s + 4
i = i + FrameSize + 8
'--------------------

Next z
'------------------------------------------------

btnRun.Enabled = True

'Preview = First Frame Picture
    Image1.Picture = imglstImageHolder.ListImages("#1#").Picture
    Image1.Left = Frame2.Width / 2 - Image1.Width / 2
    Image1.Top = Frame2.Height / 2 - Image1.Height / 2
    Label3.Caption = 1
    CurrentFrame = 1
End Sub

Private Sub btnRun_Click()

If btnRun.Tag = "&Play" Then
    
    'Playing Animation:
    btnRun.Picture = imgStop.Picture
    btnRun.Tag = "&Stop"
    Text1.Locked = True
    cmdLoad.Enabled = False
    Frame3.Enabled = False
    Frame1.Enabled = False
    
    'Animation begins from the first frame:
    CurrentFrame = 1
    Timer1.Interval = Val(imglstImageHolder.ListImages("#" & CurrentFrame & "#").Tag) * 10
    Image1.Picture = imglstImageHolder.ListImages("#" & CurrentFrame & "#").Picture
    Image1.Left = Frame2.Width / 2 - Image1.Width / 2
    Image1.Top = Frame2.Height / 2 - Image1.Height / 2
    Label3.Caption = CurrentFrame
    
    Timer1.Enabled = True

Else
    
    'Stoping Animation
    Timer1.Enabled = False
    Timer1.Interval = 0
    
    btnRun.Picture = imgPlay.Picture
    btnRun.Tag = "&Play"
    Text1.Locked = False
    cmdLoad.Enabled = True
    Frame3.Enabled = True
    Frame1.Enabled = True
    
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
End 'Terminating Program
End Sub

Private Sub Label7_Click()
'Email Address: neenojee@hotmail.com
'Subject: Naveed Animation Files Display

'I know this is not a fair method :)
Shell "start mailto:neenojee@hotmail.com?subject=Naveed%20Animation%20Files%20Display", vbHide

End Sub

Private Sub vlstFrameList_ItemClick(ByVal Item As MSComctlLib.ListItem)
'Still preview of selected frame:
On Error GoTo ErrorReturn

    CurrentFrame = vlstFrameList.SelectedItem.Index
    Image1.Picture = imglstImageHolder.ListImages("#" & CurrentFrame & "#").Picture
    Image1.Left = Frame2.Width / 2 - Image1.Width / 2
    Image1.Top = Frame2.Height / 2 - Image1.Height / 2
    Label3.Caption = CurrentFrame

ErrorReturn:
Exit Sub
End Sub

Private Sub Timer1_Timer()
'Animation By Timer Control!!!
CurrentFrame = CurrentFrame + 1
If CurrentFrame > imglstImageHolder.ListImages.Count Then CurrentFrame = 1
Timer1.Interval = Val(imglstImageHolder.ListImages("#" & CurrentFrame & "#").Tag) * 10
Image1.Picture = imglstImageHolder.ListImages("#" & CurrentFrame & "#").Picture
Image1.Left = Frame2.Width / 2 - Image1.Width / 2
Image1.Top = Frame2.Height / 2 - Image1.Height / 2
Label3.Caption = CurrentFrame
End Sub

Private Sub Form_Load()
vlstFrameList.ColumnHeaders.Add , , "Frame#", vlstFrameList.Width / 5
vlstFrameList.ColumnHeaders.Add , , "Delay", vlstFrameList.Width / 5
vlstFrameList.ColumnHeaders.Add , , "Size", vlstFrameList.Width / 5
vlstFrameList.ColumnHeaders.Add , , "Height x Width", vlstFrameList.Width / 5 * 2

ListView2.ColumnHeaders.Add , , "Block", ListView2.Width / 3
ListView2.ColumnHeaders.Add , , "Start", ListView2.Width / 3
ListView2.ColumnHeaders.Add , , "Size", ListView2.Width / 3
ListView2.ListItems.Add , , "RIFF"
ListView2.ListItems.Add , , "RATE"
ListView2.ListItems.Add , , "LIST"

Advertise   'Naveed IconEX 5.00 (Second Edition - XP-Look)
frmHelpAbout.Show 1
End Sub

'Drag-and-Drop file browsing:
Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Text1.Text = Data.Files(1)
cmdLoad_Click   'Auto-Loading
End Sub

'---------------------------------------------
