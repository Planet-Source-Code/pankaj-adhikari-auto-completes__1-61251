VERSION 5.00
Begin VB.UserControl ExtedListbox 
   BackColor       =   &H80000005&
   ClientHeight    =   2205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2730
   ScaleHeight     =   2205
   ScaleWidth      =   2730
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "ExtedListbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'FIXIT: Use Option Explicit to avoid implicitly creating variables of type Variant         FixIT90210ae-R383-H1984
Dim IsSelect As Boolean
Dim Ind As Integer
Dim Tx As String
Dim press As Boolean
Dim Pres As Integer
'Event Declarations:
Event Click() 'MappingInfo=Text1,Text1,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=Text1,Text1,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=Text1,Text1,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event Change() 'MappingInfo=Text1,Text1,-1,Change
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=Text1,Text1,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event Scroll() 'MappingInfo=List1,List1,-1,Scroll
Attribute Scroll.VB_Description = "Occurs when you reposition the scroll box on a control."

Private Sub List1_Click()
Dim I As Integer
For I = 0 To List1.ListCount - 1
If List1.Selected(I) = True Then
Text1.Text = List1.List(I)
I = List1.ListCount - 1
End If
Next I
List1.Visible = False
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or keyindex = 32 Then
Dim I As Integer
For I = 0 To List1.ListCount - 1
If List1.Selected(I) = True Then
Text1.Text = List1.List(I)
I = List1.ListCount - 1
End If
Next I
List1.Visible = False
End If
End Sub

Private Sub Text1_GotFocus()
UserControl.Height = Text1.Height + List1.Height

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    RaiseEvent KeyDown(KeyCode, Shift)
If KeyCode = 38 Or KeyCode = 40 Then
    If IsSelect = True Then
      If KeyCode = 40 Then
        Dim chx As String
        chx = Text1.Text
        If Ind < List1.ListCount - 1 Then
        List1.Selected(Ind + 1) = True
        Ind = Ind + 1
        If Pres = 0 Or Pres <> Ind Then
    press = True
    Text1.Text = chx
    Text1.SelStart = Len(Text1.Text)
    List1.Visible = True
    End If
        End If
        End If
        '--------------------
        If KeyCode = 38 Then
        chx = Text1.Text
        If Ind > 0 Then
        List1.Selected(Ind - 1) = True
        Ind = Ind - 1
        If Pres = 0 Or Pres <> Ind Then
    press = True
    Text1.Text = chx
    Text1.SelStart = Len(Text1.Text)
    List1.Visible = True
    End If
        End If
        End If
     End If
    Exit Sub
 End If
 If KeyCode = 32 Then
    If IsSelect = True Then
    Text1.Text = List1.List(Ind)
    List1.Visible = False
    Text1.Text = Text1.Text
    Text1.SelStart = Len(Text1.Text)
    End If
  Exit Sub
 End If
 If KeyCode = 13 Then
    KeyCode = 0
    Exit Sub
 End If
  Tx = Text1.Text & Chr(KeyCode)
Checkit Tx
'KeyCode = 0
End Sub
''
Public Sub Apper(Ty As Integer)
If Ty = 1 Then
'FIXIT: Text1.Appearance property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
Text1.Appearance = 1
'FIXIT: List1.Appearance property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
List1.Appearance = 1
Else
'FIXIT: Text1.Appearance property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
Text1.Appearance = 0
'FIXIT: List1.Appearance property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
List1.Appearance = 0
End If
End Sub
Public Sub Txtheight(Hi As Integer)
Text1.Height = Hi
On Error Resume Next
List1.Height = UserControl.Height - Text1.Height
If Error Then Err.Clear
List1.Top = Text1.Top + Text1.Height
End Sub
Sub Checkit(Tmx As String)
Dim chx As String
chx = Text1.Text
Dim L As Integer
Dim C As Integer
Dim Loop1 As Integer
L = Len(Tmx)
C = List1.ListCount
For Loop1 = 0 To C - 1
If Len(List1.List(Loop1)) >= L Then
'FIXIT: Replace 'Mid' function with 'Mid$' function                                        FixIT90210ae-R9757-R1B8ZE
'FIXIT: Replace 'UCase' function with 'UCase$' function                                    FixIT90210ae-R9757-R1B8ZE
'FIXIT: Replace 'UCase' function with 'UCase$' function                                    FixIT90210ae-R9757-R1B8ZE
    If UCase(Tmx) = UCase(Mid(List1.List(Loop1), 1, L)) Then
    List1.Selected(Loop1) = True
    If Pres = 0 Or Pres <> Loop1 Then
    press = True
    Text1.Text = chx
    Text1.SelStart = Len(Text1.Text)
    End If
    List1.Visible = True
    IsSelect = True
    Ind = Loop1
    Exit Sub
    Else
    IsSelect = False
    List1.Visible = False
    End If
 Else
    IsSelect = False
    List1.Visible = False
 End If
 Next Loop1
 End Sub

Private Sub Text1_LostFocus()
List1.Visible = False
UserControl.Height = Text1.Height
End Sub

Private Sub UserControl_Initialize()
List1.Clear
List1.AddItem "Pankaj"

End Sub

Private Sub UserControl_Resize()
Text1.Width = UserControl.Width
List1.Width = UserControl.Width
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,Alignment
Public Property Get Alignment() As Integer
Attribute Alignment.VB_Description = "Returns/sets the alignment of a CheckBox or OptionButton, or a control's text."
    Alignment = Text1.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As Integer)
    Text1.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

Private Sub Text1_Change()
    RaiseEvent Change
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = Text1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Text1.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,FontBold
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets bold font styles."
    FontBold = Text1.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    Text1.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,FontItalic
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets italic font styles."
    FontItalic = Text1.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    Text1.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,FontName
Public Property Get FontName() As String
Attribute FontName.VB_Description = "Specifies the name of the font that appears in each row for the given level."
    FontName = Text1.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    Text1.FontName() = New_FontName
    PropertyChanged "FontName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,FontSize
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Specifies the size (in points) of the font that appears in each row for the given level."
    FontSize = Text1.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    Text1.FontSize() = New_FontSize
    PropertyChanged "FontSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,FontStrikethru
Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_Description = "Returns/sets strikethrough font styles."
    FontStrikethru = Text1.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    Text1.FontStrikethru() = New_FontStrikethru
    PropertyChanged "FontStrikethru"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = Text1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Text1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=List1,List1,-1,AddItem
'FIXIT: Declare 'index' with an early-bound data type                                      FixIT90210ae-R1672-R1B8ZE
Public Sub AddItem(ByVal item As String, Optional ByVal index As Variant)
Attribute AddItem.VB_Description = "Adds an item to a Listbox or ComboBox control or a row to a Grid control."
    List1.AddItem item, index
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=List1,List1,-1,Clear
Public Sub Clear()
Attribute Clear.VB_Description = "Clears the contents of a control or the system Clipboard."
    List1.Clear
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=List1,List1,-1,List
Public Property Get List(ByVal index As Integer) As String
Attribute List.VB_Description = "Returns/sets the items contained in a control's list portion."
    List = List1.List(index)
End Property

Public Property Let List(ByVal index As Integer, ByVal New_List As String)
    List1.List(index) = New_List
    PropertyChanged "List"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=List1,List1,-1,ListCount
Public Property Get ListCount() As Integer
Attribute ListCount.VB_Description = "Returns the number of items in the list portion of a control."
    ListCount = List1.ListCount
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=List1,List1,-1,ListIndex
Public Property Get ListIndex() As Integer
Attribute ListIndex.VB_Description = "Returns/sets the index of the currently selected item in the control."
    ListIndex = List1.ListIndex
End Property

Public Property Let ListIndex(ByVal New_ListIndex As Integer)
    List1.ListIndex() = New_ListIndex
    PropertyChanged "ListIndex"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,MaxLength
Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Returns/sets the maximum number of characters that can be entered in a control."
    MaxLength = Text1.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    Text1.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,MultiLine
Public Property Get MultiLine() As Boolean
Attribute MultiLine.VB_Description = "Returns/sets a value that determines whether a control can accept multiple lines of text."
    MultiLine = Text1.MultiLine
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,PasswordChar
Public Property Get PasswordChar() As String
Attribute PasswordChar.VB_Description = "Returns/sets a value that determines whether characters typed by a user or placeholder characters are displayed in a control."
    PasswordChar = Text1.PasswordChar
End Property

Public Property Let PasswordChar(ByVal New_PasswordChar As String)
    Text1.PasswordChar() = New_PasswordChar
    PropertyChanged "PasswordChar"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=List1,List1,-1,RemoveItem
Public Sub RemoveItem(ByVal index As Integer)
Attribute RemoveItem.VB_Description = "Removes an item from a ListBox or ComboBox control or a row from a Grid control."
    List1.RemoveItem index
End Sub

'FIXIT: List1_Scroll event has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
Private Sub List1_Scroll()
    RaiseEvent Scroll
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,ScrollBars
Public Property Get ScrollBars() As Integer
Attribute ScrollBars.VB_Description = "Returns/sets a value indicating whether an object has vertical or horizontal scroll bars."
'FIXIT: ScrollBars property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
    ScrollBars = Text1.ScrollBars
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=List1,List1,-1,SelCount
Public Property Get SelCount() As Integer
Attribute SelCount.VB_Description = "Returns the number of selected items in a ListBox control."
    SelCount = List1.SelCount
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=List1,List1,-1,Selected
Public Property Get Selected(ByVal index As Integer) As Boolean
Attribute Selected.VB_Description = "Returns/sets the selection status of an item in a control."
    Selected = List1.Selected(index)
End Property

Public Property Let Selected(ByVal index As Integer, ByVal New_Selected As Boolean)
    List1.Selected(index) = New_Selected
    PropertyChanged "Selected"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,SelLength
Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Returns/sets the number of characters selected."
    SelLength = Text1.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    Text1.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,SelStart
Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Returns/sets the starting point of text selected."
    SelStart = Text1.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    Text1.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,SelText
Public Property Get SelText() As String
Attribute SelText.VB_Description = "Returns/sets the string containing the currently selected text."
    SelText = Text1.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    Text1.SelText() = New_SelText
    PropertyChanged "SelText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
    Text = Text1.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    Text1.Text() = New_Text
    PropertyChanged "Text"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,ToolTipText
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
    ToolTipText = Text1.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    Text1.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
'Public Function Checkit(Tmx As String) As Variant

'End Function

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim index As Integer

    Text1.Alignment = PropBag.ReadProperty("Alignment", 0)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Text1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Text1.FontBold = PropBag.ReadProperty("FontBold", 0)
    Text1.FontItalic = PropBag.ReadProperty("FontItalic", 0)
    Text1.FontName = PropBag.ReadProperty("FontName", "")
    Text1.FontSize = PropBag.ReadProperty("FontSize", 0)
    Text1.FontStrikethru = PropBag.ReadProperty("FontStrikethru", 0)
    Text1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
    List1.List(index) = PropBag.ReadProperty("List" & index, "")
    List1.ListIndex = PropBag.ReadProperty("ListIndex", 0)
    Text1.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    Text1.PasswordChar = PropBag.ReadProperty("PasswordChar", "")
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
    List1.Selected(index) = PropBag.ReadProperty("Selected" & index, 0)
    Text1.SelLength = PropBag.ReadProperty("SelLength", 0)
    Text1.SelStart = PropBag.ReadProperty("SelStart", 0)
    Text1.SelText = PropBag.ReadProperty("SelText", "")
    Text1.Text = PropBag.ReadProperty("Text", "")
    Text1.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    Text1.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
'FIXIT: Text1.Appearance property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
    Text1.Appearance = PropBag.ReadProperty("Appearance", 1)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim index As Integer

    Call PropBag.WriteProperty("Alignment", Text1.Alignment, 0)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", Text1.Font, Ambient.Font)
    Call PropBag.WriteProperty("FontBold", Text1.FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", Text1.FontItalic, 0)
    Call PropBag.WriteProperty("FontName", Text1.FontName, "")
    Call PropBag.WriteProperty("FontSize", Text1.FontSize, 0)
    Call PropBag.WriteProperty("FontStrikethru", Text1.FontStrikethru, 0)
    Call PropBag.WriteProperty("ForeColor", Text1.ForeColor, &H80000008)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
    Call PropBag.WriteProperty("List" & index, List1.List(index), "")
    Call PropBag.WriteProperty("ListIndex", List1.ListIndex, 0)
    Call PropBag.WriteProperty("MaxLength", Text1.MaxLength, 0)
    Call PropBag.WriteProperty("PasswordChar", Text1.PasswordChar, "")
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
    Call PropBag.WriteProperty("Selected" & index, List1.Selected(index), 0)
    Call PropBag.WriteProperty("SelLength", Text1.SelLength, 0)
    Call PropBag.WriteProperty("SelStart", Text1.SelStart, 0)
    Call PropBag.WriteProperty("SelText", Text1.SelText, "")
    Call PropBag.WriteProperty("Text", Text1.Text, "")
    Call PropBag.WriteProperty("ToolTipText", Text1.ToolTipText, "")
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H80000005)
    Call PropBag.WriteProperty("BorderStyle", Text1.BorderStyle, 1)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
'FIXIT: Text1.Appearance property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
    Call PropBag.WriteProperty("Appearance", Text1.Appearance, 1)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = Text1.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    Text1.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

Private Sub Text1_Click()
    RaiseEvent Click
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,Appearance
Public Property Get Appearance() As Integer
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
'FIXIT: Appearance property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
'FIXIT: Text1.Appearance property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
    Appearance = Text1.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Integer)
'FIXIT: Text1.Appearance property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
    Text1.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Property Get IsVisible() As Boolean
    IsVisible = List1.Visible
End Property
