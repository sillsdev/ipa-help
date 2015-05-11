VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsXMLWordList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const XMLProcInstruction = "version='1.0'" ' encoding='iso-8859-1'"

Private Const TagRootElement = "IPAWordListDef"
Private Const TagGeneralInfo = "GeneralInfo"
Private Const TagPhoneticFont = "PhoneticFont"
Private Const TagOrthoFont = "OrthoFont"
Private Const TagGlossFont = "GlossFont"
Private Const TagDialectFont = "DialectFont"
Private Const TagFontName = "Name"
Private Const TagFontSize = "Size"
Private Const TagFontBold = "Bold"
Private Const TagFontItalic = "Italic"
Private Const TagSoundPath = "SoundPath"
Private Const TagGraphicPath = "GraphicPath"
Private Const TagAllCategories = "Categories"
Private Const TagOneCategory = "Category"
Private Const TagWord = "Word"
Private Const TagPitch = "Pitch"
Private Const TagPhonetic = "Phonetic"
Private Const TagOrtho = "Orthographic"
Private Const TagGloss = "Gloss"
Private Const TagDialect = "Dialect"
Private Const TagSoundFile = "SoundFile"
Private Const TagGraphicFile = "GraphicFile"

Enum WordRecord
  Pitch = 0
  phonetic = 1
  Ortho = 2
  Gloss = 3
  Dialect = 4
  SoundFile = 5
  GraphicFile = 6
End Enum

Private sXMLFilePath As String
Private bInitFailure As Boolean
Private xmlWLDoc As DOMDocument
Private RootNode As IXMLDOMElement
Private node As IXMLDOMNode
Private NodeList As IXMLDOMNodeList

Public Sub AddCategory(sCategoryName$, Optional sWords, Optional sCategoryFollowing)

  Dim i As Integer
  Dim node2 As IXMLDOMNode
  Dim node3 As IXMLDOMNode
  Dim node4 As IXMLDOMNode
  
  '*********************************************************
  '* If category already exists then don't add a new node
  '* for it. Just point to it and add the words to it.
  '*********************************************************
  Set node = GetCategoryNode(sCategoryName)
  If (node Is Nothing) Then
    Set node = xmlWLDoc.createElement(TagOneCategory)
    node.Text = sCategoryName
  End If
  
  '*********************************************************
  '* Don't loop through words if there aren't any specified.
  '*********************************************************
  If Not (IsMissing(sWords)) Then
    For i = 0 To UBound(sWords, 1)
      Set node2 = xmlWLDoc.createElement(TagWord)
      node.appendChild node2
    
      Set node3 = xmlWLDoc.createElement(TagPitch)
      node3.Text = sWords(i, Pitch)
      node2.appendChild node3
    
      Set node3 = xmlWLDoc.createElement(TagPhonetic)
      node3.Text = sWords(i, phonetic)
      node2.appendChild node3

      Set node3 = xmlWLDoc.createElement(TagOrtho)
      node3.Text = sWords(i, Ortho)
      node2.appendChild node3
    
      Set node3 = xmlWLDoc.createElement(TagGloss)
      node3.Text = sWords(i, Gloss)
      node2.appendChild node3

      Set node3 = xmlWLDoc.createElement(TagDialect)
      node3.Text = sWords(i, Dialect)
      node2.appendChild node3

      Set node3 = xmlWLDoc.createElement(TagSoundFile)
      node3.Text = sWords(i, SoundFile)
      node2.appendChild node3
  
      Set node3 = xmlWLDoc.createElement(TagGraphicFile)
      node3.Text = sWords(i, GraphicFile)
      node2.appendChild node3
    Next
  End If

  '*********************************************************
  '* Now point to the category's node and it's parent.
  '*********************************************************
  Set node2 = xmlWLDoc.selectSingleNode("//" & TagAllCategories)
  Set node3 = GetCategoryNode(sCategoryName)
  
  '*********************************************************
  '* If pointing to the category was unsuccessful, it means
  '* the category is new. Otherwise, replace the category
  '* node and all it's children.
  '*********************************************************
  If (node3 Is Nothing) Then
    '*******************************************************
    '* If caller did not specify a category before which
    '* the added list should be inserted, then add the list
    '* to the end of all the categories. Otherwise, try to
    '* insert the added list before the one specified.
    '*******************************************************
    If (IsMissing(sCategoryFollowing)) Then
      node2.appendChild node
    Else
      Set node4 = GetCategoryNode(CStr(sCategoryFollowing))
      node2.insertBefore node, IIf(node4 Is Nothing, Nothing, node4)
    End If
  Else
    node2.replaceChild node, node3
  End If
 
End Sub

Public Property Get CategoryNames() As String()

  Dim i As Integer
  Dim sCategories() As String
  Dim nl As IXMLDOMNodeList
  
  Set NodeList = xmlWLDoc.selectNodes("//" & TagAllCategories & "/" & TagOneCategory)
  
  If (NodeList.length > 0) Then
    ReDim sCategories(0 To NodeList.length - 1)
    For i = 0 To NodeList.length - 1
      sCategories(i) = NodeList.item(i).childNodes.item(0).Text
    Next
  End If
  
  CategoryNames = sCategories
 
End Property

Public Function ChangeCategoryName(sOldName$, sNewName$) As Boolean
  
  ChangeCategoryName = False
  Set node = GetCategoryNode(sOldName)
  If (node Is Nothing) Then Exit Function
  If (node.childNodes(0) Is Nothing) Then Exit Function
  node.childNodes(0).Text = sNewName
  ChangeCategoryName = True
  
End Function

Public Property Get DialectFontBold() As Boolean

  On Error Resume Next
  DialectFontBold = GetFontBold(TagDialectFont)
  
End Property

Public Property Let DialectFontBold(bBold As Boolean)

  On Error Resume Next
  Call LetFontBold(TagDialectFont, bBold)
  
End Property

Public Property Get DialectFontItalic() As Boolean

  On Error Resume Next
  DialectFontItalic = GetFontItalic(TagDialectFont)
  
End Property

Public Property Let DialectFontItalic(bItalic As Boolean)

  On Error Resume Next
  Call LetFontItalic(TagDialectFont, bItalic)
  
End Property

Public Property Get DialectFontName() As String

  On Error Resume Next
  DialectFontName = GetFontName(TagDialectFont)
  
End Property

Public Property Let DialectFontName(sName$)

  On Error Resume Next
  Call LetFontName(TagDialectFont, sName)
  
End Property

Public Property Get DialectFontSize() As Integer

  On Error Resume Next
  DialectFontSize = GetFontSize(TagDialectFont)
  
End Property

Public Property Let DialectFontSize(iSize%)

  On Error Resume Next
  Call LetFontSize(TagDialectFont, iSize)
  
End Property

Public Function EmptyCategory(sCategoryName$) As Boolean
  
  Dim i As Integer
  Dim node1 As IXMLDOMNode
  Dim cNodes As IXMLDOMNodeList

  EmptyCategory = False
  Set node1 = GetCategoryNode(sCategoryName)
  If (node1 Is Nothing) Then Exit Function
  If Not (node1.hasChildNodes) Then Exit Function
  Set cNodes = node1.childNodes
  
  For i = cNodes.length - 1 To 1 Step -1
    node1.removeChild cNodes(i)
  Next

  EmptyCategory = True

End Function

Private Function GetFontBold(FontTag$) As Boolean
  
  Dim sGIPath As String
  
  On Error Resume Next
  sGIPath = "//" & TagGeneralInfo & "/"
  Set node = xmlWLDoc.selectSingleNode(sGIPath & FontTag & "/" & TagFontBold)
  If Not (node Is Nothing) Then GetFontBold = CBool(node.Text)

End Function

Private Function GetFontItalic(FontTag$) As Boolean
  
  Dim sGIPath As String
  
  On Error Resume Next
  sGIPath = "//" & TagGeneralInfo & "/"
  Set node = xmlWLDoc.selectSingleNode(sGIPath & FontTag & "/" & TagFontItalic)
  If Not (node Is Nothing) Then GetFontItalic = CBool(node.Text)

End Function

Private Function GetFontName(FontTag$) As String
  
  Dim sGIPath As String
  
  On Error Resume Next
  sGIPath = "//" & TagGeneralInfo & "/"
  Set node = xmlWLDoc.selectSingleNode(sGIPath & FontTag & "/" & TagFontName)
  If Not (node Is Nothing) Then GetFontName = node.Text

End Function

Private Function GetFontSize(FontTag$) As Integer

  Dim sGIPath As String
  
  On Error Resume Next
  GetFontSize = 10
  sGIPath = "//" & TagGeneralInfo & "/"
  Set node = xmlWLDoc.selectSingleNode(sGIPath & FontTag & "/" & TagFontSize)
  If Not (node Is Nothing) Then GetFontSize = node.Text
  
End Function

Private Function GetPath(PathTag$) As String

  On Error Resume Next
  GetPath = ""
  Set node = xmlWLDoc.selectSingleNode("//" & TagGeneralInfo & "/" & PathTag)
  If Not (node Is Nothing) Then GetPath = node.Text
  
End Function

Public Property Get GraphicPath() As String

  On Error Resume Next
  GraphicPath = GetPath(TagGraphicPath)

End Property

Public Property Let GraphicPath(sPath$)

  On Error Resume Next
  Call LetPath(TagGraphicPath, sPath)

End Property

Private Function GetCategoryNode(sCategoryName$) As IXMLDOMNode

  Dim i As Integer
  
  Set NodeList = xmlWLDoc.selectNodes("//" & TagAllCategories & "/" & TagOneCategory)
  
  If (NodeList.length > 0) Then
    For i = 0 To NodeList.length - 1
      If (NodeList.item(i).childNodes.item(0).Text = sCategoryName) Then
        Set GetCategoryNode = NodeList.item(i)
        Exit Function
      End If
    Next
  End If
  
  Set GetCategoryNode = Nothing
  
End Function
  
Private Function GetCategoryNodeList(sCategoryName$) As IXMLDOMNodeList

  Dim i As Integer
  
  Set NodeList = xmlWLDoc.selectNodes("//" & TagAllCategories & "/" & TagOneCategory)
  
  If (NodeList.length > 0) Then
    For i = 0 To NodeList.length - 1
      If (NodeList.item(i).childNodes.item(0).Text = sCategoryName) Then
        Set GetCategoryNodeList = NodeList.item(i).selectNodes("Word")
        Exit Function
      End If
    Next
  End If
  
  Set GetCategoryNodeList = Nothing
  
End Function

Public Property Get GlossFontBold() As Boolean

  On Error Resume Next
  GlossFontBold = GetFontBold(TagGlossFont)
  
End Property

Public Property Let GlossFontBold(bBold As Boolean)

  On Error Resume Next
  Call LetFontBold(TagGlossFont, bBold)
  
End Property

Public Property Get GlossFontItalic() As Boolean

  On Error Resume Next
  GlossFontItalic = GetFontItalic(TagGlossFont)
  
End Property

Public Property Let GlossFontItalic(bItalic As Boolean)

  On Error Resume Next
  Call LetFontItalic(TagGlossFont, bItalic)
  
End Property

Public Property Get GlossFontName() As String

  On Error Resume Next
  GlossFontName = GetFontName(TagGlossFont)
  
End Property

Public Property Let GlossFontName(sName$)

  On Error Resume Next
  Call LetFontName(TagGlossFont, sName)
  
End Property

Public Property Get GlossFontSize() As Integer

  On Error Resume Next
  GlossFontSize = GetFontSize(TagGlossFont)
  
End Property

Public Property Let GlossFontSize(iSize%)

  On Error Resume Next
  Call LetFontSize(TagGlossFont, iSize)
  
End Property

Public Property Get ID() As String

  On Error Resume Next
  
  If (RootNode.Attributes.length > 0) Then
    ID = RootNode.Attributes.item(0).Text
  Else
    ID = ""
  End If

End Property

Public Property Let ID(sID$)

  On Error Resume Next
  
  If (RootNode.Attributes.length > 0) Then
    RootNode.Attributes.item(0).Text = sID
  Else
    Dim attrib As IXMLDOMAttribute
    Set attrib = xmlWLDoc.createAttribute("ID")
    RootNode.Attributes.setNamedItem attrib
    RootNode.Attributes.item(0).Text = sID
  End If

End Property

Public Property Get InitFailure() As Boolean

  On Error Resume Next
  InitFailure = bInitFailure
  
End Property

Public Property Get IsFileValidWL() As Boolean

  On Error Resume Next
  IsFileValidWL = (RootNode.baseName = TagRootElement)
  
End Property

Private Sub LetFontBold(FontTag$, bBold As Boolean)

  Dim sGIPath As String
  
  On Error Resume Next
  sGIPath = "//" & TagGeneralInfo & "/"
  Set node = xmlWLDoc.selectSingleNode(sGIPath & FontTag & "/" & TagFontBold)
  If Not (node Is Nothing) Then node.Text = (Str$(bBold))

End Sub

Private Sub LetFontItalic(FontTag$, bItalic As Boolean)

  Dim sGIPath As String
  
  On Error Resume Next
  sGIPath = "//" & TagGeneralInfo & "/"
  Set node = xmlWLDoc.selectSingleNode(sGIPath & FontTag & "/" & TagFontItalic)
  If Not (node Is Nothing) Then node.Text = (Str$(bItalic))

End Sub

Private Sub LetFontName(FontTag$, sName$)

  Dim sGIPath As String
  
  On Error Resume Next
  sGIPath = "//" & TagGeneralInfo & "/"
  Set node = xmlWLDoc.selectSingleNode(sGIPath & FontTag & "/" & TagFontName)
  If Not (node Is Nothing) Then node.Text = Trim$(sName)

End Sub

Private Sub LetFontSize(FontTag$, iSize%)

  Dim sGIPath As String
  
  On Error Resume Next
  If (IsNumeric(iSize)) Then
    sGIPath = "//" & TagGeneralInfo & "/"
    Set node = xmlWLDoc.selectSingleNode(sGIPath & FontTag & "/" & TagFontSize)
    If Not (node Is Nothing) Then node.Text = Trim$(Str$(iSize))
  End If
  
End Sub

Private Sub LetPath(PathTag$, sPath$)

  On Error Resume Next
  Set node = xmlWLDoc.selectSingleNode("//" & TagGeneralInfo & "/" & PathTag)
  If Not (node Is Nothing) Then node.Text = sPath

End Sub

Public Function Load(sFileSpec$) As Boolean

  On Error Resume Next
  
  If (xmlWLDoc.Load(sFileSpec)) Then
    Set RootNode = xmlWLDoc.documentElement
    Load = True
    sXMLFilePath = sFileSpec
  Else
    Load = False
  End If
  
End Function

Public Sub LoadNew(sFileSpec$)

  Dim node2 As IXMLDOMNode
  Dim node3 As IXMLDOMNode
  
  On Error Resume Next
  
  If (Len(Trim$(sFileSpec)) = 0) Then
    MsgBox "No file specification for new XML file.", _
           vbOKOnly + vbExclamation, "XML WordList Class"
    Exit Sub
  End If
  
  sXMLFilePath = sFileSpec
  
  Set node = xmlWLDoc.createProcessingInstruction("xml", XMLProcInstruction)
  Set node = xmlWLDoc.insertBefore(node, xmlWLDoc.childNodes.item(0))

  Set RootNode = xmlWLDoc.createElement(TagRootElement)
  Set xmlWLDoc.documentElement = RootNode

  Set node = xmlWLDoc.createElement(TagGeneralInfo)
  RootNode.appendChild node

  Call WriteFontBlock(node, TagPhoneticFont)
  Call WriteFontBlock(node, TagOrthoFont)
  Call WriteFontBlock(node, TagGlossFont)
  Call WriteFontBlock(node, TagDialectFont)
  
  Set node2 = xmlWLDoc.createElement(TagSoundPath)
  node.appendChild node2

  Set node2 = xmlWLDoc.createElement(TagGraphicPath)
  node.appendChild node2

  Set node = xmlWLDoc.createElement(TagAllCategories)
  RootNode.appendChild node
  
End Sub

Public Property Get OrthoFontBold() As Boolean

  On Error Resume Next
  OrthoFontBold = GetFontBold(TagOrthoFont)
  
End Property

Public Property Let OrthoFontBold(bBold As Boolean)

  On Error Resume Next
  Call LetFontBold(TagOrthoFont, bBold)
  
End Property

Public Property Get OrthoFontItalic() As Boolean

  On Error Resume Next
  OrthoFontItalic = GetFontItalic(TagOrthoFont)
  
End Property

Public Property Let OrthoFontItalic(bItalic As Boolean)

  On Error Resume Next
  Call LetFontItalic(TagOrthoFont, bItalic)
  
End Property

Public Property Get OrthoFontName() As String

  On Error Resume Next
  OrthoFontName = GetFontName(TagOrthoFont)
  
End Property

Public Property Let OrthoFontName(sName$)

  On Error Resume Next
  Call LetFontName(TagOrthoFont, sName)
  
End Property

Public Property Get OrthoFontSize() As Integer

  On Error Resume Next
  OrthoFontSize = GetFontSize(TagOrthoFont)
  
End Property

Public Property Let OrthoFontSize(iSize%)

  On Error Resume Next
  Call LetFontSize(TagOrthoFont, iSize)
  
End Property

Public Property Get PhoneticFontBold() As Boolean

  On Error Resume Next
  PhoneticFontBold = GetFontBold(TagPhoneticFont)
  
End Property

Public Property Let PhoneticFontBold(bBold As Boolean)

  On Error Resume Next
  Call LetFontBold(TagPhoneticFont, bBold)
  
End Property

Public Property Get PhoneticFontItalic() As Boolean

  On Error Resume Next
  PhoneticFontItalic = GetFontItalic(TagPhoneticFont)
  
End Property

Public Property Let PhoneticFontItalic(bItalic As Boolean)

  On Error Resume Next
  Call LetFontItalic(TagPhoneticFont, bItalic)
  
End Property

Public Property Get PhoneticFontName() As String

  On Error Resume Next
  PhoneticFontName = GetFontName(TagPhoneticFont)
  
End Property

Public Property Let PhoneticFontName(sName$)

  On Error Resume Next
  Call LetFontName(TagPhoneticFont, sName)
  
End Property

Public Property Get PhoneticFontSize() As Integer

  On Error Resume Next
  PhoneticFontSize = GetFontSize(TagPhoneticFont)
  
End Property

Public Property Let PhoneticFontSize(iSize%)

  On Error Resume Next
  Call LetFontSize(TagPhoneticFont, iSize)
  
End Property

Public Function RemoveCategory(sCategoryName$) As Boolean

  Dim node1 As IXMLDOMNode
  Dim node2 As IXMLDOMNode

  RemoveCategory = False
  Set node1 = GetCategoryNode(sCategoryName)
  If (node1 Is Nothing) Then Exit Function
  Set node2 = node1.ParentNode
  node2.removeChild node1
  RemoveCategory = True
  
End Function

Public Sub Save(Optional vFileSpec)

  On Error Resume Next
  
  If (IsMissing(vFileSpec)) Then vFileSpec = sXMLFilePath
  
  If (Len(Trim$(vFileSpec)) = 0) Then
    MsgBox "No file specification for XML file save.", _
           vbOKOnly + vbExclamation, "XML WordList Class"
  Else
    xmlWLDoc.Save vFileSpec
  End If

End Sub

Public Property Get SoundPath() As String

  On Error Resume Next
  SoundPath = GetPath(TagSoundPath)

End Property

Public Property Let SoundPath(sPath$)

  On Error Resume Next
  Call LetPath(TagSoundPath, sPath)

End Property

Public Property Get WordsInCategory(sCategoryName$) As Variant

  Dim i As Integer
  Dim sWords() As String
  Dim nl As IXMLDOMNodeList
  
  WordsInCategory = Null
  Set nl = GetCategoryNodeList(sCategoryName)
  If (nl Is Nothing) Then Exit Property
  If (nl.length = 0) Then Exit Property
  
  ReDim sWords(0 To nl.length - 1, Pitch To GraphicFile) As String
  
  For i = 0 To nl.length - 1
    sWords(i, Pitch) = nl.item(i).childNodes.item(Pitch).Text
    sWords(i, phonetic) = nl.item(i).childNodes.item(phonetic).Text
    sWords(i, Ortho) = nl.item(i).childNodes.item(Ortho).Text
    sWords(i, Gloss) = nl.item(i).childNodes.item(Gloss).Text
    sWords(i, Dialect) = nl.item(i).childNodes.item(Dialect).Text
    sWords(i, SoundFile) = nl.item(i).childNodes.item(SoundFile).Text
    sWords(i, GraphicFile) = nl.item(i).childNodes.item(GraphicFile).Text
  Next

  WordsInCategory = sWords
  
End Property

Private Sub WriteFontBlock(ParentNode As IXMLDOMNode, FontTag$)
  
  Dim node2 As IXMLDOMNode
  Dim node3 As IXMLDOMNode

  On Error Resume Next

  Set node2 = xmlWLDoc.createElement(FontTag)
  ParentNode.appendChild node2

  Set node3 = xmlWLDoc.createElement(TagFontName)
  node2.appendChild node3

  Set node3 = xmlWLDoc.createElement(TagFontSize)
  node2.appendChild node3

  Set node3 = xmlWLDoc.createElement(TagFontBold)
  node3.Text = "False"
  node2.appendChild node3

  Set node3 = xmlWLDoc.createElement(TagFontItalic)
  node3.Text = "False"
  node2.appendChild node3

End Sub

Private Sub Class_Initialize()

  sXMLFilePath = ""
  bInitFailure = False
  Set xmlWLDoc = Nothing
  Set xmlWLDoc = New DOMDocument
  
  If (xmlWLDoc Is Nothing) Then
    bInitFailure = True
  Else
    xmlWLDoc.async = False
  End If

End Sub

Private Sub Class_Terminate()

  On Error Resume Next
  Set xmlWLDoc = Nothing
  
End Sub
