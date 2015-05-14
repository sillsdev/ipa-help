Option Strict Off
Option Explicit On

Friend Class clsXMLWordList
	
	Private Const XMLProcInstruction As String = "version='1.0'" ' encoding='iso-8859-1'"
	
	Private Const TagRootElement As String = "IPAWordListDef"
	Private Const TagGeneralInfo As String = "GeneralInfo"
	Private Const TagPhoneticFont As String = "PhoneticFont"
	Private Const TagOrthoFont As String = "OrthoFont"
	Private Const TagGlossFont As String = "GlossFont"
	Private Const TagDialectFont As String = "DialectFont"
	Private Const TagFontName As String = "Name"
	Private Const TagFontSize As String = "Size"
	Private Const TagFontBold As String = "Bold"
	Private Const TagFontItalic As String = "Italic"
	Private Const TagSoundPath As String = "SoundPath"
	Private Const TagGraphicPath As String = "GraphicPath"
	Private Const TagAllCategories As String = "Categories"
	Private Const TagOneCategory As String = "Category"
	Private Const TagWord As String = "Word"
	Private Const TagPitch As String = "Pitch"
	Private Const TagPhonetic As String = "Phonetic"
	Private Const TagOrtho As String = "Orthographic"
	Private Const TagGloss As String = "Gloss"
	Private Const TagDialect As String = "Dialect"
	Private Const TagSoundFile As String = "SoundFile"
	Private Const TagGraphicFile As String = "GraphicFile"
	
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
    Private xmlWLDoc As MSXML2.DOMDocument60
    Private RootNode As MSXML2.IXMLDOMElement
    Private node As MSXML2.IXMLDOMNode
    Private NodeList As MSXML2.IXMLDOMNodeList
	
	Public Sub AddCategory(ByRef sCategoryName As String, Optional ByRef sWords As Object = Nothing, Optional ByRef sCategoryFollowing As Object = Nothing)
		
		Dim i As Short
        Dim node2 As MSXML2.IXMLDOMNode
        Dim node3 As MSXML2.IXMLDOMNode
        Dim node4 As MSXML2.IXMLDOMNode
		
		'*********************************************************
		'* If category already exists then don't add a new node
		'* for it. Just point to it and add the words to it.
		'*********************************************************
		node = GetCategoryNode(sCategoryName)
		If (node Is Nothing) Then
			node = xmlWLDoc.createElement(TagOneCategory)
			node.Text = sCategoryName
		End If
		
		'*********************************************************
		'* Don't loop through words if there aren't any specified.
		'*********************************************************
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If Not (IsNothing(sWords)) Then
			For i = 0 To UBound(sWords, 1)
				node2 = xmlWLDoc.createElement(TagWord)
				node.appendChild(node2)
				
				node3 = xmlWLDoc.createElement(TagPitch)
				'UPGRADE_WARNING: Couldn't resolve default property of object sWords(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				node3.Text = sWords(i, WordRecord.Pitch)
				node2.appendChild(node3)
				
				node3 = xmlWLDoc.createElement(TagPhonetic)
				'UPGRADE_WARNING: Couldn't resolve default property of object sWords(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				node3.Text = sWords(i, WordRecord.phonetic)
				node2.appendChild(node3)
				
				node3 = xmlWLDoc.createElement(TagOrtho)
				'UPGRADE_WARNING: Couldn't resolve default property of object sWords(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				node3.Text = sWords(i, WordRecord.Ortho)
				node2.appendChild(node3)
				
				node3 = xmlWLDoc.createElement(TagGloss)
				'UPGRADE_WARNING: Couldn't resolve default property of object sWords(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				node3.Text = sWords(i, WordRecord.Gloss)
				node2.appendChild(node3)
				
				node3 = xmlWLDoc.createElement(TagDialect)
				'UPGRADE_WARNING: Couldn't resolve default property of object sWords(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				node3.Text = sWords(i, WordRecord.Dialect)
				node2.appendChild(node3)
				
				node3 = xmlWLDoc.createElement(TagSoundFile)
				'UPGRADE_WARNING: Couldn't resolve default property of object sWords(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				node3.Text = sWords(i, WordRecord.SoundFile)
				node2.appendChild(node3)
				
				node3 = xmlWLDoc.createElement(TagGraphicFile)
				'UPGRADE_WARNING: Couldn't resolve default property of object sWords(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				node3.Text = sWords(i, WordRecord.GraphicFile)
				node2.appendChild(node3)
			Next 
		End If
		
		'*********************************************************
		'* Now point to the category's node and it's parent.
		'*********************************************************
		node2 = xmlWLDoc.selectSingleNode("//" & TagAllCategories)
		node3 = GetCategoryNode(sCategoryName)
		
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
			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If (IsNothing(sCategoryFollowing)) Then
				node2.appendChild(node)
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object sCategoryFollowing. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				node4 = GetCategoryNode(CStr(sCategoryFollowing))
				node2.insertBefore(node, IIf(node4 Is Nothing, Nothing, node4))
			End If
		Else
			node2.replaceChild(node, node3)
		End If
		
	End Sub
	
	Public ReadOnly Property CategoryNames() As String()
		Get
			
			Dim i As Short
			Dim sCategories() As String
            Dim nl As MSXML2.IXMLDOMNodeList
			
			NodeList = xmlWLDoc.selectNodes("//" & TagAllCategories & "/" & TagOneCategory)
			
			If (NodeList.length > 0) Then
				ReDim sCategories(NodeList.length - 1)
				For i = 0 To NodeList.length - 1
					sCategories(i) = NodeList.item(i).childNodes.item(0).Text
				Next 
			End If
			
			CategoryNames = VB6.CopyArray(sCategories)
			
		End Get
	End Property
	
	
	Public Property DialectFontBold() As Boolean
		Get
			
			On Error Resume Next
			DialectFontBold = GetFontBold(TagDialectFont)
			
		End Get
		Set(ByVal Value As Boolean)
			
			On Error Resume Next
			Call LetFontBold(TagDialectFont, Value)
			
		End Set
	End Property
	
	
	Public Property DialectFontItalic() As Boolean
		Get
			
			On Error Resume Next
			DialectFontItalic = GetFontItalic(TagDialectFont)
			
		End Get
		Set(ByVal Value As Boolean)
			
			On Error Resume Next
			Call LetFontItalic(TagDialectFont, Value)
			
		End Set
	End Property
	
	
	Public Property DialectFontName() As String
		Get
			
			On Error Resume Next
			DialectFontName = GetFontName(TagDialectFont)
			
		End Get
		Set(ByVal Value As String)
			
			On Error Resume Next
			Call LetFontName(TagDialectFont, Value)
			
		End Set
	End Property
	
	
	Public Property DialectFontSize() As Short
		Get
			
			On Error Resume Next
			DialectFontSize = GetFontSize(TagDialectFont)
			
		End Get
		Set(ByVal Value As Short)
			
			On Error Resume Next
			Call LetFontSize(TagDialectFont, Value)
			
		End Set
	End Property
	
	
	Public Property GraphicPath() As String
		Get
			
			On Error Resume Next
			GraphicPath = GetPath(TagGraphicPath)
			
		End Get
		Set(ByVal Value As String)
			
			On Error Resume Next
			Call LetPath(TagGraphicPath, Value)
			
		End Set
	End Property
	
	
	Public Property GlossFontBold() As Boolean
		Get
			
			On Error Resume Next
			GlossFontBold = GetFontBold(TagGlossFont)
			
		End Get
		Set(ByVal Value As Boolean)
			
			On Error Resume Next
			Call LetFontBold(TagGlossFont, Value)
			
		End Set
	End Property
	
	
	Public Property GlossFontItalic() As Boolean
		Get
			
			On Error Resume Next
			GlossFontItalic = GetFontItalic(TagGlossFont)
			
		End Get
		Set(ByVal Value As Boolean)
			
			On Error Resume Next
			Call LetFontItalic(TagGlossFont, Value)
			
		End Set
	End Property
	
	
	Public Property GlossFontName() As String
		Get
			
			On Error Resume Next
			GlossFontName = GetFontName(TagGlossFont)
			
		End Get
		Set(ByVal Value As String)
			
			On Error Resume Next
			Call LetFontName(TagGlossFont, Value)
			
		End Set
	End Property
	
	
	Public Property GlossFontSize() As Short
		Get
			
			On Error Resume Next
			GlossFontSize = GetFontSize(TagGlossFont)
			
		End Get
		Set(ByVal Value As Short)
			
			On Error Resume Next
			Call LetFontSize(TagGlossFont, Value)
			
		End Set
	End Property
	
	
	Public Property ID() As String
		Get
			
			On Error Resume Next
			
			If (RootNode.Attributes.length > 0) Then
				ID = RootNode.Attributes.item(0).Text
			Else
				ID = ""
			End If
			
		End Get
		Set(ByVal Value As String)
			
			On Error Resume Next
			
            Dim attrib As MSXML2.IXMLDOMAttribute
			If (RootNode.Attributes.length > 0) Then
				RootNode.Attributes.item(0).Text = Value
			Else
				attrib = xmlWLDoc.createAttribute("ID")
				'UPGRADE_WARNING: Couldn't resolve default property of object attrib. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				RootNode.Attributes.setNamedItem(attrib)
				RootNode.Attributes.item(0).Text = Value
			End If
			
		End Set
	End Property
	
	Public ReadOnly Property InitFailure() As Boolean
		Get
			
			On Error Resume Next
			InitFailure = bInitFailure
			
		End Get
	End Property
	
	Public ReadOnly Property IsFileValidWL() As Boolean
		Get
			
			On Error Resume Next
			IsFileValidWL = (RootNode.baseName = TagRootElement)
			
		End Get
	End Property
	
	
	Public Property OrthoFontBold() As Boolean
		Get
			
			On Error Resume Next
			OrthoFontBold = GetFontBold(TagOrthoFont)
			
		End Get
		Set(ByVal Value As Boolean)
			
			On Error Resume Next
			Call LetFontBold(TagOrthoFont, Value)
			
		End Set
	End Property
	
	
	Public Property OrthoFontItalic() As Boolean
		Get
			
			On Error Resume Next
			OrthoFontItalic = GetFontItalic(TagOrthoFont)
			
		End Get
		Set(ByVal Value As Boolean)
			
			On Error Resume Next
			Call LetFontItalic(TagOrthoFont, Value)
			
		End Set
	End Property
	
	
	Public Property OrthoFontName() As String
		Get
			
			On Error Resume Next
			OrthoFontName = GetFontName(TagOrthoFont)
			
		End Get
		Set(ByVal Value As String)
			
			On Error Resume Next
			Call LetFontName(TagOrthoFont, Value)
			
		End Set
	End Property
	
	
	Public Property OrthoFontSize() As Short
		Get
			
			On Error Resume Next
			OrthoFontSize = GetFontSize(TagOrthoFont)
			
		End Get
		Set(ByVal Value As Short)
			
			On Error Resume Next
			Call LetFontSize(TagOrthoFont, Value)
			
		End Set
	End Property
	
	
	Public Property PhoneticFontBold() As Boolean
		Get
			
			On Error Resume Next
			PhoneticFontBold = GetFontBold(TagPhoneticFont)
			
		End Get
		Set(ByVal Value As Boolean)
			
			On Error Resume Next
			Call LetFontBold(TagPhoneticFont, Value)
			
		End Set
	End Property
	
	
	Public Property PhoneticFontItalic() As Boolean
		Get
			
			On Error Resume Next
			PhoneticFontItalic = GetFontItalic(TagPhoneticFont)
			
		End Get
		Set(ByVal Value As Boolean)
			
			On Error Resume Next
			Call LetFontItalic(TagPhoneticFont, Value)
			
		End Set
	End Property
	
	
	Public Property PhoneticFontName() As String
		Get
			
			On Error Resume Next
			PhoneticFontName = GetFontName(TagPhoneticFont)
			
		End Get
		Set(ByVal Value As String)
			
			On Error Resume Next
			Call LetFontName(TagPhoneticFont, Value)
			
		End Set
	End Property
	
	
	Public Property PhoneticFontSize() As Short
		Get
			
			On Error Resume Next
			PhoneticFontSize = GetFontSize(TagPhoneticFont)
			
		End Get
		Set(ByVal Value As Short)
			
			On Error Resume Next
			Call LetFontSize(TagPhoneticFont, Value)
			
		End Set
	End Property
	
	
	Public Property SoundPath() As String
		Get
			
			On Error Resume Next
			SoundPath = GetPath(TagSoundPath)
			
		End Get
		Set(ByVal Value As String)
			
			On Error Resume Next
			Call LetPath(TagSoundPath, Value)
			
		End Set
	End Property
	
	Public ReadOnly Property WordsInCategory(ByVal sCategoryName As String) As Object
		Get
			
			Dim i As Short
            Dim sWords(,) As String
            Dim nl As MSXML2.IXMLDOMNodeList
			
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object WordsInCategory. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			WordsInCategory = System.DBNull.Value
			nl = GetCategoryNodeList(sCategoryName)
			If (nl Is Nothing) Then Exit Property
			If (nl.length = 0) Then Exit Property
			
            ReDim sWords(nl.length - 1, WordRecord.GraphicFile)
            For i = 0 To nl.length - 1
                sWords(i, WordRecord.Pitch) = nl.item(i).childNodes.item(WordRecord.Pitch).Text
                sWords(i, WordRecord.phonetic) = nl.item(i).childNodes.item(WordRecord.phonetic).Text
                sWords(i, WordRecord.Ortho) = nl.item(i).childNodes.item(WordRecord.Ortho).Text
                sWords(i, WordRecord.Gloss) = nl.item(i).childNodes.item(WordRecord.Gloss).Text
                sWords(i, WordRecord.Dialect) = nl.item(i).childNodes.item(WordRecord.Dialect).Text
                sWords(i, WordRecord.SoundFile) = nl.item(i).childNodes.item(WordRecord.SoundFile).Text
                sWords(i, WordRecord.GraphicFile) = nl.item(i).childNodes.item(WordRecord.GraphicFile).Text
            Next
			
			'UPGRADE_WARNING: Couldn't resolve default property of object WordsInCategory. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			WordsInCategory = VB6.CopyArray(sWords)
			
		End Get
	End Property
	
	Public Function ChangeCategoryName(ByRef sOldName As String, ByRef sNewName As String) As Boolean
		
		ChangeCategoryName = False
		node = GetCategoryNode(sOldName)
		If (node Is Nothing) Then Exit Function
		If (node.childNodes(0) Is Nothing) Then Exit Function
		node.childNodes(0).Text = sNewName
		ChangeCategoryName = True
		
	End Function
	
	Public Function EmptyCategory(ByRef sCategoryName As String) As Boolean
		
		Dim i As Short
        Dim node1 As MSXML2.IXMLDOMNode
        Dim cNodes As MSXML2.IXMLDOMNodeList
		
		EmptyCategory = False
		node1 = GetCategoryNode(sCategoryName)
		If (node1 Is Nothing) Then Exit Function
		If Not (node1.hasChildNodes) Then Exit Function
		cNodes = node1.childNodes
		
		For i = cNodes.length - 1 To 1 Step -1
			node1.removeChild(cNodes(i))
		Next 
		
		EmptyCategory = True
		
	End Function
	
	Private Function GetFontBold(ByRef FontTag As String) As Boolean
		
		Dim sGIPath As String
		
		On Error Resume Next
		sGIPath = "//" & TagGeneralInfo & "/"
		node = xmlWLDoc.selectSingleNode(sGIPath & FontTag & "/" & TagFontBold)
		If Not (node Is Nothing) Then GetFontBold = CBool(node.Text)
		
	End Function
	
	Private Function GetFontItalic(ByRef FontTag As String) As Boolean
		
		Dim sGIPath As String
		
		On Error Resume Next
		sGIPath = "//" & TagGeneralInfo & "/"
		node = xmlWLDoc.selectSingleNode(sGIPath & FontTag & "/" & TagFontItalic)
		If Not (node Is Nothing) Then GetFontItalic = CBool(node.Text)
		
	End Function
	
	Private Function GetFontName(ByRef FontTag As String) As String
		
		Dim sGIPath As String
		
		On Error Resume Next
		sGIPath = "//" & TagGeneralInfo & "/"
		node = xmlWLDoc.selectSingleNode(sGIPath & FontTag & "/" & TagFontName)
		If Not (node Is Nothing) Then GetFontName = node.Text
		
	End Function
	
	Private Function GetFontSize(ByRef FontTag As String) As Short
		
		Dim sGIPath As String
		
		On Error Resume Next
		GetFontSize = 10
		sGIPath = "//" & TagGeneralInfo & "/"
		node = xmlWLDoc.selectSingleNode(sGIPath & FontTag & "/" & TagFontSize)
		If Not (node Is Nothing) Then GetFontSize = CShort(node.Text)
		
	End Function
	
	Private Function GetPath(ByRef PathTag As String) As String
		
		On Error Resume Next
		GetPath = ""
		node = xmlWLDoc.selectSingleNode("//" & TagGeneralInfo & "/" & PathTag)
		If Not (node Is Nothing) Then GetPath = node.Text
		
	End Function
	
    Private Function GetCategoryNode(ByRef sCategoryName As String) As MSXML2.IXMLDOMNode

        Dim i As Short

        NodeList = xmlWLDoc.selectNodes("//" & TagAllCategories & "/" & TagOneCategory)

        If (NodeList.length > 0) Then
            For i = 0 To NodeList.length - 1
                If (NodeList.item(i).childNodes.item(0).text = sCategoryName) Then
                    GetCategoryNode = NodeList.item(i)
                    Exit Function
                End If
            Next
        End If

        'UPGRADE_NOTE: Object GetCategoryNode may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        GetCategoryNode = Nothing

    End Function
	
    Private Function GetCategoryNodeList(ByRef sCategoryName As String) As MSXML2.IXMLDOMNodeList

        Dim i As Short

        NodeList = xmlWLDoc.selectNodes("//" & TagAllCategories & "/" & TagOneCategory)

        If (NodeList.length > 0) Then
            For i = 0 To NodeList.length - 1
                If (NodeList.item(i).childNodes.item(0).text = sCategoryName) Then
                    GetCategoryNodeList = NodeList.item(i).selectNodes("Word")
                    Exit Function
                End If
            Next
        End If

        'UPGRADE_NOTE: Object GetCategoryNodeList may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        GetCategoryNodeList = Nothing

    End Function
	
	Private Sub LetFontBold(ByRef FontTag As String, ByRef bBold As Boolean)
		
		Dim sGIPath As String
		
		On Error Resume Next
		sGIPath = "//" & TagGeneralInfo & "/"
		node = xmlWLDoc.selectSingleNode(sGIPath & FontTag & "/" & TagFontBold)
		If Not (node Is Nothing) Then node.Text = (Str(bBold))
		
	End Sub
	
	Private Sub LetFontItalic(ByRef FontTag As String, ByRef bItalic As Boolean)
		
		Dim sGIPath As String
		
		On Error Resume Next
		sGIPath = "//" & TagGeneralInfo & "/"
		node = xmlWLDoc.selectSingleNode(sGIPath & FontTag & "/" & TagFontItalic)
		If Not (node Is Nothing) Then node.Text = (Str(bItalic))
		
	End Sub
	
	Private Sub LetFontName(ByRef FontTag As String, ByRef sName As String)
		
		Dim sGIPath As String
		
		On Error Resume Next
		sGIPath = "//" & TagGeneralInfo & "/"
		node = xmlWLDoc.selectSingleNode(sGIPath & FontTag & "/" & TagFontName)
		If Not (node Is Nothing) Then node.Text = Trim(sName)
		
	End Sub
	
	Private Sub LetFontSize(ByRef FontTag As String, ByRef iSize As Short)
		
		Dim sGIPath As String
		
		On Error Resume Next
		If (IsNumeric(iSize)) Then
			sGIPath = "//" & TagGeneralInfo & "/"
			node = xmlWLDoc.selectSingleNode(sGIPath & FontTag & "/" & TagFontSize)
			If Not (node Is Nothing) Then node.Text = Trim(Str(iSize))
		End If
		
	End Sub
	
	Private Sub LetPath(ByRef PathTag As String, ByRef sPath As String)
		
		On Error Resume Next
		node = xmlWLDoc.selectSingleNode("//" & TagGeneralInfo & "/" & PathTag)
		If Not (node Is Nothing) Then node.Text = sPath
		
	End Sub
	
	Public Function Load(ByRef sFileSpec As String) As Boolean
		
		On Error Resume Next
		
		If (xmlWLDoc.Load(sFileSpec)) Then
			RootNode = xmlWLDoc.documentElement
			Load = True
			sXMLFilePath = sFileSpec
		Else
			Load = False
		End If
		
	End Function
	
	Public Sub LoadNew(ByRef sFileSpec As String)
		
        Dim node2 As MSXML2.IXMLDOMNode
        Dim node3 As MSXML2.IXMLDOMNode
		
		On Error Resume Next
		
		If (Len(Trim(sFileSpec)) = 0) Then
			MsgBox("No file specification for new XML file.", MsgBoxStyle.OKOnly + MsgBoxStyle.Exclamation, "XML WordList Class")
			Exit Sub
		End If
		
		sXMLFilePath = sFileSpec
		
		node = xmlWLDoc.createProcessingInstruction("xml", XMLProcInstruction)
		node = xmlWLDoc.insertBefore(node, xmlWLDoc.childNodes.item(0))
		
		RootNode = xmlWLDoc.createElement(TagRootElement)
		xmlWLDoc.documentElement = RootNode
		
		node = xmlWLDoc.createElement(TagGeneralInfo)
		RootNode.appendChild(node)
		
		Call WriteFontBlock(node, TagPhoneticFont)
		Call WriteFontBlock(node, TagOrthoFont)
		Call WriteFontBlock(node, TagGlossFont)
		Call WriteFontBlock(node, TagDialectFont)
		
		node2 = xmlWLDoc.createElement(TagSoundPath)
		node.appendChild(node2)
		
		node2 = xmlWLDoc.createElement(TagGraphicPath)
		node.appendChild(node2)
		
		node = xmlWLDoc.createElement(TagAllCategories)
		RootNode.appendChild(node)
		
	End Sub
	
	Public Function RemoveCategory(ByRef sCategoryName As String) As Boolean
		
        Dim node1 As MSXML2.IXMLDOMNode
        Dim node2 As MSXML2.IXMLDOMNode
		
		RemoveCategory = False
		node1 = GetCategoryNode(sCategoryName)
		If (node1 Is Nothing) Then Exit Function
		node2 = node1.ParentNode
		node2.removeChild(node1)
		RemoveCategory = True
		
	End Function
	
	Public Sub Save(Optional ByRef vFileSpec As Object = Nothing)
		
		On Error Resume Next
		
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		'UPGRADE_WARNING: Couldn't resolve default property of object vFileSpec. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (IsNothing(vFileSpec)) Then vFileSpec = sXMLFilePath
		
		'UPGRADE_WARNING: Couldn't resolve default property of object vFileSpec. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (Len(Trim(vFileSpec)) = 0) Then
			MsgBox("No file specification for XML file save.", MsgBoxStyle.OKOnly + MsgBoxStyle.Exclamation, "XML WordList Class")
		Else
			xmlWLDoc.Save(vFileSpec)
		End If
		
	End Sub
	
    Private Sub WriteFontBlock(ByRef ParentNode As MSXML2.IXMLDOMNode, ByRef FontTag As String)

        Dim node2 As MSXML2.IXMLDOMNode
        Dim node3 As MSXML2.IXMLDOMNode

        On Error Resume Next

        node2 = xmlWLDoc.createElement(FontTag)
        ParentNode.appendChild(node2)

        node3 = xmlWLDoc.createElement(TagFontName)
        node2.appendChild(node3)

        node3 = xmlWLDoc.createElement(TagFontSize)
        node2.appendChild(node3)

        node3 = xmlWLDoc.createElement(TagFontBold)
        node3.text = "False"
        node2.appendChild(node3)

        node3 = xmlWLDoc.createElement(TagFontItalic)
        node3.text = "False"
        node2.appendChild(node3)

    End Sub
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		
		sXMLFilePath = ""
		bInitFailure = False
		'UPGRADE_NOTE: Object xmlWLDoc may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		xmlWLDoc = Nothing
        xmlWLDoc = New MSXML2.DOMDocument60
		
		If (xmlWLDoc Is Nothing) Then
			bInitFailure = True
		Else
			xmlWLDoc.async = False
		End If
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		
		On Error Resume Next
		'UPGRADE_NOTE: Object xmlWLDoc may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		xmlWLDoc = Nothing
		
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
End Class