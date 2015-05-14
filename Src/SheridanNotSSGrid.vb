Public Class SheridanNotSSGrid

    Friend Bookmark As Object
    Friend Rows As Integer
    Friend Redraw As Boolean
    Friend AllowUpdate As Boolean
    Friend AllowAddNew As Boolean
    Friend AllowDelete As Boolean
    Friend Col As Integer
    Private _rowBookmark As Object
    Private _styleSets As Object

    Property SelBookmarks As Object

    Property Columns(ByVal p1 As String) As Object
        Get
            Return _Columns
        End Get
        Set(ByVal value As Object)
            _Columns = value
        End Set
    End Property

    Property Columns As Object

    Property Cols As Integer

    Property RowBookmark(ByVal i As Short) As Object
        Get
            Return _rowBookmark
        End Get
        Set(ByVal value As Object)
            _rowBookmark = value
        End Set
    End Property

    Property OcxState As AxHost.State

    Property RowHeight As Object

    Property StyleSets(ByVal p1 As Object) As Object
        Get
            Return _styleSets
        End Get
        Set(ByVal value As Object)
            _styleSets = value
        End Set
    End Property

    Function AddItemBookmark(ByVal i As Object) As Object
        Throw New NotImplementedException
    End Function

    Function AddItem(ByVal p1 As Object) As Integer
        Throw New NotImplementedException
    End Function

    Sub RemoveAll()
        Throw New NotImplementedException
    End Sub

    Sub CtlUpdate()
        Throw New NotImplementedException
    End Sub

End Class
