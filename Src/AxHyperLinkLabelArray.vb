'UPGRADE_WARNING: The entire project must be compiled once before a form with an ActiveX Control Array can be displayed

Imports System.ComponentModel

<ProvideProperty("Index",GetType(AxavHyperLink.AxHyperLinkLabel))> Public Class AxHyperLinkLabelArray
	Inherits Microsoft.VisualBasic.Compatibility.VB6.BaseOcxArray
	Implements IExtenderProvider

	Public Sub New()
		MyBase.New()
	End Sub

	Public Sub New(ByVal Container As IContainer)
		MyBase.New(Container)
	End Sub

	Public Shadows Event [ClickEvent] (ByVal sender As System.Object, ByVal e As System.EventArgs)
	Public Shadows Event [Error] (ByVal sender As System.Object, ByVal e As AxavHyperLink.__HyperLinkLabel_ErrorEvent)
	Public Shadows Event [Hover] (ByVal sender As System.Object, ByVal e As System.EventArgs)

	<System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)> Public Function CanExtend(ByVal target As Object) As Boolean Implements IExtenderProvider.CanExtend
		If TypeOf target Is AxavHyperLink.AxHyperLinkLabel Then
			Return BaseCanExtend(target)
		End If
	End Function

	Public Function GetIndex(ByVal o As AxavHyperLink.AxHyperLinkLabel) As Short
		Return BaseGetIndex(o)
	End Function

	<System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)> Public Sub SetIndex(ByVal o As AxavHyperLink.AxHyperLinkLabel, ByVal Index As Short)
		BaseSetIndex(o, Index)
	End Sub

	<System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)> Public Function ShouldSerializeIndex(ByVal o As AxavHyperLink.AxHyperLinkLabel) As Boolean
		Return BaseShouldSerializeIndex(o)
	End Function

	<System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)> Public Sub ResetIndex(ByVal o As AxavHyperLink.AxHyperLinkLabel)
		BaseResetIndex(o)
	End Sub

	Default Public ReadOnly Property Item(ByVal Index As Short) As AxavHyperLink.AxHyperLinkLabel
		Get
			Item = CType(BaseGetItem(Index), AxavHyperLink.AxHyperLinkLabel)
		End Get
	End Property

	Protected Overrides Function GetControlInstanceType() As System.Type
		Return GetType(AxavHyperLink.AxHyperLinkLabel)
	End Function

	Protected Overrides Sub HookUpControlEvents(ByVal o As Object)
		Dim ctl As AxavHyperLink.AxHyperLinkLabel = CType(o, AxavHyperLink.AxHyperLinkLabel)
		MyBase.HookUpControlEvents(o)
		If Not ClickEventEvent Is Nothing Then
			AddHandler ctl.ClickEvent, New System.EventHandler(AddressOf HandleClickEvent)
		End If
		If Not ErrorEvent Is Nothing Then
			AddHandler ctl.Error, New AxavHyperLink.__HyperLinkLabel_ErrorEventHandler(AddressOf HandleError)
		End If
		If Not HoverEvent Is Nothing Then
			AddHandler ctl.Hover, New System.EventHandler(AddressOf HandleHover)
		End If
	End Sub

	Private Sub HandleClickEvent (ByVal sender As System.Object, ByVal e As System.EventArgs) 
		RaiseEvent [ClickEvent] (sender, e)
	End Sub

	Private Sub HandleError (ByVal sender As System.Object, ByVal e As AxavHyperLink.__HyperLinkLabel_ErrorEvent) 
		RaiseEvent [Error] (sender, e)
	End Sub

	Private Sub HandleHover (ByVal sender As System.Object, ByVal e As System.EventArgs) 
		RaiseEvent [Hover] (sender, e)
	End Sub

End Class

