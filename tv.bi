Function TreeViewInsertItem(BYVAL hTreeView AS long, BYVAL hParent AS long,sItem AS STRING) AS LONG
    Dim tTVInsert   AS TV_INSERTSTRUCT
    Dim tTVItem     AS TV_ITEM

    If hParent THEN
        tTVItem.mask        = TVIF_CHILDREN OR TVIF_HANDLE
        tTVItem.hItem       = cast(HTREEITEM ,hParent)
        tTVItem.cchildren   = 1
        TreeView_SetItem( cast(hwnd,hTreeView), @tTVItem)
    END IF

    tTVInsert.hParent              = cast(HTREEITEM ,hParent)
    tTVInsert.Item.mask            = TVIF_TEXT Or TVIF_STATE
    tTVInsert.Item.state           = TVIS_EXPANDED
    tTVInsert.Item.pszText         = STRPTR(sItem)
    tTVInsert.Item.cchTextMax      = LEN(sItem)
    FUNCTION = cast(long,TreeView_InsertItem(cast(hwnd,hTreeView), @tTVInsert))
END Function

Function TreeView_GetItemText(ByVal hTreeView As Long, ByVal hItem As Long)As String
    Dim ItemText As zstring*32765
    Dim Item As TV_ITEM
    Item.hItem=cast(HTREEITEM ,hItem)
    Item.Mask=TVIF_TEXT
    Item.cchTextMax=32765
    Item.pszText=@ItemText
    SendMessage cast(hwnd,hTreeView),TVM_GETITEM,0,cast(LPARAM ,@Item)
    Function=ItemText
End Function

Function TreeView_ChangeChildState(ByVal hTreeView As Long, ByVal hItem As Long, ByVal iState As Long) As Long
	Dim hTmpItem As HTREEITEM
	Dim TVITEM As TV_ITEM
	hTmpItem = TreeView_GetChild(cast(hwnd,hTreeView),hItem)
	While hTmpItem
		TVITEM.mask = TVIF_STATE Or TVIF_HANDLE
		TVITEM.hItem = hTmpItem
		TVITEM.stateMask = TVIS_STATEIMAGEMASK
		TVITEM.state = (iState+1) Shl 12
		sendmessage cast(hwnd,hTreeView),TVM_SetItem,0,cast(LPARAM ,@TVITEM)
		TreeView_ChangeChildState hTreeView, cast(long,hTmpItem), iState
		hTmpItem = treeview_getnextitem(cast(hwnd,htreeview),hTmpItem,TVGN_NEXT)
	wend
	function=1
End Function

Function TreeView_ChangeParentState(ByVal hTreeView As Long, ByVal hItem As Long, ByVal iState As Long) As Long
	Dim hTmpItem As HTREEITEM
	Dim TVITEM As TV_ITEM
	hTmpItem = TreeView_GetParent(cast(hwnd,hTreeView),hItem)
	If hTmpItem then
		TVITEM.mask = TVIF_STATE Or TVIF_HANDLE
		TVITEM.hItem = hTmpItem
		TVITEM.stateMask = TVIS_STATEIMAGEMASK
		TVITEM.state = (iState+1) Shl 12
		sendmessage cast(hwnd,hTreeView),TVM_SetItem,0,cast(LPARAM ,@TVITEM)
		TreeView_ChangeParentState hTreeView, cast(long,hTmpItem),0
	End If
	function=1
End Function

Function TreeView_SetCheckState(ByVal hTreeView As Long, ByVal hItem As Long, ByVal iState As Long) As Long
	Dim TVITEM As TV_ITEM
	TVITEM.mask = TVIF_STATE Or TVIF_HANDLE
	TVITEM.hItem = cast(HTREEITEM,hItem)
	TVITEM.stateMask = TVIS_STATEIMAGEMASK
	TVITEM.state = (iState+1) Shl 12
	sendmessage cast(hwnd,hTreeView),TVM_SetItem,0,cast(LPARAM ,@TVITEM)
	function=1
End Function

Function TreeView_GetCheckState( ByVal hWnd1 As Long, ByVal hItem As Long ) As Long
    Dim tvItem As TV_ITEM
    tvItem.mask         = TVIF_HANDLE Or TVIF_STATE
    tvItem.hItem        = cast(HTREEITEM,hItem)
    tvItem.stateMask    = TVIS_STATEIMAGEMASK
    sendmessage cast(hwnd,hwnd1),TVM_GetItem,0,cast(LPARAM ,@tvItem)
    tvItem.state Shr=12
    tvItem.state-=1
    Function = tvItem.state Xor 1
End Function
