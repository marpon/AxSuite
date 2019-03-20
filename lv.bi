#include once "strwrap.bi"


'CONST LVN_GETDISPINFO     = LVN_FIRST - 50
'const LVN_COLUMNCLICK     = LVN_FIRST - 8
const LVM_SORTITEMSEX     = LVM_FIRST + 81

'TYPE LV_ITEM
'    mask       AS uinteger
'    iItem      AS LONG
'    iSubItem   AS LONG
'    STATE      AS uinteger
'    stateMask  AS uinteger
'    pszText    AS zstring PTR
'    cchTextMax AS LONG
'    iImage     AS LONG
'    lParam     AS LONG
'    iIndent    AS LONG
'    iGroupId   AS LONG
'    cColumns   AS uinteger        ' tile view columns
'    puColumns  AS uinteger PTR
'END TYPE

TYPE LV_DISPINFO
    hdr  AS NMHDR
    item AS LV_ITEM
END TYPE

Sub LvInit
	dim iccex AS InitCommonControlsEx
	iccex.dwSize = SIZEOF(iccex)
	iccex.dwICC  = ICC_LISTVIEW_CLASSES
	InitCommonControlsEx(@iccex)
End Sub

SUB LVSetHeader(BYVAL lvg AS uinteger,BYVAL strItem AS STRING)
  dim szItem        AS zstring * MAX_PATH   ' working variable
  dim tlvc          AS LVCOLUMN             ' specifies or receives the attributes of a listview column
  dim tlvi          AS LVITEM               ' specifies or receives the attributes of a listview item
  dim iItem         AS LONG
  dim nItem         AS LONG

  nItem=str_numparse(stritem,"|")

  FOR iItem=1 TO nItem
      szItem =str_parse(stritem,"|",iItem)
      tlvc.mask     = LVCF_FMT OR LVCF_WIDTH OR LVCF_TEXT OR LVCF_SUBITEM
      tlvc.fmt      = LVCFMT_LEFT
      tlvc.cx       = 96
      tlvc.pszText  = VARPTR(szItem)
      tlvc.iSubItem = iitem-1
      SendMessage cast(hwnd,lvg), LVM_SETCOLUMN, iitem-1, cast(LPARAM ,VARPTR(tlvc))
  NEXT
END Sub

Sub LVSetColFmt(BYVAL lvg AS uinteger,ByVal col As LONG,BYVAL fmt AS LONG)
  dim szItem        AS zstring * MAX_PATH   ' working variable
  dim tlvc          AS LVCOLUMN             ' specifies or receives the attributes of a listview column
  'dim tlvi          AS LV_ITEM               ' specifies or receives the attributes of a listview item

	tlvc.mask     = LVCF_FMT OR LVCF_SUBITEM
	tlvc.fmt      = fmt
	tlvc.iSubItem = col
	SendMessage cast(hwnd,lvg), LVM_SETCOLUMN, col, cast(LPARAM ,VARPTR(tlvc))

End Sub

SUB LVSetMaxCols(BYVAL rr AS uinteger,BYVAL cols AS LONG)
  dim tlvc  AS LVCOLUMN
  dim col   AS LONG

  FOR col=0 TO cols-1:SendMessage cast(hwnd,rr), LVM_INSERTCOLUMN, 0, cast(LPARAM , VARPTR(tlvc)):NEXT
END SUB

SUB LVSetMaxRows(BYVAL rr as uinteger,BYVAL rows AS LONG)
  dim irow AS LONG
  dim tlvi AS LVITEM
  dim glvnumrows as long
  glvnumrows=sendmessage(cast(hwnd,rr),LVM_GETITEMCOUNT,0,0)
  IF glvnumrows<=rows THEN
    FOR irow=glvnumrows TO rows-1
      tlvi.iItem = irow
      SendMessage cast(hwnd,rr), LVM_INSERTITEM, irow, cast(LPARAM , VARPTR(tlvi))
    NEXT
  ELSE
    FOR irow=rows TO glvnumrows-1
      SendMessage cast(hwnd,rr),LVM_DELETEITEM,0,cast(LPARAM ,rows)
    NEXT
  END IF
END SUB

SUB LVSetValue(BYVAL rr AS uinteger,BYVAL item AS LONG,BYVAL column AS LONG,BYVAL strs AS STRING)
    dim ms_lvi AS LVITEM
    dim psztext as zstring*max_path
    psztext=strs
    ms_lvi.iitem = item
    ms_lvi.iSubItem = column
    ms_lvi.pszText = @pszText
    ms_lvi.lparam = item
    SendMessage cast(hwnd,rr), LVM_SETITEMTEXT, item, cast(LPARAM ,VARPTR(ms_lvi))
END SUB

FUNCTION LVGetValue(BYVAL rr AS uinteger, BYVAL item AS LONG, BYVAL column AS LONG)AS STRING
    dim ms_lvi AS LV_ITEM
    dim psztext as zstring*max_path
    ms_lvi.iSubItem = Column
    ms_lvi.cchTextMax = max_path
    ms_lvi.pszText = @pszText
    SendMessage cast(hwnd,rr), LVM_GETITEMTEXT, item, cast(LPARAM ,@ms_lvi)
    FUNCTION = psztext
END FUNCTION

function lvFindItem(hList AS LONG, Value AS STRING)as integer
    DIM lvf AS LVFINDINFO
    DIM i AS integer
    DIM szStr AS zstring * 300
    DIM iStatus AS INTEGER

    lvf.flags = LVFI_STRING OR LVFI_Partial
    szStr = ucase$(Value)

    lvf.psz = @szStr
    i = SendMessage(cast(hwnd,hList),LVM_FINDITEM, -1, cast(LPARAM ,@lvf))
    IF i >= 0 THEN
        istatus = ListView_EnsureVisible(cast(hwnd,hList), i, true)
    END IF
    function=i
END function

SUB LVSetItem(hList AS LONG, byval item as long,Rec AS STRING )
    DIM z AS INTEGER
    DIM iStatus AS INTEGER
    DIM szStr AS zstring* 300
    DIM lvi AS LV_ITEM
    dim x AS LONG

    'this will be the next record

    lvi.iItem = item
    lvi.mask = LVIF_TEXT
    lvi.stateMask = LVIS_FOCUSED '%LVIS_SELECTED
    lvi.pszText = @szStr

    FOR z = 0 TO str_numparse(Rec,"|")-1
        szStr =str_parse(Rec,"|",z+1)
        lvi.iSubItem = z

        lvi.lParam = lvi.iItem
        IF z = 0 THEN
            lvi.mask = LVIF_TEXT OR LVIF_PARAM OR LVIF_STATE
            iStatus = ListView_setItem (cast(hwnd,hList), @lvi)
        ELSE
            lvi.mask = LVIF_TEXT
            iStatus = ListView_SetItem (cast(hwnd,hList), @lvi)
        END IF
    NEXT
END SUB


SUB LVAddItem(hList AS LONG, Rec AS STRING )
    DIM z AS INTEGER
    DIM iStatus AS INTEGER
    DIM szStr AS zstring* 300
    DIM lvi AS LV_ITEM
    dim x AS LONG

    'this will be the next record

    lvi.iItem = sendmessage(cast(hwnd,hlist),lvm_GetItemCount,0,0)'ListView_GetItemCount(hList) '+ 1

    lvi.mask = LVIF_TEXT
    lvi.stateMask = LVIS_FOCUSED '%LVIS_SELECTED
    lvi.pszText = @szStr

    FOR z = 0 TO str_numparse(Rec,"|")-1
        szStr =str_parse(Rec,"|",z+1)
        lvi.iSubItem = z

        lvi.lParam = lvi.iItem
        IF z = 0 THEN
            lvi.mask = LVIF_TEXT OR LVIF_PARAM OR LVIF_STATE
            iStatus = ListView_InsertItem (cast(hwnd,hList), @lvi)
        ELSE
            lvi.mask = LVIF_TEXT
            iStatus = ListView_SetItem (cast(hwnd,hList), @lvi)
        END IF
    NEXT
END SUB