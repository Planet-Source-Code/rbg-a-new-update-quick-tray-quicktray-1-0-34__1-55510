Attribute VB_Name = "modListView"
Option Explicit

Public objFind As LV_FINDINFO
Public objItem As LV_ITEM

Public sOrder As Boolean

Public Type POINTAPI
 x As Long
 y As Long
End Type

Public Type LV_FINDINFO
 flags As Long
 psz As String
 lParam As Long
 pt As POINTAPI
 vkDirection As Long
End Type

Public Type LV_ITEM
 mask As Long
 iItem As Long
 iSubItem As Long
 state As Long
 stateMask As Long
 pszText As String
 cchTextMax As Long
 iImage As Long
 lParam As Long
 iIndent As Long
End Type

Public Const LVFI_PARAM As Long = &H1
Public Const LVIF_TEXT As Long = &H1
Public Const LVM_FIRST As Long = &H1000
Public Const LVM_FINDITEM As Long = (LVM_FIRST + 13)
Public Const LVM_GETITEMTEXT As Long = (LVM_FIRST + 45)
Public Const LVM_SORTITEMS As Long = (LVM_FIRST + 48)

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Function CompareDates(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal hWnd As Long) As Long
Dim dDate1 As Date
Dim dDate2 As Date
 dDate1 = ListView_GetItemDate(hWnd, lParam1)
 dDate2 = ListView_GetItemDate(hWnd, lParam2)
 Select Case sOrder
  Case True: 'sort descending
   If dDate1 < dDate2 Then
    CompareDates = 0
   ElseIf dDate1 = dDate2 Then
    CompareDates = 1
   Else: CompareDates = 2
   End If
  Case Else: 'sort ascending
   If dDate1 > dDate2 Then
    CompareDates = 0
   ElseIf dDate1 = dDate2 Then
    CompareDates = 1
   Else: CompareDates = 2
   End If
  End Select
End Function

Public Function CompareValues(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal hWnd As Long) As Long
Dim val1 As Long
Dim val2 As Long
 val1 = ListView_GetItemValueStr(hWnd, lParam1)
 val2 = ListView_GetItemValueStr(hWnd, lParam2)
 Select Case sOrder
  Case True: 'sort descending
   If val1 < val2 Then
    CompareValues = 0
   ElseIf val1 = val2 Then
    CompareValues = 1
   Else: CompareValues = 2
   End If
  Case Else: 'sort ascending
   If val1 > val2 Then
    CompareValues = 0
   ElseIf val1 = val2 Then
    CompareValues = 1
   Else: CompareValues = 2
   End If
  End Select
End Function

Public Function FARPROC(ByVal pfn As Long) As Long
 FARPROC = pfn
End Function

Public Function ListView_GetItemDate(hWnd As Long, lParam As Long) As Date
Dim hIndex As Long
Dim r As Long
 objFind.flags = LVFI_PARAM
 objFind.lParam = lParam
 hIndex = SendMessage(hWnd, LVM_FINDITEM, -1, objFind)
 objItem.mask = LVIF_TEXT
 objItem.iSubItem = 1
 objItem.pszText = Space$(32)
 objItem.cchTextMax = Len(objItem.pszText)
 r = SendMessage(hWnd, LVM_GETITEMTEXT, hIndex, objItem)
 If r > 0 Then ListView_GetItemDate = CDate(Left$(objItem.pszText, r))
End Function

Public Function ListView_GetItemValueStr(hWnd As Long, lParam As Long) As Long
Dim hIndex As Long
Dim r As Long
 objFind.flags = LVFI_PARAM
 objFind.lParam = lParam
 hIndex = SendMessage(hWnd, LVM_FINDITEM, -1, objFind)
 objItem.mask = LVIF_TEXT
 objItem.iSubItem = 2
 objItem.pszText = Space$(32)
 objItem.cchTextMax = Len(objItem.pszText)
 r = SendMessage(hWnd, LVM_GETITEMTEXT, hIndex, objItem)
 If r > 0 Then If IsNumeric((Trim$(objItem.pszText))) Then ListView_GetItemValueStr = CLng(Left$(Trim$(objItem.pszText), r))
End Function
