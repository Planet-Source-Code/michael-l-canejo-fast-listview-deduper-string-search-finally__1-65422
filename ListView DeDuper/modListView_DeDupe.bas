Attribute VB_Name = "modListView"

Option Explicit

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type LVFINDINFO
    flags As Long
    psz As String
    lParam As Long
    pt As POINTAPI
    vkDirection As Long
End Type

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
                                                  ByVal hwnd As Long, _
                                                  ByVal wMsg As Long, _
                                                  ByVal wParam As Long, _
                                                  lParam As Any) _
                                                  As Long
Private Const LVM_FIRST = &H1000
Private Const LVM_FINDITEM = (LVM_FIRST + 13)
Private Const LVFI_PARAM = &H1
Private Const LVFI_STRING = &H2
Private Const LVFI_PARTIAL = &H8
Private Const LVFI_WRAP = &H20
Private Const LVFI_NEARESTXY = &H40
Private Const LVM_DELETEITEM = (LVM_FIRST + 8)


Public Function DeDupe_ListView_API(lsv As ListView) As Long

 Dim x As Long, lRet As Long, lngDupesFound As Long

 For x = 1 To lsv.ListItems.Count
  
  Do While lRet <> -1 _
           And x <= lsv.ListItems.Count

   lRet = FindIn_ListView(lsv, lsv.ListItems.Item(x), x - 1) 'find string in listview, returns index

   If lRet <> -1 Then
    If lRet + 1 = x Then Exit Do 'Dont remove original
    lsv.ListItems.Remove x 'remove dupe
    lngDupesFound = lngDupesFound + 1
   End If
  Loop
  
 Next
 
 DeDupe_ListView_API = lngDupesFound
 
End Function

Public Function DeDupe_ListView_Col(lsv As ListView) As Long

 Dim cCol    As Collection
 Dim lCt     As Long
 Dim lngDupesFound As Long
 
 On Error Resume Next
 
 Set cCol = New Collection
 
 lsv.Sorted = False
 
 For lCt = lsv.ListItems.Count To 1 Step -1
  cCol.Add 1, lsv.ListItems.Item(lCt).Text
  If Err.Number = 457 Then
   lsv.ListItems.Remove lCt
   lngDupesFound = lngDupesFound + 1
   Err.Clear
  End If
 Next
 
 DeDupe_ListView_Col = lngDupesFound
 
End Function

Public Function FindIn_ListView(lsv As ListView, strFindString As String, _
                                Optional lngStartPos As Long = 1) As Long

 Dim LFI As LVFINDINFO, lRet As Long
 
 LFI.flags = LVFI_PARTIAL Or LVFI_WRAP
 LFI.psz = strFindString
 
 lRet = SendMessage(lsv.hwnd, LVM_FINDITEM, lngStartPos, LFI) 'return index of found string
 
 FindIn_ListView = lRet
 
End Function
