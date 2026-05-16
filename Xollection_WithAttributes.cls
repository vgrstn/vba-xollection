VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Xollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Extended Collection Class."

'------------------------------------------------------------------------------
' MIT License
'
' Copyright (c) 2025 Vincent van Geerestein
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
'------------------------------------------------------------------------------

Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Author: Vincent van Geerestein
' E-mail: vincent@vangeerestein.com
' Description: Extended Collection Class
' Enumeration: For Each item in object
' Add-in: RubberDuck (https://rubberduckvba.com/)
' Version: 2024.11.24
'
' Methods
' object.Add item [, ...]   Adds an item to the Xollection
' object.Remove index       Removes an item from the Xollection
' object.RemoveAll          Removes all the items from the Xollection
' object.Exists(index)      Returns True if an item exists in the Xollection
' object.Count              Returns the number of items in the Xollection
' object.Items([base])      Returns an array containing all the Xollection items
' object.Keys([base])       Returns an array containing all Xollection keys
' object.Sort()             Sorts the items in the Xollection by their keys
' object.Position(key)      Returns the position of an item in the Xollection
'
' Properties
' object.Item(index)        Sets or returns an item in the Xollection
' object.Key(index)         Sets or returns a key of an item in the Xollection

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Compiler directives
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Use API or the faster Variant ByRef approach for getting Collection Keys.
#Const API = False

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private API declarations
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' https://learn.microsoft.com/en-us/windows-hardware/drivers/ddi/wdm/nf-wdm-rtlmovememory
Private Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
    pDst As Any, _
    pSrc As Any, _
    ByVal NBytes As Long _
)

#If API Then
' https://docs.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-sysreallocstring
Private Declare PtrSafe Function SysReAllocString Lib "oleaut32.dll" ( _
    ByVal pBSTR As LongPtr, _
    Optional ByVal pszStrPtr As LongPtr _
) As Long
#End If

#If API Then
' https://docs.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-variantcopy
Private Declare PtrSafe Function VarCopyToPtr Lib "oleaut32.dll" Alias "VariantCopy" ( _
    ByVal pvDest As LongPtr, _
    ByVal pvSrc As Variant _
) As Long
#End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private declarations
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

#If API Then
' The null pointer constant (not VB defined).
Private Const vbNullPtr As LongPtr = 0
#Else
' The variable subtype constant for a LongPtr (not VB defined).
#If Win64 Then
Private Const vbLongPtr As Integer = VBA.vbLongLong
#Else
Private Const vbLongPtr As Integer = VBA.vbLong
#End If
#End If

' Selected memory sizes of intrinsic data types.
' https://docs.microsoft.com/en-us/office/vba/Language/Reference/User-Interface-Help/data-type-summary
Private Enum VARSIZE
    vbSizeInteger = 2
#If Win64 Then
    vbSizeLongPtr = 8
#Else
    vbSizeLongPtr = 4
#End If
End Enum

' Selected memory offsets for a VB Collection.
Private Enum MEMOFFSET
#If Win64 Then
    vbOffsetCollectionHead = 40         ' *head (LongPtr)
    vbOffsetCollectionTail = 48         ' *tail (LongPtr)
    vbOffsetCollectionKey = 24          ' *key  (LongPtr)
    vbOffsetCollectionPrev = 32         ' *prev (LongPtr)
    vbOffsetCollectionNext = 40         ' *next (LongPtr)
#Else
    vbOffsetCollectionHead = 24
    vbOffsetCollectionTail = 28
    vbOffsetCollectionKey = 16
    vbOffsetCollectionPrev = 20
    vbOffsetCollectionNext = 24
#End If
End Enum

' Selected VB errors.
Private Enum VBERROR
    vbErrorInvalidProcedureCall = 5
    vbErrorSubscriptOutOfRange = 9
End Enum

' Variant ByRef construct for memory access by address.
#If API = False Then
Private Const VT_BYREF As Integer = &H4000
Private Type CONSTRUCT
    vt As Variant
    ref As Variant
End Type
Private VarByRef As CONSTRUCT
#End If

Private Type TPRIVATE
    coll As Collection
End Type
Private this As TPRIVATE


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Class_Initialize()

    Randomize

    Set this.coll = New Collection

#If API = False Then
    InitializeVarByRef
#End If

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Enumerator() As IEnumVARIANT
Attribute Enumerator.VB_UserMemId = -4
' Enumerate is triggered by For Each and sets the IEnumVARIANT interface.

    ' [_NewEnum] is the name of the hidden method to enumerate a VB Collection.
    Set Enumerator = this.coll.[_NewEnum]

End Property


Public Sub Add( _
    ByVal Item As Variant, _
    Optional ByVal Key As Variant, _
    Optional ByVal before As Variant, _
    Optional ByVal after As Variant)
Attribute Add.VB_Description = "Adds an item to the Xollection."

    Select Case True
    Case VBA.IsMissing(before) And VBA.IsMissing(after)
        If VBA.IsMissing(Key) Then
            this.coll.Add Item
        Else
            this.coll.Add Item, Key
        End If
    Case VBA.IsMissing(before)
        If VBA.IsMissing(Key) Then
            this.coll.Add Item, after:=after
        Else
            this.coll.Add Item, Key, after:=after
        End If
    Case VBA.IsMissing(after)
        If VBA.IsMissing(Key) Then
            this.coll.Add Item, before:=before
        Else
            this.coll.Add Item, Key, before:=before
        End If
    Case Else
        ' Before and After can't be provided both.
        Err.Raise vbErrorInvalidProcedureCall
    End Select

End Sub


Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of items in the Xollection."

    Count = this.coll.Count

End Property


Public Function Exists(ByVal index As Variant) As Boolean
Attribute Exists.VB_Description = "Returns True if an item exists in the object or False otherwise."
' VB has no neat way of coding this function other than via error trapping.

    On Error GoTo ErrorHandler
    this.coll.Item index
    Exists = True
    Exit Function

ErrorHandler:
    If Err.Number = vbErrorInvalidProcedureCall Then
        Err.Clear
    Else
        Err.Raise Err.Number
    End If

End Function


Public Property Get Item(ByVal index As Variant) As Variant
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_Description = "Sets or returns an item in the Xollection."

    AssignVariant Item, this.coll.Item(index)

End Property
Public Property Let Item(ByVal index As Variant, ByVal RHS As Variant)
' An item is updated and its position is retained as well as its key.

    If this.coll.Count = 0 Then Err.Raise vbErrorInvalidProcedureCall, , "Xollection is empty"

#If API Then
    CollItem(this.coll, index) = RHS
#Else
    CollItem(this.coll, index, VarByRef.vt, VarByRef.ref) = RHS
#End If

End Property


Public Function Items(Optional ByVal base As Long) As Variant()
Attribute Items.VB_Description = "Returns an array containing all the items in the Xollection."
' Returns an empty array for an empty collection.

    If this.coll.Count = 0 Then Exit Function

    Dim List() As Variant: ReDim List(base To base + this.coll.Count - 1)
    Dim i As Long: i = base
    Dim Item As Variant
    For Each Item In this.coll
        If VBA.IsObject(Item) Then
            Set List(i) = Item
        Else
            List(i) = Item
        End If
        i = i + 1
    Next

    Items = List

End Function


Public Property Get Key(ByVal index As Variant) As String
Attribute Key.VB_Description = "Sets or returns the key of an item the Xollection."
' Returns "" for an item that doesn't have a set key.

    If this.coll.Count = 0 Then Err.Raise vbErrorInvalidProcedureCall, , "Xollection is empty"
    If VBA.IsNumeric(index) Then
        If index <> Int(index) Then Err.Raise vbErrorInvalidProcedureCall, , "Index is invalid"
        If index < 1 Or index > this.coll.Count Then Err.Raise vbErrorSubscriptOutOfRange
    ElseIf VBA.VarType(index) = VBA.vbString Then
        If Exists(index) = False Then Err.Raise vbErrorInvalidProcedureCall, , "ey doesn't exist"
    Else
        Err.Raise vbErrorInvalidProcedureCall, , "Index is invalid"
    End If

#If API Then
    Key = CollItemKey(this.coll, index)
#Else
    Key = CollItemKey(this.coll, index, VarByRef.vt)
#End If

End Property
Public Property Let Key(ByVal index As Variant, ByVal RHS As String)
' The key of an item is updated and its position is retained.

    If this.coll.Count = 0 Then Err.Raise vbErrorInvalidProcedureCall, , "Xollection is empty"

     ' Raise error if the item doesn't exist.
    Dim pos As Long
    If VBA.VarType(index) = VBA.vbString Then
        If Exists(index) = False Then Err.Raise vbErrorInvalidProcedureCall, , "Key doesn't exist"
#If API Then
        pos = CollItemPos(this.coll, index)
#Else
        pos = CollItemPos(this.coll, index, VarByRef.vt)
#End If
    Else
        If index < 1 Or index > this.coll.Count Then Err.Raise vbErrorSubscriptOutOfRange
        pos = index
    End If

    ' Remove the item and replace with same item but new key.
    Dim Item As Variant: AssignVariant Item, this.coll.Item(index)
    If this.coll.Count = pos Then
        this.coll.Remove index
        this.coll.Add Item, RHS
    Else
        this.coll.Remove index
        this.coll.Add Item, RHS, before:=pos
    End If

End Property


Public Function Keys(Optional ByVal base As Long) As String()
Attribute Keys.VB_Description = "Returns an array containing all the keys in the Xollection."
' Returns an empty string array for an empty Collection.

    If this.coll.Count = 0 Then Exit Function

#If API Then
    Keys = CollKeys(this.coll, base)
#Else
    Keys = CollKeys(this.coll, base, VarByRef.vt)
#End If

End Function


Public Function Position(ByVal Key As String) As Long
Attribute Position.VB_Description = "Returns the index of an item in the Xollection specified by its key."
' Returns 0 if no item with doesn't exist.

    If this.coll.Count = 0 Then Err.Raise vbErrorInvalidProcedureCall, , "Xollection is empty"
    If Me.Exists(Key) = False Then Exit Function

#If API Then
    Position = CollItemPos(this.coll, Key)
#Else
    Position = CollItemPos(this.coll, Key, VarByRef.vt)
#End If

End Function


Public Sub Remove(ByVal index As Variant)
Attribute Remove.VB_Description = "Removes an item from the Xollection."

    this.coll.Remove index

End Sub


Public Sub RemoveAll()
Attribute RemoveAll.VB_Description = "Removes all the items from the Xollection."

    Do While this.coll.Count <> 0
        this.coll.Remove 1
    Loop

End Sub


Public Sub Sort()
Attribute Sort.VB_Description = "Sorts the items in the Xollection."

    If this.coll.Count <= 1 Then Exit Sub

    ' Extract keys array.
    Dim Keys() As String: Keys = Me.Keys(1)

    ' Initialize index array.
    Dim idx() As Long: ReDim idx(LBound(Keys) To UBound(Keys))
    Dim i As Long
    For i = LBound(Keys) To UBound(Keys)
        idx(i) = i
    Next

    ' Sort keys array by index.
    QuickSortByIndex Keys, idx

    ' Extract items array.
    Dim Items() As Variant: Items = Me.Items(LBound(Keys))

    ' Replace the current Collection with the sorted Collection
    Set this.coll = New Collection
    For i = LBound(idx) To UBound(idx)
        If Len(Keys(idx(i))) = 0 Then
            this.coll.Add Items(idx(i))
        Else
            this.coll.Add Items(idx(i)), Keys(idx(i))
        End If
    Next

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub AssignVariant(ByRef variable As Variant, ByVal value As Variant)
Attribute AssignVariant.VB_Description = "Assigns a value or a Xollection reference to a variable."
' Prevents indexing a collection item twice at the expense of a stack frame.

    If VBA.IsObject(value) Then
        Set variable = value
    Else
        variable = value
    End If

End Sub


#If API = False Then
Private Sub InitializeVarByRef()
Attribute InitializeVarByRef.VB_Description = "Initializes the Variant ByRef construct."

    VarByRef.vt = VarPtr(VarByRef.ref)
    CopyMemory VarByRef.vt, VBA.vbInteger Or VT_BYREF, vbSizeInteger

End Sub
#End If


#If API Then
Private Function CollKeys( _
    ByVal coll As Collection, _
    ByVal base As Long _
) As String()
Attribute CollKeys.VB_Description = "Returns an array containing the keys of all the items in a VB Collection."
' Returns an empty string array for an empty Collection.

    ' The null pointer constant (not VB defined).
    Const vbNullPtr As LongPtr = 0

    If coll.Count = 0 Then Exit Function

    ' Traverse the double linked list and copy all the key strings.
    Dim Keys() As String: ReDim Keys(base To base + coll.Count - 1)
    Dim NodePtr As LongPtr, KeyPtr As LongPtr
    NodePtr = ObjPtr(coll) + vbOffsetCollectionHead
    Dim i As Long
    For i = base To base + coll.Count - 1
        CopyMemory NodePtr, ByVal NodePtr, vbSizeLongPtr
        CopyMemory KeyPtr, ByVal NodePtr + vbOffsetCollectionKey, vbSizeLongPtr
        If KeyPtr <> vbNullPtr Then
            SysReAllocString VarPtr(Keys(i)), KeyPtr
        End If
        NodePtr = NodePtr + vbOffsetCollectionNext
    Next

    ' Return all the item keys.
    CollKeys = Keys

End Function
#Else
Private Function CollKeys( _
    ByVal coll As Collection, _
    ByVal base As Long, _
    ByRef vt As Variant _
) As String()
Attribute CollKeys.VB_Description = "Returns an array containing the keys of all the items in a VB Collection."

    ' Traverse the double linked list and copy all the key strings.
    Dim Keys() As String: ReDim Keys(base To base + coll.Count - 1)
    Dim NodePtr As LongPtr
    NodePtr = ObjPtr(coll) + vbOffsetCollectionHead - vbOffsetCollectionNext
    Dim i As Long
    For i = base To base + coll.Count - 1
        VarByRef.ref = NodePtr + vbOffsetCollectionNext
        vt = vbLongPtr Or VT_BYREF
        NodePtr = VarByRef.ref
        VarByRef.ref = NodePtr + vbOffsetCollectionKey
        vt = VBA.vbString Or VT_BYREF
        Keys(i) = VarByRef.ref
    Next

    ' Return all the item keys.
    CollKeys = Keys

End Function
#End If


#If API Then
Private Function CollItemKey( _
    ByVal coll As Collection, _
    ByVal pos As Long _
) As String
Attribute CollItemKey.VB_Description = "Returns the key of an item in a VB Collection specified by its position."
' Returns "" for an item in the Collection that doesn't have a key.

    ' Traverse the double linked list depending on the shortest path to the node position.
    Dim i As Long, NodePtr As LongPtr
    If pos <= coll.Count \ 2 Then
        NodePtr = ObjPtr(coll) + vbOffsetCollectionHead - vbOffsetCollectionNext
        For i = 1 To pos
            NodePtr = NodePtr + vbOffsetCollectionNext
            CopyMemory NodePtr, ByVal NodePtr, vbSizeLongPtr
        Next
    Else
        NodePtr = ObjPtr(coll) + vbOffsetCollectionTail - vbOffsetCollectionPrev
        For i = coll.Count To pos Step -1
            NodePtr = NodePtr + vbOffsetCollectionPrev
            CopyMemory NodePtr, ByVal NodePtr, vbSizeLongPtr
        Next
    End If

    ' Return the key for the node.
    Dim KeyPtr As LongPtr
    CopyMemory KeyPtr, ByVal NodePtr + vbOffsetCollectionKey, vbSizeLongPtr
    If KeyPtr <> vbNullPtr Then
        SysReAllocString VarPtr(CollItemKey), KeyPtr
    End If

End Function
#Else
Private Function CollItemKey( _
    ByVal coll As Collection, _
    ByVal pos As Long, _
    ByRef vt As Variant _
) As String
Attribute CollItemKey.VB_Description = "Returns the key of an item in a VB Collection specified by its position."
' Returns "" for an item in the Collection that doesn't have a key.

    ' Traverse the double linked list depending on the shortest path to the node position.
    Dim i As Long, NodePtr As LongPtr
    If pos <= coll.Count \ 2 Then
        NodePtr = ObjPtr(coll) + vbOffsetCollectionHead - vbOffsetCollectionNext
        For i = 1 To pos
            VarByRef.ref = NodePtr + vbOffsetCollectionNext
            vt = vbLongPtr Or VT_BYREF
            NodePtr = VarByRef.ref
        Next
    Else
        NodePtr = ObjPtr(coll) + vbOffsetCollectionTail - vbOffsetCollectionPrev
        For i = coll.Count To pos Step -1
            VarByRef.ref = NodePtr + vbOffsetCollectionPrev
            vt = vbLongPtr Or VT_BYREF
            NodePtr = VarByRef.ref
        Next
    End If

    ' Return the key for the node.
    VarByRef.ref = NodePtr + vbOffsetCollectionKey
    vt = VBA.vbString Or VT_BYREF
    CollItemKey = VarByRef.ref

End Function
#End If


#If API Then
Private Function CollItemPos( _
    ByVal coll As Collection, _
    ByVal Key As String _
) As Long
Attribute CollItemPos.VB_Description = "Returns the position of an item in a VB Collection specified by its key."
' Returns 0 if no item with the key exists in the Collection.

    ' Traverse the double linked list until a matching key has been found.
    Dim NodePtr As LongPtr, KeyPtr As LongPtr, ItemKey As String
    NodePtr = ObjPtr(coll) + vbOffsetCollectionHead - vbOffsetCollectionNext
    Dim i As Long
    For i = 1 To coll.Count
        NodePtr = NodePtr + vbOffsetCollectionNext
        CopyMemory NodePtr, ByVal NodePtr, vbSizeLongPtr
        CopyMemory KeyPtr, ByVal NodePtr + vbOffsetCollectionKey, vbSizeLongPtr
        If KeyPtr <> vbNullPtr Then
            SysReAllocString VarPtr(ItemKey), KeyPtr
            If StrComp(ItemKey, Key, VBA.vbTextCompare) = 0 Then
                CollItemPos = i
                Exit Function
            End If
        End If
    Next

End Function
#Else
Private Function CollItemPos( _
    ByVal coll As Collection, _
    ByVal Key As String, _
    ByRef vt As Variant _
) As Long
Attribute CollItemPos.VB_Description = "Returns the position of an item in a VB Collection specified by its key."
' Returns 0 if no item with the key exist in the Collection.

    ' Traverse the double linked list until a matching key has been found.
    Dim NodePtr As LongPtr, ItemKey As String
    NodePtr = ObjPtr(coll) + vbOffsetCollectionHead - vbOffsetCollectionNext
    Dim i As Long
    For i = 1 To coll.Count
        VarByRef.ref = NodePtr + vbOffsetCollectionNext
        vt = vbLongPtr Or VT_BYREF
        NodePtr = VarByRef.ref
        VarByRef.ref = NodePtr + vbOffsetCollectionKey
        vt = VBA.vbString Or VT_BYREF
        ItemKey = VarByRef.ref
        If StrComp(ItemKey, Key, VBA.vbTextCompare) = 0 Then
            CollItemPos = i
            Exit Function
        End If
    Next

End Function
#End If


#If API Then
Private Property Let CollItem( _
    ByVal coll As Collection, _
    ByVal index As Variant, _
    ByVal RHS As Variant _
)
Attribute CollItem.VB_Description = "Sets the value of an item in a VB Collection specified by its index."

   If VBA.VarType(index) = VBA.vbString Then
        If Me.Exists(index) = False Then Err.Raise vbErrorInvalidProcedureCall
        Dim NodePtr As LongPtr, KeyPtr As LongPtr, ItemKey As String
        NodePtr = ObjPtr(coll) + vbOffsetCollectionHead - vbOffsetCollectionNext
        Dim i As Long
        For i = 1 To coll.Count
            NodePtr = NodePtr + vbOffsetCollectionNext
            CopyMemory NodePtr, ByVal NodePtr, vbSizeLongPtr
            CopyMemory KeyPtr, ByVal NodePtr + vbOffsetCollectionKey, vbSizeLongPtr
            If KeyPtr <> vbNullPtr Then
                SysReAllocString VarPtr(ItemKey), KeyPtr
                If StrComp(ItemKey, index, VBA.vbTextCompare) = 0 Then
                    Exit For
                End If
            End If
        Next
    ElseIf VBA.IsNumeric(index) Then
        If index <= coll.Count \ 2 Then
            NodePtr = ObjPtr(coll) + vbOffsetCollectionHead - vbOffsetCollectionNext
            For i = 1 To index
                NodePtr = NodePtr + vbOffsetCollectionNext
                CopyMemory NodePtr, ByVal NodePtr, vbSizeLongPtr
            Next
        Else
            NodePtr = ObjPtr(coll) + vbOffsetCollectionTail - vbOffsetCollectionPrev
            For i = coll.Count To index Step -1
                NodePtr = NodePtr + vbOffsetCollectionPrev
                CopyMemory NodePtr, ByVal NodePtr, vbSizeLongPtr
            Next
        End If
    Else
        Err.Raise vbErrorInvalidProcedureCall, , "Index is invalid"
    End If

    ' Update the variant item in the node.
    VarCopyToPtr NodePtr, RHS

End Property
#Else
Private Property Let CollItem( _
    ByVal coll As Collection, _
    ByVal index As Variant, _
    ByRef vt As Variant, _
    ByRef ref As Variant, _
    ByVal RHS As Variant _
)
Attribute CollItem.VB_Description = "Sets the value of an item in a VB Collection specified by its index."

   If VBA.VarType(index) = VBA.vbString Then
        If Me.Exists(index) = False Then Err.Raise vbErrorInvalidProcedureCall
        Dim NodePtr As LongPtr, ItemKey As String
        NodePtr = ObjPtr(coll) + vbOffsetCollectionHead - vbOffsetCollectionNext
        Dim i As Long
        For i = 1 To coll.Count
            VarByRef.ref = NodePtr + vbOffsetCollectionNext
            vt = vbLongPtr Or VT_BYREF
            NodePtr = VarByRef.ref
            VarByRef.ref = NodePtr + vbOffsetCollectionKey
            vt = VBA.vbString Or VT_BYREF
            ItemKey = VarByRef.ref
            If StrComp(ItemKey, index, VBA.vbTextCompare) = 0 Then
                Exit For
            End If
        Next
    ElseIf VBA.IsNumeric(index) Then
        If index <= coll.Count \ 2 Then
            NodePtr = ObjPtr(coll) + vbOffsetCollectionHead - vbOffsetCollectionNext
            For i = 1 To index
                VarByRef.ref = NodePtr + vbOffsetCollectionNext
                vt = vbLongPtr Or VT_BYREF
                NodePtr = VarByRef.ref
            Next
        Else
            NodePtr = ObjPtr(coll) + vbOffsetCollectionTail - vbOffsetCollectionPrev
            For i = coll.Count To index Step -1
                VarByRef.ref = NodePtr + vbOffsetCollectionPrev
                vt = vbLongPtr Or VT_BYREF
                NodePtr = VarByRef.ref
            Next
        End If
    Else
        Err.Raise vbErrorInvalidProcedureCall, , "Index is invalid"
    End If

    ' Update the variant item in the node.
    VarByRef.ref = NodePtr
    vt = vbVariant Or VT_BYREF
    If VBA.IsObject(RHS) Then
        Set ref = RHS
    Else
        ref = RHS
    End If

End Property
#End If


Private Sub QuickSortByIndex( _
    ByRef arr As Variant, _
    ByRef idx As Variant _
)
Attribute QuickSortByIndex.VB_Description = "Sorts an array by index."
' Iterative algorithm after Hardcore Visual Basic version 5.0 (Bruce McKinney).

    ' Threshold for applying InsertionSort (taken from Java).
    Const INSERTION_SORT_THRESHOLD As Long = 47

    ' An empty stack has one item with prevents a test for a push operation.
    Dim Stack As Collection: Set Stack = New Collection: Stack.Add Empty

    Dim lower As Long: lower = LBound(arr)
    Dim upper As Long: upper = UBound(arr)
    Dim i As Long, j As Long, Pivot As Variant, x As Long, index As Long
    Do
      Do
        If lower < upper Then
            ' Swap from ends until first and last meet in the middle
            If upper - lower < INSERTION_SORT_THRESHOLD Then
                InsertionSortByIndex arr, lower, upper, idx
            Else
                ' Start with random partitioning.
                index = Int((upper - lower + 1) * Rnd + lower)
                x = idx(upper): idx(upper) = idx(index): idx(index) = x
                ' Swap high values below the split for low values above.
                i = lower
                j = upper
                ' Save current pivot value for efficiency.
                Pivot = arr(idx(upper))

                Do
                    ' Find any low value > split.
                    For i = i To j - 1
                        If Pivot < arr(idx(i)) Then Exit For
                    Next
                    ' Find any high value < split.
                    For j = j To i + 1 Step -1
                        If Pivot > arr(idx(j)) Then Exit For
                    Next
                    ' Swap too high low value for too low high value.
                    If i < j Then
                        x = idx(i): idx(i) = idx(j): idx(j) = x
                    End If
                Loop While i < j

                If i <> upper Then
                    x = idx(i): idx(i) = idx(upper): idx(upper) = x
                End If

                ' Push range markers of larger part for later sorting.
                If i - lower < upper - i Then
                    Stack.Add i + 1, before:=1
                    Stack.Add upper, before:=1
                    upper = i - 1
                Else
                    Stack.Add lower, before:=1
                    Stack.Add i - 1, before:=1
                    lower = i + 1
                End If

                ' Exit from inner loop to process smaller part.
                Exit Do
            End If
        End If

        If Stack.Count = 1 Then Exit Sub

        ' Pop range markers for next partition.
        upper = Stack.Item(1): Stack.Remove 1
        lower = Stack.Item(1): Stack.Remove 1
      Loop
    Loop

End Sub


Private Sub InsertionSortByIndex( _
    ByRef arr As Variant, _
    ByVal lower As Long, _
    ByVal upper As Long, _
    ByRef idx As Variant _
)
Attribute InsertionSortByIndex.VB_Description = "Sorts a subarray by index."

    Dim Item As Variant, index As Long
    Dim i As Long, j As Long
    For i = lower + 1 To upper
        index = idx(i)
        Item = arr(index)
        For j = i To lower + 1 Step -1
            If Item > arr(idx(j - 1)) Then Exit For
            idx(j) = idx(j - 1)
        Next
        idx(j) = index
    Next

End Sub
