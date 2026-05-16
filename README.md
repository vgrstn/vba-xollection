# vba-xollection
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
![Platform](https://img.shields.io/badge/Platform-VBA%20(Excel%2C%20Access%2C%20Word%2C%20Outlook%2C%20PowerPoint)-blue)
![Architecture](https://img.shields.io/badge/Architecture-x86%20%7C%20x64-lightgrey)
![Rubberduck](https://img.shields.io/badge/Rubberduck-Ready-orange)

VBA class module for an extended Collection — key read-back, in-place item update, key reassignment, `For Each` support, sorted enumeration, and position lookup, all backed by direct traversal of the Collection's internal doubly-linked list.

---

## 📦 Features

- **Key read-back** — `Key(index)` and `Keys([base])` return key strings for any item or all items
- **In-place item update** — `Item(index) = value` writes directly to the node; position and key are preserved
- **Key reassignment** — `Key(index) = newKey` changes a key while retaining position and item
- **Position lookup** — `Position(key)` returns the 1-based index for a key in O(n)
- **Existence check** — `Exists(index)` returns `True` for both numeric positions and key strings
- **Sorted enumeration** — `Sort` rebuilds the Collection in ascending key order
- **`For Each`** — delegates to the native Collection `IEnumVARIANT` via `[_NewEnum]`
- **Full Collection API** — `Add`, `Remove`, `RemoveAll`, `Count`, `Item`, `Items`, `Keys`
- **Fast pointer reads** — Variant ByRef construct (no kernel transitions, no extra stack frames) by default
- **API mode** — switch to `CopyMemory` + `SysReAllocString` + `VarCopyToPtr` with `#Const API = True`
- x86 / x64 compatible via `LongPtr` and `#If Win64`
- Pure VBA, zero dependencies, Rubberduck-friendly annotations

---

## 📁 Files

| File | Type | Description |
|---|---|---|
| `Xollection.cls` | Class | Extended Collection class with [Rubberduck](https://rubberduckvba.com/) annotations |
| `Xollection_WithAttributes.cls` | Class | Ready-to-import version with VB attributes baked in — no Rubberduck required |

Import `Xollection_WithAttributes.cls` if you are not using Rubberduck.

---

## 🚀 Quick Start

```vb
Dim xc As New Xollection

xc.Add "apple",  "a"
xc.Add "banana", "b"
xc.Add "cherry", "c"

Debug.Print xc.Count          ' -> 3
Debug.Print xc.Item("b")      ' -> banana    (default member)
Debug.Print xc("b")           ' -> banana    (default member shorthand)
Debug.Print xc.Key(2)         ' -> b
Debug.Print xc.Position("c")  ' -> 3

xc.Item("b") = "blueberry"    ' update in-place — position and key are retained
Debug.Print xc.Item(2)        ' -> blueberry

xc.Key(1) = "apricot"         ' reassign key — item and position are retained
Debug.Print xc.Key(1)         ' -> apricot

Dim v As Variant
For Each v In xc              ' enumerates in insertion order
    Debug.Print v
Next

xc.Sort                       ' sort ascending by key
Debug.Print xc.Key(1)         ' -> apricot  (a < b < c lexicographically)

Dim k() As String:  k = xc.Keys(1)   ' 1-based string array of all keys
Dim it() As Variant: it = xc.Items(1) ' 1-based Variant array of all items
```

---

## ⚙️ Public Interface

| Member | Description |
|---|---|
| `Add(Item [, Key] [, Before] [, After])` | Adds an item to the Xollection, forwarding all arguments to the underlying VB Collection. |
| `Remove(index)` | Removes an item by numeric position or key string. |
| `RemoveAll()` | Removes all items. |
| `Exists(index) As Boolean` | Returns `True` if a numeric position or key string is valid. Uses error trapping on the underlying Collection. |
| `Count As Long` | Returns the number of items. |
| `Item(index) As Variant` *(default member)* | Gets or sets the item at a numeric position or key string. Setting writes directly to the node — position and key are preserved. |
| `Key(index) As String` | Gets or sets the key at a numeric position or existing key string. Setting removes and re-adds the item at the same position with the new key. Returns `""` for unkeyed items. |
| `Items([base]) As Variant()` | Returns a `base`-indexed array of all items in insertion order. Returns an empty array for an empty Xollection. |
| `Keys([base]) As String()` | Returns a `base`-indexed string array of all keys. Unkeyed items produce `""`. Returns an empty array for an empty Xollection. |
| `Position(Key) As Long` | Returns the 1-based position of the first item with the given key. Returns 0 if the key does not exist. |
| `Sort()` | Rebuilds the Collection in ascending key order. Items without keys sort to the front. |
| `Enumerator As IEnumVARIANT` | Exposes `For Each` by delegating to the native Collection `[_NewEnum]`. |

---

## ⚙️ Compiler directive

| Directive | Default | Description |
|---|---|---|
| `#Const API` | `False` | `False` — Variant ByRef construct (fast, zero extra API imports). `True` — `CopyMemory` + `SysReAllocString` + `VarCopyToPtr`. |

`#Const API = False` is the default and recommended setting. The Variant ByRef construct avoids kernel-mode transitions (no `CopyMemory` per access) and avoids extra stack frames, making pointer reads faster than the API path. `#Const API = True` uses explicit Windows API calls and is easier to follow under a debugger or static analyser.

---

## 🧠 How it works

VBA's built-in `Collection` stores items in a doubly-linked list of heap-allocated nodes. Each node contains the item `Variant`, a BSTR key pointer, and forward/backward node pointers. The COM interface deliberately exposes no way to read a key or to write an item without removing and re-adding it.

`Xollection` keeps one private `Collection` for all storage and COM lifecycle, then operates directly on its internal memory for the features the COM interface withholds:

- Key reads traverse the linked list from `pvHead` or `pvTail` and dereference each node's BSTR pointer.
- In-place item writes navigate to the node and overwrite its `Variant` slot directly.
- Key reassignment falls back to `Remove` + `Add` at the same position — the only operation that allocates a new node.

All pointer reads use the Variant ByRef construct by default: a one-time setup that lets VBA dereference arbitrary addresses by reassigning the VT tag of a module-level `Variant`, with zero API calls per access.

---

## 🧠 Implementation notes

### VB Collection — doubly-linked list layout

The `Collection` COM object contains two pointer fields: `pvHead` (pointing to the first node) and `pvTail` (pointing to the last). Each node has this memory layout:

```
x64 (byte offsets from node base):
  [ 0 – 23]  Item   (Variant, 24 bytes on x64)
  [24]        pvKey  (LongPtr → BSTR; null if no key was provided)
  [32]        pvPrev (LongPtr → previous node; null at head)
  [40]        pvNext (LongPtr → next node; null at tail)

x86 (byte offsets from node base):
  [ 0 – 15]  Item   (Variant, 16 bytes on x86)
  [16]        pvKey  (LongPtr → BSTR)
  [20]        pvPrev
  [24]        pvNext
```

`pvHead` lives at `ObjPtr(coll) + vbOffsetCollectionHead`; `pvTail` at `+ vbOffsetCollectionTail`. These offsets are encoded in the private `MEMOFFSET` enum and were determined by inspection. All pointer arithmetic uses `LongPtr` and `#If Win64` guards for x86/x64 portability.

### `InitializeVarByRef` — the VT-tag write trick

```vb
Private Sub InitializeVarByRef()
    VarByRef.vt = VarPtr(VarByRef.ref)
    CopyMemory VarByRef.vt, VBA.vbInteger Or VT_BYREF, vbSizeInteger
End Sub
```

`VarByRef` is a module-level `CONSTRUCT` UDT with two `Variant` fields, `vt` and `ref`.

**Step 1** — `VarByRef.vt = VarPtr(VarByRef.ref)` stores the address of `VarByRef.ref` in the data bytes of `vt` (as `VT_I4` on x86 or `VT_I8` on x64).

**Step 2** — `CopyMemory VarByRef.vt, vbInteger Or VT_BYREF, 2` overwrites only the first 2 bytes of `vt` (the VT tag) with `0x4002` (`VT_BYREF|VT_I2`). The address in the data bytes is untouched.

After setup, `VarByRef.vt` is a `VT_BYREF|VT_I2` Variant — a pointer-to-Integer whose data holds `VarPtr(VarByRef.ref)`. That means it points at the first 2 bytes of `VarByRef.ref`, which are `ref`'s own VT tag.

**The lever**: assigning a 2-byte integer through a `ByRef vt As Variant` parameter that holds `VarByRef.vt` rewrites the VT tag of `VarByRef.ref` without touching its data bytes. The data bytes carry the address; the VT tag is swapped at runtime to tell VBA how to interpret it:

```vb
VarByRef.ref = someAddress       ' data bytes of ref = address (VT tag = VT_I4/VT_I8 for now)
vt = vbLongPtr Or VT_BYREF      ' writes VT_BYREF|LongPtr into ref's VT tag bytes
NodePtr = VarByRef.ref           ' VBA dereferences: reads LongPtr at someAddress
```

One `CopyMemory` call at class initialization; **zero API calls per pointer read or write** thereafter. Each pointer-chase costs two assignments and one read.

This construct is narrower in scope than DMA's SafeArray approach — DMA redirects a typed array's `pvData` to access memory with correct stride for any element size; the Variant ByRef construct handles one scalar at a time with runtime VT switching.

### `CollKeys` — full key traversal

The **pre-start trick** initialises `NodePtr` one step before the first node so the loop body always advances before reading:

```vb
NodePtr = ObjPtr(coll) + vbOffsetCollectionHead - vbOffsetCollectionNext
For i = base To base + coll.Count - 1
    VarByRef.ref = NodePtr + vbOffsetCollectionNext   ' address of Next field
    vt = vbLongPtr Or VT_BYREF                        ' dereference → NodePtr = *Next
    NodePtr = VarByRef.ref
    VarByRef.ref = NodePtr + vbOffsetCollectionKey    ' address of Key field
    vt = vbString Or VT_BYREF                         ' dereference → BSTR copy
    Keys(i) = VarByRef.ref
Next
```

Two pointer-chases per node: one for `pvNext`, one for `pvKey`. No special-case for the first iteration.

When `vt = vbString Or VT_BYREF`, `VarByRef.ref` becomes `VT_BYREF|VT_BSTR`. VBA reads the BSTR pointer stored at the key-field address, dereferences it, and produces a `String` copy. If the key field is null (item was added without a key), the BSTR pointer value is zero; VBA returns `""` rather than faulting — so unkeyed items silently produce empty key strings.

**API mode** uses `SysReAllocString(VarPtr(Keys(i)), KeyPtr)` instead: reusing the existing BSTR allocation in `Keys(i)` each iteration rather than creating a fresh one. Requires an explicit null guard (`If KeyPtr <> vbNullPtr`) before each call.

### `CollItemKey` — bidirectional single-key lookup

The API version picks the shorter path:

```vb
If pos <= coll.Count \ 2 Then
    ' walk forward from pvHead, following pvNext
Else
    ' walk backward from pvTail, following pvPrev
End If
```

Maximum path length is `Count / 2` nodes; average is `Count / 4`. The non-API version always walks forward from `pvHead` — O(n) for any position. This asymmetry means the non-API path is slower for positions in the second half of a large collection, but for the operations `Xollection` exposes publicly (`Key Get`, `Key Let`, `Position`), the difference is rarely measurable in practice.

`@Ignore NonReturningFunction` is applied because the return value is written via the VT_BYREF trick (`CollItemKey = VarByRef.ref`) rather than a normal assignment, which Rubberduck incorrectly flags as a missing return path.

### `CollItemPos` — position search by key

```vb
NodePtr = ObjPtr(coll) + vbOffsetCollectionHead - vbOffsetCollectionNext
For i = 1 To coll.Count
    ...advance NodePtr via pvNext...
    ItemKey = VarByRef.ref          ' read key string via VT_BYREF|VT_BSTR
    If StrComp(ItemKey, Key, vbTextCompare) = 0 Then
        CollItemPos = i
        Exit Function
    End If
Next
```

Always walks forward from head, comparing keys with `vbTextCompare` — matching VB Collection's own case-insensitive key semantics. Returns on first match; returns 0 if no match. In non-API mode, unkeyed nodes produce `""` and a `StrComp("", key)` naturally mismatches any non-empty key, so no null guard is needed. In API mode, null BSTR pointers are skipped explicitly before calling `SysReAllocString`.

### `CollItem` Let — direct in-place Variant write

The standard VB Collection offers no way to update an item's value without changing its position. `CollItem Let` navigates to the node and overwrites its Item slot:

**API path:**
```vb
VarCopyToPtr NodePtr, RHS
```
`VarCopyToPtr` is `VariantCopy` from oleaut32, called with a raw destination address. It handles COM `AddRef` for objects and BSTR duplication for strings.

**Non-API path:**
```vb
VarByRef.ref = NodePtr
vt = vbVariant Or VT_BYREF
If VBA.IsObject(RHS) Then
    Set ref = RHS
Else
    ref = RHS
End If
```
`vt = vbVariant Or VT_BYREF` makes `VarByRef.ref` a `VT_BYREF|VT_VARIANT` Variant pointing to the node's Item slot at `NodePtr`. The `ByRef ref As Variant` parameter shares this memory location. Assigning `Set ref = RHS` or `ref = RHS` writes through the pointer — VBA handles COM reference counting and BSTR duplication correctly for both branches.

`CollItem` is a `Property Let` with additional parameters before `RHS`. Extra parameters in a VBA property accessor appear in the call-site argument list: `CollItem(coll, index, VarByRef.vt, VarByRef.ref) = RHS`. Passing `VarByRef.vt` and `VarByRef.ref` as `ByRef Variant` parameters lets assignments to `vt` and `ref` inside the function drive the VT-tag trick and the Variant write on the module-level construct.

For numeric index navigation, the API path uses bidirectional traversal (head-or-tail) identical to `CollItemKey`. The non-API path always traverses forward from head. String index navigation is always forward in both modes.

### `Key` Let — remove-and-re-add

The VB Collection COM interface provides no way to change a stored key in-place. `Key Let` works around this:

```vb
' Resolve position from key or numeric index
Dim pos As Long: pos = CollItemPos(this.coll, index, ...)
' Capture the item value
Dim Item As Variant: AssignVariant Item, this.coll.Item(index)
' Swap at same position with new key
If this.coll.Count = pos Then
    this.coll.Remove index
    this.coll.Add Item, RHS           ' append to end
Else
    this.coll.Remove index
    this.coll.Add Item, RHS, before:=pos
End If
```

The `Count = pos` branch handles the tail edge case: after removing the last item, `Count` decrements by one, so `before:=pos` would be out of range. Appending with plain `Add` is correct in that case.

`Key Let` is the only operation that allocates a new node. All other read/write operations work on existing nodes.

### `Sort` — iterative QuickSort by key

Sort is an index sort — `Keys` and `Items` remain in original order; only `idx` is permuted — to avoid swapping large Variants:

```vb
Keys = Me.Keys(1)           ' extract all keys ("" for unkeyed items)
ReDim idx(1 To Count)       ' identity permutation: idx(i) = i
QuickSortByIndex Keys, idx  ' permute idx so Keys(idx(i)) is ascending
Items = Me.Items(1)
Set this.coll = New Collection
For i = LBound(idx) To UBound(idx)
    If Len(Keys(idx(i))) = 0 Then
        this.coll.Add Items(idx(i))
    Else
        this.coll.Add Items(idx(i)), Keys(idx(i))
    End If
Next
```

`QuickSortByIndex` is an iterative implementation from Bruce McKinney's *Hardcore Visual Basic 5.0*:

- A `Collection` is used as the explicit partition stack — push two `Long` range markers per deferred partition, pop two per iteration.
- **Random pivot**: `idx(upper)` is swapped with a random element in `[lower, upper]` before each partition. `Randomize` is called in `Class_Initialize` to seed `Rnd`.
- **Insertion sort threshold**: subarrays of fewer than 47 elements switch to `InsertionSortByIndex`. This is the threshold used in Java's dual-pivot quicksort; below it, insertion sort's O(n²) cost is negligible and avoids the stack overhead of tiny partitions.
- **Comparison on `arr(idx(i))`**: only `idx` entries are swapped — `arr` (the keys) is never modified.

Items without keys produce `""` from `Keys()`. Empty string is lexicographically smallest, so unkeyed items accumulate at the front. `Len(Keys(idx(i))) = 0` re-adds them without a key string.

### `Enumerator` — delegated `_NewEnum`

```vb
Public Property Get Enumerator() As IEnumVARIANT
    Set Enumerator = this.coll.[_NewEnum]
End Property
```

VBA's `For Each` requires `IEnumVARIANT` at dispatch ID −4 (COM `_NewEnum`). Rather than implementing a custom enumerator, `Enumerator` returns the underlying Collection's own `IEnumVARIANT` directly. The `@Enumerator` annotation (baked in as `Attribute Enumerator.VB_UserMemId = -4` in the `_WithAttributes` file) registers this property at dispatch ID −4 so `For Each` finds it.

Items are enumerated in the current insertion order of the linked list — which is the sorted order after `Sort()` is called.

### `AssignVariant` — avoiding double index

```vb
Private Sub AssignVariant(ByRef variable As Variant, ByVal value As Variant)
    If VBA.IsObject(value) Then
        Set variable = value
    Else
        variable = value
    End If
End Sub
```

VBA requires `Set` for object assignment and plain `=` for values, with no unified syntax. Without this helper, every `IsObject`/assignment pair would call `this.coll.Item(index)` twice — once to test the type and once to store the result — incurring two COM dispatches and two linked-list traversals. `AssignVariant` captures the item in one `ByVal Variant` parameter (one dispatch) then branches on type once. The extra stack frame is the cost.

### Two compile modes

| Feature | `#Const API = False` (default) | `#Const API = True` |
|---|---|---|
| Pointer reads | VT-tag write, zero API calls | `CopyMemory` (kernel32) per read |
| Key string copy | `VT_BYREF\|VT_BSTR` — no allocation | `SysReAllocString` (oleaut32) |
| Item write | `VT_BYREF\|VT_VARIANT` assignment | `VarCopyToPtr` (oleaut32) |
| Null key guard | implicit (`""` on null BSTR) | explicit `KeyPtr <> vbNullPtr` |
| `CollItemKey` direction | head-only (forward) | bidirectional (head or tail) |
| Extra declarations | none | `SysReAllocString`, `VarCopyToPtr`, `vbNullPtr` |

`#Const API = False` is faster for key reads (no kernel round-trip per pointer chase) and requires less boilerplate (null key guard is implicit). `#Const API = True` uses explicit Windows API calls throughout and is easier to audit or step through in a debugger.

---

## 📄 License

MIT © 2025 Vincent van Geerestein
