# vba-enumerator-late-binding
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
![Platform](https://img.shields.io/badge/Platform-VBA%20(Excel%2C%20Access%2C%20Word%2C%20Outlook%2C%20PowerPoint)-blue)
![Architecture](https://img.shields.io/badge/Architecture-x86%20%7C%20x64-lightgrey)
![Rubberduck](https://img.shields.io/badge/Rubberduck-Ready-orange)

VBA standard module for IEnumVARIANT interface implementation — no typelib required, late binding via CallByName.

Implements the full `IEnumVARIANT` interface (`Next`, `Skip`, `Reset`, `Clone`) in a standard module using `AddressOf` and a heap-allocated vtable. Items are retrieved one by one via `CallByName`, so the iterable Class does not need to implement any interface.

---

## 📦 Features

- **`For Each` without a typelib** — pure VBA, no external dependencies
- **Late binding** — items are retrieved via `CallByName`, no `IEnumerator` interface required
- **Nested loops** — works correctly for nested `For Each` with mixed objects and mixed enumerators
- **Ascending iteration** — enumerates from `base` to `base + count - 1`
- **Fast variant copy** — uses a Variant ByRef construct (5× faster than `VariantCopy` API); switch to API mode with `#Const API = True`
- **Full COM lifecycle** — `QueryInterface`, `AddRef`, `Release`, `Clone` all correctly implemented; `KeepAlive` static collection keeps the iterable alive for the lifetime of the enumerator
- x86 / x64 compatible via `LongPtr` and `#If Win64`
- Pure VBA, zero dependencies, Rubberduck-friendly annotations

---

## 📁 Files

| File | Type | Description |
|---|---|---|
| `EnumeratorLateBinding.bas` | Module | `Enumerate(iterable, callback, count [, base])` — the main entry point |
| `CEnumTestLate.cls` | Example | Simple iterable Class using late binding |
| `EnumTestLate.bas` | Example | `For Each` tests and performance timings |

Each file has a corresponding `_WithAttributes` version (e.g. `EnumeratorLateBinding_WithAttributes.bas`) with [Rubberduck](https://rubberduckvba.com/) annotations removed and VB attributes baked in. Import the `_WithAttributes` files if you are not using Rubberduck.

> **Note:** `EnumTestLate.bas` uses a `Stopwatch` module for timing measurements. Remove or replace those calls if you do not have it.

---

## ⚙️ Public Interface

### `EnumeratorLateBinding` module

| Member | Description |
|---|---|
| `Enumerate(iterable, callback, count [, base])` | Returns a synthetic `IEnumVARIANT` for the iterable object. `callback` is the name of the indexed property (e.g. `"Item"`). `count` is the number of items. `base` is the first index (default 1). Raises an error if `iterable` is `Nothing` or `callback` is not a valid property. |

---

## 🚀 Quick Start

**1. Add the `Enumerate` function to your Class:**

```vb
'@Enumerator
Public Function Enumerate() As IEnumVARIANT
    Set Enumerate = EnumeratorLateBinding.Enumerate(Me, "Item", this.Count, 1)
End Function
```

**2. Expose an indexed `Item` property:**

```vb
Public Property Get Item(ByVal index As Long) As Variant
    Item = this.Items(index)
End Property
```

**3. Use `For Each`:**

```vb
Dim obj As MyClass
Set obj = New MyClass
' ... populate obj ...

Dim v As Variant
For Each v In obj
    Debug.Print v
Next
```

---

## ⏱️ Performance

Timings for `n = 10,000` items (Immediate Window):

| Method | Time (ms) |
|---|---|
| Custom enumerator — `For Each` (VarByRef) | 18.86 |
| Custom enumerator — `For Each` (API) | 43.0 |
| VB Collection — `For Each` | 0.21 |
| VB Array — `For Each` | 0.14 |
| VB Array — `For i` | 0.08 |

The overhead versus a native VB Collection is primarily the `CallByName` dispatch per element — roughly 7× slower than the early-bound `IEnumerator.Item(i)` call in the early-binding version (18.86 ms vs 2.69 ms). The VarByRef variant copy is ~5× faster than the `VariantCopy` API.

---

## ⚙️ Compiler directive

| Directive | Default | Description |
|---|---|---|
| `#Const API` | `False` | Copy Variants via Variant ByRef construct (fast) or `VariantCopy` API (slow) |

`#Const API = True` uses `VarCopyToPtr` (`VariantCopy` from oleaut32) and `SysReAllocString` to recover the callback name. `#Const API = False` uses the Variant ByRef construct — approximately 5× faster for the Variant copy. See performance table above.

---

## 🧠 How it works

VBA's `For Each` requires the iterable object to expose `IEnumVARIANT` via `_NewEnum`. Rather than using a typelib to define `IEnumVARIANT`, this module:

1. Allocates a block of heap memory (`CoTaskMemAlloc`) large enough to hold the enumerator state (`TENUM` UDT)
2. Builds a vtable of function pointers (`AddressOf`) for the seven COM methods: `QueryInterface`, `AddRef`, `Release`, `Next`, `Skip`, `Reset`, `Clone`
3. Writes the vtable pointer into the first field of `TENUM` — making it a valid COM object
4. Overwrites the return value of `Enumerate` with the heap pointer — returning the synthetic object as `IEnumVARIANT`
5. Stores the callback name as a BSTR (`SysAllocString`) in the `TENUM` struct; freed on `Release` via `SysFreeString`
6. Uses `CallByName` at each iteration to retrieve the item by index — no interface required on the iterable Class
7. Keeps the iterable object alive via a `Static Collection` keyed by the heap address, compensating for reference count changes when the local `TENUM` goes out of scope

Based on work by Dexter Freivald (32-bit, late binding) and ideas from *Hardcore Visual Basic 5.0* by Bruce McKinney.

---

## 🧠 Implementation notes

### `TENUM` — synthetic COM object layout

```vb
Private Type TENUM
    pvTable  As LongPtr     ' MUST be first — COM reads vtable pointer at offset 0
    caller   As Object      ' late-bound reference to the iterable object
    callback As LongPtr     ' BSTR pointer to the callback property name
    nRef     As Long        ' COM reference count
    First    As Long        ' index of first item
    Last     As Long        ' index of last item
    Current  As Long        ' index of current position
End Type
```

`pvTable` is first because COM requires a pointer to the vtable at offset 0 of any COM object. `caller As Object` is a late-bound reference — no interface is required on the iterable Class. `callback As LongPtr` holds a heap-allocated BSTR pointer to the property name string (see below). The `Step` field present in the early-binding version is absent — this enumerator is ascending-only; `First` maps to `base` and `Last` to `base + count - 1`.

### vtable — built once per session

```vb
Static vTable(0 To 6) As LongPtr
If vTable(0) = vbNullPtr Then
    vTable(0) = VBA.CLngPtr(AddressOf IUnknown_QueryInterface)
    ...
    vTable(6) = VBA.CLngPtr(AddressOf IEnumVARIANT_Clone)
End If
```

The `Static` array persists for the lifetime of the VBA session. `vTable(0) = vbNullPtr` is the once-only sentinel — subsequent calls to `Enumerate` reuse the same vtable without rebuilding it. Slots 0–2 are the three `IUnknown` methods; slots 3–6 are the four `IEnumVARIANT` methods.

### Return value trick

```vb
CopyMemory ByVal VarPtr(Enumerate), MemoryBlock, vbSizeLongPtr
```

`Enumerate` returns `IEnumVARIANT`. VBA stores the return value as an object pointer at `VarPtr(Enumerate)`. Overwriting those bytes with the heap block address makes VBA believe that address is a valid COM object — which it is, because the vtable pointer sits at offset 0. This is why `Enumerate` carries `'@Ignore NonReturningFunction`; the return value is set by raw memory write, not by a `Set Enumerate = ...` assignment.

### BSTR callback storage

```vb
.callback = SysAllocString(StrPtr(callback))
```

The callback property name is stored as a heap-allocated BSTR. `SysAllocString(StrPtr(callback))` allocates a new BSTR on the COM heap and copies the string content from the VBA `String` argument. The raw pointer is stored in `callback As LongPtr`; VBA has no knowledge of this allocation and will not free it automatically. `SysFreeString` in `IUnknown_Release` frees the BSTR when the enumerator's reference count reaches zero:

```vb
If obj.nRef = 0 Then
    Set KeepAlive(VarPtr(obj)) = Nothing
    SysFreeString obj.callback
    CoTaskMemFree VarPtr(obj)
End If
```

`SysAllocString` and `SysFreeString` are declared `Public` so the caller can use them directly for BSTR management if needed.

### Callback validation

Before storing the callback name, `Enumerate` validates it by calling it once:

```vb
On Error Resume Next
VBA.CallByName iterable, callback, VBA.VbGet, base
If Err.Number = 0 Then
    ' store and continue
Else
    Err.Raise vbErrorInvalidProcedureCall, , "Call back property is invalid"
End If
```

This surfaces a bad property name at construction time with a clear error, rather than silently failing on the first iteration.

### `KeepAlive` — reference count management

`CopyMemory ByVal MemoryBlock, obj, LenB(obj)` copies the `TENUM` struct to the heap as raw bytes — it copies the `caller` pointer without calling `AddRef`. When `obj` goes out of scope at function exit, VBA calls `Release` on `obj.caller`, which could destroy the iterable even though the heap block still holds a raw (untracked) copy of the pointer.

`KeepAlive` compensates:

```vb
Set KeepAlive(MemoryBlock) = obj.caller   ' hold one tracked reference
```

A `Static Collection` inside the property holds the reference, keyed by the heap block address. When `IUnknown_Release` sees `nRef = 0`, it removes the entry and frees the block:

```vb
Set KeepAlive(VarPtr(obj)) = Nothing   ' release — iterable may now be destroyed
SysFreeString obj.callback             ' free callback BSTR
CoTaskMemFree VarPtr(obj)              ' free heap block
```

The same pattern applies in `IEnumVARIANT_Clone`: the cloned enumerator's `caller` reference is also registered with `KeepAlive`.

### `IEnumVARIANT_Next` — hot path

```vb
For i = obj.Current To obj.Last
    ProcName = PeekStr(obj.callback, VarByRef.vt)
    CopyVarByRef rgVar, VBA.CallByName(obj.caller, ProcName, VBA.VbGet, i), VarByRef.vt, VarByRef.ref
    NumberFetched = NumberFetched + 1
    If NumberFetched = celt Then Exit For
    rgVar = rgVar + vbSizeVariant
Next
obj.Current = obj.Current + NumberFetched
```

Per iteration: `PeekStr` recovers the callback name from the heap-allocated BSTR (no allocation — reads directly from the heap via the Variant ByRef construct), then `CallByName` dispatches a late-bound property call and one Variant copy follows. `CallByName` is the dominant cost — the late-bound dispatch through `IDispatch.Invoke` is roughly 7× slower than the early-bound `.Item(i)` call in the early-binding version. `obj.Current` is updated in one step after the loop.

In API mode, `SysReAllocString(VarPtr(ProcName), obj.callback)` is used instead of `PeekStr`. This copies the BSTR content into `ProcName`'s existing allocation, reusing it across calls rather than creating a fresh allocation each iteration. `VarCopyToPtr` then copies the returned Variant to `rgVar`.

### `PeekStr` — reading a String from an address

```vb
Private Function PeekStr(ByVal address As LongPtr, ByRef vt As Variant) As String
    VarByRef.ref = VarPtr(address)
    vt = VBA.vbString Or VT_BYREF
    PeekStr = VarByRef.ref
End Function
```

`address` is a `ByVal LongPtr` parameter — its value is the BSTR pointer (`obj.callback`), and it lives on the stack for the duration of the call. `VarByRef.ref = VarPtr(address)` points the VT_BYREF construct at the stack copy of the BSTR pointer value. Setting `vt = vbString Or VT_BYREF` tells VBA to interpret the construct as a pointer to a BSTR pointer — reading through it recovers the string without any heap allocation. The `vt` and `ref` parameters are received `ByRef` from the `VarByRef` global, so VBA writes through the same global construct rather than creating a local copy.

`PeekStr` is the counterpart to `CopyLngByRef`: where `CopyLngByRef` writes a `Long` to an address, `PeekStr` reads a `String` from one.

### `IEnumVARIANT_Skip` — ascending-only

```vb
obj.Current = obj.Current + celt
If obj.Current <= obj.Last Then
    IEnumVARIANT_Skip = S_OK
Else
    obj.Current = obj.Last
    IEnumVARIANT_Skip = S_FALSE
End If
```

Without a `Step` field, Skip is ascending-only. On overshoot, `Current` is clamped to `Last` (rather than placed one past the end as in the early-binding version). VBA's `For Each` never calls `Skip`, so this difference is not observable in practice.

### `IEnumVARIANT_Clone` — shared BSTR

```vb
Dim Copy As TENUM: Copy = obj
Copy.nRef = 1
```

`Copy = obj` copies all fields as raw bytes: VBA automatically calls `AddRef` on the embedded `caller` object during the UDT copy, but `callback` is copied as a raw `LongPtr` — both the original and the clone hold the same BSTR pointer. `SysFreeString` is called in `IUnknown_Release` for whichever enumerator is released first; the other then holds a dangling BSTR pointer. In practice this is benign — VBA's `For Each` releases each enumerator as soon as the corresponding loop exits, so the clone is always released before or simultaneously with the original. The clone starts at `nRef = 1` regardless of the original's count, and captures the enumeration position at the moment of cloning.

### GUID comparison — field-by-field

```vb
' Const IID_IUnknown As String = "{00000000-0000-0000-C000-000000000046}"
IsIID_IUnknown = (id.Data1 = &H0) And (id.Data2 = &H0) And ... And (id.Data4(7) = &H46)
```

`QueryInterface` checks for `IID_IUnknown` and `IID_IEnumVARIANT`. Comparing the `GUID` UDT field-by-field avoids string parsing and is direct integer comparison. The commented-out `Const` documents the expected GUID string without runtime cost.

### `VARSIZE` — x86/x64 portability

```vb
#If Win64 Then
    vbSizeLongPtr = 8
    vbSizeVariant = 24
#Else
    vbSizeLongPtr = 4
    vbSizeVariant = 16
#End If
```

A `Variant` is 16 bytes on x86 and 24 bytes on x64 due to pointer-size alignment in the data union. All `CopyMemory` sizes and pointer arithmetic use these constants — no hard-coded values appear in the code.

---

## 📄 License

MIT © 2025 Vincent van Geerestein
