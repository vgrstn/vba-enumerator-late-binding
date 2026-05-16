# vba-enumerator-late-binding
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
![Platform](https://img.shields.io/badge/Platform-VBA%20(Excel%2C%20Access%2C%20Word%2C%20Outlook%2C%20PowerPoint)-blue)
![Architecture](https://img.shields.io/badge/Architecture-x86%20%7C%20x64-lightgrey)
![Rubberduck](https://img.shields.io/badge/Rubberduck-Ready-orange)

VBA standard module for `IEnumVARIANT` interface implementation — no typelib required, late binding via `CallByName`.

Implements the full `IEnumVARIANT` interface (`Next`, `Skip`, `Reset`, `Clone`) in a standard module using `AddressOf` and a heap-allocated vtable. Items are retrieved one by one via `VBA.CallByName`, so the iterable Class does not need to implement any interface.

> For the early-binding variant (faster, requires `IEnumerator` interface) see [vba-enumerator](https://github.com/vgrstn/vba-enumerator).

---

## 📦 Features

- **`For Each` without a typelib** — pure VBA, no external dependencies
- **No interface required** — items retrieved via `VBA.CallByName`; the iterable Class exposes any indexed property or method
- **Nested loops** — works correctly for nested `For Each` with mixed objects and mixed enumerators
- **Fast variant copy** — uses a Variant ByRef construct (~5× faster than `VariantCopy` API); switch to API mode with `#Const API = True`
- **Full COM lifecycle** — `QueryInterface`, `AddRef`, `Release`, `Clone` all correctly implemented
- **BSTR callback storage** — property name heap-allocated via `SysAllocString`, freed on `Release` via `SysFreeString`; no string allocation in the hot path (`PeekStr`)
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
| `Enumerate(iterable, callback, count [, base])` | Returns a synthetic `IEnumVARIANT` for the iterable object. `callback` is the name of the indexed property or method (e.g. `"Item"`). `count` is the number of items. `base` is the first index (default `1`). Raises an error if `iterable` is `Nothing` or `callback` is not callable on `iterable`. |

---

## 🚀 Quick Start

**1. Expose an indexed property in your Class:**

```vb
Public Property Get Item(ByVal index As Long) As Variant
    If VBA.IsObject(this.Items(index)) Then
        Set Item = this.Items(index)
    Else
        Item = this.Items(index)
    End If
End Property
```

**2. Add the `Enumerate` function:**

```vb
'@Enumerator
Public Function Enumerate() As IEnumVARIANT
    Set Enumerate = EnumeratorLateBinding.Enumerate(Me, "Item", this.Count)
End Function
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
| Custom enumerator — `For Each` (VarByRef) | 19.58 |
| Custom enumerator — `For Each` (API) | 47.23 |
| VB Collection — `For Each` | 0.21 |
| VB Array — `For Each` | 0.14 |
| VB Array — `For i` | 0.08 |

The dominant cost is `VBA.CallByName` dispatch per element — roughly 7× slower than the early-bound `IEnumerator.Item(i)` call in [vba-enumerator](https://github.com/vgrstn/vba-enumerator) (2.89 ms). The VarByRef variant copy is ~5× faster than the `VariantCopy` API.

---

## ⚙️ Compiler directive

| Directive | Default | Description |
|---|---|---|
| `#Const API` | `False` | Copy Variants via Variant ByRef construct (fast) or `VariantCopy` API (slow) |

`#Const API = True` uses `VarCopyToPtr` (`VariantCopy` from oleaut32) and `SysReAllocString` to read the callback name. `#Const API = False` uses the Variant ByRef construct and `PeekStr` — approximately 5× faster for the Variant copy. See performance table above.

---

## 🧠 How it works

VBA's `For Each` requires the iterable object to expose `IEnumVARIANT` via `_NewEnum`. Rather than using a typelib to define `IEnumVARIANT`, this module:

1. Allocates a block of heap memory (`CoTaskMemAlloc`) large enough to hold the enumerator state (`TENUM` UDT)
2. Builds a vtable of function pointers (`AddressOf`) for the seven COM methods: `QueryInterface`, `AddRef`, `Release`, `Next`, `Skip`, `Reset`, `Clone`
3. Writes the vtable pointer into the first field of `TENUM` — making it a valid COM object
4. Overwrites the return value of `Enumerate` with the heap pointer — returning the synthetic object as `IEnumVARIANT`
5. Stores the callback property name as a BSTR on the COM heap (`SysAllocString`); reads it without allocation on every iteration (`PeekStr`)
6. Keeps the iterable object alive via a `Static Collection` keyed by the heap address, compensating for reference count changes when the local `TENUM` goes out of scope

Based on work by Dexter Freivald (32-bit, late binding) and ideas from *Hardcore Visual Basic 5.0* by Bruce McKinney.

---

## 🧠 Implementation notes

### `TENUM` — synthetic COM object layout

```vb
Private Type TENUM
    pvTable  As LongPtr     ' MUST be first — COM reads vtable pointer at offset 0
    caller   As Object      ' late-bound reference to the iterable object
    callback As LongPtr     ' BSTR holding the property/method name
    nRef     As Long        ' COM reference count
    First    As Long        ' index of first item
    Last     As Long        ' index of last item
    Current  As Long        ' index of current position
End Type
```

`pvTable` is first because COM requires a pointer to the vtable at offset 0 of any COM object. `caller As Object` is a late-bound reference — no interface is required. `callback As LongPtr` holds a heap-allocated BSTR (see BSTR lifecycle below). Unlike the early-binding version there is no `Step` field — this module is ascending-only.

### vtable — built once per session

```vb
Static vTable(0 To 6) As LongPtr
If vTable(0) = vbNullPtr Then
    vTable(0) = VBA.CLngPtr(AddressOf IUnknown_QueryInterface)
    ...
    vTable(6) = VBA.CLngPtr(AddressOf IEnumVARIANT_Clone)
End If
```

The `Static` array persists for the lifetime of the VBA session. Subsequent calls to `Enumerate` reuse the same vtable without rebuilding it. Slots 0–2 are the three `IUnknown` methods; slots 3–6 are the four `IEnumVARIANT` methods.

### Return value trick

```vb
CopyMemory ByVal VarPtr(Enumerate), MemoryBlock, vbSizeLongPtr
```

`Enumerate` returns `IEnumVARIANT`. VBA stores the return value as an object pointer at `VarPtr(Enumerate)`. Overwriting those bytes with the heap block address makes VBA believe that address is a valid COM object — which it is, because the vtable pointer sits at offset 0. This is why `Enumerate` carries `'@Ignore NonReturningFunction`; the return value is set by raw memory write, not by a `Set Enumerate = ...` assignment.

### BSTR lifecycle

`callback As LongPtr` is not managed by VBA — it must be allocated and freed manually at every exit path:

- **`Enumerate`**: `SysAllocString(StrPtr(callback))` allocates a fresh BSTR. If `CoTaskMemAlloc` subsequently fails, `SysFreeString(obj.callback)` cleans up before `Err.Raise`
- **`IUnknown_Release`**: `SysFreeString obj.callback` before `CoTaskMemFree` when `nRef = 0`
- **`IEnumVARIANT_Clone`**: `Copy = obj` copies `callback` as a raw pointer (no `SysAllocString`). Immediately: `Copy.callback = SysAllocString(obj.callback)` gives the clone its own independent BSTR so that `SysFreeString` in one enumerator's `Release` does not leave the other with a dangling pointer. If `CoTaskMemAlloc` fails, `SysFreeString Copy.callback` cleans up before returning `E_OUTOFMEMORY`

Callback validation happens at construction time — `VBA.CallByName` is called once inside `On Error Resume Next`. A bad property name raises `vbErrorInvalidProcedureCall` immediately, before any BSTR is allocated.

### `KeepAlive` — reference count management

`CopyMemory ByVal MemoryBlock, obj, LenB(obj)` copies the `TENUM` struct to the heap as raw bytes — it copies the `caller` pointer without calling `AddRef`. When `obj` goes out of scope at function exit, VBA decrements the reference count on `obj.caller`, which could destroy the iterable. `KeepAlive` compensates:

```vb
Set KeepAlive(MemoryBlock) = obj.caller   ' hold one tracked reference
```

A `Static Collection` inside the property holds the reference, keyed by the heap block address. When `IUnknown_Release` sees `nRef = 0`:

```vb
Set KeepAlive(VarPtr(obj)) = Nothing   ' release tracked reference
SysFreeString obj.callback             ' free BSTR
CoTaskMemFree VarPtr(obj)              ' free heap block
```

The same pattern applies in `IEnumVARIANT_Clone`: the cloned enumerator's `caller` reference is registered too.

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

Per iteration: one `PeekStr` (no allocation), one `CallByName` dispatch, one Variant copy. `obj.Current` is updated in one step after the loop. VBA's `For Each` always requests one item at a time (`celt = 1`), so the inner `Exit For` fires immediately.

### `PeekStr` — BSTR read without allocation

```vb
Private Function PeekStr(ByVal address As LongPtr, ByRef vt As Variant) As String
    VarByRef.ref = VarPtr(address)
    vt = VBA.vbString Or VT_BYREF
    PeekStr = VarByRef.ref
End Function
```

`address` is a stack copy of `obj.callback`. `VarPtr(address)` points the Variant ByRef construct at this stack slot. Setting `vt = vbString Or VT_BYREF` tells VBA to read it as a pointer to a BSTR pointer — producing the `String` with no `SysAllocString` call. This is the non-API counterpart to `SysReAllocString` used in API mode.

### Variant ByRef construct

```vb
Private Type CONSTRUCT
    vt  As Variant
    ref As Variant
End Type
Private VarByRef As CONSTRUCT
```

`InitializeVarByRef` sets `vt` to `VT_INTEGER | VT_BYREF` pointing at `ref`. `CopyVarByRef` then writes a Variant to any address without an API call; `CopyLngByRef` writes a `Long` (for `pceltFetched`); `PeekStr` reads a `String`. All helpers receive `vt` and `ref` as `ByRef Variant` parameters to avoid a global UDT indirection per call.

### `IEnumVARIANT_Skip` — `Select Case True`

```vb
Select Case True
Case celt = 0:  IEnumVARIANT_Skip = S_OK
Case celt < 0:  IEnumVARIANT_Skip = E_INVALIDARG
Case celt <= obj.Last - obj.Current + 1
    obj.Current = obj.Current + celt
    IEnumVARIANT_Skip = S_OK
Case Else
    obj.Current = obj.Last + 1
    IEnumVARIANT_Skip = S_FALSE
End Select
```

`obj.Current` is only mutated after the bounds check — state is never corrupted on an invalid or overshooting call. `celt = 0` is an explicit fast-path. On overshoot, `Current` is placed one step past `Last`, matching the post-exhaustion state that `Next` leaves.

### `IEnumVARIANT_Clone` — independent BSTR

```vb
Dim Copy As TENUM: Copy = obj      ' UDT copy — VBA AddRefs caller; callback copied as raw LongPtr
Copy.callback = SysAllocString(obj.callback)
Copy.nRef = 1
```

`Copy = obj` AddRefs `caller` automatically but copies `callback` as a raw pointer. `SysAllocString(obj.callback)` immediately gives the clone its own independent BSTR so that `SysFreeString` in either enumerator's `Release` does not leave the other with a dangling pointer. The clone starts at `nRef = 1` and captures the enumeration position at the moment of cloning.

### GUID comparison — field-by-field

```vb
' Const IID_IUnknown As String = "{00000000-0000-0000-C000-000000000046}"
IsIID_IUnknown = (id.Data1 = &H0) And (id.Data2 = &H0) And ... And (id.Data4(7) = &H46)
```

Direct integer comparison on the `GUID` UDT fields — no string parsing. The commented `Const` documents the expected GUID string without runtime cost.

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

A `Variant` is 16 bytes on x86 and 24 bytes on x64. All `CopyMemory` sizes and pointer arithmetic use these constants — no hard-coded values appear in the code.

---

## 📄 License

MIT © 2025 Vincent van Geerestein
