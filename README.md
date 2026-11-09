# vba-enumerator-late-binding
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
![Platform](https://img.shields.io/badge/Platform-VBA%20(Excel%2C%20Access%2C%20Word%2C%20Outlook%2C%20PowerPoint)-blue)
![Architecture](https://img.shields.io/badge/Architecture-x86%20%7C%20x64-lightgrey)
![Rubberduck](https://img.shields.io/badge/Rubberduck-Ready-orange)

VBA standard module for `IEnumVARIANT` interface implementation — no typelib required, late binding via the **default member** of the iterable object.

Implements the full `IEnumVARIANT` interface (`Next`, `Skip`, `Reset`, `Clone`) in a standard module using `AddressOf` and a heap-allocated vtable. Items are retrieved one by one by a late-bound call of the default member, `iterable(index)`, so the iterable object does not need to implement any interface. Any object with an indexed default member can be enumerated — also one you can't change, such as a VB Collection or a COM object from another library.

> [!IMPORTANT]
> **Breaking change (2026-11-08):** `Enumerate(iterable, callback, count [, base])` is now **`Enumerate(iterable, count [, base])`**. The callback name is gone; the items come from the default member. See [Why the default member](#-why-the-default-member-and-not-callbyname).

> For the early-binding variant (about 3× faster, requires the `IEnumerator` interface) see [vba-enumerator](https://github.com/vgrstn/vba-enumerator).

---

## 📦 Features

- **`For Each` without a typelib** — pure VBA, no external dependencies
- **No interface required** — items retrieved via the default member, `iterable(index)`
- **Nested loops** — works correctly for nested `For Each` with mixed objects and mixed enumerators
- **Empty objects** — `count = 0` gives no loop and never calls the default member
- **Errors reach the caller** — an error raised in the default member or in the loop body propagates as usual
- **Fast variant copy** — uses a Variant ByRef construct (~2.5× faster than the `VariantCopy` API here); switch to API mode with `#Const API = True`
- **Reused memory blocks** — released enumerators go on a stack, so after the first `For Each` no memory is allocated
- **Full COM lifecycle** — `QueryInterface`, `AddRef`, `Release`, `Clone` all correctly implemented
- x86 / x64 compatible via `LongPtr` and `#If Win64`
- Pure VBA, zero dependencies, Rubberduck-friendly annotations

---

## 📁 Files

| File | Type | Description |
|---|---|---|
| `EnumeratorLateBinding.bas` | Module | `Enumerate(iterable, count [, base])` — the main entry point |
| `CEnumTestLate.cls` | Example | Simple iterable Class with an indexed default member |
| `EnumTestLate.bas` | Example | `For Each` examples and performance timings |

Each module and class has a corresponding `_WithAttributes` version (e.g. `EnumeratorLateBinding_WithAttributes.bas`) with [Rubberduck](https://rubberduckvba.com/) annotations removed and VB attributes baked in. Import the `_WithAttributes` files if you are not using Rubberduck.

> **Note:** `EnumTestLate.bas` uses a `Stopwatch` module for timing measurements. Remove or replace those calls if you do not have it.

---

## ⚙️ Public Interface

### `EnumeratorLateBinding` module

| Member | Description |
|---|---|
| `Enumerate(iterable, count [, base])` | Returns a synthetic `IEnumVARIANT` for the iterable object. The items are `iterable(index)` for `index = base To base + count - 1`; `base` defaults to `1`. Raises error 424 if `iterable` is `Nothing`. An object without a default member raises error 451 at the first item; with `count = 0` nothing is called. |

---

## 🚀 Quick Start

**1. Give your Class an indexed default member:**

```vb
'@DefaultMember
Public Property Get Item(ByVal index As Long) As Variant
Attribute Item.VB_UserMemId = 0
    If VBA.IsObject(this.Items(index)) Then
        Set Item = this.Items(index)
    Else
        Item = this.Items(index)
    End If
End Property
```

(The `Attribute` line is what the `_WithAttributes` files contain; with Rubberduck, the `'@DefaultMember` annotation sets it.)

**2. Add the `Enumerate` function:**

```vb
'@Enumerator
Public Function Enumerate() As IEnumVARIANT
    Set Enumerate = EnumeratorLateBinding.Enumerate(Me, this.Count)
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

## 🤔 Why the default member and not `CallByName`

Until 2026-11-08 `Enumerate` took the name of a callback, and `Next` called it for every item with `CallByName(caller, name, VbGet, i)`. That was dropped:

1. **Speed.** `CallByName` looks the name up with `IDispatch::GetIDsOfNames` for every item and then calls `Invoke`; the default member is called with `Invoke` at `DISPID_VALUE` right away. Per call: `iterable(index)` 0.67 µs, `CallByName` 1.36 µs (and `iterable.Item(index)` 1.24 µs). A `For Each` is now 1.75× faster per item.
2. **Simplicity.** The callback name had to stay in the enumerator object as a string for the whole enumeration, and that brought three bugs:
   - a cache of the name in `Next` compared BSTR addresses, so a new name at the address of a released one returned the items of the *old* callback, without an error;
   - a UDT with a `String` is converted to ANSI when it is passed to a `Declare`, so `CopyMemory` copied a temporary string;
   - the string had to be copied by `Clone` and freed by `Release`.

   Without the name, the enumerator state holds only pointers and `Long`s.
3. **The alternative was slower.** Resolving the DISPID once and calling `Invoke` through `DispCallFunc` was tried in [vba-enumerator-dispatch](https://github.com/vgrstn/vba-enumerator-dispatch) (retired): every item then takes a `Declare` call of about 1.4 µs, which is slower than `CallByName`.

A class without an indexed default member can get one, or can implement `IEnumerator` and use the faster [vba-enumerator](https://github.com/vgrstn/vba-enumerator).

---

## ⏱️ Performance

Timings for `n = 10,000` items:

| Method | Time (ms) |
|---|---|
| Custom enumerator — `For Each` (VarByRef) | 9.40 |
| Custom enumerator — `For Each` (API) | 22.54 |
| Iterable object — `For i`, `iterable(i)` late bound | 6.71 |
| Iterable object — `For i`, `CallByName(iterable, "Item", VbGet, i)` | 14.46 |
| Iterable object — `For i`, `Item(i)` early bound | 0.85 |
| Iterable object — `For Each` over an `Items` array | 0.31 |
| VB Collection — `For Each` | 0.25 |
| VB Array — `For Each` | 0.12 |
| VB Array — `For i` | 0.09 |

Most of the time goes to the late-bound call of the default member (0.67 µs per item); `Next` and the item write add about 0.27 µs. With `CallByName` it was 16.41 ms for 10,000 items. The early-binding [vba-enumerator](https://github.com/vgrstn/vba-enumerator) takes 3.01 ms.

Starting a `For Each` takes about 4 µs (API 8 µs), against 0.2 µs for a VB Collection. It used to take 27 µs, before the memory blocks were reused and the pointer writes moved from `CopyMemory` to the Variant ByRef construct.

---

## ⚙️ Compiler directive

| Directive | Default | Description |
|---|---|---|
| `#Const API` | `False` | Copy Variants and pointers via the Variant ByRef construct (fast) or via the `VariantCopy` and `CopyMemory` APIs (slow) |

---

## 🧠 How it works

VBA's `For Each` requires the iterable object to expose `IEnumVARIANT` via `_NewEnum`. Rather than using a typelib to define `IEnumVARIANT`, this module:

1. Takes a block of heap memory, large enough to hold the enumerator state (`TENUM` UDT), from a stack of free blocks — or allocates one with `CoTaskMemAlloc` when the stack is empty
2. Builds a vtable of function pointers (`AddressOf`) for the seven COM methods: `QueryInterface`, `AddRef`, `Release`, `Next`, `Skip`, `Reset`, `Clone`
3. Writes the vtable pointer into the first field of `TENUM` — making it a valid COM object
4. Overwrites the return value of `Enumerate` with the heap pointer — returning the synthetic object as `IEnumVARIANT`
5. Hands the reference to the iterable object over to the heap block, which keeps it alive until the enumeration is released

Based on work by Dexter Freivald (32-bit, late binding) and ideas from *Hardcore Visual Basic 5.0* by Bruce McKinney.

---

## 🧠 Implementation notes

### `TENUM` — synthetic COM object layout

```vb
Private Type TENUM
    pvTable  As LongPtr     ' MUST be first — COM reads vtable pointer at offset 0
    caller   As Object      ' late-bound reference to the iterable object
    nRef     As Long        ' COM reference count
    First    As Long        ' index of first item
    Last     As Long        ' index of last item
    Current  As Long        ' index of current position
End Type
```

`pvTable` is first because COM requires a pointer to the vtable at offset 0 of any COM object. `caller As Object` is a late-bound reference — no interface is required. Unlike the early-binding version there is no `Step` field — this module is ascending-only, from `base` to `base + count - 1`.

### Reference counting — the heap block owns the reference

`CopyMemory ByVal MemoryBlock, obj, LenB(obj)` copies the `TENUM` struct to the heap as raw bytes — it copies the `caller` pointer without calling `AddRef`. The local `obj` holds the one counted reference, taken by `Set .caller = iterable`. Clearing that pointer in `obj` without a `Release` hands the reference over to the heap block:

```vb
CopyPtrByRef VarPtr(obj.caller), vbNullPtr, VarByRef.vt, VarByRef.ref
```

When `IUnknown_Release` sees `nRef = 0`, it releases the reference held by the block and puts the block on the stack of free blocks:

```vb
Set obj.caller = Nothing                ' release — iterable may now be destroyed
FreeBlocks.Add VarPtr(obj), Before:=1   ' (or Add when the stack is empty)
```

`IEnumVARIANT_Clone` does the same with its copy: `Copy = obj` AddRefs `caller`, and that reference is handed over to the clone's block.

> Earlier versions kept the iterable object alive with a `KeepAlive` Collection keyed by the block address. Its `On Error` handling cleared the `Err` object when an enumerator was released while an error was being raised, so an error in the callback or in the body of a `For Each` reached the caller one level up as `Err.Number = 0`.

### Memory blocks — a stack for reuse

A `Declare` call takes about 1.4 µs, and `CoTaskMemAlloc` with `CoTaskMemFree` about 6 µs. Released blocks are therefore never freed but kept in a Collection used as a stack (`FreeBlocks`, top at position 1, no key). `TakeBlock` pops a block, or allocates a new one when the stack is empty. Only as many blocks are allocated as the most enumerations that were active at the same time.

### `IEnumVARIANT_Next` — hot path

```vb
For i = obj.Current To obj.Last
    CopyVarByRef rgVar, obj.caller(i), VarByRef.vt, VarByRef.ref
    NumberFetched = NumberFetched + 1
    If NumberFetched = celt Then Exit For
    rgVar = rgVar + vbSizeVariant
Next
obj.Current = obj.Current + NumberFetched
```

Per iteration: one late-bound call of the default member and one Variant copy to the destination address. VBA's `For Each` always requests one item at a time (`celt = 1`). `celt = 0` returns `S_OK` at once; before that check was added, `celt = 0` fetched items without stopping and wrote beyond the caller's array.

### Variant ByRef construct

```vb
Private Type CONSTRUCT
    vt  As Variant
    ref As Variant
End Type
Private VarByRef As CONSTRUCT
```

`InitializeVarByRef` sets `vt` to `VT_INTEGER | VT_BYREF` pointing at `ref`. `CopyVarByRef` writes a Variant to any address without an API call, `CopyLngByRef` a `Long` (for `pceltFetched`), and `CopyPtrByRef` a pointer. On x64 a `LongLong` by reference is not supported (error 458), so `CopyPtrByRef` copies the 8 bytes as a `Currency`:

```vb
VarByRef.ref = value            ' the pointer as a LongLong
vt = VBA.vbCurrency             ' the same 8 bytes, now read as a Currency
Dim Bits As Currency: Bits = ref
VarByRef.ref = address
vt = VBA.vbCurrency Or VT_BYREF
ref = Bits                      ' the 8 bytes written as they are
```

Only the type is changed, so the bits are copied and nothing is calculated. (An earlier version wrote `CCur(value) * 0.0001@`; that is exact for every user-mode address, but silently wrong above 922,337,203,685,477.)

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

`obj.Current` is only mutated after the bounds check. On overshoot, `Current` is placed one step past `Last`, matching the post-exhaustion state that `Next` leaves.

### GUID comparison and `VARSIZE`

`IsIID_IUnknown` and `IsIID_IEnumVARIANT` compare the `GUID` UDT fields directly — no string parsing. A `Variant` is 16 bytes on x86 and 24 bytes on x64; all sizes and pointer arithmetic use the `VARSIZE` constants.

---

## 📝 Changes

**2026-11-09**
- Fixed: `CopyPtrByRef` copies the 8 bytes of a pointer as a `Currency` by changing the Variant type only. The previous `CCur(value) * 0.0001@` was exact for every user-mode address, but silently wrong above 922,337,203,685,477.
- Module header rewritten: what the module does and needs at the top, the performance notes in one paragraph, a corrected note on the alternatives.

**2026-11-08**
- **Breaking:** `Enumerate(iterable, callback, count [, base])` became `Enumerate(iterable, count [, base])`; the items come from the default member (see above).
- Fixed: errors raised in the callback or in the body of a `For Each` were lost one level up (`Err.Number = 0`); the `KeepAlive` Collection is gone.
- Fixed: an empty object (`count = 0`) raised error 5 "Call back property is invalid" (the callback test in `Enumerate`, which also cleared the caller's `Err`).
- Fixed: after a loop with one callback, a loop with another callback could return the items of the first one (name cache keyed by BSTR address).
- Fixed: `Next` with `celt = 0` wrote beyond the caller's array.
- Faster start of a `For Each` (27 → 4 µs): memory blocks are reused through a stack, and pointers are written with the Variant ByRef construct.

---

## 📄 License

MIT © 2025-2026 Vincent van Geerestein
