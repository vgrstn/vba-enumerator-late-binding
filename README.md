# vba-enumerator-late-binding
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
![Platform](https://img.shields.io/badge/Platform-VBA%20(Excel%2C%20Access%2C%20Word%2C%20Outlook%2C%20PowerPoint)-blue)
![Architecture](https://img.shields.io/badge/Architecture-x86%20%7C%20x64-lightgrey)
![Rubberduck](https://img.shields.io/badge/Rubberduck-Ready-orange)

VBA standard module that adds `For Each` support to any Class using late binding and a synthetic `IEnumVARIANT` COM object — no typelib, no interface required.

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

> **Note:** `EnumTestLate.bas` uses a `Stopwatch` module for timing measurements. Remove or replace those calls if you do not have it.

---

## ⚙️ Public Interface

### `Enumerator` module

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

The overhead versus a native VB Collection is primarily the `CallByName` dispatch per element. The VarByRef variant copy is ~5× faster than the `VariantCopy` API.

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

## 📄 License

MIT © 2025 Vincent van Geerestein
