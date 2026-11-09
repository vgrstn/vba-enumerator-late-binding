Attribute VB_Name = "EnumeratorLateBinding"
Attribute VB_Description = "For Each enumerator for VBA objects without a typelib (late binding)."

'------------------------------------------------------------------------------
' MIT License
'
' Copyright (c) 2025-2026 Vincent van Geerestein
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
' Description: For Each enumerator for VBA objects using Late Binding
' Add-in: RubberDuck (https://rubberduckvba.com/)
' Version: 2026.11.09
'
' Enumerate gives any VBA class For Each support, without a typelib or a
' reference, in 32-bit and 64-bit Office (VBA7). The module builds a complete
' IEnumVARIANT object (Next, Skip, Reset and Clone) in memory and gets the
' items one by one through the default member of the object (late binding).
' No reference and no interface are required.
'
' Methods
' Enumerate(iterable, count [, base])   Returns the IEnumVARIANT interface for
'                                       an iterable object
'
' The items are the values of the default member of the iterable object, for
' the indices base To base + count - 1, called by late binding: iterable(index).
' Any object with an indexed default member can be enumerated, also an object
' that can't implement an interface.
'
' Enumerator works correctly for nested loops with mixed objects as well as for
' nested loops with mixed enumerators. An error raised in the default member or
' in the body of the loop reaches the caller as usual. An object without a
' default member raises error 451 at the first item, and an empty object
' (count = 0) never calls it. The iterable object stays alive until the
' enumeration ends.
'
' Code to be included in the iterable object:
'
' '@DefaultMember
' Public Property Get Item(ByVal index As Long) As Variant
'    (return the item at index, with Set for an object)
' End Property
'
' '@Enumerator
' Public Function Enumerate() As IEnumVARIANT
'    Set Enumerate = EnumeratorLateBinding.Enumerate(Me, count [, base])
' End Function
'
' Why the default member and not a callback name with CallByName (until
' version 2026.11.08 the callback name was the second argument of Enumerate):
' - Speed. CallByName looks the name up with IDispatch::GetIDsOfNames for every
'   item, then calls IDispatch::Invoke; the default member is called with
'   Invoke at DISPID_VALUE right away. Per call 0.67 탎 for iterable(index)
'   against 1.36 탎 for CallByName (and 1.24 탎 for iterable.Item(index)).
' - Simplicity. The callback name had to be kept in the IEnumVARIANT object as
'   a String for the whole enumeration. That brought three bugs: a cache of the
'   name in IEnumVARIANT_Next compared BSTR addresses, so a new name at the
'   address of a released one gave the items of the old callback; a UDT with a
'   String is converted to ANSI when it is passed to a Declare, so CopyMemory
'   copied a temporary string; and the String had to be copied by Clone and
'   freed by Release. Without the name the object holds only pointers and Longs.
' - Resolving the DISPID once and calling Invoke with the DispCallFunc API was
'   tried as well (EnumerateDispatch, retired): every item then takes a Declare
'   call of about 1.4 탎, which makes it slower than CallByName.
' A class without an indexed default member can get one, or can implement the
' IEnumerator interface and use the faster Enumerator module.
'
' Timings (ms) for n = 10,000, see BenchmarkEnumerator
' Iterable object with default member - late binding        9.40 (API 22.54 ms)
' Iterable object - iterable(index), late binding (For)     6.71
' Iterable object - CallByName(Item, index) (For)          14.46
' Iterable object - Item(index), early binding (For)        0.85
' Iterable object - Items array (For Each)                  0.31
' VB Collection - VB enumerator (For Each)                  0.25
' VB Array - VB enumerator (For Each)                       0.12
' VB Array - VB loop (For)                                  0.09
'
' Iterable object with IEnumerator - early binding          3.01 (API 16.04 ms)
'
' Most of the time goes to the late-bound call of the default member. With the
' IEnumerator interface (Enumerator module) the enumeration is about 3 times
' faster.
'
' No API is called per item: every Declare call takes about 1.4 탎. A Variant
' is copied with the Variant ByRef construct (5 times faster than VarCopyToPtr)
' and pointers are written with it too; on 64 bits a LongLong can't be written
' by reference (error 458), so its 8 bytes are copied as a Currency. The memory
' blocks of the IEnumVARIANT objects are reused through a stack instead of
' allocated and freed for every enumeration (CoTaskMemAlloc and CoTaskMemFree
' together take about 6 탎), so starting a For Each takes about 4 탎 (API 8 탎),
' against 0.2 탎 for a VB Collection. A new block is only allocated when more
' enumerations are active at the same time than ever before; the blocks are
' never freed.
'
' 29 unit tests are in TestEnumLateBinding, the timings in BenchmarkEnumerator.
'
' The original ideas for a custom enumerator using a typelib and redefining
' the IEnumVARIANT interface routines in a standard module originate from
' Hardcore Visual Basic 5.0 by Bruce McKinney.
'
' The implementation without using a typelib is based on work by Dexter
' Freivald whose original code was for 32 bits and was using late binding.
' https://www.vbforums.com/showthread.php?854963-VB6-IEnumVARIANT-For-Each-support-without-a-typelib
'
' Alternatives are to copy the items to an embedded VB Collection and expose
' its enumerator, or to return the items as an array, as the Items and Keys of
' a Scripting.Dictionary do. Both hand over all items at once, whereas
' Enumerate hands them over one by one.

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Compiler Directives
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' The IEnumVARIANT_Next function is critical to the overall performance and it
' scales with the number of items enumerated. IEnumVARIANT_Next needs to copy a
' variant to a memory address. Depending on the API compiler directive it uses
' the VarCopyToPtr API or the 5x faster CopyVarByRef Variant ByRef method.
' The same directive chooses CopyMemory or the Variant ByRef construct for
' writing pointers. The late binding approach calls the default member of the
' iterable object: iterable(index).

#Const API = False

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private API Declarations
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' https://learn.microsoft.com/en-us/windows-hardware/drivers/ddi/wdm/nf-wdm-rtlmovememory
Private Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
    pDst As Any, _
    pSrc As Any, _
    ByVal NBytes As Long _
)

' https://docs.microsoft.com/en-us/windows/win32/api/combaseapi/nf-combaseapi-cotaskmemalloc
Private Declare PtrSafe Function CoTaskMemAlloc Lib "ole32.dll" ( _
    ByVal cb As Long _
) As LongPtr

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

' The vbNullPtr constant is the null pointer (not VB defined).
Private Const vbNullPtr As LongPtr = 0

' The memory size of intrinsic data types.
Private Enum VARSIZE
    vbSizeInteger = 2
    vbSizeLong = 4
#If Win64 Then
    vbSizeLongPtr = 8
    vbSizeVariant = 24
#Else
    vbSizeLongPtr = 4
    vbSizeVariant = 16
#End If
End Enum

' Selected HRESULT constants.
Private Enum HRESULT
    S_OK = &H0                          ' Operation successful, returns True
    S_FALSE = &H1                       ' Operation successful, returns False
    E_NOTIMPL = &H80004001              ' Not implemented
    E_NOINTERFACE = &H80004002          ' No such interface supported
    E_POINTER = &H80004003              ' Pointer that is not valid
    E_OUTOFMEMORY = &H8007000E          ' Failed to allocate necessary memory
    E_INVALIDARG = &H80070057           ' One of the arguments is not valid
End Enum

' Selected VBA errors.
Private Enum VBERROR
    vbErrorInvalidProcedureCall = 5
    vbErrorOutOfMemory = 7
    vbErrorObjectRequired = 424
End Enum

' GUID is the UDT for the Global Unique Identifier.
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

' The IEnumVARIANT status is captured in an UDT.
Private Type TENUM
    pvTable As LongPtr
    caller As Object
    nRef As Long
    First As Long
    Last As Long
    Current As Long
End Type

' The memory blocks of released IEnumVARIANT objects are kept for reuse in a
' stack (top at position 1): Enumerate and Clone take a block from the stack and
' only allocate a new one when it is empty, IUnknown_Release puts it back.
Private FreeBlocks As Collection

#If API = False Then
' Variant ByRef construct for memory access by address.
Private Const VT_BYREF As Integer = &H4000
Private Type CONSTRUCT
    vt As Variant
    ref As Variant
End Type
Private VarByRef As CONSTRUCT
#End If


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function Enumerate( _
    ByVal iterable As Object, _
    ByVal count As Long, _
    Optional ByVal base As Long = 1 _
) As IEnumVARIANT

    If iterable Is Nothing Then Err.Raise vbErrorObjectRequired

    ' Initialize the vTable with the redefined IUnknown/IEnumVARIANT functions.
    Static vTable(0 To 6) As LongPtr
    If vTable(0) = vbNullPtr Then
        vTable(0) = VBA.CLngPtr(AddressOf IUnknown_QueryInterface)
        vTable(1) = VBA.CLngPtr(AddressOf IUnknown_AddRef)
        vTable(2) = VBA.CLngPtr(AddressOf IUnknown_Release)
        vTable(3) = VBA.CLngPtr(AddressOf IEnumVARIANT_Next)
        vTable(4) = VBA.CLngPtr(AddressOf IEnumVARIANT_Skip)
        vTable(5) = VBA.CLngPtr(AddressOf IEnumVARIANT_Reset)
        vTable(6) = VBA.CLngPtr(AddressOf IEnumVARIANT_Clone)
#If API = False Then
        ' Initialize the Variant ByRef construct.
        InitializeVarByRef
#End If
        ' Initialize the stack of blocks for TENUM.
        Set FreeBlocks = New Collection
    End If

    ' Construct the synthetic IEnumVARIANT object.
    Dim obj As TENUM
    With obj
        .pvTable = VarPtr(vTable(0))
        Set .caller = iterable
        .First = base
        .Last = base + count - 1
        .nRef = 1
        .Current = .First
    End With

    Dim MemoryBlock As LongPtr: MemoryBlock = TakeBlock(LenB(obj))
    If MemoryBlock = vbNullPtr Then Err.Raise vbErrorOutOfMemory
    CopyMemory ByVal MemoryBlock, obj, LenB(obj)

    ' The memory block now holds the reference to the iterable object.
    ' Clear it in obj without a release, IUnknown_Release releases it when done.
#If API Then
    CopyMemory ByVal VarPtr(Enumerate), MemoryBlock, vbSizeLongPtr
    CopyMemory ByVal VarPtr(obj.caller), vbNullPtr, vbSizeLongPtr
#Else
    CopyPtrByRef VarPtr(Enumerate), MemoryBlock, VarByRef.vt, VarByRef.ref
    CopyPtrByRef VarPtr(obj.caller), vbNullPtr, VarByRef.vt, VarByRef.ref
#End If

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function IUnknown_QueryInterface( _
    ByRef obj As TENUM, _
    ByRef riid As GUID, _
    ByVal ppvObj As LongPtr _
) As Long
Attribute IUnknown_QueryInterface.VB_Description = "Queries a COM object for a pointer to one of its interfaces."

    If ppvObj = vbNullPtr Then
        IUnknown_QueryInterface = E_POINTER
        Exit Function
    End If

    If IsIID_IUnknown(riid) Or IsIID_IEnumVARIANT(riid) Then
#If API Then
        CopyMemory ByVal ppvObj, VarPtr(obj), vbSizeLongPtr
#Else
        CopyPtrByRef ppvObj, VarPtr(obj), VarByRef.vt, VarByRef.ref
#End If
        IUnknown_AddRef obj
        IUnknown_QueryInterface = S_OK
    Else
        IUnknown_QueryInterface = E_NOINTERFACE
    End If

End Function


Private Function IUnknown_AddRef(ByRef obj As TENUM) As Long
Attribute IUnknown_AddRef.VB_Description = "Increments the reference count for an interface pointer to a COM object."

    obj.nRef = obj.nRef + 1
    IUnknown_AddRef = obj.nRef

End Function


Private Function IUnknown_Release(ByRef obj As TENUM) As Long
Attribute IUnknown_Release.VB_Description = "Decrements the reference count for an interface pointer to a COM object."

    obj.nRef = obj.nRef - 1
    IUnknown_Release = obj.nRef

    If obj.nRef = 0 Then
        Set obj.caller = Nothing
        ' Put the memory block on the stack for reuse.
        If FreeBlocks.count = 0 Then
            FreeBlocks.Add VarPtr(obj)
        Else
            FreeBlocks.Add VarPtr(obj), Before:=1
        End If
    End If

End Function


Private Function IEnumVARIANT_Next( _
    ByRef obj As TENUM, _
    ByVal celt As Long, _
    ByVal rgVar As LongPtr, _
    ByVal pceltFetched As LongPtr _
) As Long
Attribute IEnumVARIANT_Next.VB_Description = "Retrieves the next item in the enumeration sequence."

    If rgVar = vbNullPtr Then
        IEnumVARIANT_Next = E_INVALIDARG
        Exit Function
    End If

    ' Set pceltFetched to 0 if the pointer is provided.
    If pceltFetched <> vbNullPtr Then
#If API Then
        CopyMemory ByVal pceltFetched, 0, vbSizeLong
#Else
        CopyLngByRef pceltFetched, 0, VarByRef.vt, VarByRef.ref
#End If
    End If

    ' Only continue with loop for celt > 0.
    If celt <= 0 Then
        If celt < 0 Then
            IEnumVARIANT_Next = E_INVALIDARG
        Else
            IEnumVARIANT_Next = S_OK
        End If
        Exit Function
    End If

    ' Get the next item(s) from the iterable object.
    Dim i As Long, NumberFetched As Long
    For i = obj.Current To obj.Last
#If API Then
        If VarCopyToPtr(rgVar, obj.caller(i)) <> S_OK Then Err.Raise Err.LastDllError
#Else
        ' Call default member of caller object.
        CopyVarByRef rgVar, obj.caller(i), VarByRef.vt, VarByRef.ref
#End If
        NumberFetched = NumberFetched + 1
        If NumberFetched = celt Then Exit For
        ' Advance the pointer to the next element in the destination array.
        rgVar = rgVar + vbSizeVariant
    Next
    obj.Current = obj.Current + NumberFetched

    ' Set pceltFetched to NumberFetched if the pointer is provided.
    If pceltFetched <> vbNullPtr Then
#If API Then
        CopyMemory ByVal pceltFetched, NumberFetched, vbSizeLong
#Else
        CopyLngByRef pceltFetched, NumberFetched, VarByRef.vt, VarByRef.ref
#End If
    End If

    ' Return S_OK if the number of fetched items matches the requested number.
    If NumberFetched = celt Then
        IEnumVARIANT_Next = S_OK
    Else
        IEnumVARIANT_Next = S_FALSE
    End If

End Function


Private Function IEnumVARIANT_Skip(ByRef obj As TENUM, ByVal celt As Long) As Long
Attribute IEnumVARIANT_Skip.VB_Description = "Skips over a number of elements in the enumeration sequence."

    Select Case True
    Case celt = 0
        IEnumVARIANT_Skip = S_OK
    Case celt < 0
        IEnumVARIANT_Skip = E_INVALIDARG
    Case celt <= obj.Last - obj.Current + 1
        obj.Current = obj.Current + celt
        IEnumVARIANT_Skip = S_OK
    Case Else
        ' For overshoot set one-past-end.
        obj.Current = obj.Last + 1
        IEnumVARIANT_Skip = S_FALSE
    End Select

End Function


Private Function IEnumVARIANT_Reset(ByRef obj As TENUM) As Long
Attribute IEnumVARIANT_Reset.VB_Description = "Resets the enumeration sequence to the beginning."

    obj.Current = obj.First
    IEnumVARIANT_Reset = S_OK

End Function


Private Function IEnumVARIANT_Clone(ByRef obj As TENUM, ByVal ppEnum As LongPtr) As Long
Attribute IEnumVARIANT_Clone.VB_Description = "Creates a copy of the current state of enumeration."

    If ppEnum = vbNullPtr Then
        IEnumVARIANT_Clone = E_INVALIDARG
        Exit Function
    End If

    ' UDT assignment AddRefs caller.
    Dim Copy As TENUM: Copy = obj
    Copy.nRef = 1

    Dim MemoryBlock As LongPtr: MemoryBlock = TakeBlock(LenB(obj))
    If MemoryBlock = vbNullPtr Then
        IEnumVARIANT_Clone = E_OUTOFMEMORY
        Exit Function
    End If
    CopyMemory ByVal MemoryBlock, Copy, LenB(obj)
    IEnumVARIANT_Clone = S_OK

#If API Then
    CopyMemory ByVal ppEnum, MemoryBlock, vbSizeLongPtr
    CopyMemory ByVal VarPtr(Copy.caller), vbNullPtr, vbSizeLongPtr
#Else
    CopyPtrByRef ppEnum, MemoryBlock, VarByRef.vt, VarByRef.ref
    CopyPtrByRef VarPtr(Copy.caller), vbNullPtr, VarByRef.vt, VarByRef.ref
#End If

End Function


Private Function TakeBlock(ByVal NBytes As Long) As LongPtr
Attribute TakeBlock.VB_Description = "Returns a free memory block from the stack, or a new one if the stack is empty."

    If FreeBlocks.count = 0 Then
        TakeBlock = CoTaskMemAlloc(NBytes)
    Else
        TakeBlock = FreeBlocks.Item(1)
        FreeBlocks.Remove 1
    End If

End Function


Private Function IsIID_IUnknown(ByRef id As GUID) As Boolean
Attribute IsIID_IUnknown.VB_Description = "Returns True if id is IID_IUnknown GUID."

'    Const IID_IUnknown As String = "{00000000-0000-0000-C000-000000000046}"
    IsIID_IUnknown = _
        (id.Data1 = &H0) And _
        (id.Data2 = &H0) And _
        (id.Data3 = &H0) And _
        (id.Data4(0) = &HC0) And _
        (id.Data4(1) = &H0) And _
        (id.Data4(2) = &H0) And _
        (id.Data4(3) = &H0) And _
        (id.Data4(4) = &H0) And _
        (id.Data4(5) = &H0) And _
        (id.Data4(6) = &H0) And _
        (id.Data4(7) = &H46)

End Function


Private Function IsIID_IEnumVARIANT(ByRef id As GUID) As Boolean
Attribute IsIID_IEnumVARIANT.VB_Description = "Returns True if id is IID_IEnumVARIANT GUID."

'    Const IID_IEnumVARIANT As String = "{00020404-0000-0000-C000-000000000046}"
    IsIID_IEnumVARIANT = _
        (id.Data1 = &H20404) And _
        (id.Data2 = &H0) And _
        (id.Data3 = &H0) And _
        (id.Data4(0) = &HC0) And _
        (id.Data4(1) = &H0) And _
        (id.Data4(2) = &H0) And _
        (id.Data4(3) = &H0) And _
        (id.Data4(4) = &H0) And _
        (id.Data4(5) = &H0) And _
        (id.Data4(6) = &H0) And _
        (id.Data4(7) = &H46)

End Function


#If API = False Then
Private Sub InitializeVarByRef()
Attribute InitializeVarByRef.VB_Description = "Initializes the Variant ByRef construct."

    VarByRef.vt = VarPtr(VarByRef.ref)
    CopyMemory VarByRef.vt, VBA.vbInteger Or VT_BYREF, vbSizeInteger

End Sub
#End If


#If API = False Then
Private Sub CopyVarByRef( _
    ByVal address As LongPtr, _
    ByVal value As Variant, _
    ByRef vt As Variant, _
    ByRef ref As Variant _
)
Attribute CopyVarByRef.VB_Description = "Copies a Variant to a memory address using the Variant ByRef construct."

    VarByRef.ref = address
    vt = VBA.vbVariant Or VT_BYREF
    If VBA.IsObject(value) Then
        Set ref = value
    Else
        ref = value
    End If

End Sub
#End If


#If API = False Then
Private Sub CopyLngByRef( _
    ByVal address As LongPtr, _
    ByVal value As Long, _
    ByRef vt As Variant, _
    ByRef ref As Variant _
)
Attribute CopyLngByRef.VB_Description = "Copies a Long to a memory address using the Variant ByRef construct."

    VarByRef.ref = address
    vt = VBA.vbLong Or VT_BYREF
    ref = value

End Sub
#End If


#If API = False Then
Private Sub CopyPtrByRef( _
    ByVal address As LongPtr, _
    ByVal value As LongPtr, _
    ByRef vt As Variant, _
    ByRef ref As Variant _
)
Attribute CopyPtrByRef.VB_Description = "Copies a pointer to a memory address using the Variant ByRef construct."

#If Win64 Then
    ' A LongLong by reference is not supported (error 458). The 8 bytes of the
    ' LongLong are read as a Currency by changing the type only, and that bit
    ' pattern value is copied by reference, nothing is calculated.
    VarByRef.ref = value
    vt = VBA.vbCurrency
    Dim Bits As Currency: Bits = ref

    VarByRef.ref = address
    vt = VBA.vbCurrency Or VT_BYREF
    ref = Bits
#Else
    VarByRef.ref = address
    vt = VBA.vbLong Or VT_BYREF
    ref = value
#End If

End Sub
#End If
