VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' Dictionary v1.4.1
' (c) Tim Hall - https://github.com/timhall/VBA-Dictionary
'
' Drop-in replacement for Scripting.Dictionary on Mac
'
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

#Const UseScriptingDictionaryIfAvailable = True

#If Mac Or Not UseScriptingDictionaryIfAvailable Then

' dict_KeyValue 0: FormattedKey, 1: OriginalKey, 2: Value
Private dict_pKeyValues As Collection
Private dict_pKeys() As Variant
Private dict_pItems() As Variant
Private dict_pObjectKeys As Collection
Private dict_pCompareMode As CompareMethod

#Else

Private dict_pDictionary As Object

#End If

' --------------------------------------------- '
' Types
' --------------------------------------------- '

Public Enum CompareMethod
    BinaryCompare = VBA.vbBinaryCompare
    TextCompare = VBA.vbTextCompare
    DatabaseCompare = VBA.vbDatabaseCompare
End Enum

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Public Property Get CompareMode() As CompareMethod
Attribute CompareMode.VB_Description = "Set or get the string comparison method."
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
1         CompareMode = dict_pCompareMode
#Else
2         CompareMode = dict_pDictionary.CompareMode
#End If
End Property
Public Property Let CompareMode(value As CompareMethod)
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
1         If Me.Count > 0 Then
              ' Can't change CompareMode for Dictionary that contains data
              ' http://msdn.microsoft.com/en-us/library/office/gg278481(v=office.15).aspx
2             Err.Raise 5 ' Invalid procedure call or argument
3         End If

4         dict_pCompareMode = value
#Else
5         dict_pDictionary.CompareMode = value
#End If
End Property

Public Property Get Count() As Long
Attribute Count.VB_Description = "Get the number of items in the dictionary.\n"
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
1         Count = dict_pKeyValues.Count
#Else
2         Count = dict_pDictionary.Count
#End If
End Property

Public Property Get Item(Key As Variant) As Variant
Attribute Item.VB_Description = "Set or get the item for a given key."
Attribute Item.VB_UserMemId = 0
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
          Dim dict_KeyValue As Variant
1         dict_KeyValue = dict_GetKeyValue(Key)

2         If Not IsEmpty(dict_KeyValue) Then
3             If VBA.IsObject(dict_KeyValue(2)) Then
4                 Set Item = dict_KeyValue(2)
5             Else
6                 Item = dict_KeyValue(2)
7             End If
8         Else
              ' Not found -> Returns Empty
9         End If
#Else
10        If VBA.IsObject(dict_pDictionary.Item(Key)) Then
11            Set Item = dict_pDictionary.Item(Key)
12        Else
13            Item = dict_pDictionary.Item(Key)
14        End If
#End If
End Property
Public Property Let Item(Key As Variant, value As Variant)
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
1         If Me.Exists(Key) Then
2             dict_ReplaceKeyValue dict_GetKeyValue(Key), Key, value
3         Else
4             dict_AddKeyValue Key, value
5         End If
#Else
6         dict_pDictionary.Item(Key) = value
#End If
End Property
Public Property Set Item(Key As Variant, value As Variant)
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
1         If Me.Exists(Key) Then
2             dict_ReplaceKeyValue dict_GetKeyValue(Key), Key, value
3         Else
4             dict_AddKeyValue Key, value
5         End If
#Else
6         Set dict_pDictionary.Item(Key) = value
#End If
End Property

Public Property Let Key(Previous As Variant, Updated As Variant)
Attribute Key.VB_Description = "Change a key to a different key."
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
          Dim dict_KeyValue As Variant
1         dict_KeyValue = dict_GetKeyValue(Previous)

2         If Not VBA.IsEmpty(dict_KeyValue) Then
3             dict_ReplaceKeyValue dict_KeyValue, Updated, dict_KeyValue(2)
4         End If
#Else
5         dict_pDictionary.Key(Previous) = Updated
#End If
End Property

' ============================================= '
' Public Methods
' ============================================= '

''
' Add an item with the given key
'
' @param {Variant} Key
' @param {Variant} Item
' --------------------------------------------- '
Public Sub Add(Key As Variant, Item As Variant)
Attribute Add.VB_Description = "Add a new key and item to the dictionary."
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
1         If Not Me.Exists(Key) Then
2             dict_AddKeyValue Key, Item
3         Else
              ' This key is already associated with an element of this collection
4             Err.Raise 457
5         End If
#Else
6         dict_pDictionary.Add Key, Item
#End If
End Sub

''
' Check if an item exists for the given key
'
' @param {Variant} Key
' @return {Boolean}
' --------------------------------------------- '
Public Function Exists(Key As Variant) As Boolean
Attribute Exists.VB_Description = "Determine if a given key is in the dictionary."
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
1         Exists = Not IsEmpty(dict_GetKeyValue(Key))
#Else
2         Exists = dict_pDictionary.Exists(Key)
#End If
End Function

''
' Get an array of all items
'
' @return {Variant}
' --------------------------------------------- '
Public Function Items() As Variant
Attribute Items.VB_Description = "Get an array containing all items in the dictionary."
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
1         If Me.Count > 0 Then
2             Items = dict_pItems
3         Else
              ' Split("") creates initialized empty array that matches Dictionary Keys and Items
4             Items = VBA.Split("")
5         End If
#Else
6         Items = dict_pDictionary.Items
#End If
End Function

''
' Get an array of all keys
'
' @return {Variant}
' --------------------------------------------- '
Public Function Keys() As Variant
Attribute Keys.VB_Description = "Get an array containing all keys in the dictionary."
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
1         If Me.Count > 0 Then
2             Keys = dict_pKeys
3         Else
              ' Split("") creates initialized empty array that matches Dictionary Keys and Items
4             Keys = VBA.Split("")
5         End If
#Else
6         Keys = dict_pDictionary.Keys
#End If
End Function

''
' Remove an item for the given key
'
' @param {Variant} Key
' --------------------------------------------- '
Public Sub Remove(Key As Variant)
Attribute Remove.VB_Description = "Remove a given key from the dictionary."
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
          Dim dict_KeyValue As Variant
1         dict_KeyValue = dict_GetKeyValue(Key)

2         If Not VBA.IsEmpty(dict_KeyValue) Then
3             dict_RemoveKeyValue dict_KeyValue
4         Else
              ' Application-defined or object-defined error
5             Err.Raise 32811
6         End If
#Else
7         dict_pDictionary.Remove Key
#End If
End Sub

''
' Remove all items
' --------------------------------------------- '
Public Sub RemoveAll()
Attribute RemoveAll.VB_Description = "Remove all information from the dictionary."
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
1         Set dict_pKeyValues = New Collection

2         Erase dict_pKeys
3         Erase dict_pItems
#Else
4         dict_pDictionary.RemoveAll
#End If
End Sub

' ============================================= '
' Private Functions
' ============================================= '

#If Mac Or Not UseScriptingDictionaryIfAvailable Then

Private Function dict_GetKeyValue(dict_Key As Variant) As Variant
1         On Error Resume Next
2         dict_GetKeyValue = dict_pKeyValues(dict_GetFormattedKey(dict_Key))
3         Err.Clear
End Function

Private Sub dict_AddKeyValue(dict_Key As Variant, dict_Value As Variant, Optional dict_Index As Long = -1)
1         If Me.Count = 0 Then
2             ReDim dict_pKeys(0 To 0)
3             ReDim dict_pItems(0 To 0)
4         Else
5             ReDim Preserve dict_pKeys(0 To UBound(dict_pKeys) + 1)
6             ReDim Preserve dict_pItems(0 To UBound(dict_pItems) + 1)
7         End If

          Dim dict_FormattedKey As String
8         dict_FormattedKey = dict_GetFormattedKey(dict_Key)

9         If dict_Index >= 0 And dict_Index < dict_pKeyValues.Count Then
              ' Shift keys/items after + including index into empty last slot
              Dim dict_i As Long
10            For dict_i = UBound(dict_pKeys) To dict_Index + 1 Step -1
11                dict_pKeys(dict_i) = dict_pKeys(dict_i - 1)
12                If VBA.IsObject(dict_pItems(dict_i - 1)) Then
13                    Set dict_pItems(dict_i) = dict_pItems(dict_i - 1)
14                Else
15                    dict_pItems(dict_i) = dict_pItems(dict_i - 1)
16                End If
17            Next dict_i

              ' Add key/item at index
18            dict_pKeys(dict_Index) = dict_Key
19            If VBA.IsObject(dict_Value) Then
20                Set dict_pItems(dict_Index) = dict_Value
21            Else
22                dict_pItems(dict_Index) = dict_Value
23            End If

              ' Add key-value at proper index
24            dict_pKeyValues.Add Array(dict_FormattedKey, dict_Key, dict_Value), dict_FormattedKey, Before:=dict_Index + 1
25        Else
              ' Add key-value as last item
26            If VBA.IsObject(dict_Key) Then
27                Set dict_pKeys(UBound(dict_pKeys)) = dict_Key
28            Else
29                dict_pKeys(UBound(dict_pKeys)) = dict_Key
30            End If
31            If VBA.IsObject(dict_Value) Then
32                Set dict_pItems(UBound(dict_pItems)) = dict_Value
33            Else
34                dict_pItems(UBound(dict_pItems)) = dict_Value
35            End If

36            dict_pKeyValues.Add Array(dict_FormattedKey, dict_Key, dict_Value), dict_FormattedKey
37        End If
End Sub

Private Sub dict_ReplaceKeyValue(dict_KeyValue As Variant, dict_Key As Variant, dict_Value As Variant)
          Dim dict_Index As Long
          Dim dict_i As Integer

1         dict_Index = dict_GetKeyIndex(dict_KeyValue(1))

          ' Remove existing dict_Value
2         dict_RemoveKeyValue dict_KeyValue, dict_Index

          ' Add new dict_Key dict_Value back
3         dict_AddKeyValue dict_Key, dict_Value, dict_Index
End Sub

Private Sub dict_RemoveKeyValue(dict_KeyValue As Variant, Optional ByVal dict_Index As Long = -1)
          Dim dict_i As Long
1         If dict_Index = -1 Then
2             dict_Index = dict_GetKeyIndex(dict_KeyValue(1))
3         End If

4         If dict_Index >= 0 And dict_Index <= UBound(dict_pKeys) Then
              ' Shift keys/items after index down
5             For dict_i = dict_Index To UBound(dict_pKeys) - 1
6                 dict_pKeys(dict_i) = dict_pKeys(dict_i + 1)

7                 If VBA.IsObject(dict_pItems(dict_i + 1)) Then
8                     Set dict_pItems(dict_i) = dict_pItems(dict_i + 1)
9                 Else
10                    dict_pItems(dict_i) = dict_pItems(dict_i + 1)
11                End If
12            Next dict_i

              ' Resize keys/items to remove empty slot
13            If UBound(dict_pKeys) = 0 Then
14                Erase dict_pKeys
15                Erase dict_pItems
16            Else
17                ReDim Preserve dict_pKeys(0 To UBound(dict_pKeys) - 1)
18                ReDim Preserve dict_pItems(0 To UBound(dict_pItems) - 1)
19            End If
20        End If

21        dict_pKeyValues.Remove dict_KeyValue(0)
22        dict_RemoveObjectKey dict_KeyValue(1)
End Sub

Private Function dict_GetFormattedKey(dict_Key As Variant) As String
1         If VBA.IsObject(dict_Key) Then
2             dict_GetFormattedKey = dict_GetObjectKey(dict_Key)
3         ElseIf VarType(dict_Key) = VBA.vbBoolean Then
4             dict_GetFormattedKey = IIf(dict_Key, "-1__-1", "0__0")
5         ElseIf VarType(dict_Key) = VBA.vbString Then
6             dict_GetFormattedKey = dict_Key

7             If Me.CompareMode = CompareMethod.BinaryCompare Then
                  ' Collection does not have method of setting key comparison
                  ' So case-sensitive keys aren't supported by default
                  ' -> Approach: Append lowercase characters to original key
                  '    AbC -> AbC___b_, abc -> abc__abc, ABC -> ABC_____
                  Dim dict_Lowercase As String
8                 dict_Lowercase = ""

                  Dim dict_i As Integer
                  Dim dict_Char As String
                  Dim dict_Ascii As Integer
9                 For dict_i = 1 To VBA.Len(dict_GetFormattedKey)
10                    dict_Char = VBA.Mid$(dict_GetFormattedKey, dict_i, 1)
11                    dict_Ascii = VBA.Asc(dict_Char)
12                    If dict_Ascii >= 97 And dict_Ascii <= 122 Then
13                        dict_Lowercase = dict_Lowercase & dict_Char
14                    Else
15                        dict_Lowercase = dict_Lowercase & "_"
16                    End If
17                Next dict_i

18                If dict_Lowercase <> "" Then
19                    dict_GetFormattedKey = dict_GetFormattedKey & "__" & dict_Lowercase
20                End If
21            End If
22        Else
              ' For numbers, add duplicate to distinguish from strings
              ' ->  123  -> "123__123"
              '    "123" -> "123"
23            dict_GetFormattedKey = VBA.CStr(dict_Key) & "__" & CStr(dict_Key)
24        End If
End Function

Private Function dict_GetObjectKey(dict_ObjKey As Variant) As String
          Dim dict_i As Integer
1         For dict_i = 1 To dict_pObjectKeys.Count
2             If dict_pObjectKeys.Item(dict_i) Is dict_ObjKey Then
3                 dict_GetObjectKey = "__object__" & dict_i
4                 Exit Function
5             End If
6         Next dict_i

7         dict_pObjectKeys.Add dict_ObjKey
8         dict_GetObjectKey = "__object__" & dict_pObjectKeys.Count
End Function

Private Sub dict_RemoveObjectKey(dict_ObjKey As Variant)
          Dim dict_i As Integer
1         For dict_i = 1 To dict_pObjectKeys.Count
2             If dict_pObjectKeys.Item(dict_i) Is dict_ObjKey Then
3                 dict_pObjectKeys.Remove dict_i
4                 Exit Sub
5             End If
6         Next dict_i
End Sub

Private Function dict_GetKeyIndex(dict_Key As Variant) As Long
          Dim dict_i As Long
1         For dict_i = 0 To UBound(dict_pKeys)
2             If VBA.IsObject(dict_pKeys(dict_i)) And VBA.IsObject(dict_Key) Then
3                 If dict_pKeys(dict_i) Is dict_Key Then
4                     dict_GetKeyIndex = dict_i
5                     Exit For
6                 End If
7             ElseIf VBA.IsObject(dict_pKeys(dict_i)) Or VBA.IsObject(dict_Key) Then
                  ' Both need to be objects to check equality, skip
8             ElseIf dict_pKeys(dict_i) = dict_Key Then
9                 dict_GetKeyIndex = dict_i
10                Exit For
11            End If
12        Next dict_i
End Function

#End If

Private Sub Class_Initialize()
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
1         Set dict_pKeyValues = New Collection

2         Erase dict_pKeys
3         Erase dict_pItems
4         Set dict_pObjectKeys = New Collection
#Else
5         Set dict_pDictionary = CreateObject("Scripting.Dictionary")
#End If
End Sub

Private Sub Class_Terminate()
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
1         Set dict_pKeyValues = Nothing
2         Set dict_pObjectKeys = Nothing
#Else
3         Set dict_pDictionary = Nothing
#End If
End Sub
