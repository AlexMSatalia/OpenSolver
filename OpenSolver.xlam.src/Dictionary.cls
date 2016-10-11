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
' Dictionary v1.2.0
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

' KeyValue 0: FormattedKey, 1: OriginalKey, 2: Value
Private pKeyValues As Collection
Private pKeys() As Variant
Private pItems() As Variant
Private pCompareMode As CompareMethod

#Else

Private pDictionary As Object

#End If

' --------------------------------------------- '
' Types
' --------------------------------------------- '

Public Enum CompareMethod
    BinaryCompare = vbBinaryCompare
    TextCompare = vbTextCompare
    DatabaseCompare = vbDatabaseCompare
End Enum

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Public Property Get CompareMode() As CompareMethod
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
1         CompareMode = pCompareMode
#Else
2         CompareMode = pDictionary.CompareMode
#End If
End Property
Public Property Let CompareMode(value As CompareMethod)
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
1         If Me.Count > 0 Then
              ' Can't change CompareMode for Dictionary that contains data
              ' http://msdn.microsoft.com/en-us/library/office/gg278481(v=office.15).aspx
2             Err.Raise 5 ' Invalid procedure call or argument
3         End If
          
4         pCompareMode = value
#Else
5         pDictionary.CompareMode = value
#End If
End Property

Public Property Get Count() As Long
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
1         Count = pKeyValues.Count
#Else
2         Count = pDictionary.Count
#End If
End Property

Public Property Get Item(Key As Variant) As Variant
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
          Dim KeyValue As Variant
1         KeyValue = GetKeyValue(Key)
          
2         If Not IsEmpty(KeyValue) Then
3             If IsObject(KeyValue(2)) Then
4                 Set Item = KeyValue(2)
5             Else
6                 Item = KeyValue(2)
7             End If
8         Else
              ' Not found -> Returns Empty
9         End If
#Else
10        If IsObject(pDictionary.Item(Key)) Then
11            Set Item = pDictionary.Item(Key)
12        Else
13            Item = pDictionary.Item(Key)
14        End If
#End If
End Property
Public Property Let Item(Key As Variant, value As Variant)
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
1         If Me.Exists(Key) Then
2             ReplaceKeyValue GetKeyValue(Key), Key, value
3         Else
4             AddKeyValue Key, value
5         End If
#Else
6         pDictionary.Item(Key) = value
#End If
End Property
Public Property Set Item(Key As Variant, value As Variant)
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
1         If Me.Exists(Key) Then
2             ReplaceKeyValue GetKeyValue(Key), Key, value
3         Else
4             AddKeyValue Key, value
5         End If
#Else
6         Set pDictionary.Item(Key) = value
#End If
End Property

Public Property Let Key(Previous As Variant, Updated As Variant)
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
          Dim KeyValue As Variant
1         KeyValue = GetKeyValue(Previous)
          
2         If Not IsEmpty(KeyValue) Then
3             ReplaceKeyValue KeyValue, Updated, KeyValue(2)
4         End If
#Else
5         pDictionary.Key(Previous) = Updated
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
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
1         If Not Me.Exists(Key) Then
2             AddKeyValue Key, Item
3         Else
              ' This key is already associated with an element of this collection
4             Err.Raise 457
5         End If
#Else
6         pDictionary.Add Key, Item
#End If
End Sub

''
' Check if an item exists for the given key
'
' @param {Variant} Key
' @return {Boolean}
' --------------------------------------------- '
Public Function Exists(Key As Variant) As Boolean
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
1         Exists = Not IsEmpty(GetKeyValue(Key))
#Else
2         Exists = pDictionary.Exists(Key)
#End If
End Function

''
' Get an array of all items
'
' @return {Variant}
' --------------------------------------------- '
Public Function Items() As Variant
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
1         If Me.Count > 0 Then
2             Items = pItems
3         Else
              ' Split("") creates initialized empty array that matches Dictionary Keys and Items
4             Items = Split("")
5         End If
#Else
6         Items = pDictionary.Items
#End If
End Function

''
' Get an array of all keys
'
' @return {Variant}
' --------------------------------------------- '
Public Function Keys() As Variant
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
1         If Me.Count > 0 Then
2             Keys = pKeys
3         Else
              ' Split("") creates initialized empty array that matches Dictionary Keys and Items
4             Keys = Split("")
5         End If
#Else
6         Keys = pDictionary.Keys
#End If
End Function

''
' Remove an item for the given key
'
' @param {Variant} Key
' --------------------------------------------- '
Public Sub Remove(Key As Variant)
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
          Dim KeyValue As Variant
1         KeyValue = GetKeyValue(Key)
          
2         If Not IsEmpty(KeyValue) Then
3             RemoveKeyValue KeyValue
4         Else
              ' Application-defined or object-defined error
5             Err.Raise 32811
6         End If
#Else
7         pDictionary.Remove Key
#End If
End Sub

''
' Remove all items
' --------------------------------------------- '
Public Sub RemoveAll()
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
1         Set pKeyValues = New Collection
          
2         Erase pKeys
3         Erase pItems
#Else
4         pDictionary.RemoveAll
#End If
End Sub

' ============================================= '
' Private Functions
' ============================================= '

#If Mac Or Not UseScriptingDictionaryIfAvailable Then

Private Function GetKeyValue(Key As Variant) As Variant
1         On Error Resume Next
2         GetKeyValue = pKeyValues(GetFormattedKey(Key))
3         Err.Clear
End Function

Private Sub AddKeyValue(Key As Variant, value As Variant, Optional Index As Long = -1)
1         If Me.Count = 0 Then
2             ReDim pKeys(0 To 0)
3             ReDim pItems(0 To 0)
4         Else
5             ReDim Preserve pKeys(0 To UBound(pKeys) + 1)
6             ReDim Preserve pItems(0 To UBound(pItems) + 1)
7         End If
          
          Dim FormattedKey As String
8         FormattedKey = GetFormattedKey(Key)
          
9         If Index > 0 And Index <= pKeyValues.Count Then
              Dim i As Long
10            For i = UBound(pKeys) To Index Step -1
11                pKeys(i) = pKeys(i - 1)
12                If IsObject(pItems(i - 1)) Then
13                    Set pItems(i) = pItems(i - 1)
14                Else
15                    pItems(i) = pItems(i - 1)
16                End If
17            Next i
              
18            pKeys(Index - 1) = Key
19            If IsObject(value) Then
20                Set pItems(Index - 1) = value
21            Else
22                pItems(Index - 1) = value
23            End If
              
24            pKeyValues.Add Array(FormattedKey, Key, value), FormattedKey, Before:=Index
25        Else
26            pKeys(UBound(pKeys)) = Key
27            If IsObject(value) Then
28                Set pItems(UBound(pItems)) = value
29            Else
30                pItems(UBound(pItems)) = value
31            End If
              
32            pKeyValues.Add Array(FormattedKey, Key, value), FormattedKey
33        End If
End Sub

Private Sub ReplaceKeyValue(KeyValue As Variant, Key As Variant, value As Variant)
          Dim Index As Long
          Dim i As Integer
          
1         For i = 0 To UBound(pKeys)
2             If pKeys(i) = KeyValue(1) Then
3                 Index = i + 1
4                 Exit For
5             End If
6         Next i
          
          ' Remove existing value
7         RemoveKeyValue KeyValue, Index
          
          ' Add new key value back
8         AddKeyValue Key, value, Index
End Sub

Private Sub RemoveKeyValue(KeyValue As Variant, Optional ByVal Index As Long = -1)
          Dim i As Long
1         If Index = -1 Then
2             For i = 0 To UBound(pKeys)
3                 If pKeys(i) = KeyValue(1) Then
4                     Index = i
5                 End If
6             Next i
7         Else
8             Index = Index - 1
9         End If
          
10        If Index >= 0 And Index <= UBound(pKeys) Then
11            For i = Index To UBound(pKeys) - 1
12                pKeys(i) = pKeys(i + 1)
                  
13                If IsObject(pItems(i + 1)) Then
14                    Set pItems(i) = pItems(i + 1)
15                Else
16                    pItems(i) = pItems(i + 1)
17                End If
18            Next i
              
19            If UBound(pKeys) = 0 Then
20                Erase pKeys
21                Erase pItems
22            Else
23                ReDim Preserve pKeys(0 To UBound(pKeys) - 1)
24                ReDim Preserve pItems(0 To UBound(pItems) - 1)
25            End If
26        End If
          
27        pKeyValues.Remove KeyValue(0)
End Sub

Private Function GetFormattedKey(Key As Variant) As String
1         GetFormattedKey = CStr(Key)
2         If Me.CompareMode = CompareMethod.BinaryCompare Then
              ' Collection does not have method of setting key comparison
              ' So case-sensitive keys aren't supported by default
              ' -> Approach: Append lowercase characters to original key
              '    AbC -> AbC__b, abc -> abc__abc, ABC -> ABC
              '    Won't work in very strange cases, but should work for now
              '    AbBb -> AbBb__bb matches AbbB -> AbbB__bb
              Dim Lowercase As String
3             Lowercase = ""
              
              Dim i As Integer
              Dim Ascii As Integer
              Dim Char As String
4             For i = 1 To Len(GetFormattedKey)
5                 Char = VBA.Mid$(GetFormattedKey, i, 1)
6                 Ascii = Asc(Char)
7                 If Ascii >= 97 And Ascii <= 122 Then
8                     Lowercase = Lowercase & Char
9                 End If
10            Next i
              
11            If Lowercase <> "" Then
12                GetFormattedKey = GetFormattedKey & "__" & Lowercase
13            End If
14        End If
End Function

#End If

Private Sub Class_Initialize()
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
1         Set pKeyValues = New Collection
          
2         Erase pKeys
3         Erase pItems
#Else
4         Set pDictionary = CreateObject("Scripting.Dictionary")
#End If
End Sub

Private Sub Class_Terminate()
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
1         Set pKeyValues = Nothing
#Else
2         Set pDictionary = Nothing
#End If
End Sub

