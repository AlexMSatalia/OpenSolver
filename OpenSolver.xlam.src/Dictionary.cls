VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' JWD: Obtained from https://sysmod.wordpress.com/2011/11/02/dictionary-class-in-vba-instead-of-scripting-dictionary/
' This is an implementation of a Dictionary class so that Mac can use it as well

'Collection methods: Add, Count, Item, Remove
'Dictionary : .Add(Key as string, Item as variant), .CompareMode, .Count, .Exists(Key); _
   .Item(Key) - writeable, .Items, .Keys, .Remove(Key), .RemoveAll
'plus KeyValuePairs collection, KeyValuePair(Index as long), Tag as variant
' 25-11-2011 KeyValuePair helper object

Public KeyValuePairs As Collection ' open access but allows iteration
Public Tag As Variant            ' read/write unrestricted

Private Sub Class_Initialize()
   Set KeyValuePairs = New Collection
End Sub

Private Sub Class_Terminate()
   Set KeyValuePairs = Nothing
End Sub

' in Scripting.Dictionary this is writeable, here we have only vbtextCompare because we are using a Collection
Public Property Get CompareMode() As VbCompareMethod
   CompareMode = vbTextCompare   '=1; vbBinaryCompare=0
End Property

Public Property Let Item(Key As String, Item As Variant)    ' dic.Item(Key) = value ' update a scalar value for an existing key
   Let KeyValuePairs.Item(Key).value = Item
End Property

Public Property Set Item(Key As String, Item As Variant)    ' Set dic.Item(Key) = value ' update an object value for an existing key
   Set KeyValuePairs.Item(Key).value = Item
End Property

Public Property Get Item(Key As String) As Variant
   AssignVariable Item, KeyValuePairs.Item(Key).value
End Property

' Collection parameter order is Add(Item,Key); Dictionary is Add(Key,Item) so always used named arguments
Public Sub Add(Key As String, Item As Variant)
   Dim oKVP As KeyValuePair
   Set oKVP = New KeyValuePair
   oKVP.Key = Key
   If IsObject(Item) Then
      Set oKVP.value = Item
   Else
      Let oKVP.value = Item
   End If
   KeyValuePairs.Add Item:=oKVP, Key:=Key
End Sub

Public Property Get Exists(Key As String) As Boolean
   On Error Resume Next
   Exists = TypeName(KeyValuePairs.Item(Key)) > ""  ' we can have blank key, empty item
End Property

Public Sub Remove(Key As String)
   'show error if not there rather than On Error Resume Next
   KeyValuePairs.Remove Key
End Sub

Public Sub RemoveAll()
   Set KeyValuePairs = Nothing
   Set KeyValuePairs = New Collection
End Sub

Public Property Get Count() As Long
   Count = KeyValuePairs.Count
End Property

Public Property Get Items() As Variant     ' for compatibility with Scripting.Dictionary
Dim vlist As Variant, i As Long
If Me.Count > 0 Then
   ReDim vlist(0 To Me.Count - 1) ' to get a 0-based array same as scripting.dictionary
   For i = LBound(vlist) To UBound(vlist)
      AssignVariable vlist(i), KeyValuePairs.Item(1 + i).value ' could be scalar or array or object
   Next i
   Items = vlist
End If
End Property

Public Property Get Keys() As String()
Dim vlist() As String, i As Long
If Me.Count > 0 Then
   ReDim vlist(0 To Me.Count - 1)
   For i = LBound(vlist) To UBound(vlist)
      vlist(i) = KeyValuePairs.Item(1 + i).Key   '
   Next i
   Keys = vlist
End If
End Property

Public Property Get KeyValuePair(Index As Long) As Variant  ' returns KeyValuePair object
    Set KeyValuePair = KeyValuePairs.Item(1 + Index)            ' collections are 1-based
End Property

Private Sub AssignVariable(variable As Variant, value As Variant)
   If IsObject(value) Then
      Set variable = value
   Else
      Let variable = value
   End If
End Sub

Public Sub DebugPrint()
   Dim lItem As Long, lIndex As Long, vItem As Variant, oKVP As KeyValuePair
   lItem = 0
   For Each oKVP In KeyValuePairs
      lItem = lItem + 1
      Debug.Print lItem; oKVP.Key; " "; TypeName(oKVP.value);
      If InStr(1, TypeName(oKVP.value), "()") > 0 Then
         vItem = oKVP.value
         Debug.Print "("; CStr(LBound(vItem)); " to "; CStr(UBound(vItem)); ")";
         For lIndex = LBound(vItem) To UBound(vItem)
            Debug.Print " (" & CStr(lIndex) & ")"; TypeName(vItem(lIndex)); "="; vItem(lIndex);
         Next
         Debug.Print
      Else
         Debug.Print "="; oKVP.value
      End If
   Next
End Sub

'NB VBA Collection object index is 1-based, scripting.dictionary items array is 0-based
'cf Scripting.Dictionary Methods s.Add(Key, Item), s.CompareMode, s.Count, s.Exists(Key); _
   s.Item(Key) - updateable, s.Items, s.Key(Key), s.Keys, s.Remove(Key), s.RemoveAll
'Scripting.Dictionary has no index number; you can index the 0-based variant array of Items returned
'  unlike Collections which can be indexed starting at 1
'Efficient iteration is For Each varPair in thisdic.KeyValuePairs
'Another difference I introduce is that in a scripting.dictionary, the doc says
'  If key is not found when changing an item, a new key is created with the specified newitem.
'  If key is not found when attempting to return an existing item, a new key is created and its corresponding item is left empty.
'but I want to raise an error when addressing a key that does not exist
'similarly, the scripting.dictionary will create separate integer and string keys for eg 2


