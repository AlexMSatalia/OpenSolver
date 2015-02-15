VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyValuePair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' JWD: Obtained from https://sysmod.wordpress.com/2011/11/02/dictionary-class-in-vba-instead-of-scripting-dictionary/
' This is an implementation of a Dictionary class so that Mac can use it as well

'Unrestricted class just to hold pairs of values together and permit Dictionary object updating
Public Key As String
Public value As Variant
