Attribute VB_Name = "modErrorNumbers"
Option Explicit

' =======================================
' Common
' =======================================
Public Enum CommonError
    UnknownError = 1000
    NotImplemented = 1001
End Enum

' =======================================
' Dictionary
' =======================================
Public Enum DictError
    KeyNotFound = 1100
    DuplicateKey = 1101
    IndexOutOfRange = 1102
End Enum

' =======================================
' WBS Tree
' =======================================
Public Enum WbsTreeError
    NodeNotFound = 1200
    DuplicateNodeId = 1201
    InvalidStructure = 1202
End Enum

