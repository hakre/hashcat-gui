Attribute VB_Name = "mCollection"
Option Explicit
'
' Collection related functions
'

'
' returns the number of elements in an array,
' returns 0 if there are no elements (or array is not initialized)
' returns -1 if parameter was not an array
'
Public Function COL_arrayCountElements(aArray As Variant) As Long
Dim sType As String
Dim iUBound As Long
Dim bInitialized As Boolean

    sType = TypeName(aArray)
    If Right(sType, 2) <> "()" Then
        COL_arrayCountElements = -1
        Exit Function
    End If

    On Error Resume Next
        iUBound = UBound(aArray)
        If Err.Number = 0 Then
            bInitialized = True
        End If
    On Error GoTo 0
    
    If bInitialized Then
        COL_arrayCountElements = iUBound - LBound(aArray) + 1
    End If

End Function

Public Function COL_arrayHasElements(aArray As Variant) As Boolean
    If COL_arrayCountElements(aArray) > 0 Then
        COL_arrayHasElements = True
    End If
End Function
' find sItem in array
'
' returns: 0 if not found, 1-based index position if found
Public Function COL_arrayIndex(vItem As Variant, ByRef aParams As Variant) As Long
Dim i As Long, m As Long

    'check if aParams is actually an array
    If Not IsArray(aParams) Then
        'it's not an array, so we can not find anything.
        Debug.Print "ERROR: Illegal Function Call: COL_arrayIndex on non-array"
        Exit Function
    End If

    m = COL_arrayCountElements(aParams)
    If m > 0 Then
        For i = 1 To m
            If aParams(i - 1) = vItem Then
                COL_arrayIndex = i
                Exit Function
            End If
        Next i
    End If

End Function

' convert a sNumbered string (comma seperated values) to an array of long values
' preventing invalid values (like 0)
Public Function COL_arrayNumbered(sNumbered As String) As Long()
Dim aIndexies() As String
Dim i As Long, m As Long
Dim aValues() As Long
Dim sValue As String, iValue As Long
Dim iCount As Long

    aIndexies = Split(sNumbered, ",")
    m = UBound(aIndexies)
           
    If m > -1 Then
        ReDim aValues(0 To m) As Long
    
        For i = 0 To m Step 1
            sValue = Trim(aIndexies(i))
            If sValue <> "" Then
                iValue = Val(sValue)
                aValues(iCount) = iValue
                iCount = iCount + 1
            End If
        Next i
        
        If iCount - 1 < m Then
            ReDim Preserve aValues(0 To iCount - 1)
        End If
    End If
    
COL_arrayNumbered = aValues
End Function
' find sItem in array of parameters
'
' returns: 0 if not found, 1-based index position if found
Public Function COL_index(vItem As Variant, ParamArray vParams() As Variant) As Long
Dim i As Long, m As Long
Dim aParams() As Variant

    aParams = vParams 'FIXME re-check if this is really necessary
    
    m = COL_arrayCountElements(aParams)
    If m > 0 Then
        For i = 1 To m
            If vParams(i - 1) = vItem Then
                COL_index = i
                Exit Function
            End If
        Next i
    End If
    
End Function

Public Function COL_objExists(cCol As Collection, oObject As Object) As Boolean
Dim iIndex As Long

    iIndex = COL_objIndex(cCol, oObject)
    COL_objExists = Not CBool(iIndex = 0)
    
End Function
'
' Index of an Object in a Collection
'
' @return 0 if not found, 1-n: index of oObject in cCol
'
Public Function COL_objIndex(cCol As Collection, oObject As Object) As Long
Dim oItem As Object
Dim iIndex As Long

    For Each oItem In cCol
        iIndex = iIndex + 1
        If oItem Is oObject Then
            COL_objIndex = iIndex
            Exit Function
        End If
    Next

End Function

'
' Index of an Object in a Collection that has a property sPropname containing string sSearch
'
Public Function COL_objSearchString(cCol As Collection, sSearch As String, sPropname As String) As Long
Dim oItem As Object
Dim vValue As Variant
Dim iIndex As Long

    For Each oItem In cCol
        iIndex = iIndex + 1
        vValue = CallByName(oItem, sPropname, VbGet)
        If vValue = sSearch Then
            COL_objSearchString = iIndex
            Exit Function
        End If
    Next
    
End Function


'returns the number of occurences of sSearch in sString
Public Function COL_stringCount(sString As String, sSearch As String) As Long
Dim p As Long, o As Long, c As Long

    o = 1
    Do
        p = InStr(o, sString, sSearch)
        If p = 0 Or p = Null Then
            Exit Do
        End If
        o = p + Len(sSearch)
        c = c + 1
    Loop

COL_stringCount = c
End Function


