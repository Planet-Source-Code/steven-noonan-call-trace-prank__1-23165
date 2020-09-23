Attribute VB_Name = "Module1"
Function mReplaceCharacter(strOrigChar, strReplaceChar, strString)
'****************************************************************
' Name: mReplaceCharacter
' Description:Replaces all instances of substring A with sub
'     string B in a string
' By: Ian Ippolito
' Inputs:strString==string to do replacing on
'strOrigChar==orig substring
'strReplaceChar==substring to replace orig substring

' Returns:strString after replacing all instances of strOrigChar with strReplaceChar
' Assumes:None
' Side Effects:None
'****************************************************************
       
'     '**********************************
'     'changes all strOrigChar
'     ' to
'     ' in strString
'     '**********************************
    Dim strResult
    strResult = ""
    '     'traverse string
    Dim intIndex
    For intIndex = 1 To Len(strString)
    
        If (Mid(strString, intIndex, Len(strOrigChar)) = strOrigChar) Then
            '*************
            'match found
            '*************
            strResult = strResult + strReplaceChar
            intIndex = intIndex + Len(strOrigChar) - 1
        Else
            '*************
            'no match
            '*************
             strResult = strResult + Mid(strString, intIndex, 1)
        End If
        
    Next

    mReplaceCharacter = strResult
    
End Function
