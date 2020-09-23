Attribute VB_Name = "modSplit"

'***************************************************************
' Name: My own Split function for Vb5
'
' Description: It's my version of the 'Split' Vb 6
'              function for Vb5 Users ...
'
' By: Jérémy cluzel
'     jcluzel@hotmail.com
'
' Inputs:   chaine: string to separate
'           separ: string used as separator, may be more than
'                  one charactere long. Ex: '|', ';', or 'XX'...
'           tableau() : array of string used to store elements
'                       after separation
'           nb_elem : integer used to store the number
'                     of elements (in the array above...)
'
' Returns:  0, if everything works...
'           the number of the error generated else ...
'
'Assumes:None
'
'Side Effects: If you want to use it under Vb6 rename it
'              to 'split_' or whatever you want...
'
'***************************************************************

Option Explicit
Public Function split_(chaine As String, separ As String, tableau() As String, nb_elem As Integer) As Integer
    On Error GoTo erreur
    Dim pos_act As Long, pos_occur As Long
    If Right(chaine, 1) <> separ Then chaine = chaine & separ
    Do
        pos_act = pos_occur + Len(separ)
        pos_occur = InStr(pos_act, chaine, separ)
        If pos_occur <> 0 Then
            ReDim Preserve tableau(nb_elem)
            tableau(nb_elem) = Mid(chaine, pos_act, pos_occur - pos_act)
            nb_elem = nb_elem + 1
        End If
    Loop Until pos_occur = 0
    split_ = 0
Exit Function

erreur:
    'split_ = Err.Number
End Function


