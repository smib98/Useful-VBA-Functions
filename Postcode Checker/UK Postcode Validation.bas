Public Function ValidateUKPostcode(strInput As String)

Dim RgExp As Variant

Set RgExp = CreateObject("VBScript.RegExp")

ValidateUKPostcode = ""

If strInput = "" Then

    ValidateUKPostcode = "Not Supplied"
    
    Exit Function

End If
    
    RgExp.Pattern = "(?:(?:A[BL]|B[ABDHLNRST]?|" _
                & "C[ABFHMORTVW]|D[ADEGHLNTY]|E[CHNX]?|F[KY]|G[LUY]?|" _
                & "H[ADGPRSUX]|I[GMPV]|JE|K[ATWY]|L[ADELNSU]?|M[EKL]?|" _
                & "N[EGNPRW]?|O[LX]|P[AEHLOR]|R[GHM]|S[AEGKLMNOPRSTWY]?|" _
                & "T[ADFNQRSW]|UB|W[ACDFNRSV]?|YO|ZE)" _
                & "\d(?:\d|[A-Z])? \d[A-Z]{2})"
    
    If RgExp.test(strInput) = True Then
    
        ValidateUKPostcode = "Valid"
        
    Else
        strInput = UCase(Replace(strInput, " ", ""))
        strInput = Replace(strInput, "_", "")
        strInput = Replace(strInput, ",", "")
        strInput = Replace(strInput, "+", "")
        strInput = Replace(strInput, "-", "")
        strInput = Replace(strInput, ":", "")
        strInput = Replace(strInput, "=", "")
        strInput = Replace(strInput, "/", "")
        strInput = Replace(strInput, "*", "")
        strInput = Replace(strInput, "?", "")
        
        If Len(strInput) = 0 Then
        
            ValidateUKPostcode = "Not Supplied"
            Exit Function
            
        ElseIf IsNumeric(strInput) Then
        
            ValidateUKPostcode = "All Numbers"
            Exit Function
            
        ElseIf Len(strInput) < 6 Then
        
            ValidateUKPostcode = "Too Short"
            Exit Function
                    
        End If
        If Mid(strInput, Len(strInput) - 2, 1) = "O" Then strInput = _
        Left(strInput, Len(strInput) - 3) & "0" & Right(strInput, 2)
        If Mid(strInput, 2, 1) = "0" Then strInput = _
        Left(strInput, 1) & "O" & Right(strInput, Len(strInput) - 2)
        
        If Left(strInput, 1) = "0" Then strInput = _
        "O" & Right(strInput, Len(strInput) - 1)
        If Mid(strInput, Len(strInput) - 2, 1) = "l" Then strInput = _
        Left(strInput, Len(strInput) - 3) & "1" & Right(strInput, 2)
         If Mid(strInput, 3, 1) = "l" Then strInput = _
        Left(strInput, 2) & "1" & Right(strInput, Len(strInput) - 3)
        If Mid(strInput, Len(strInput) - 3, 1) = "S" Then strInput = _
        Left(strInput, Len(strInput) - 3) & "5" & Right(strInput, 2)
        Select Case Len(strInput)
                 
        Case 6
             
            If RgExp.test(Left(strInput, 3) & " " & Right(strInput, 3)) = True Then
                 
                 ValidateUKPostcode = "Valid"
                 
            Else
            
            ValidateUKPostcode = "Invalid"
                 
            End If
         
        Case 7
        
             If RgExp.test(Left(strInput, 4) & " " & Right(strInput, 3)) = True Then
             
                 ValidateUKPostcode = "Valid"
                 
            Else
            
                ValidateUKPostcode = "Invalid"
         
            End If
            
        Case Else
        
            ValidateUKPostcode = "Invalid"
             
        End Select
        
    End If
    


End Function
