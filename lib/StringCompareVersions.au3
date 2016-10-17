;===============================================================================
;
; FunctionName:  _StringCompareVersions()
; Description:    Compare 2 strings of the FileGetVersion format [a.b.c.d].
; Syntax:          _StringCompareVersions( $s_Version1, [$s_Version2] )
; Parameter(s):  $s_Version1          - The string being compared
;                  $s_Version2        - The string to compare against
;                                         [Optional] : Default = 0.0.0.0
; Requirement(s):   None
; Return Value(s):  0 - Strings are the same (if @error=0)
;                 -1 - First string is (<) older than second string
;                  1 - First string is (>) newer than second string
;                  0 and @error<>0 - String(s) are of incorrect format:
;                        @error 1 = 1st string; 2 = 2nd string; 3 = both strings.
; Author(s):        PeteW
; Note(s):        Comparison checks that both strings contain numeric (decimal) data.
;                  Supplied strings are contracted or expanded (with 0s)
;                    MostSignificant_Major.MostSignificant_minor.LeastSignificant_major.LeastSignificant_Minor
;
;===============================================================================

Func _StringCompareVersions($s_Version1, $s_Version2 = "0.0.0.0")

; Confirm strings are of correct basic format. Set @error to 1,2 or 3 if not.
    SetError((StringIsDigit(StringReplace($s_Version1, ".", ""))=0) + 2 * (StringIsDigit(StringReplace($s_Version2, ".", ""))=0))
    If @error>0 Then Return 0; Ought to Return something!

    Local $i_Index, $i_Result, $ai_Version1, $ai_Version2

; Split into arrays by the "." separator
    $ai_Version1 = StringSplit($s_Version1, ".")
    $ai_Version2 = StringSplit($s_Version2, ".")
    $i_Result = 0; Assume strings are equal

; Ensure strings are of the same (correct) format:
;  Short strings are padded with 0s. Extraneous components of long strings are ignored. Values are Int.
    If $ai_Version1[0] <> 4 Then ReDim $ai_Version1[5]
    For $i_Index = 1 To 4
        $ai_Version1[$i_Index] = Int($ai_Version1[$i_Index])
    Next

    If $ai_Version2[0] <> 4 Then ReDim $ai_Version2[5]
    For $i_Index = 1 To 4
        $ai_Version2[$i_Index] = Int($ai_Version2[$i_Index])
    Next

    For $i_Index = 1 To 4
        If $ai_Version1[$i_Index] < $ai_Version2[$i_Index] Then; Version1 older than Version2
            $i_Result = -1
        ElseIf $ai_Version1[$i_Index] > $ai_Version2[$i_Index] Then; Version1 newer than Version2
            $i_Result = 1
        EndIf
   ; Bail-out if they're not equal
        If $i_Result <> 0 Then ExitLoop
    Next

    Return $i_Result

EndFunc ;==>_StringCompareVersions