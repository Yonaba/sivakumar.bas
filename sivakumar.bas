Attribute VB_Name = "sivakumar"
'  sivakumar.bas
'    Version: 1.0.093020
'    A VBA module to calculate the rainy season onset/ending dates and the count of occurences of dry spells
'    R.Yonaba <roland.yonaba@gmail.com>
'    License: MIT-LICENSE <https://opensource.org/licenses/MIT>
'    Academic paper reference:
'      Sivakumar, M., 1988. Predicting rainy season potential from the onset of rains in Southern Sahelian and Sudanian climatic zones of West Africa.
'      Agricultural and forest meteorology 42, 295–305.

Option Explicit

'Function sivakumar_onset(rng, rf_thsld = 1, dry_spell = 7)
'Computes the rainy season onset date following the Sivakumar criterion
'Argument rng (required) (range): an array of 365 (or 366) daily rainfall values
'Argument rf_thsld (optional) (double): minimum rainfall threshold. Defaults to 1 mm
'Argument dry_spell (optional) (integer): dry spell length observed on 30 consecutive days validating the rainy season onset condition. Defaults to 7 days
'Returns (integer): the rainy season onset date
Public Function sivakumar_onset(rng As Range, Optional rf_thsld As Double = 1, Optional dry_spell As Integer = 7) As Integer
    Dim roll20 As Variant
    Dim n As Integer, n_start As Integer
    n = WorksheetFunction.CountA(rng)
    
    If (n = 365) Then
        n_start = 120
    ElseIf (n = 366) Then
        n_start = 121
    Else
        'MsgBox "The given range should feature 365 (or 366) values.", vbExclamation, "Not enough values"
        roll20 = CVErr(xlErrNA)
    End If
    
    Dim i As Integer, j As Integer, spl As Integer, cur_spl As Integer, n30 As Integer
    Dim cur_roll As Double
    
    If rng(n_start, 1) >= 20 Then
      roll20 = n_start
    ElseIf rng(n_start + 1) >= 20 Or (rng(n_start) + rng(n_start + 1)) >= 20 Then
      roll20 = n_start + 1
    Else
      For i = (n_start + 2) To n
          cur_roll = rng(i, 1) + rng(i - 1, 1) + rng(i - 2, 1)
          spl = 0
          cur_spl = 0
          If (cur_roll >= 20) Then
              n30 = WorksheetFunction.Min(i + 30, n)
              For j = i + 1 To n30
                  If rng(j, 1) <= rf_thsld Then
                      cur_spl = cur_spl + 1
                  Else
                      spl = WorksheetFunction.Max(spl, cur_spl)
                      cur_spl = 0
                  End If
              Next j
              If spl <= dry_spell Then
                  roll20 = i
                  Exit For
              End If
          End If
      Next i
    End If
    sivakumar_onset = roll20
End Function

'Function sivakumar_ending(rng, rf_thsld = 1, dry_spell = 20)
'Computes the ending date of the rainy season following the Sivakumar criterion
'Argument rng (required) (range): an array of 365 (or 366) daily rainfall values
'Argument rf_thsld (optional) (double): minimum rainfall threshold. Defaults to 1 mm
'Argument dry_spell (optional) (entier): dry spell length observed after the 1st September validating the rainy season ending condition. Defaults to 20 days
'Returns (integer): the rainy season ending date
Public Function sivakumar_ending(rng As Range, Optional rf_thsld As Double = 1, Optional dry_spell As Integer = 20) As Integer
    Dim n As Integer, n_start As Integer, offset As Integer
    n = WorksheetFunction.CountA(rng)
    
    If (n = 365) Then
        n_start = 244
    ElseIf (n = 366) Then
        n_start = 245
    Else
        'MsgBox "The given range should feature 365 (or 366) values.", vbExclamation, "Not enough values"
        offset = CVErr(xlErrNA)
    End If
    
    Dim i As Integer, j As Integer, n20 As Integer
    Dim has_rf As Boolean
    
    For i = n_start + 1 To n
      n20 = WorksheetFunction.Min(i + dry_spell, n)
      has_rf = False
      For j = i + 1 To n20
        If rng(j, 1) > rf_thsld Then has_rf = True
      Next j
      If has_rf = False Then
        offset = i
        Exit For
      End If
    Next i
    
    sivakumar_ending = offset
End Function

'Function count_dry_spells(rng, spell_length, rf_thsld = 1, onset = 1, offset = 365 or 366)
'Counts the number of occurrence of dry spells of a given duration
'Argument rng (required) (range): an array of 365 (or 366) daily rainfall values
'Argument spell_length (required) (integer): the dry spells length for which the count of occurrences is required
'Argument rf_thsld (optional) (double): minimum rainfall threshold. Defaults to 1 mm
'Argument onset (optional) (integer): the starting date (in the year) from which dry spells are considered. Typically, should be the rainy season onset. Defaults to 1 (January 1st).
'Argument ending (optional) (integer): the ending date (in the year) up to which dry spells are considered. Typically, should be the rainy season ending. Defaults to 365 (or 366) (December 31st).
'Returns (integer): the number of occurrences of dry spells of a given duration
Public Function count_dry_spells(rng As Range, spell_length As Integer, Optional rf_thsld As Double = 1, Optional onset As Integer = 1, Optional ending As Integer = 365) As Integer
    Dim n As Integer, n_start As Integer, count As Integer
    n = WorksheetFunction.CountA(rng)
    
    If (n = 366) Then
        ending = n
    ElseIf (n <> 365) Then
        'MsgBox "The given range should feature 365 (or 366) values.", vbExclamation, "Not enough values"
        count = CVErr(xlErrNA)
    End If
    
    count = 0
    
    Dim i As Integer, cur_spl As Integer
    cur_spl = 0
    For i = onset To ending
      If rng(i, 1) <= rf_thsld Then
        cur_spl = cur_spl + 1
      Else
        count = IIf(cur_spl = spell_length, count + 1, count)
        cur_spl = 0
      End If
    Next i
    
    count_dry_spells = count
End Function

'Function longest_dry_spell(rng, rf_thsld = 1, onset = 1, offset = 365 or 366)
'Compute the duration of the longest dry spell
'Argument rng (required) (range): an array of 365 (or 366) daily rainfall values
'Argument rf_thsld (optional) (double): minimum rainfall threshold. Defaults to 1 mm
'Argument onset (optional) (integer): the starting date (in the year) from which dry spells are considered. Typically, should be the rainy season onset. Defaults to 1 (January 1st).
'Argument ending (optional) (integer): the ending date (in the year) up to which dry spells are considered. Typically, should be the rainy season ending. Defaults to 365 (or 366) (December 31st).
'Returns (integer): the duration of the longest dry spell
Public Function longest_dry_spell(rng As Range, Optional rf_thsld As Double = 1, Optional onset As Integer = 1, Optional ending As Integer = 365) As Integer
    Dim n As Integer, n_start As Integer, max_len As Integer
    n = WorksheetFunction.CountA(rng)
    
    If (n = 366) Then
        ending = n
    ElseIf (n <> 365) Then
        'MsgBox "The given range should feature 365 (or 366) values.", vbExclamation, "Not enough values"
        max_len = CVErr(xlErrNA)
    End If
    
    max_len = 0
        
    Dim i As Integer, cur_spl As Integer
    cur_spl = 0
    For i = onset To ending
      If rng(i, 1) <= rf_thsld Then
        cur_spl = cur_spl + 1
      Else
        max_len = WorksheetFunction.Max(max_len, cur_spl)
        cur_spl = 0
      End If
    Next i
    
    longest_dry_spell = max_len
End Function
