Attribute VB_Name = "Utils"
Option Explicit

Public Declare Function ShellExecute Lib _
     "shell32.dll" Alias "ShellExecuteA" _
     (ByVal hwnd As Long, ByVal lpOperation _
     As String, ByVal lpFile As String, ByVal _
     lpParameters As String, ByVal lpDirectory _
     As String, ByVal nShowCmd As Long) As Long


Public Function LeftChr(strLine As String, cutChar As String) As String
'Quebra uma string, retornando os caracteres à esquerda de um determinado caractere (primeira ocorrencia)
            Dim i As Integer
            For i = 1 To Len(strLine)
                If (Asc(Mid(strLine, i, 1)) = Asc(cutChar)) Then
                    LeftChr = Left(strLine, i - 1)
                    Exit For
                End If
            Next i
End Function

Public Function LeftLastChr(strLine As String, cutChar As String) As String
'Quebra uma string, retornando os caracteres à esquerda de um determinado caractere (ultima ocorrencia)
            Dim i As Integer
            For i = 1 To Len(strLine)
                If (Asc(Mid(strLine, i, 1)) = Asc(cutChar)) Then
                    LeftLastChr = Left(strLine, i - 1)
                End If
            Next i
End Function

Public Function RightChr(strLine As String, cutChar As String) As String
'Quebra uma string, retornando os caracteres à direita de um determinado caractere
            Dim i As Integer
            For i = 1 To Len(strLine)
                If (Asc(Mid(strLine, i, 1)) = Asc(cutChar)) Then
                    RightChr = Right(strLine, Len(strLine) - i)
                    Exit For
                End If
            Next i
End Function

Public Function RightLastChr(strLine As String, cutChar As String) As String
'Quebra uma string, retornando os caracteres à direita de um determinado caractere (ultima ocorrencia)
            Dim i As Integer
            For i = 1 To Len(strLine)
                If (Asc(Mid(strLine, i, 1)) = Asc(cutChar)) Then
                    RightLastChr = Right(strLine, Len(strLine) - i)
                End If
            Next i
End Function


Public Function DecimalIngles(ByVal linha As String, casasDecimais As Integer) As String
'Padroniza numeros em string (1,234,567,890.43) no formato amigavel com os programas (1234567890.43)
    Dim c As String, nl As String
    Dim pFix As String, pDec As String
    Dim i As Integer, ponto As Integer
    Dim compLinha As Integer
    
    compLinha = Len(linha)
    nl = ""
    
    For i = 1 To compLinha
        c = Mid(linha, i, 1)
        If (c <> ".") And (c <> ",") Then
            nl = nl + c
        Else
            pFix = nl
            ponto = Len(nl)
        End If
    Next i
    
    If ponto <> 0 Then
        pDec = Mid(nl, (ponto + 1), casasDecimais)
    Else
        pFix = nl
        pDec = ""
    End If
    If pFix = "" Then pFix = "0"

    If casasDecimais > 0 Then
            'DecimalIngles = nl & "." & String(casasDecimais, "0")
            DecimalIngles = pFix & "." & pDec & String(casasDecimais - Len(pDec), "0")
    Else
            DecimalIngles = pFix
    End If

End Function

Public Function TiraPontos(linha As String) As String
'Retira os pontos de uma dada string
    Dim nlp As String, i As Integer
    nlp = ""
    
    For i = 1 To Len(linha)
        If (Mid(linha, i, 1) <> ".") Then nlp = nlp + Mid(linha, i, 1)
    Next i
    
    TiraPontos = nlp
End Function

Public Function strDec(n As Double, comprimento As Byte) As String
'Retorna a parte decimal de um numero em forma de string, com comprimento fixo
    Dim num As Double
    Dim sn As String
    
    num = Abs(n) 'tira o sinal
    num = num - Fix(num) 'tira a parte inteira
    num = Round(num, comprimento)
    
    If num <> 0 Then
        sn = Mid(Str(num), 3, comprimento)
    Else
        sn = ""
    End If
    
    strDec = Left(sn & "000000000000000", comprimento)
End Function

Public Function strToDec(expressao As String) As Double
    Dim nl As String, i As Integer
    Dim ponto As Integer
    Dim val As Double, decval As Double
    
    nl = Trim(expressao)
    If Len(nl) = 0 Then Exit Function

    For i = 1 To Len(nl) - 1
        If (Mid(nl, i, 1) = ".") Or (Mid(nl, i, 1) = ",") Then
            ponto = i
        End If
    Next i
    
    If ponto <> 0 Then
        val = CDec(TiraPontos(Left(nl, ponto - 1)))
        decval = (CDec(Mid(nl, (ponto + 1), Len(nl) - ponto)) / (10 ^ (Len(nl) - ponto)))
        
        If val < 0 Then
            val = val - decval
        Else
            val = val + decval
        End If
    Else
        val = CLng(nl)
    End If
    
    strToDec = val
End Function

Public Function decToStr(ByVal valor As Double, casasDecimais As Byte) As String
    decToStr = CStr(Fix(valor)) & "," & CStr(strDec(valor, casasDecimais))
End Function

Public Function decToStrIngles(ByVal valor As Double, casasDecimais As Byte) As String
    decToStrIngles = CStr(Fix(valor)) & "." & CStr(strDec(valor, casasDecimais))
End Function


Public Function dec2dms(valor As Double, tipo As Byte) As String
'Converte uma fracao decimal (-4,5) para uma notação de coordenada (-4º30'00")
    Dim min As String, segi As String
    Dim seg As Double, val As Double
    Dim txt As String, stIntVal As String, lgtVal As Integer
    
    val = Round(valor, 6)
    'val = valor
    min = CStr(Abs(Fix(((val - Fix(val)) * 60))))
    seg = ((val * 60) - Fix(val * 60)) * 60
    segi = Abs(Fix(seg))
        
    If tipo = 1 Or tipo = 2 Then 'latitude ou longitude
        
        stIntVal = CStr(Fix(Abs(val)))
        lgtVal = Len(stIntVal)
        
        If lgtVal < 3 Then
            txt = Left("000", 2 - lgtVal)
        End If
                
        txt = txt & stIntVal & "°" & _
                Left("000", 2 - Len(min)) & min & "'" & _
                Left("000", 2 - Len(segi)) & segi & "."
        txt = txt & strDec(seg, 6) & Chr(34)
                
        If tipo = 1 Then
            If valor >= 0 Then
                dec2dms = txt & "N"
            Else
                dec2dms = txt & "S"
            End If
        Else
            If valor >= 0 Then
                dec2dms = txt & "E"
            Else
                dec2dms = txt & "W"
            End If
        End If
    ElseIf tipo = 3 Then
        'DMS sem casas decimais
        dec2dms = Fix(val) & "°" & _
                Left("000", 2 - Len(min)) & min & "'" & _
                Left("000", 2 - Len(segi)) & segi & Chr(34)
    ElseIf tipo = 4 Then
        'DMS com 6 casas decimais
        dec2dms = Fix(val) & "°" & _
                Left("000", 2 - Len(min)) & min & "'" & _
                Left("000", 2 - Len(segi)) & segi & "." & _
                strDec(seg, 6) & Chr(34)
    Else
        'geral - DMS com 2 casas decimais
        dec2dms = Fix(val) & "°" & _
                Left("000", 2 - Len(min)) & min & "'" & _
                Left("000", 2 - Len(segi)) & segi & "." & _
                strDec(seg, 2) & Chr(34)
    End If
End Function

Public Function dms2dec(dms As String) As Double
'Converte uma notação de ângulo/coordenada (12º34'56.78) em numero decimal (12,58)
    Dim nl As String, c As String
    Dim i As Integer
    Dim compLinha As Integer
    Dim v As Double
    Dim item As Byte
    Dim negativo As Boolean
    
    'Exemplos:
    '45°67'89.987654"W
    '45°67'89.98W
    '-45:67:89.98
    '45º
    '46º67'
    '4:56'00.00
    'conversao = (((Sec / 60) + Min) / 60) + Grau
    
    nl = ""
    compLinha = Len(dms)
    item = 1
    negativo = False
    
    For i = 1 To compLinha
        c = Mid(dms, i, 1)
        If (c = "º") Or (c = "°") Or (c = "^") Or ((c = ":") And (item = 1)) Then
            v = CDec(nl)
            nl = ""
            item = 2
        ElseIf (c = "'") Or ((c = ":") And (item = 2)) Then
            v = v + (CDec(nl) / 60)
            nl = ""
            item = 3
        'ElseIf (c = ".") Or (c = ",") Or ((c = ":") And (item = 3)) Then
            'v = v + (CDec(nl) / 3600)
            'nl = ""
            'item = 4
        ElseIf (c = "N") Or (c = "E") Or (c = "O") Then
            negativo = False
        ElseIf (c = "W") Or (c = "S") Or (c = "L") Then
            negativo = True
        ElseIf ((c = "-") And (item = 1)) Then
            negativo = True
        ElseIf ((c = Chr(34)) Or (c = " ")) Then
            'ignorar
        Else
            nl = nl + c
        End If
    Next i
    
    If negativo Then
        v = -Abs(v) - (strToDec(nl) / 3600)
    Else
        v = Abs(v) + (strToDec(nl) / 3600)
    End If
    
    dms2dec = v
End Function

Public Function isNumerico(value As String) As Boolean
    Dim l_strOut As String
    Dim i As Integer

    For i = 1 To Len(value)
        If Not InStr("0123456789-.,", Mid(value, i, 1)) > 0 Then
            isNumerico = False
            Exit Function
        End If
    Next
    isNumerico = True
End Function

Public Function UTF8_Encode(ByVal sStr As String) As String
Dim sUtf8$, l&, lChar&
    For l& = 1 To Len(sStr)
        lChar& = AscW(Mid(sStr, l&, 1))
        If lChar& < 128 Then
            sUtf8$ = sUtf8$ + Mid(sStr, l&, 1)
        ElseIf ((lChar& > 127) And (lChar& < 2048)) Then
            sUtf8$ = sUtf8$ + Chr(((lChar& \ 64) Or 192))
            sUtf8$ = sUtf8$ + Chr(((lChar& And 63) Or 128))
        Else
            sUtf8$ = sUtf8$ + Chr(((lChar& \ 144) Or 234))
            sUtf8$ = sUtf8$ + Chr((((lChar& \ 64) And 63) Or 128))
            sUtf8$ = sUtf8$ + Chr(((lChar& And 63) Or 128))
        End If
    Next l&
    UTF8_Encode = sUtf8$
End Function


