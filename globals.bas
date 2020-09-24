Attribute VB_Name = "Globals"
Option Explicit
Global BA(7) As Integer
Global IPType As String
Global iRange As Integer
Global iMaskEndPosition As Integer
Global spacelen As String
Sub Main()
  On Error Resume Next
  Dim C As Integer
  spacelen = vbTab & vbTab
  BA(0) = 1
  For C = 1 To UBound(BA)
    BA(C) = 2 ^ C
  Next C
  frmmain.Show
End Sub
Public Function ConvertBin(iValue As Integer) As String
  On Error Resume Next
  Dim C As Integer
  Dim TempValue As Integer
  Dim tempdata As String
  If iValue = Null Then ConvertBin = "00000000": Exit Function
  TempValue = 0
  For C = 7 To 0 Step -1
    If TempValue + BA(C) <= iValue Then
      tempdata = tempdata & "1"
      TempValue = TempValue + BA(C)
    Else
      tempdata = tempdata & "0"
    End If
  Next C
  ConvertBin = tempdata
End Function

Public Function GetBinNetID(strip As String, StrSubnetMask As String) As String
  On Error Resume Next
  Dim pos As Integer, tempnetid As String, X As String, Y As String, z As String
  pos = 1
  Do While pos <> Len(strip) + 1
    If Mid(strip, pos, 1) <> "." Then
      X = Mid(strip, pos, 1)
      Y = Mid(StrSubnetMask, pos, 1)
      z = (CInt(X) * CInt(Y))
      tempnetid = tempnetid & z
    Else
      tempnetid = tempnetid & "."
    End If
    pos = pos + 1
  Loop
  GetBinNetID = tempnetid
End Function

Public Function ConvertBinToIP(strBin As String) As String
  On Error Resume Next
  Dim pos As Integer, binarray, tempnetid As String, ix As Integer, X As Integer, Y As Integer, z As String
  strBin = strBin & "."
  binarray = Split(strBin, ".")
  For ix = 0 To UBound(binarray) - 1
    X = 0
    For Y = 7 To 0 Step -1
      If Mid(StrReverse(binarray(ix)), Y + 1, 1) = "1" Then
        X = X + BA(Y)
      Else
        X = X
      End If
    Next Y
    z = z & CStr(X) & "."
  Next ix
  ConvertBinToIP = Left(z, Len(z) - 1)
End Function

Public Function GetIPClass(strip As String) As String
  On Error Resume Next
  Dim tempip, X As Integer
  strip = strip & "."
  tempip = Split(strip, ".")

  Select Case tempip(0)
    Case 0 To 127
      GetIPClass = "A"
      If tempip(0) = 10 Then
        IPType = "Reserved"
        Exit Function
      ElseIf tempip(0) = 127 Then
        IPType = "Loopback"
        GetIPClass = "Loopback"
        Exit Function
      End If
      IPType = "Public"
    Case 128 To 191
      GetIPClass = "B"
      If tempip(0) = 172 Then
        Select Case tempip(1)
          Case 16 To 31
            IPType = "Resreved"
        End Select
        Exit Function
      End If
      IPType = "Public"
    Case 192 To 223
      GetIPClass = "C"
      If tempip(0) = 192 And tempip(1) = 168 Then
        IPType = "Reserved"
        Exit Function
      End If
      IPType = "Public"
    Case 224 To 239
      GetIPClass = "D"
      IPType = "Multicast(RFC 1112)"
    Case 240 To 255
      GetIPClass = "E"
      IPType = "Experemential"
  End Select
End Function
Public Function GetBits(strmask As String) As Single
  On Error Resume Next
  Dim tempdata, ix As Integer, pos As Integer, itemp As Single
  strmask = strmask & "."
  tempdata = Split(strmask, ".")
  For ix = 0 To UBound(tempdata) - 1
    Select Case tempdata(ix)
      Case "255"
        itemp = itemp + 8
      Case "128"
        itemp = itemp + 1
      Case "192"
        itemp = itemp + 2
      Case "224"
        itemp = itemp + 3
      Case "240"
        itemp = itemp + 4
      Case "248"
        itemp = itemp + 5
      Case "252"
        itemp = itemp + 6
      Case "254"
        itemp = itemp + 7
    End Select
  Next ix
  GetBits = itemp
End Function

Public Function GetRange(strmask As String) As Integer
  'uses the couchie method
  Dim tempdata, ix As Integer, itemp As Integer, ipclass As String
  strmask = strmask & "."
  tempdata = Split(strmask, ".")
  For ix = 0 To UBound(tempdata) - 1
    Select Case tempdata(ix)
      Case "128"
        iRange = 256 - CInt(tempdata(ix))
        Exit For
      Case "192"
        iRange = 256 - CInt(tempdata(ix))
        Exit For
      Case "224"
        iRange = 256 - CInt(tempdata(ix))
        Exit For
      Case "240"
        iRange = 256 - CInt(tempdata(ix))
        Exit For
      Case "248"
        iRange = 256 - CInt(tempdata(ix))
        Exit For
      Case "252"
        iRange = 256 - CInt(tempdata(ix))
        Exit For
      Case "254"
        iRange = 256 - CInt(tempdata(ix))
        Exit For
      Case Else
        iRange = 256
    End Select
  Next ix
  GetRange = iRange
End Function
Public Function GetPosNetworks(strmask As String) As Integer
  Dim tempdata, ix As Integer, itemp As Integer, ipclass As String
  strmask = strmask & "."
  tempdata = Split(strmask, ".")
  For ix = 0 To UBound(tempdata) - 1
    Select Case tempdata(ix)
      Case "128"
        GetPosNetworks = 2 ^ 1
      Case "192"
        GetPosNetworks = 2 ^ 2
      Case "224"
        GetPosNetworks = 2 ^ 3
      Case "240"
        GetPosNetworks = 2 ^ 4
      Case "248"
        GetPosNetworks = 2 ^ 5
      Case "252"
        GetPosNetworks = 2 ^ 6
      Case "254"
        GetPosNetworks = 2 ^ 7
        'Case Else
        'GetPosNetworks = 1
    End Select
  Next ix
End Function
Public Function GetPosHosts(strBinaryMask As String) As String
  Dim tempdata, ix As Integer, itemp As Integer, ipclass As String, pos As Integer, BitCount As Integer
  strBinaryMask = strBinaryMask & "."
  tempdata = Split(strBinaryMask, ".")
  For ix = 0 To UBound(tempdata) - 1
    pos = 1
    Do While pos <= Len(tempdata(ix)) + 1
      If Mid(tempdata(ix), pos, 1) = "0" Then
        BitCount = BitCount + 1
      End If
      pos = pos + 1
    Loop
  Next ix
  'Takes 2 to the BitCount power that is not used by NetID - 2 (for netid and broadcast)
  GetPosHosts = Format((2 ^ BitCount) - 2, "###,###,###,###")
End Function
Private Function GetMaskEndPosition(strmask As String) As Integer
  Dim tmpmask, X As Integer
  strmask = strmask & "."
  tmpmask = Split(strmask, ".")
  For X = 0 To UBound(tmpmask) - 1
    Select Case tmpmask(X)
      Case "255"
        GetMaskEndPosition = X + 1
        iMaskEndPosition = X + 1
      Case Else
        Exit For
    End Select
  Next X
End Function
Public Sub LoadNetID(strNetID As String, strmask As String, ipRange As Integer)
  Dim inet As Integer, ibroad As Integer, X As Integer, ipleft As String, iptemp, imaskend As Integer
  imaskend = GetMaskEndPosition(strmask)
  strNetID = Mid(strNetID, 1, InStrRev(strNetID, "/", Len(strNetID)) - 1)
  strNetID = strNetID & "."
  iptemp = Split(strNetID, ".")
  If ipRange = 0 Then ipRange = 256
  For X = 0 To imaskend - 1
    ipleft = ipleft & iptemp(X) & "."
  Next X
  frmmain.lstNetIDs.Clear
  If ipRange <> 256 Then
    For X = 0 To 255 Step ipRange
      With frmmain.lstNetIDs
        iptemp = Split(ipleft & "x", ".")
        Select Case UBound(iptemp) + 1
          Case 1
            .AddItem ipleft & X & ".0.0.0" & spacelen & ipleft & X + (ipRange - 1) & ".255.255.255"
          Case 2
            .AddItem ipleft & X & ".0.0" & spacelen & ipleft & X + (ipRange - 1) & ".255.255"
          Case 3
            .AddItem ipleft & X & ".0" & spacelen & ipleft & X + (ipRange - 1) & ".255"
          Case 4
            .AddItem ipleft & X & spacelen & ipleft & X + (ipRange - 1)
        End Select
      End With
      DoEvents
    Next X
  Else
    With frmmain.lstNetIDs
      iptemp = Split(ipleft & "x", ".")
      Select Case UBound(iptemp) + 1
        Case 1
          .AddItem ipleft & "0.0.0.0" & spacelen & ipleft & (ipRange - 1) & ".255.255.255"
        Case 2
          .AddItem ipleft & "0.0.0" & spacelen & ipleft & (ipRange - 1) & ".255.255"
        Case 3
          .AddItem ipleft & "0.0" & spacelen & ipleft & (ipRange - 1) & ".255"
        Case 4
          .AddItem ipleft & "0" & spacelen & ipleft & (ipRange - 1)
      End Select
    End With
    DoEvents
  End If
End Sub
Public Sub defaultmask(strip As String)
  Dim X As Integer
  With frmmain
    Select Case GetIPClass(strip)
      Case "A"
        .txtsm(0).Text = "255"
        For X = 1 To .txtsm.Count - 1
          .txtsm(X).Text = "0"
        Next X
      Case "B"
        .txtsm(0).Text = "255"
        .txtsm(1).Text = "255"
        For X = 2 To .txtsm.Count - 1
          .txtsm(X).Text = "0"
        Next X
      Case "C"
        .txtsm(0).Text = "255"
        .txtsm(1).Text = "255"
        .txtsm(2).Text = "255"
        For X = 3 To .txtsm.Count - 1
          .txtsm(X).Text = "0"
        Next X
      Case Else
        For X = 0 To .txtsm.Count - 1
          .txtsm(X).Text = "255"
        Next X
    End Select
  End With
End Sub

Public Function IsGoodIP(strip As String) As String
  Dim tempdata, tempinfo As String, X As Integer, Y As Integer, tempgood As String, temptype As String
  Dim tempclass As String
  tempclass = GetIPClass(strip)
  strip = Left$(strip, Len(strip) - 1)
  tempgood = ""
  If tempclass = "D" Or tempclass = "E" Then IsGoodIP = "No (Invalid Class)": Exit Function
  For X = 0 To frmmain.lstNetIDs.ListCount - 1
    tempinfo = frmmain.lstNetIDs.List(X) & spacelen
    tempdata = Split(tempinfo, spacelen)
    For Y = 0 To UBound(tempdata) - 1
      If strip = tempdata(Y) Then
        tempgood = "No"
        If Y = 0 Then
          temptype = "Network ID"
        Else
          temptype = "Broadcast ID"
        End If
      End If
    Next Y
    If tempgood = "No" Then Exit For
  Next X
  If tempgood = "No" Then
    IsGoodIP = tempgood & " (" & temptype & ")"
  Else
    IsGoodIP = "Yes"
  End If
End Function
Public Sub HighlightNetworkID(strNetID As String)
  Dim X As Integer, tempdata
  For X = 0 To frmmain.lstNetIDs.ListCount - 1
    tempdata = Split(frmmain.lstNetIDs.List(X), spacelen)
    If strNetID = tempdata(0) Then
      frmmain.lstNetIDs.ListIndex = X
      Exit For
    End If
  Next X
  frmmain.txtip(0).SetFocus
End Sub
Public Sub Disable()
Dim X As Integer
  With frmmain
    .Frame1.Enabled = False
    .Frame2.Enabled = False
    .Frame3.Enabled = False
    For X = 2 To .lbl.Count - 1
      .lbl(X).Enabled = False
    Next X
  End With
End Sub
Public Sub Enable()
Dim X As Integer
  With frmmain
    .Frame1.Enabled = True
    .Frame2.Enabled = True
    .Frame3.Enabled = True
    For X = 2 To .lbl.Count - 1
      .lbl(X).Enabled = True
    Next X
  End With
End Sub
