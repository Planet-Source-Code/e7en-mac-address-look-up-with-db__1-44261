Attribute VB_Name = "modMAC"
Dim sMACcodes() As String

Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2003 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const NO_ERROR = 0

Private Declare Function inet_addr Lib "wsock32.dll" _
  (ByVal s As String) As Long

Private Declare Function SendARP Lib "iphlpapi.dll" _
  (ByVal DestIP As Long, _
   ByVal SrcIP As Long, _
   pMacAddr As Long, _
   PhyAddrLen As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (dst As Any, _
   src As Any, _
   ByVal bcount As Long)
   


Public Function GetMAC(sIP As String) As String

   Dim sRemoteMacAddress As String
   
   If Len(sIP) > 0 Then
   
      If GetRemoteMACAddress(sIP, sRemoteMacAddress) Then
         GetMAC = sRemoteMacAddress
      Else
         GetMAC = "(SendARP call failed)"
      End If
      
   End If

End Function


Private Function GetRemoteMACAddress(ByVal sRemoteIP As String, sRemoteMacAddress As String) As Boolean

   Dim dwRemoteIP As Long
   Dim pMacAddr As Long
   Dim bpMacAddr() As Byte
   Dim PhyAddrLen As Long
   Dim cnt As Long
   Dim Tmp As String
    
  'convert the string IP into
  'an unsigned long value containing
  'a suitable binary representation
  'of the Internet address given
   dwRemoteIP = inet_addr(sRemoteIP)
   
   If dwRemoteIP <> 0 Then
   
     'set PhyAddrLen to 6
      PhyAddrLen = 6
   
     'retrieve the remote MAC address
      If SendARP(dwRemoteIP, 0&, pMacAddr, PhyAddrLen) = NO_ERROR Then
      
         If pMacAddr <> 0 And PhyAddrLen <> 0 Then
      
           'returned value is a long pointer
           'to the mac address, so copy data
           'to a byte array
            ReDim bpMacAddr(0 To PhyAddrLen - 1)
            CopyMemory bpMacAddr(0), pMacAddr, ByVal PhyAddrLen
          
           'loop through array to build string
            For cnt = 0 To PhyAddrLen - 1
               
               If bpMacAddr(cnt) = 0 Then
                  Tmp = Tmp & "00:"
               Else
                  Tmp = Tmp & Hex$(bpMacAddr(cnt)) & ":"
               End If
         
            Next
           
           'remove the trailing dash
           'added above and return True
            If Len(Tmp) > 0 Then
               sRemoteMacAddress = Left$(Tmp, Len(Tmp) - 1)
               GetRemoteMACAddress = True
            End If

            Exit Function
         
         Else
            GetRemoteMACAddress = False
         End If
            
      Else
         GetRemoteMACAddress = False
      End If  'SendARP
      
   Else
      GetRemoteMACAddress = False
   End If  'dwRemoteIP

End Function
'--end block--'

Public Sub LoadMACdb()
Dim sTmpMACcodes As String
Dim lcount As Integer
Dim ret As Long

  Open App.Path & "\Resources\MAC codes.txt" For Binary As #1
  sTmpMACcodes = Space(LOF(1))
  Get #1, , sTmpMACcodes
  Close #1
 
ret = split_(sTmpMACcodes, vbCrLf, sMACcodes(), lcount)
End Sub

Public Function MAClookup(sMAC As String) As String
Dim X As Long

For X = LBound(sMACcodes) To UBound(sMACcodes)
    If Not Mid(sMACcodes(X), 1, 1) = "#" Then
        If Mid(sMACcodes(X), 1, 8) = Mid(sMAC, 1, 8) Then
            MAClookup = Mid(sMACcodes(X), 10)
            Exit Function
        End If
    End If
Next
End Function


