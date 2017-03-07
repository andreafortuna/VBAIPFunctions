'============================================================================
'| This program is free software: you can redistribute it and/or modify     |
'| it under the terms of the GNU General Public License as published by     |
'| the Free Software Foundation, either version 3 of the License, or        |
'| (at your option) any later version.                                      |
'| This program is distributed in the hope that it will be useful,          |
'| but WITHOUT ANY WARRANTY; without even the implied warranty of           |
'| MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the            |
'| GNU General Public License for more details.                             |
'| You should have received a copy of the GNU General Public License        |
'| along with this program.  If not, see <http://www.gnu.org/licenses/>.    |
'============================================================================
'| IP Functions v2.0.3 (20160810)                                           |
'| https://github.com/andreafortuna/VBAIPFunctions                          |
'============================================================================
'|  Andrea Fortuna                                                          |
'|  https://andreafortuna.org                                               |
'|  andrea@andreafortuna.org                                                |
'============================================================================

'===================================================
'   IpIsValid
'===================================================
' Returns true if an ip address is formated exactly
Function IpIsValid(ByVal ip As String) As Boolean
    IpIsValid = (IpBinToStr(IpStrToBin(ip)) = ip)
End Function

'===================================================
'   IpStrToBin
'===================================================
' Text IP address to binary
Function IpStrToBin(ByVal ip As String) As Double
    Dim pos As Integer
    ip = ip + "."
    IpStrToBin = 0
    While ip <> ""
        pos = InStr(ip, ".")
        IpStrToBin = IpStrToBin * 256 + Val(Left(ip, pos - 1))
        ip = Mid(ip, pos + 1)
    Wend
End Function

'===================================================
'   IpBinToStr
'===================================================
' Binary IP address to text
Function IpBinToStr(ByVal ip As Double) As String
    Dim divEnt As Double
    Dim i As Integer
    i = 0
    IpBinToStr = ""
    While i < 4
        If IpBinToStr <> "" Then IpBinToStr = "." + IpBinToStr
        divEnt = Int(ip / 256)
        IpBinToStr = Format(ip - (divEnt * 256)) + IpBinToStr
        ip = divEnt
        i = i + 1
    Wend
End Function

'===================================================
'   IpAdd
'===================================================
Function IpAdd(ByVal ip As String, offset As Double) As String
    IpAdd = IpBinToStr(IpStrToBin(ip) + offset)
End Function

'===================================================
'   IpAnd
'===================================================
' IP logical AND
Function IpAnd(ByVal ip1 As String, ByVal ip2 As String) As String
    ' compute logical AND from right to left
    Dim result As String
    While ((ip1 <> "") And (ip2 <> ""))
        Call IpBuild(IpParse(ip1) And IpParse(ip2), result)
    Wend
    IpAnd = result
End Function

'===================================================
'   IpAdd2
'===================================================
Function IpAdd2(ByVal ip As String, offset As Double) As String
    Dim result As String
    While (ip <> "")
        offset = IpBuild(IpParse(ip) + offset, result)
    Wend
    IpAdd2 = result
End Function

'===================================================
'   IpGetByte
'===================================================
Function IpGetByte(ByVal ip As String, pos As Integer) As Integer
    pos = 4 - pos
    For i = 0 To pos
        IpGetByte = IpParse(ip)
    Next
End Function

'===================================================
'   IpSetByte
'===================================================
Function IpSetByte(ByVal ip As String, pos As Integer, newvalue As Integer) As String
    Dim result As String
    Dim byteval As Double
    i = 4
    While (ip <> "")
        byteval = IpParse(ip)
        If (i = pos) Then byteval = newvalue
        Call IpBuild(byteval, result)
        i = i - 1
    Wend
    IpSetByte = result
End Function

'===================================================
'   IpMask
'===================================================
' Returns netmask from a subnet
Function IpMask(ByVal ip As String) As String
    IpMask = IpBinToStr(IpMaskBin(ip))
End Function

'===================================================
'   IpWildMask
'===================================================
' Returns Wildcard mask from a subnet
Function IpWildMask(ByVal ip As String) As String
    IpWildMask = IpBinToStr(((2 ^ 32) - 1) - IpMaskBin(ip))
End Function

'===================================================
'   IpInvertMask
'===================================================
' Returns Wildcard mask from a subnet mask
Function IpInvertMask(ByVal mask As String) As String
    IpInvertMask = IpBinToStr(((2 ^ 32) - 1) - IpStrToBin(mask))
End Function

'===================================================
'   IpMaskLen
'===================================================
Function IpMaskLen(ByVal ipmaskstr As String) As Integer
    Dim notMask As Double
    notMask = 2 ^ 32 - 1 - IpStrToBin(ipmaskstr)
    zeroBits = 0
    Do While notMask <> 0
        notMask = Int(notMask / 2)
        zeroBits = zeroBits + 1
    Loop
    IpMaskLen = 32 - zeroBits
End Function

'===================================================
'   IpWithoutMask
'===================================================
Function IpWithoutMask(ByVal ip As String) As String
    Dim p As Integer
    p = InStr(ip, "/")
    If (p = 0) Then
        p = InStr(ip, " ")
    End If
    If (p = 0) Then
        IpWithoutMask = ip
    Else
        IpWithoutMask = Left(ip, p - 1)
    End If
End Function

'===================================================
'   IpSubnetLen
'===================================================
' Return the mask len from a subnet
Function IpSubnetLen(ByVal ip As String) As Integer
    Dim p As Integer
    p = InStr(ip, "/")
    If (p = 0) Then
        p = InStr(ip, " ")
        If (p = 0) Then
            IpSubnetLen = 32
        Else
            IpSubnetLen = IpMaskLen(Mid(ip, p + 1))
        End If
    Else
        IpSubnetLen = Val(Mid(ip, p + 1))
    End If
End Function

'===================================================
'   IpSubnetSize
'===================================================
' Returns the number of IPs in a subnet
Function IpSubnetSize(ByVal subnet As String) As Double
    IpSubnetSize = 2 ^ (32 - IpSubnetLen(subnet))
End Function

'===================================================
'   IpClearHostBits
'===================================================
Function IpClearHostBits(ByVal net As String) As String
    Dim ip As String
    ip = IpWithoutMask(net)
    IpClearHostBits = IpAnd(ip, IpMask(net)) + Mid(net, Len(ip) + 1)
End Function

'===================================================
'   IpIsInSubnet
'===================================================
' returns TRUE if "ip" is in "subnet"
Function IpIsInSubnet(ByVal ip As String, ByVal subnet As String) As Boolean
    'IpIsInSubnet = (IpAnd(ip, IpMask(subnet)) = IpWithoutMask(subnet))
    ' the following line also works with non standard subnet notation:
    IpIsInSubnet = (IpAnd(ip, IpMask(subnet)) = (IpAnd(IpWithoutMask(subnet), IpMask(subnet))))
End Function

'===================================================
'   IpSubnetVLookup
'===================================================
' tries to match an IP address against a list of subnets
Function IpSubnetVLookup(ByVal ip As String, table_array As Range, index_number As Integer) As String
    Dim previousMatch As String
    previousMatch = "0.0.0.0/0"
    IpSubnetVLookup = "Not Found"
    For a = 1 To table_array.Areas.Count
        For i = 1 To table_array.Areas(a).Rows.Count
            Dim subnet As String
            subnet = table_array.Areas(a).Cells(i, 1)
            If IpIsInSubnet(ip, subnet) And (IpSubnetLen(subnet) > IpSubnetLen(previousMatch)) Then
                previousMatch = subnet
                IpSubnetVLookup = table_array.Areas(a).Cells(i, index_number)
            End If
        Next i
    Next a
End Function

'===================================================
'   IpSubnetMatch
'===================================================
Function IpSubnetMatch(ByVal ip As String, table_array As Range) As Integer
    Dim previousMatch As String
    previousMatch = "0.0.0.0/0"
    IpSubnetMatch = 0
    For i = 1 To table_array.Rows.Count
        Dim subnet As String
        subnet = table_array.Cells(i, 1)
        If IpIsInSubnet(ip, subnet) And (IpSubnetLen(subnet) > IpSubnetLen(previousMatch)) Then
            previousMatch = subnet
            IpSubnetMatch = i
        End If
    Next i
End Function

'===================================================
'   IpSubnetIsInSubnet
'===================================================
Function IpSubnetIsInSubnet(ByVal subnet1 As String, ByVal subnet2 As String) As Boolean
    Dim Mask1 As Double
    Dim Mask2 As Double
    Mask1 = IpMaskBin(subnet1)
    Mask2 = IpMaskBin(subnet2)
    If (Mask1 < Mask2) Then
        IpSubnetIsInSubnet = False
    ElseIf IpIsInSubnet(IpWithoutMask(subnet1), subnet2) Then
        IpSubnetIsInSubnet = True
    Else
        IpSubnetIsInSubnet = False
    End If
End Function

'===================================================
'   IpSubnetInSubnetVLookup
'===================================================
Function IpSubnetInSubnetVLookup(ByVal subnet As String, table_array As Range, index_number As Integer) As String
    IpSubnetInSubnetVLookup = "Not Found"
    For i = 1 To table_array.Rows.Count
        If IpSubnetIsInSubnet(subnet, table_array.Cells(i, 1)) Then
            IpSubnetInSubnetVLookup = table_array.Cells(i, index_number)
            Exit For
        End If
    Next i
End Function

'===================================================
'   IpSubnetInSubnetMatch
'===================================================
Function IpSubnetInSubnetMatch(ByVal subnet As String, table_array As Range) As Integer
    IpSubnetInSubnetMatch = 0
    For i = 1 To table_array.Rows.Count
        If IpSubnetIsInSubnet(subnet, table_array.Cells(i, 1)) Then
            IpSubnetInSubnetMatch = i
            Exit For
        End If
    Next i
End Function

'===================================================
'   IpFindOverlappingSubnets
'===================================================
Function IpFindOverlappingSubnets(subnets_array As Range) As Variant
    Dim result_array() As Variant
    ReDim result_array(1 To subnets_array.Rows.Count, 1 To 1)
    For i = 1 To subnets_array.Rows.Count
        result_array(i, 1) = ""
        For j = 1 To subnets_array.Rows.Count
            If (i <> j) And IpSubnetIsInSubnet(subnets_array.Cells(i, 1), subnets_array.Cells(j, 1)) Then
                result_array(i, 1) = subnets_array.Cells(j, 1)
                Exit For
            End If
        Next j
    Next i
    IpFindOverlappingSubnets = result_array
End Function

'===================================================
'   IpSortArray
'===================================================
Function IpSortArray(ip_array As Range, Optional descending As Boolean = False) As Variant
    Dim s As Integer
    Dim t As Integer
    t = 0
    s = ip_array.Rows.Count
    Dim list() As Double
    ReDim list(1 To s)
    For i = 1 To s
        If (ip_array.Cells(i, 1) <> 0) Then
            t = t + 1
            list(t) = IpStrToBin(ip_array.Cells(i, 1))
        End If
    Next i
    For i = t - 1 To 1 Step -1
        For j = 1 To i
            If ((list(j) > list(j + 1)) Xor descending) Then
                Dim swap As Double
                swap = list(j)
                list(j) = list(j + 1)
                list(j + 1) = swap
            End If
        Next j
    Next i
    Dim resultArray() As Variant
    ReDim resultArray(1 To s, 1 To 1)
    For i = 1 To t
        resultArray(i, 1) = IpBinToStr(list(i))
    Next i
    IpSortArray = resultArray
End Function

'===================================================
'   IpSubnetSortArray
'===================================================
Function IpSubnetSortArray(ip_array As Range, Optional descending As Boolean = False) As Variant
    Dim s As Integer
    Dim t As Integer
    t = 0
    s = ip_array.Rows.Count
    Dim list() As String
    ReDim list(1 To s)
    For i = 1 To s
        If (ip_array.Cells(i, 1) <> 0) Then
            t = t + 1
            list(t) = ip_array.Cells(i, 1)
        End If
    Next i
    For i = t - 1 To 1 Step -1
        For j = 1 To i
            Dim m, n As Double
            m = IpStrToBin(list(j))
            n = IpStrToBin(list(j + 1))
            If (((m > n) Or ((m = n) And (IpMaskBin(list(j)) < IpMaskBin(list(j + 1))))) Xor descending) Then
                Dim swap As String
                swap = list(j)
                list(j) = list(j + 1)
                list(j + 1) = swap
            End If
        Next j
    Next i
    Dim resultArray() As Variant
    ReDim resultArray(1 To s, 1 To 1)
    For i = 1 To t
        resultArray(i, 1) = list(i)
    Next i
    IpSubnetSortArray = resultArray
End Function

'===================================================
'   IpParseRoute
'===================================================
Function IpParseRoute(ByVal route As String, ByRef nexthop As String)
    slash = InStr(route, "/")
    sp = InStr(route, " ")
    If ((slash = 0) And (sp > 0)) Then
        temp = Mid(route, sp + 1)
        sp = InStr(sp + 1, route, " ")
    End If
    If (sp = 0) Then
        IpParseRoute = route
        nexthop = ""
    Else
        IpParseRoute = Left(route, sp - 1)
        nexthop = Mid(route, sp + 1)
    End If
End Function

'===================================================
'   IpSubnetSortJoinArray
'===================================================
Function IpSubnetSortJoinArray(ip_array As Range) As Variant
    Dim s As Integer
    Dim t As Integer
    Dim a As String
    Dim b As String
    Dim nexthop1 As String
    Dim nexthop2 As String
    t = 0
    s = ip_array.Rows.Count
    Dim list() As String
    ReDim list(1 To s)
    For i = 1 To s
        If (ip_array.Cells(i, 1) <> 0) Then
            t = t + 1
            a = IpParseRoute(ip_array.Cells(i, 1), nexthop1)
            list(t) = IpClearHostBits(a) + " " + nexthop1
        End If
    Next i
    For i = t - 1 To 1 Step -1
        For j = 1 To i
            Dim m, n As Double
            a = IpParseRoute(list(j), nexthop1)
            b = IpParseRoute(list(j + 1), nexthop2)
            m = IpStrToBin(IpWithoutMask(a))
            n = IpStrToBin(IpWithoutMask(b))
            If ((m > n) Or ((m = n) And (IpMaskBin(a) < IpMaskBin(b)))) Then
                Dim swap As String
                swap = list(j)
                list(j) = list(j + 1)
                list(j + 1) = swap
            End If
        Next j
    Next i
    i = 1
    While (i < t)
        remove_next = False
        a = IpParseRoute(list(i), nexthop1)
        b = IpParseRoute(list(i + 1), nexthop2)
        If (IpSubnetIsInSubnet(a, b) And (nexthop1 = nexthop2)) Then
            list(i) = list(i + 1)
            remove_next = True
        ElseIf (IpSubnetIsInSubnet(b, a) And (nexthop1 = nexthop2)) Then
            remove_next = True
        ElseIf ((IpSubnetLen(a) = IpSubnetLen(b)) And (nexthop1 = nexthop2)) Then
            bigsubnet = Replace(IpWithoutMask(a) + "/" + Str(IpSubnetLen(a) - 1), " ", "")
            If (InStr(a, "/") = 0) Then
                bigsubnet = IpWithoutMask(a) & " " & IpMask(bigsubnet)
            Else
            End If
            If (IpSubnetIsInSubnet(b, bigsubnet)) Then
                list(i) = bigsubnet & " " & nexthop1
                remove_next = True
            End If
        End If
        
        If (remove_next) Then
            For j = i + 1 To t - 1
                list(j) = list(j + 1)
            Next j
            t = t - 1
            If (i > 1) Then i = i - 1
        Else
            i = i + 1
        End If
    Wend
    Dim resultArray() As Variant
    ReDim resultArray(1 To s, 1 To 1)
    For i = 1 To t
        resultArray(i, 1) = list(i)
    Next i
    IpSubnetSortJoinArray = resultArray
End Function

'===================================================
'   IpDivideSubnet
'===================================================
Function IpDivideSubnet(ByVal subnet As String, n As Integer, index As Integer)
    Dim ip As String
    Dim slen As Integer
    ip = IpAnd(IpWithoutMask(subnet), IpMask(subnet))
    slen = IpSubnetLen(subnet) + n
    If (slen > 32) Then
        IpDivideSubnet = "ERR subnet lenght > 32"
        Exit Function
    End If
    If (index >= 2 ^ n) Then
        IpDivideSubnet = "ERR index out of range"
        Exit Function
    End If
    ip = IpBinToStr(IpStrToBin(ip) + (2 ^ (32 - slen)) * index)
    IpDivideSubnet = Replace(ip + "/" + Str(slen), " ", "")
End Function

'===================================================
'   IpIsPrivate
'===================================================
Function IpIsPrivate(ByVal ip As String) As Boolean
    IpIsPrivate = (IpIsInSubnet(ip, "10.0.0.0/8") Or IpIsInSubnet(ip, "172.16.0.0/12") Or IpIsInSubnet(ip, "192.168.0.0/16"))
End Function

'===================================================
'   IpDiff
'===================================================
Function IpDiff(ByVal ip1 As String, ByVal ip2 As String) As Double
    Dim mult As Double
    mult = 1
    IpDiff = 0
    While ((ip1 <> "") Or (ip2 <> ""))
        IpDiff = IpDiff + mult * (IpParse(ip1) - IpParse(ip2))
        mult = mult * 256
    Wend
End Function

'===================================================
'   IpParse
'===================================================
Function IpParse(ByRef ip As String) As Integer
    On Error Resume Next
    Dim pos As Integer
    pos = InStrRev(ip, ".")
    If pos = 0 Then
        IpParse = Val(ip)
        ip = ""
    Else
        IpParse = Val(Mid(ip, pos + 1))
        ip = Left(ip, pos - 1)
    End If
End Function

'===================================================
'   IpBuild
'===================================================
Function IpBuild(ip_byte As Double, ByRef ip As String) As Double
    If ip <> "" Then ip = "." + ip
    ip = Format(ip_byte And 255) + ip
    IpBuild = ip_byte \ 256
End Function

'===================================================
'   IpMaskBin
'===================================================
Function IpMaskBin(ByVal ip As String) As Double
    Dim bits As Integer
    bits = IpSubnetLen(ip)
    IpMaskBin = (2 ^ bits - 1) * 2 ^ (32 - bits)
End Function

Function Hex2Bin(ByVal strHex As String) As Double
    Dim v As Double
    For i = 1 To Len(strHex)
        v = 16 * v + Val("&H" & Mid$(strHex, i, 1))
    Next
    Hex2Bin = v
End Function
