# VBAIPFunctions
VBA functions for IP manipulation and IP/Subnet lookup

## Functions List


###IpIsValid
Returns true if an ip address is formated exactly as it should be:
no space, no extra zero, no incorrect value

###IpStrToBin
Converts a text IP address to binary

ex:

  IpStrToBin("1.2.3.4") returns 16909060

###IpBinToStr
Converts a binary IP address to text

ex:

  IpBinToStr(16909060) returns "1.2.3.4"

###IpAdd

ex:

  IpAdd("192.168.1.1"; 4) returns "192.168.1.5"
  
  IpAdd("192.168.1.1"; 256) returns "192.168.2.1"

###IpAnd
IP logical AND

ex:

  IpAnd("192.168.1.1"; "255.255.255.0") returns "192.168.1.0"
  

###IpAdd2
another implementation of IpAdd which not use the binary representation

###IpGetByte
get one byte from an ip address given its position

ex:

  IpGetByte("192.168.1.1"; 1) returns 192

###IpSetByte
set one byte in an ip address given its position and value

ex:

  IpSetByte("192.168.1.1"; 4; 20) returns "192.168.1.20"

###IpMask
returns an IP netmask from a subnet
both notations are accepted

ex:

  IpMask("192.168.1.1/24") returns "255.255.255.0"
  
  IpMask("192.168.1.1 255.255.255.0") returns "255.255.255.0"

###IpWildMask
returns an IP Wildcard (inverse) mask from a subnet
both notations are accepted

ex:

  IpWildMask("192.168.1.1/24") returns "0.0.0.255"
  
  IpWildMask("192.168.1.1 255.255.255.0") returns "0.0.0.255"

###IpInvertMask
returns an IP Wildcard (inverse) mask from a subnet mask
or a subnet mask from a wildcard mask

ex:

  IpWildMask("255.255.255.0") returns "0.0.0.255"
  
  IpWildMask("0.0.0.255") returns "255.255.255.0"

###IpMaskLen
returns prefix length from a mask given by a string notation (xx.xx.xx.xx)

ex:

  IpMaskLen("255.255.255.0") returns 24 which is the number of bits of the subnetwork prefix

###IpWithoutMask
removes the netmask notation at the end of the IP

ex:

  IpWithoutMask("192.168.1.1/24") returns "192.168.1.1"
  
  IpWithoutMask("192.168.1.1 255.255.255.0") returns "192.168.1.1"
  
###IpSubnetLen
get the mask len from a subnet

ex:
  IpSubnetLen("192.168.1.1/24") returns 24
  
  IpSubnetLen("192.168.1.1 255.255.255.0") returns 24

###IpSubnetSize
returns the number of addresses in a subnet

ex:

  IpSubnetSize("192.168.1.32/29") returns 8
  
  IpSubnetSize("192.168.1.0 255.255.255.0") returns 256

###IpClearHostBits
set to zero the bits in the host part of an address

ex:

  IpClearHostBits("192.168.1.1/24") returns "192.168.1.0/24"
  
  IpClearHostBits("192.168.1.193 255.255.255.128") returns "192.168.1.128 255.255.255.128"

###IpIsInSubnet
returns TRUE if "ip" is in "subnet"
subnet must have the / mask notation (xx.xx.xx.xx/yy)

ex:

  IpIsInSubnet("192.168.1.35"; "192.168.1.32/29") returns TRUE
  
  IpIsInSubnet("192.168.1.35"; "192.168.1.32 255.255.255.248") returns TRUE
  
  IpIsInSubnet("192.168.1.41"; "192.168.1.32/29") returns FALSE
  

###IpSubnetVLookup
tries to match an IP address against a list of subnets in the left-most
column of table_array and returns the value in the same row based on the
index_number

this function selects the smallest matching subnet

"ip" is the value to search for in the subnets in the first column of
     the table_array
     
"table_array" is one or more columns of data

"index_number" is the column number in table_array from which the matching
     value must be returned. The first column which contains subnets is 1.
     
note: add the subnet 0.0.0.0/0 at the end of the array if you want the
function to return a default value

###IpSubnetMatch
tries to match an IP address against a list of subnets in the left-most
column of table_array and returns the row number
this function selects the smallest matching subnet

"ip" is the value to search for in the subnets in the first column of
     the table_array
     
"table_array" is one or more columns of data

returns 0 if the IP address is not matched.

###IpSubnetIsInSubnet
returns TRUE if "subnet1" is in "subnet2"
subnets must have the / mask notation (xx.xx.xx.xx/yy)

ex:

  IpSubnetIsInSubnet("192.168.1.35/30"; "192.168.1.32/29") returns TRUE
  
  IpSubnetIsInSubnet("192.168.1.41/30"; "192.168.1.32/29") returns FALSE
  
  IpSubnetIsInSubnet("192.168.1.35/28"; "192.168.1.32/29") returns FALSE
  

###IpSubnetInSubnetVLookup
tries to match a subnet against a list of subnets in the left-most
column of table_array and returns the value in the same row based on the
index_number
the value matches if 'subnet' is equal or included in one of the subnets
in the array

"subnet" is the value to search for in the subnets in the first column of
     the table_array
     
"table_array" is one or more columns of data

"index_number" is the column number in table_array from which the matching
     value must be returned. The first column which contains subnets is 1.
     
note: add the subnet 0.0.0.0/0 at the end of the array if you want the
function to return a default value

###IpSubnetInSubnetMatch
tries to match a subnet against a list of subnets in the left-most
column of table_array and returns the row number
the value matches if 'subnet' is equal or included in one of the subnets
in the array

"subnet" is the value to search for in the subnets in the first column of
     the table_array
     
"table_array" is one or more columns of data

returns 0 if the subnet is not included in any of the subnets from the list

###IpFindOverlappingSubnets
this function must be used in an array formula
it will find in the list of subnets which subnets overlap

"SubnetsArray" is single column array containing a list of subnets, the
list may be sorted or not
the return value is also a array of the same size
if the subnet on line x is included in a larger subnet from another line,
this function returns an array in which line x contains the value of the
larger subnet
if the subnet on line x is distinct from any other subnet in the array,
then this function returns on line x an empty cell
if there are no overlapping subnets in the input array, the returned array
is empty

###IpSortArray
this function must be used in an array formula

"ip_array" is a single column array containing ip addresses
the return value is also a array of the same size containing the same
addresses sorted in ascending or descending order

"descending" is an optional parameter, if set to True the addresses are
sorted in descending order

###IpSubnetSortArray
this function must be used in an array formula

"ip_array" is a single column array containing ip subnets in "prefix/len"
or "prefix mask" notation
the return value is also an array of the same size containing the same
subnets sorted in ascending or descending order

"descending" is an optional parameter, if set to True the subnets are
sorted in descending order

###IpParseRoute
this function is used by IpSubnetSortJoinArray to extract the subnet
and next hop in route
the supported formats are

10.0.0.0 255.255.255.0 1.2.3.4

10.0.0.0/24 1.2.3.4

the next hop can be any character sequence, and not only an IP

###IpSubnetSortJoinArray
this fuction car sort and summarize subnets or ip routes
it must be used in an array formula

"ip_array" is a single column array containing ip subnets in "prefix/len"
or "prefix mask" notation

the return value is also an array of the same size containing the same
subnets sorted in ascending order
any consecutive subnets of the same size will be summarized when it is
possible
each line may contain any character sequence after the subnet, such as
a next hop or any parameter of an ip route
in this case, only subnets with the same parameters will be summarized

###IpDivideSubnet
divide a network in smaller subnets

"n" is the value that will be added to the subnet length

"SubnetSeqNbr" is the index of the smaller subnet to return

ex:

  IpDivideSubnet("1.2.3.0/24"; 2; 0) returns "1.2.3.0/26"
  
  IpDivideSubnet("1.2.3.0/24"; 2; 1) returns "1.2.3.64/26"
  

###IpIsPrivate
returns TRUE if "ip" is in one of the private IP address ranges

ex:

  IpIsPrivate("192.168.1.35") returns TRUE
  
  IpIsPrivate("209.85.148.104") returns FALSE
  

###IpDiff
difference between 2 IP addresses

ex:

  IpDiff("192.168.1.7"; "192.168.1.1") returns 6
  

###IpParse
Parses an IP address by iteration from right to left
Removes one byte from the right of "ip" and returns it as an integer

ex:

  if ip="192.168.1.32"
  
  IpParse(ip) returns 32 and ip="192.168.1" when the function returns

###IpBuild
Builds an IP address by iteration from right to left
Adds "ip_byte" to the left the "ip"

If "ip_byte" is greater than 255, only the lower 8 bits are added to "ip"
and the remaining bits are returned to be used on the next IpBuild call

ex 1:

  if ip="168.1.1"
  
  IpBuild(192, ip) returns 0 and ip="192.168.1.1"
  
ex 2:

  if ip="1"
  
  IpBuild(258, ip) returns 1 and ip="2.1"
  

###IpMaskBin
returns binary IP mask from an address with / notation (xx.xx.xx.xx/yy)

ex:

  IpMask("192.168.1.1/24") returns 4294967040 which is the binary
  representation of "255.255.255.0"
