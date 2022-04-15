# VBAIPFunctions
VBA functions for IP manipulation and IP/Subnet lookup

## Functions List

Unless otherwise specified, IP address arguments are given as strings in standard IPv4 d.d.d.d decimal notation, and subnet arguments 
are given in either CIDR notation, or a base address and mask pair, separated by a space.

### IpIsValid(_ip_)

Returns true if an ip address is formated exactly as it should be:
no space, no extra zero, no incorrect value

### IpStrToBin(_ip_)

Converts a text IP address to binary.

ex:

  IpStrToBin("1.2.3.4") returns 16909060

### IpBinToStr(_ipnum_)

Converts a binary IP address to text.

ex:

  IpBinToStr(16909060) returns "1.2.3.4"

### IpAdd(_ip_; _offset_)

ex:

  IpAdd("192.168.1.1"; 4) returns "192.168.1.5"
  
  IpAdd("192.168.1.1"; 256) returns "192.168.2.1"

### IpAnd(_ip-1_; _ip-2_)

IP logical AND

ex:

  IpAnd("192.168.1.1"; "255.255.255.0") returns "192.168.1.0"
  

### IpAdd2(_ip_; _offset_)

Another implementation of IpAdd which does not use the binary representation internally.

### IpGetByte(_ip_; _position_)

Get one byte from an ip address given its position, counted from the most significant byte starting as one.

ex:

  IpGetByte("192.168.1.1"; 1) returns 192

### IpSetByte(_ip_; _position_; _value_)

Set one byte in an ip address given its position, counted from the most significant byte as one, and new value.

ex:

  IpSetByte("192.168.1.1"; 4; 20) returns "192.168.1.20"

### IpMask(_subnet_)

Returns an IP netmask from a subnet.
Accepts either CIDR notation or address-and-mask.

ex:

  IpMask("192.168.1.1/24") returns "255.255.255.0"
  
  IpMask("192.168.1.1 255.255.255.0") returns "255.255.255.0"

### IpWildMask(_subnet_)

Returns an IP Wildcard (inverse) mask from a subnet.
Accepts either CIDR notation or address-and-mask.

ex:

  IpWildMask("192.168.1.1/24") returns "0.0.0.255"
  
  IpWildMask("192.168.1.1 255.255.255.0") returns "0.0.0.255"

### IpInvertMask(_ip_)

Inverts all bits of an IP address. Typically used to generate a wild card mask from a subnet mask, or vice versa. 

ex:

  IpWildMask("255.255.255.0") returns "0.0.0.255"
  
  IpWildMask("0.0.0.255") returns "255.255.255.0"

### IpMaskLen(_mask_)

Returns prefix length from a mask given by a string notation (xx.xx.xx.xx).

ex:

  IpMaskLen("255.255.255.0") returns 24 which is the number of bits of the subnetwork prefix

### IpWithoutMask(_subnet_)

Removes the mask notation from a subnet string, returning the base ip.

ex:

  IpWithoutMask("192.168.1.1/24") returns "192.168.1.1"
  
  IpWithoutMask("192.168.1.1 255.255.255.0") returns "192.168.1.1"
  
### IpSubnetLen(_subnet_)

Returns the mask len from a subnet string.

ex:
  IpSubnetLen("192.168.1.1/24") returns 24
  
  IpSubnetLen("192.168.1.1 255.255.255.0") returns 24

### IpSubnetSize(_subnet_)

Returns the number of addresses in a subnet.

ex:

  IpSubnetSize("192.168.1.32/29") returns 8
  
  IpSubnetSize("192.168.1.0 255.255.255.0") returns 256

### IpClearHostBits(_subnet_)

Returns a subnet string with with the host bits of the base address set to zero.

ex:

  IpClearHostBits("192.168.1.1/24") returns "192.168.1.0/24"
  
  IpClearHostBits("192.168.1.193 255.255.255.128") returns "192.168.1.128 255.255.255.128"

### IpIsInSubnet(_ip_; _subnet_)

Returns TRUE if the given IP is in the given subnet.

ex:

  IpIsInSubnet("192.168.1.35"; "192.168.1.32/29") returns TRUE
  
  IpIsInSubnet("192.168.1.35"; "192.168.1.32 255.255.255.248") returns TRUE
  
  IpIsInSubnet("192.168.1.41"; "192.168.1.32/29") returns FALSE
  

### IpSubnetVLookup(_ip_; _range_; _index_)

Tries to match an IP address against a list of subnets in the left-most
column of _range_ and returns the value in the same row based on _index_.

this function selects the smallest matching subnet

"ip" is the value to search for in the subnets in the first column of
     the table_array
     
"range" is one or more columns of data

"index" is the column number in table_array from which the matching
     value must be returned. The first column which contains subnets is 1.
     
note: Add the subnet 0.0.0.0/0 at the end of the array if you want the
function to return a default value.

### IpSubnetMatch(_ip_; _range_)

Tries to match an IP address against a list of subnets in the left-most
column of _range_ and returns the row number.
This function selects the smallest matching subnet.

"ip" is the value to search for in the subnets in the first column of
     the table_array
     
"table_array" is one or more columns of data

returns 0 if the IP address is not matched.

### IpSubnetIsInSubnet(_subnet-1_; _subnet-2_)

Returns TRUE if "subnet1" is in "subnet2".
Subnets must have the / mask notation (xx.xx.xx.xx/yy)

ex:

  IpSubnetIsInSubnet("192.168.1.35/30"; "192.168.1.32/29") returns TRUE
  
  IpSubnetIsInSubnet("192.168.1.41/30"; "192.168.1.32/29") returns FALSE
  
  IpSubnetIsInSubnet("192.168.1.35/28"; "192.168.1.32/29") returns FALSE
  

### IpSubnetInSubnetVLookup(_subnet_; _range_; _index_)

Tries to match a subnet against a list of subnets in the left-most
column of _range_ and returns the value in the same row based on _index_.
The value matches if _subnet_ is equal to or included in one of the subnets
in the array.

"subnet" is the value to search for in the subnets in the first column of
     the table_array
     
"table_array" is one or more columns of data

"index_number" is the column number in table_array from which the matching
     value must be returned. The first column which contains subnets is 1.
     
note: Add the subnet 0.0.0.0/0 at the end of the array if you want the
function to return a default value.

### IpSubnetInSubnetMatch(_subnet_; _range_)

Tries to match a subnet against a list of subnets in the left-most
column of table_array and returns the row number
the value matches if _subnet_ is equal to or included in one of the subnets
in the array.

"subnet" is the value to search for in the subnets in the first column of
     the table_array
     
"table_array" is one or more columns of data

returns 0 if the subnet is not included in any of the subnets from the list

### IpFindOverlappingSubnets(_subnets_)

This function must be used in an array formula.
It will find which subnets overlap in the given list of subnets.

_subnets_ is single column array containing a list of subnets, the
list may be sorted or not.
The return value is also an array of the same size.
If the subnet on line x is included in a larger subnet from another line,
this function returns an array in which line x contains the value of the
larger subnet.
If the subnet on line x is distinct from any other subnet in the array,
then this function returns an empty cell on line x.
If there are no overlapping subnets in the input array, the returned array
is empty.

### IpSortArray(_ips_[; _descending_])

This function must be used in an array formula.

_ips_ is a single column array containing ip addresses.
The return value is also a array of the same size containing the same
addresses sorted in ascending or descending order.

_descending_ is an optional parameter, if set to True the addresses are
sorted in descending order

### IpSubnetSortArray(_ips_[; _descending_])

This function must be used in an array formula.

_ips_ is a single column array containing ip subnets in "prefix/len"
or "prefix mask" notation.
The return value is also an array of the same size containing the same
subnets sorted in ascending or descending order.

_descending_ is an optional parameter, if set to True the subnets are
sorted in descending order.

### IpParseRoute - _internal utility subroutine_

This function is used by IpSubnetSortJoinArray to extract the subnet
and next hop in route.

The supported formats are:

  10.0.0.0 255.255.255.0 1.2.3.4

  10.0.0.0/24 1.2.3.4

the next hop can be any character sequence, and not only an IP

### IpSubnetSortJoinArray(_ips_)

This fuction can sort and summarize subnets or ip routes.
It must be used in an array formula.

_ips_ is a single column array containing ip subnets in "prefix/len"
or "prefix mask" notation

The return value is also an array of the same size containing the same
subnets sorted in ascending order.
Any consecutive subnets of the same size will be summarized when possible.
Each line may contain any character sequence after the subnet, such as
a next hop or any parameter of an ip route.
In this case, only subnets with the same parameters will be summarized.

### IpDivideSubnet(_subnet_; _n_; _index_)

Divide a network in smaller subnets.

_n_ is the value that will be added to the subnet length

_index_ is the index of the smaller subnet to return

ex:

  IpDivideSubnet("1.2.3.0/24"; 2; 0) returns "1.2.3.0/26"
  
  IpDivideSubnet("1.2.3.0/24"; 2; 1) returns "1.2.3.64/26"
  

### IpIsPrivate(_ip_)

Returns TRUE if _ip_ is in one of the private IP address ranges.

ex:

  IpIsPrivate("192.168.1.35") returns TRUE
  
  IpIsPrivate("209.85.148.104") returns FALSE
  

### IpDiff(_ip-1_; _ip-2_)

Difference between 2 IP addresses.

ex:

  IpDiff("192.168.1.7"; "192.168.1.1") returns 6
  

### IpParse(_ip-by-ref_)

Parses an IP address by iteration from right to left.
Removes one byte from the right of _ip-by-ref_ and returns it as an integer.

ex:

  if ip="192.168.1.32"
  
  IpParse(ip) returns 32 and ip="192.168.1" when the function returns

### IpBuild(_byte_; _ip-by-ref_)

Builds an IP address by iteration from right to left.
Adds _byte_ to the left of _ip_.

If "ip_byte" is greater than 255, only the lower 8 bits are added to "ip"
and the remaining bits are returned to be used on the next IpBuild call

ex 1:

  if ip="168.1.1"
  
  IpBuild(192, ip) returns 0 and ip="192.168.1.1"
  
ex 2:

  if ip="1"
  
  IpBuild(258, ip) returns 1 and ip="2.1"
  

### IpMaskBin(_subnet_)

Returns binary IP mask from an address with / notation (xx.xx.xx.xx/yy).

ex:

  IpMask("192.168.1.1/24") returns 4294967040 which is the binary
  representation of "255.255.255.0"
