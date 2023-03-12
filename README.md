# Excel-UTF-8
Excel forumlae for encoding/decoding UTF-8 codepoints

Here are 2 formulae to aid in working with UTF-8.  1 converts a code point into a UTF-8 hex string and the other converts a UTF-8 hex string into a code point.

I often find myself using a spreadsheet to model a computing concept before I commit to coding the final solution.  Spreadsheets allow me to break down large problems into smaller steps that can be stored in cells.

When I found myself needing to perform a lot of conversions between code points, UTF-8 and then back again, I wrote these formulae to expediate the process.

I hope that they can be of use to someone.


Code Point to UTF-8

Eg: 119075 => 0xF0 0x9D 0x84 0xA3

=IF(OR(ISNA(A1),ISERR(A1)),"Invalid",IF(UNICODE(A1)<128,"0x"&DEC2HEX(UNICODE(A1),2),IF(UNICODE(A1)<2048,"0x"&DEC2HEX(BITAND(BITRSHIFT(UNICODE(A1),6),63)+192,2)&" 0x"&DEC2HEX(BITAND(UNICODE(A1),63)+128,2),IF(UNICODE(A1)<65536,"0x"&DEC2HEX(BITAND(BITRSHIFT(UNICODE(A1),12),63)+224,2)&" 0x"&DEC2HEX(BITAND(BITRSHIFT(UNICODE(A1),6),63)+128,2)&" 0x"&DEC2HEX(BITAND(UNICODE(A1),63)+128,2),IF(AND(UNICODE(A1)<1114111,NOT(ISNA(A1)),NOT(ISERR(A1))),"0x"&DEC2HEX(BITAND(BITRSHIFT(UNICODE(A1),18),63)+240,2)&" 0x"&DEC2HEX(BITAND(BITRSHIFT(UNICODE(A1),12),63)+128,2)&" 0x"&DEC2HEX(BITAND(BITRSHIFT(UNICODE(A1),6),63)+128,2)&" 0x"&DEC2HEX(BITAND(UNICODE(A1),63)+128,2),"Invalid")))))


UTF-8 to Code Point

Eg: 0xF0 0x9D 0x84 0xA3 => 119075


=CHOOSE((LEN(A1)+1)/5,BITAND(HEX2DEC(MID(A1,3,2)),127),BITLSHIFT(BITAND(HEX2DEC(MID(A1,3,2)),63),6)+BITAND(HEX2DEC(MID(A1,8,2)),127),BITLSHIFT(BITAND(HEX2DEC(MID(A1,3,2)),31),12)+BITLSHIFT(BITAND(HEX2DEC(MID(A1,8,2)),63),6)+BITAND(HEX2DEC(MID(A1,13,2)),127),BITLSHIFT(BITAND(HEX2DEC(MID(A1,3,2)),15),18)+BITLSHIFT(BITAND(HEX2DEC(MID(A1,8,2)),127),12)+BITLSHIFT(BITAND(HEX2DEC(MID(A1,13,2)),127),6)+BITAND(HEX2DEC(MID(A1,18,2)),127))
