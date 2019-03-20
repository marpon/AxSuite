#include once "windows.bi"

#Include Once "strwrap.bi"


function trait(st2 as string, ar() as string) as integer
   dim as integer x,n0
   dim as string s1, s2,st1
	st1=rtrim (st2,any "' ")
	'st1=left(st1,len(st1)-1)
	st1=Str_removeany(st1, "0123456789=-")
	st1=str_remove(st1, "Type(,,,,)")

   dim i1 as integer = str_numparse(st1, ",")
   'print "st1",st1
   'print "i1",i1
   redim ar(1 to i1, 1 to 2) As string
   x=0
   for n0 = 1 to i1
      s1 = str_parse(st1, ",", n0)
		s1=rtrim(s1,any "') ")
		if s1 <> "" and instr(s1, " As ") > 0 THEN
			x += 1
			s2 = trim(str_parse(ucase(s1), " AS ", 1))
			if s2 = "" THEN s2 = "ANY_VAR"
			ar(x, 2) = s2
			select CASE trim(str_parse(ucase(s1), " AS ", 2))
				CASE "LONG", "INTEGER", "INTEGER PTR", "LONG PTR"
					ar(x, 1) = "%d"
				CASE "ULONG", "UINTEGER", "ULONG PTR", "UINTEGER PTR"
					ar(x, 1) = "%u"
				CASE "SHORT", "SHORT PTR"
					ar(x, 1) = "%i"
				CASE "USHORT", "USHORT PTR"
					ar(x, 1) = "%I"
				CASE "BSTR", "BSTR PTR"
					ar(x, 1) = "%B"
				CASE "ZSTRING", "ZSTRING PTR"
					ar(x, 1) = "%s"
				CASE "WSTRING", "WSTRING PTR"
					ar(x, 1) = "%S"
				CASE "DOUBLE", "DOUBLE PTR" ,"FLOAT","FLOAT PTR","SINGLE","SINGLE PTR"
					ar(x, 1) = "%e"
				CASE "BOOL", "BOOL PTR"
					ar(x, 1) = "%b"
				CASE "VARIANT", "VARIANT PTR"
					ar(x, 1) = "%v"
				CASE "LPDISPATCH", "LPDISPATCH PTR", "IDISPATCH", "IDISPATCH PTR"
					ar(x, 1) = "%o"
				CASE "LPUNKNOWN", "LPUNKNOWN PTR", "IUNKNOWN PTR", "IUNKNOWN"
					ar(x, 1) = "%0"
	'         CASE "DATE_", "DATE_ PTR"
	'            ar(x, 1) = "%t"
				CASE "DATE", "DATE PTR"
					ar(x, 1) = "%D"
				CASE "FBSTRING", "FBSTRING PTR"
					ar(x, 1) = "%A"
				CASE ELSE
					ar(x, 1) = "%p"
			END SELECT
		end if
   NEXT

   'print "fin",i1
   function = x
END function


dim ar() as string
dim st2 as string ,scut2 as string
dim as integer x, i1

st2="blbla (Options1 As VARIANT=Type(0,0,0,0,0),ProfileName As BSTR=0) '"


scut2=mid(st2,len(str_parse(st2, "(", 1))+2)
print "scut2 >" & scut2 & "<"
i1=trait(scut2 , ar())
print "i1 ", i1

for x = 1 to i1
	print "x "; x,"ar() 1 ";ar(x,1),"ar() 2 ";ar(x,2)
NEXT
print

st2="blibli (INIFilename As BSTR,Options1 As VARIANT=Type(0,0,0,0,0))"
scut2=mid(st2,len(str_parse(st2, "(", 1))+2)
print "scut2 >" & scut2 & "<"
i1=trait(scut2 , ar())
print "i1 ", i1

for x = 1 to i1
	print "x "; x,"ar() 1 ";ar(x,1),"ar() 2 ";ar(x,2)
NEXT
print
sleep