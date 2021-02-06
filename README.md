# VBA-inject

FOLLOWING THE ANARCHY VIRUS

                by Constantin E. Climentieff aka DrMad

     INTRODUCTION
     1. DEVICE OF WINWORD 6.0 / 7.0 DOC FILES
     2. MACROS CODING
     3. FORMATION OF MACRO TITLE
     4. MACROS IMPLEMENTATION ALGORITHM
     LITERATURE
     ATTACHMENT. Demo program text

                                  ... the information provided in the article,
                                  obviously not enough for writing viruses;
                                  this is done on purpose ...

     INTRODUCTION

   So far multi-platform virus Anarchy.6093 by Jan Fakovski
remains somewhat unsurpassed. As far as is known, no
another virus is unable to parse the DOC file structure on its own
and embed executable macros into it.

   Note. There are also a couple of Win95 viruses of the same type.
and Win95.WoGob (WG), which can do the same. Perhaps this is
everything.

   For those who are not in the know, here is a snippet of the description of the Anarchy virus from
I. Danilova:

: Dangerous memory resident polymorphic stealth viruses. Viruses
: infect COM, EXE and DOC files for WinWord
: versions 6.0-7.0. When opening DOC files ... the virus analyzes
: internal format of OLE2 document, creates a new table
: macros and appends to the end of this document
: encrypted AUTOOPEN macro containing virus code
: dropper ... When infecting a DOC file, the virus uses
: some "holes" in OLE2 format, resulting in
: active virus macro AUTOOPEN check its presence
: in the system by means of WinWord [TOOLS / MACRO] not
: available possible. The virus contains some errors
: in the procedure for infecting OLE2 DOC files. Resulting in
: some files small DOC files can be destroyed
: during infection, and all DOC files for
: MSOffice'97. This virus is the first known
: a multi-platform virus that infects files of such
: various operating environments like DOS, Windows and MS Word for
: Windows ...

   I have this virus in my collection, but I have not "opened" it.
It is much more interesting to follow in the footsteps of Jan Fakovski and try
independently develop an algorithm for infecting DOC files.
     Several years ago, I had to deal with the internal
DOC file format for WinWord 6.0 / 7.0. I think I figured out how
similar viruses could be arranged. Now that dominated
version 97/2000, you can already reveal the "strrrrrh secret".

     1. DEVICE OF WINWORD 6.0 / 7.0 DOC FILES

   .DOC files and "MsWord documents" are not one and the same
. Files are complex OLE-2 objects organized by
rules of Structured Storage. Documents
Word, Excel spreadsheets and other objects are only stored inside
complexly organized Repositories.
     The format of Structured Repositories is not officially documented.
but hacker reconstructions are available, very similar to the truth
(although containing many ambiguities, inaccuracies and errors)
/ 2 /. It is at least clear that Structured Storage is very
resembles a FAT file system, only not organized on
disk, but inside a separate file: there are cluster blocks,
block description tables, directory catalogs (root and nested) and
etc.
     All files are organized in Structured Storage format.
modern versions of MS Office: both 6.0 / 7.0 and 97/2000. But it looks like
Winword 6.0 / 7.0 uses a simplified Storage format, which allows
understand it "with little blood" / 1 /. At least I was able to
it is very easy to find the actual Word document inside a DOC file,
disassemble its structure, detect embedded macro viruses,
decrypt them and delete / 3,4 /.
     The main distinguishing features of DOC files and documents in
Winword 6.0 / 7.0 format:
     - macros are written in WordBasic language (for versions 97/2000 - on
VBA, but they understand WordBasic too);
     - macros are an integral part of the document (for versions 97/2000
they are separate objects located somewhere inside the Storage, but outside
document).

     2. MACROS CODING

   It is very easy to determine how are encoded inside macros
syntactic elements of the WodrBasic language. Write and embed inside
document WinWord 6.0 / 7.0 any sufficiently long macro (for example,
consisting of only comments). Then using any editor
code (for example, using Hiew), find the Sub MAIN lines inside the file
and End Sub, and fill in all the space between them with hex codes:
01, 02, 03, 04, etc.

    Note. Codes starting with 80h are 2 bytes.
Therefore, you need to separately form a group of codes of the form: 00 80
01 80 02 80 ... There are more encoding options, but if you are not
are going to write your own "disassembler" for WordBasic,
You should have enough information about single-byte codes and double-byte types
80XXh.

    Download this document in Word (by the way, in any, you can even in
97/2000) and you will see a list of syntax corresponding to the codes
elements of the language. Here are a summary of this list.

 A. Some keywords of the language B. Some lexemes
 ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ~~~~~~~~~~~~~~~~~ ~~~
 01, 1b Sub 31 Select 44 Long 64 Line feed
 02 Not 1c Function 32 Is 45 Single 65 \ nstring
 03 And 1d If 33 Case 46 String 67 WordBasic Command
 04 Or 1e Then 34 As 47 Cdecl 68 Double precision
 05 (1f ElseIf 35 Redim 48 Alias 69 String
 06) 20 Else 36 Print 49 Any 6a "Quoted string"
 07 + 21 While 37 Input 4c Close 6b 'Comment line
 08 - 22 Wend 38 Line 4d Begin 6c Integer
 09/23 For 39 Write 4e Lib 6d Symbol
 0a * 24 To 3a Name 4f Read 6e N spaces
 0b Mod 25 Step 3b Output 54 EndIf 6f N tabs
 0c = 26 Next 3c Append 70 REM Comment line
 0d <> 28; 3d Open 73 Dotted token (+ 2 bytes)
 0e <29 Call 3f Dialog 76. Command key subfield
 0f> 2a Goto 40 Super
 10 <= 2c On 41 Declare B. Some commands
 11> = 2d Error 42 Double ~~~~~~~~~~~~~~~~~~~~
 18 Resume 2e Let 43 Integer 802b MsgBox
 19: 2f Dim 44 Long 80b7 MacroCopy
 1a End 30 Shared

     Let's try to code the following macro:

      Sub MAIN
      MsgBox "Hello!", "", 16
      End Sub

     Here's what we should get:

      0001 <Macro number>
      64 <New line>
      1B Sub
      69 <String without quotes>
      04 <Length of string without quotes> = 4
      4D M
      41 A
      49 I
      4E N
      64 <New line>
      67 <Command>
      802B MsgBox
      6A <Quoted string>
      06 <Length of quoted string> = 6
      48 H
      65 e
      6C l
      6C l
      6F o
      21!
      12 ,
      6A <Quoted string>
      00 <Length of quoted string> = 0
      12 ,
      6C <Integer constant>
      0010 <Integer constant value> = 16
      64 <New line>
      1A End
      1B Sub

     3. FORMATION OF MACRO TITLE

The format of the macro header is described in / 1,3 /, so we will not use it
reproduce. The filled-in macro header looks like this:

     SignFF db 0FFh; Header signature
     ; Macro title
     MHdr db 1; Macro header attribute
     NMcr dw 1; Number of macros
     ; Descriptor for the 1st (and only) macro
     Vers db 055h; Language version
     Key db 0; XOR key
     IntN dw 0
     ExtN dw 0
     XN dw 0FFFFh
             dd?
     MLen dd? ; Macro length (enter !!!)
     MStat dd 5; Macro status
     MPtr dd? ; Macro position (enter !!!)
     ; External name
     EHdr db 10h; External name attribute
     ESiz dw 0Dh; Name Description Size ???
     ELen db offset NRef-offset EName; Length
     EName db 'AutoOpen'
     NREf dw 1; Number of links to it
     ; Internal name
     IHdr db 11h; Internal name attribute
     INum dw 1; Number of names
     ICnt dw 0; Name number
     ILen db offset R-offset IName; Length
     IName db 'AUTOOPEN'
     R db?
     ;
     EndHdr db 40h; End of title

     4. MACROS IMPLEMENTATION ALGORITHM

   Step 0. Perform the necessary checks and make sure that this file is in
WinWord 6.0 / 7.0 format, not 97/2000.

     Step 1. Inside the DOC file contains the actual Word document,
starting with a characteristic signature. This signature is easy to find in
file at an offset multiple of 512 bytes (since large objects in
repositories are stored in Big Blocks / 2 /). The following can be found in the file
the same signatures lying at an offset not divisible by 200h are "left"
headers.

     Step 2. The Word document begins with a heading, in which:
     - at offset 0Ah is the word of flags, the least significant bit of it -
tag template;
     - at offset 118h the address is located (offset from the
document) macro header.
     - the correct macro header starts with the signature 1FFh.
     If the file is already template, or the macro header is full, then
the document probably already contains macros and shouldn't be touched.

     Step 3. If the address of the macrohead indicates "empty", then
Macro header is missing or empty. The macro header must
start in the file with an address multiple of 512 = 200h. Let's align the file to this
length and immediately write down the correctly filled macro header.

     Step 4. The actual macro code should be located somewhere after
macro header, and start at an offset multiple of 16. Therefore
align the new file length by 16 and write it to the end of the file right after
macro header macro code.

     Step 5. In the heading of the Word document at the address 0Ah, set the flag
template. At 118h, write the address (offset from the beginning of the document)
new macro header. Enter the address inside the new macro header
(offset from the beginning of the document) the macro code.

     Everything. "And ten pairs of knitting needles." J

     CONCLUSION

     I would not have dared to publish this algorithm if it were not for
there was one nuance.
     The fact is that if a file "infected" in this way is loaded into
Winword 6.0 / 7.0, then the macro will be normally accepted and executed. But if
load it, for example, in WinWord 97, then nothing will happen (although
it should be picked up, converted to VBA and executed).
Apparently, the hole is "darned", and WinWord 97 is taking apart its
Structured Storage according to all the rules (which are themselves fine-grained
and came up with, but did not execute for WinWord 6.0 / 7.0).
     Thus, for use inside viruses, this algorithm is now
not too relevant. Although extremely interesting in itself.
     Therefore - read and use.

     LITERATURE

     1. Schvartz M. ELSER. Word convertress, 1997.
     2. Schvartz M. LAOLA file system, 1997. (Translation: M.
File system Laola // Zemsky fershal, - # 4.- 2001)
     3. DrMad. A little bit of cryptology // Zemsky fershal .- # 2.- 1998
     4. DrMad. Writing an antivirus for WM.Cap // Zemsky Fershal.-
# 3.- 1999

     ATTACHMENT. Demo program text.
