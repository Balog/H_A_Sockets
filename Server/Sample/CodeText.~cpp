//---------------------------------------------------------------------------


#pragma hdrstop

#include "CodeText.h"
#include "stdlib.h"
#include "System.hpp"
#include "Math.hpp"


//---------------------------------------------------------------------------

#pragma package(smart_init)
CodeText::CodeText()
{

}
//-----------------------------------
String CodeText::Crypt(String Tex,String P)
{
String Ps;
if(Tex!="")
{
String Text1=Tex;
CodeText="";
if(P=="")
{
P=" ";
}
Ps=CreatePassword(P);

String Sect;
if(Tex.Length()<256)
{
randomize();
String STR="1234567890-=!¹;%:?*()_+/\יצףךוםדרשחץתפûגאןנמכהז‎ÿקסלטעüב‏¸. ÉÖÓÊÅÍÃØÙÇÕÚÔÛÂÀÏÐÎËÄÆÝ‗×ÑÌÈÒÜÁÞ¨qwertyuiop[]asdfghjkl'\zxcvbnm/QWERTYUIOP{}ASDFGHJKL|ZXCVBNM<>?";
for(;Tex.Length()<256;)
{

int RND=random(STR.Length()+1);
Tex=STR.SubString(RND,1)+Tex;
}

}

for(int i=0;Tex.Length()>0;i++)
{

if(Tex.Length()>=256)
{
Sect=Tex.SubString(0,256);
}
else
{
Sect=Tex+Text1.SubString(1,256-Tex.Length());
}
Tex=Tex.SubString(Sect.Length()+1,Tex.Length());
String CombSect=CombineSection(Sect, Ps);
Ps=CreatePassword(Ps+CombSect+Ps);
CodeText=CodeText+CombSect;
}
}
else
{
CodeText="";
}
return CodeText;
}
//-----------------------------------
String CodeText::EnCrypt(String Text,String Password, int Dlina)
{
String Ps;
int D=Dlina;
if(Text!="" & Dlina!=0)
{
String Text1=Text;
DecodeText="";
if(Password=="")
{
Password=" ";
}
Ps=CreatePassword(Password);

String Sect;
for(int i=0;Text.Length()>0;i++)
{
Sect=Text.SubString(0,256);
Text=Text.SubString(Sect.Length()+1,Text.Length());
String CombSect=DeCombineSection(Sect, Ps);
Ps=CreatePassword(Ps+Sect+Ps);
if(D<256)
{
DecodeText=DecodeText+CombSect.SubString(1,D);
}
else
{
DecodeText=DecodeText+CombSect;
}
D=D-256;
}
}
else
{
DecodeText="";
}
return DecodeText;
}
//-----------------------------------
String CodeText::ReCrypt(String Text,String OldPassword, String NewPassword, int Dlina)
{

String S=EnCrypt(Text,OldPassword,Dlina);
return Crypt(S,NewPassword);

}
//------------------------------------
String CodeText::CreatePassword(String ManPass)
{

int R=0;
for(int i=1;i<=ManPass.Length();i++)
{
String C1=ManPass.SubString(i,1);
char *C=C1.c_str();
byte N=*C;
srand(N+R);
R=rand()%256;
}
String Pass="";
for(int i=1;i<=16;i++)
{
int R=rand()%256;
Pass=Pass+(char)R;
}
return Pass;
}
//----------------------------------------------
String CodeText::CombineSection(String Section, String Pass)
{
//Ñהגטד ןמ סעמכבצאל
for(int i=0;i<16;i++)//Íמלונ סעמכבצא
{
char *V=new char[16];
String PS=Pass.SubString(i+1,1);
char *C=PS.c_str();
int N=*C;
if(N<0)
{
N=N+256;
}
int S=N/16;
for(int j=0;j<16;j++)
{
//A[1]=(char)S[5];
int Shift=S+j;
if(Shift>=16)
{
Shift=Shift-16;
}
V[Shift]=(char)Section[j*16+i+1];
}

for(int k=0;k<16;k++)
{
String R=V[k];
Section.Delete(k*16+i+1,1);
Section.Insert(R,k*16+i+1);
}
delete V;
}

//Ñהגטד ןמ סענמךאל

for(int i=0;i<16;i++) //םמלונ סעמכבצא
{
char *V=new char[16];
String PS=Pass.SubString(i+1,1);
char *C=PS.c_str();
int N=*C;
if(N<0)
{
N=N+256;
}
int S1=N/16;
int S=N-S1*16;

for(int j=0;j<16;j++)
{
int Shift=S+j;
if(Shift>=16)
{
Shift=Shift-16;
}
V[Shift]=(char)Section[i*16+j+1];
}

for(int k=0;k<16;k++)
{
String R=V[k];
Section.Delete(i*16+k+1,1);
Section.Insert(R,i*16+k+1);
}


delete V;
}
return Section;
}
//---------------------------------------------
String CodeText::DeCombineSection(String Section, String Pass)
{
//Ñהגטד ןמ סענמךאל

for(int i=0;i<16;i++) //םמלונ סעמכבצא
{
char *V=new char[16];
String PS=Pass.SubString(i+1,1);
char *C=PS.c_str();
int N=*C;
if(N<0)
{
N=N+256;
}
int S1=N/16;
int S=N-S1*16;

for(int j=0;j<16;j++)
{

V[j]=(char)Section[i*16+j+1];
}

for(int k=0;k<16;k++)
{
int Shift=S+k;
if(Shift>=16)
{
Shift=Shift-16;
}
String R=V[Shift];
Section.Delete(i*16+k+1,1);
Section.Insert(R,i*16+k+1);
}


delete V;
}

//Ñהגטד ןמ סעמכבצאל
for(int i=0;i<16;i++)//Íמלונ סעמכבצא
{
char *V=new char[16];
String PS=Pass.SubString(i+1,1);
char *C=PS.c_str();
int N=*C;
if(N<0)
{
N=N+256;
}
int S=N/16;
for(int j=0;j<16;j++)
{
//A[1]=(char)S[5];

V[j]=(char)Section[j*16+i+1];
}

for(int k=0;k<16;k++)
{
int Shift=S+k;
if(Shift>=16)
{
Shift=Shift-16;
}
String R=V[Shift];
Section.Delete(k*16+i+1,1);
Section.Insert(R,k*16+i+1);
}
delete V;
}

return Section;
}

