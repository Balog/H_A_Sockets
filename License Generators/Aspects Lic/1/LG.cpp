//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "LG.h"
#include "Math.hpp"
#include "CodeText.h"
#include <stdio.h> 

//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TForm1 *Form1;
//---------------------------------------------------------------------------
__fastcall TForm1::TForm1(TComponent* Owner)
        : TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TForm1::Button1Click(TObject *Sender)
{
String Text;
String Tex;
int N, N1, N2;
String Itog;
int K;
String S1, S2;
if(CheckBox1->Checked)
{
N=StrToInt(Edit1->Text);
N1=N/10;
N2=N-N1*10;
S1=Propis(N1);
S2=Propis(N2);
String S=S1+"+"+S2+"+";
K=256-S.Length()*2;
}
else
{
S1="�������������";
S2="������";
String S=S1+"+"+S2+"+";
K=256-S.Length()*2;
/*
}
else
{
Itog="�������������+������+"+Tex+"+�������������+������";
}
*/
}
Tex="";
randomize();
String STR="1234567890-=!�;%:?*()_+/\���������������������������������. �������������������������������ިqwertyuiop[]asdfghjkl'\zxcvbnm/QWERTYUIOP{}ASDFGHJKL|ZXCVBNM<>?";
for(;Tex.Length()<K;)
{

int RND=random(STR.Length()+1);
Tex=STR.SubString(RND,1)+Tex;
}

Itog=S1+"+"+S2+"+"+Tex+"+"+S1+"+"+S2;

CodeText *CT=new CodeText();
String CodeSTR=CT->Crypt(Itog,"�������1");
delete CT;
FILE *F;
String File=ExtractFilePath(Application->ExeName)+"Aspects.lic";
if ((F = fopen(File.c_str(), "wt")) == NULL)
{
 ShowMessage("���� �� ������� �������");
 return;
}
else
{
if (fputs(CodeSTR.c_str(), F) < 0)
 ShowMessage("������ ������ � ����");

}
fclose(F);
ShowMessage("���� �������� Aspects.lic ������ \r� ��������� � ����� � ����������.");
}
//---------------------------------------------------------------------------
String TForm1::Propis(int N)
{
String Res;
switch(N)
{
 case 0:
 {
  Res="����";
  break;
 }
 case 1:
 {
  Res="����";
  break;
 }
 case 2:
 {
  Res="���";
  break;
 }
 case 3:
 {
  Res="���";
  break;
 }
 case 4:
 {
  Res="������";
  break;
 }
 case 5:
 {
  Res="����";
  break;
 }
 case 6:
 {
  Res="�����";
  break;
 }
 case 7:
 {
  Res="����";
  break;
 }
 case 8:
 {
  Res="������";
  break;
 }
 case 9:
 {
  Res="������";
  break;
 }

}
return Res;
}
//---------------------------------------------------------------------------
void __fastcall TForm1::CheckBox1Click(TObject *Sender)
{
Label1->Enabled=CheckBox1->Checked;
Edit1->Enabled=CheckBox1->Checked;
UpDown1->Enabled=CheckBox1->Checked;
}
//---------------------------------------------------------------------------
