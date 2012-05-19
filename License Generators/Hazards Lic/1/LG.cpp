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
S1="Неограниченая";
S2="версия";
String S=S1+"+"+S2+"+";
K=256-S.Length()*2;
/*
}
else
{
Itog="Неограниченая+версия+"+Tex+"+Неограниченая+версия";
}
*/
}
Tex="";
randomize();
String STR="1234567890-=!№;%:?*()_+/\йцукенгшщзхъфывапролджэячсмитьбюё. ЙЦУКЕНГШЩЗХЪФЫВАПРОЛДЖЭЯЧСМИТЬБЮЁqwertyuiop[]asdfghjkl'\zxcvbnm/QWERTYUIOP{}ASDFGHJKL|ZXCVBNM<>?";
for(;Tex.Length()<K;)
{

int RND=random(STR.Length()+1);
Tex=STR.SubString(RND,1)+Tex;
}

Itog=S1+"+"+S2+"+"+Tex+"+"+S1+"+"+S2;

CodeText *CT=new CodeText();
String CodeSTR=CT->Crypt(Itog,"Опасности1");
delete CT;
FILE *F;
String File=ExtractFilePath(Application->ExeName)+"Hazards.lic";
if ((F = fopen(File.c_str(), "wt")) == NULL)
{
 ShowMessage("Файл не удается создать");
 return;
}
else
{
if (fputs(CodeSTR.c_str(), F) < 0)
 ShowMessage("Ошибка записи в файл");

}
fclose(F);
ShowMessage("Файл лицензии Hazards.lic создан \rи находится в папке с программой.");
}
//---------------------------------------------------------------------------
String TForm1::Propis(int N)
{
String Res;
switch(N)
{
 case 0:
 {
  Res="Нуль";
  break;
 }
 case 1:
 {
  Res="Один";
  break;
 }
 case 2:
 {
  Res="Два";
  break;
 }
 case 3:
 {
  Res="Три";
  break;
 }
 case 4:
 {
  Res="Четыре";
  break;
 }
 case 5:
 {
  Res="Пять";
  break;
 }
 case 6:
 {
  Res="Шесть";
  break;
 }
 case 7:
 {
  Res="Семь";
  break;
 }
 case 8:
 {
  Res="Восемь";
  break;
 }
 case 9:
 {
  Res="Девять";
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

