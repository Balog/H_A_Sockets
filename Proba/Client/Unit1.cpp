//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "Unit1.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TForm1 *Form1;
//---------------------------------------------------------------------------
__fastcall TForm1::TForm1(TComponent* Owner)
        : TForm(Owner)
{
ClientSocket1->Active=true;

}
//---------------------------------------------------------------------------

void __fastcall TForm1::Button1Click(TObject *Sender)
{
Memo1->Lines->Clear();
Param.clear();
Parameters P;
P.SParam.push_back("Вот первое значение:");
P.ActTrue="Command1";
P.RespComm=1;
Param.push_back(P);
//RunAction("Command1");
ClientSocket1->Socket->SendText(P.RespComm);


}
//---------------------------------------------------------------------------
void __fastcall TForm1::ClientSocket1Read(TObject *Sender,
      TCustomWinSocket *Socket)
{
String Mess=Socket->ReceiveText();
int N0=Mess.Pos(":");
int N=Mess.Pos(";");

int Comm;
String Nick;
String SS=Mess.SubString(N0+1,N-N0-1);
Comm=StrToInt(SS);

String S=Mess.SubString(N+1,Mess.Length());

if(Comm==Param[0].RespComm)
{
Param[0].SParam.push_back(S);
RunAction(Param[0].ActTrue);
}
}
//---------------------------------------------------------------------------

//---------------------------------------------------------------------------
void __fastcall TForm1::Command1Execute(TObject *Sender)
{
Parameters P;
P=Param[0];
if(P.ActTrue=="Command1")
{

Memo1->Lines->Add(P.SParam[0]);
Memo1->Lines->Add(P.SParam[1]);
}
Param[0].SParam.clear();
Param.clear();
Parameters P1;
P1.SParam.push_back("Вот второе значение:");
P1.ActTrue="Command2";
P1.RespComm=2;
Param.push_back(P1);
ClientSocket1->Socket->SendText(P1.RespComm);
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Command2Execute(TObject *Sender)
{
Parameters P;
P=Param[0];
if(P.ActTrue=="Command2")
{

Memo1->Lines->Add(P.SParam[0]);
Memo1->Lines->Add(P.SParam[1]);

P.SParam.clear();
Param.clear();

Parameters P1;
P1.SParam.push_back("Вот текущее время:");
P1.ActTrue="Command3";
P1.RespComm=3;
Param.push_back(P1);
ClientSocket1->Socket->SendText(P1.RespComm);
}

}
//---------------------------------------------------------------------------
void TForm1::RunAction(String NameAct)
{
for(int i=0;i<ActionManager1->ActionCount;i++)
{
if(ActionManager1->Actions[i]->Name==NameAct)
{
 ActionManager1->Actions[i]->Execute();
}
}
}
//---------------------------------------------------------------------------


void __fastcall TForm1::Command3Execute(TObject *Sender)
{
Parameters P;
P=Param[0];
if(P.ActTrue=="Command3")
{

Memo1->Lines->Add(P.SParam[0]);
Memo1->Lines->Add(P.SParam[1]);
}
}
//---------------------------------------------------------------------------

