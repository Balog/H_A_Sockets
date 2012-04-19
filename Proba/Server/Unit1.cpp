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
Memo1->Lines->Clear();
Memo1->Lines->Add("Сервер создан");
ServerSocket1->Active=true;
}
//---------------------------------------------------------------------------

void __fastcall TForm1::ServerSocket1ClientConnect(TObject *Sender,
      TCustomWinSocket *Socket)
{
Memo1->Lines->Add("Клиент подключен");
WS=Socket;
}
//---------------------------------------------------------------------------
void __fastcall TForm1::ServerSocket1ClientDisconnect(TObject *Sender,
      TCustomWinSocket *Socket)
{
Memo1->Lines->Add("Клиент отключен");
}
//---------------------------------------------------------------------------
void __fastcall TForm1::Button1Click(TObject *Sender)
{

 Comm1->Execute();
 Button1->Enabled=false;
}
//---------------------------------------------------------------------------
void __fastcall TForm1::Button2Click(TObject *Sender)
{
Comm2->Execute();
 Button2->Enabled=false;
}
//---------------------------------------------------------------------------
void __fastcall TForm1::ServerSocket1ClientRead(TObject *Sender,
      TCustomWinSocket *Socket)
{
String RText=Socket->ReceiveText();
if(RText.Length()>0)
{
int Comm=StrToInt(RText);
switch (Comm)
{
 case 1:
 {
 Memo1->Lines->Add("Принята команда 1");
 Button1->Enabled=true;
 break;
 }
 case 2:
 {
 Memo1->Lines->Add("Принята команда 2");
 Button2->Enabled=true;
 break;
 }
 case 3:
 {
  Memo1->Lines->Add("Принята команда 3");
  Comm3->Execute();
  break;
 }
}
}
}
//---------------------------------------------------------------------------
void __fastcall TForm1::Comm1Execute(TObject *Sender)
{
WS->SendText("Comm:1;"+Edit1->Text);
 Memo1->Lines->Add("Отправлено: "+Edit1->Text);
}
//---------------------------------------------------------------------------
void __fastcall TForm1::Comm2Execute(TObject *Sender)
{
WS->SendText("Comm:2;"+Edit2->Text);
 Memo1->Lines->Add("Отправлено: "+Edit2->Text);
}
//---------------------------------------------------------------------------
void __fastcall TForm1::Comm3Execute(TObject *Sender)
{
WS->SendText("Comm:3;"+Now().CurrentDateTime());
 Memo1->Lines->Add("Отправлено: "+Now().CurrentDateTime());
}
//---------------------------------------------------------------------------
