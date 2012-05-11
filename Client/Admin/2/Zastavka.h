//---------------------------------------------------------------------------

#ifndef ZastavkaH
#define ZastavkaH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <ExtCtrls.hpp>
#include <Graphics.hpp>
//#include "F_Vvedenie.h"
//#include "MainForm.h"
#include <ADODB.hpp>
#include <DB.hpp>
//#include "Form_SendFile.h"
#include <Dialogs.hpp>
#include "inifiles.hpp"
#include <jpeg.hpp>
#include <ActnList.hpp>
#include <ActnMan.hpp>
#include <ScktComp.hpp>;
#include "MDBConnector.h"
#include "ClientClass.h"
#include <vector>
using namespace std;
//---------------------------------------------------------------------------
/*
struct CDBItem
{
int Num;
String Name;
String ServerDB;
};
*/
class TZast : public TForm
{
__published:	// IDE-managed Components
        TTimer *Timer1;
        TStaticText *StaticText1;
        TStaticText *StaticText2;
        TStaticText *StaticText3;
        TStaticText *StaticText4;
        TStaticText *StaticText5;
        TStaticText *StaticText6;
        TStaticText *StaticText7;
        TStaticText *StaticText8;
        TStaticText *StaticText9;
        TStaticText *StaticText10;
        TStaticText *StaticText11;
        TStaticText *R1;
        TStaticText *R2;
        TStaticText *R3;
        TStaticText *R4;
        TStaticText *R5;
        TStaticText *R6;
        TStaticText *R7;
        TStaticText *R8;
        TStaticText *R9;
        TStaticText *R10;
        TImage *Image1;
        TStaticText *StaticText8_1;
        TStaticText *T1;
        TStaticText *T2;
        TStaticText *T3;
        TStaticText *T4;
        TStaticText *T5;
        TStaticText *T6;
        TStaticText *T7;
        TStaticText *T8;
        TStaticText *T9;
        TStaticText *T10;
        TStaticText *Y1;
        TStaticText *Y2;
        TStaticText *Y3;
        TStaticText *Y4;
        TStaticText *Y5;
        TStaticText *Y6;
        TStaticText *Y7;
        TStaticText *Y8;
        TStaticText *Y9;
        TStaticText *Y10;
        TStaticText *Y11;
        TStaticText *Y12;
        TStaticText *Y13;
        TStaticText *W1;
        TStaticText *W2;
        TStaticText *W3;
        TStaticText *W4;
        TStaticText *W5;
        TStaticText *W6;
        TStaticText *U1;
        TStaticText *U2;
        TStaticText *U3;
        TStaticText *U4;
        TStaticText *U5;
        TStaticText *U6;
        TStaticText *U7;
        TStaticText *U8;
        TStaticText *U9;
        TStaticText *U10;
        TStaticText *U11;
        TStaticText *U12;
        TStaticText *U13;
        TStaticText *U14;
        TTimer *Timer2;
        TClientSocket *ClientSocket;
        TActionManager *ActionManager1;
        TAction *RegForm_Form1;
        TAction *RegForm_Form2;
        TAction *PostRegForm_Form1;
        TAction *PostRegForm_Form2;
        TAction *PrepareConnectBase;
        TAction *ConnectBase;
        TAction *LoadLogins;
        TAction *ViewLogins;
        TAction *UpdateOtdels;
        TAction *Trigger;
        TAction *UpdateLogins;
        TAction *UpdateOtdelsMan;
        TAction *UpdateLoginsMan;
        TAction *UpdateObslOtdelMan;
        TAction *SaveLogins;
        TAction *SaveObslOtd;
        TAction *SendMergeSave;
        TAction *PrepareSaveLogins;
        TAction *PostSaveLogins;
        TAction *SaveTempPodr;
        TAction *LoadNewLogins;
        TAction *CorrectNewLogins;
        TAction *ReadOtdels;
        TAction *PrepareRead;
        TAction *PostRead;
        TAction *PrepareUpdateOtd;
        TAction *PostUpdateOtd;
        TAction *ReadOtdelsAuto;
        void __fastcall Timer1Timer(TObject *Sender);
        void __fastcall Image1Click(TObject *Sender);
        void __fastcall FormCreate(TObject *Sender);
        void __fastcall FormShow(TObject *Sender);
        void __fastcall Timer2Timer(TObject *Sender);
        void __fastcall FormDestroy(TObject *Sender);
        void __fastcall ClientSocketRead(TObject *Sender,
          TCustomWinSocket *Socket);
        void __fastcall RegForm_Form1Execute(TObject *Sender);
        void __fastcall RegForm_Form2Execute(TObject *Sender);
        void __fastcall PostRegForm_Form1Execute(TObject *Sender);
        void __fastcall PostRegForm_Form2Execute(TObject *Sender);
        void __fastcall PrepareConnectBaseExecute(TObject *Sender);
        void __fastcall ConnectBaseExecute(TObject *Sender);
        void __fastcall LoadLoginsExecute(TObject *Sender);
        void __fastcall ViewLoginsExecute(TObject *Sender);
        void __fastcall ClientSocketError(TObject *Sender,
          TCustomWinSocket *Socket, TErrorEvent ErrorEvent,
          int &ErrorCode);
        void __fastcall UpdateOtdelsExecute(TObject *Sender);
        void __fastcall TriggerExecute(TObject *Sender);
        void __fastcall UpdateLoginsExecute(TObject *Sender);
        void __fastcall UpdateOtdelsManExecute(TObject *Sender);
        void __fastcall UpdateLoginsManExecute(TObject *Sender);
        void __fastcall UpdateObslOtdelManExecute(TObject *Sender);
        void __fastcall SaveLoginsExecute(TObject *Sender);
        void __fastcall SaveObslOtdExecute(TObject *Sender);
        void __fastcall SendMergeSaveExecute(TObject *Sender);
        void __fastcall PrepareSaveLoginsExecute(TObject *Sender);
        void __fastcall PostSaveLoginsExecute(TObject *Sender);
        void __fastcall SaveTempPodrExecute(TObject *Sender);
        void __fastcall LoadNewLoginsExecute(TObject *Sender);
        void __fastcall CorrectNewLoginsExecute(TObject *Sender);
        void __fastcall ReadOtdelsExecute(TObject *Sender);
        void __fastcall PrepareReadExecute(TObject *Sender);
        void __fastcall PostReadExecute(TObject *Sender);
        void __fastcall PrepareUpdateOtdExecute(TObject *Sender);
        void __fastcall PostUpdateOtdExecute(TObject *Sender);
        void __fastcall ReadOtdelsAutoExecute(TObject *Sender);
private:	// User declarations
String Path;
TLocateOptions SO;
String MainDatabase;
String MainDatabase2;
String MainDatabase3;
void __fastcall ArchTimer(TObject *Sender, int N);
String Server;
int Port;
bool Start;

public:		// User declarations
bool Reg;
//bool LoadLogin(MDBConnector* DB);
void MergeOtdels();
void MergeLogins();
//void MergeLogins(MDBConnector* DB);
vector<String>Parameters;

int Role;
//vector<CDBItem>VDB;
        __fastcall TZast(TComponent* Owner);
//MDBConnector* Database;
//MDBConnector* Diary;
//MDBConnector* ADOUsrAspect;
//MDBConnector* ADOSvod;
String ServerName;
Client *MClient;
/*
AnsiString HardText();
bool VerifyKey();
bool Registered;
AnsiString CodeKey(AnsiString,AnsiString);
AnsiString DecodeKey(AnsiString,AnsiString);
AnsiString DecodeFile(AnsiString Patch, AnsiString Pass);
AnsiString Patch1, Patch2, Patch3;
String B1, B2, B3;
TIniFile *Ini;
TIniFile *Admin;
bool Connect;
AnsiString FullPatch;
*/
void __fastcall MyException(TObject *Sender,
                                    Exception *E);
//void TZast::CopyZn();

//bool WrMetod, MEcolog;
void CompactDB(TADOConnection * Conn, String B);
//bool Start;

int GetIDDBName(String Name);
};
//---------------------------------------------------------------------------
extern PACKAGE TZast *Zast;
//---------------------------------------------------------------------------
#endif
