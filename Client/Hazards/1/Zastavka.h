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
#include <ADODB.hpp>
#include <DB.hpp>
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
        TAction *PrepareConnectBase;
        TAction *ConnectBase;
        TAction *LoadLogins;
        TAction *ViewLogins;
        TAction *BeginWork;
        TAction *DeletePodr;
        TAction *MergeMetodika;
        TAction *ReadMetodika;
        TAction *ReadWriteDoc;
        TAction *ReadPodrazd;
        TAction *MergePodrazd;
        TAction *ReadCrit;
        TAction *MergeCrit;
        TAction *ReadSit;
        TAction *MergeSit1;
        TAction *MergeSit2;
        TAction *ReadVozd1;
        TAction *ReadVozd2;
        TAction *MergeVozd;
        TAction *MergeVozd2;
        TAction *ReadMeropr1;
        TAction *MergeMeropr;
        TAction *ReadTerr1;
        TAction *MergeTerr1;
        TAction *MergeTerr2;
        TAction *ReadDeyat1;
        TAction *MergeDeyat1;
        TAction *ReadMeropr2;
        TAction *ReadTerr2;
        TAction *ReadDeyat2;
        TAction *MergeDeyat2;
        TAction *ReadAspect1;
        TAction *ReadAspect2;
        TAction *MergeAspect1;
        TAction *MergeAspect2;
        TAction *WriteMetodika;
        TAction *WritePodr;
        TAction *MergeServerPodr;
        TAction *WriteCrit;
        TAction *MergeServerCrit;
        TAction *WriteSit;
        TAction *WriteSit2;
        TAction *WritePodr2;
        TAction *MergeServerPodr2;
        TAction *MergeServerSit;
        TAction *WriteVozd1;
        TAction *WriteVozd2;
        TAction *MergeServerVozd;
        TAction *EndMergeServerVozd;
        TAction *WriteMeropr1;
        TAction *WriteMeropr2;
        TAction *MergeServerMeropr;
        TAction *EndMergeServerMeropr;
        TAction *WriteTerr1;
        TAction *WriteTerr2;
        TAction *MergeServerTerr;
        TAction *EndMergeServerTerr;
        TAction *WriteDeyat1;
        TAction *WriteDeyat2;
        TAction *MergeServerDeyat;
        TAction *EndMergeServerDeyat;
        TAction *WriteAspect1;
        TAction *WriteAspect2;
        TAction *MergeServerAspect;
        TAction *EndMergeServerAspect;
        TAction *ContReport;
        TAction *ContSvodReport;
        TAction *StartLoadPodr;
        TAction *StartLoadObslOtd;
        TAction *StartMergeLoginsPodr;
        TAction *CompareMSpecAspects;
        TAction *SaveAspectsMSpec;
        TAction *EndsaveAspectsMSpec;
        TAction *CompareMSpecAspects2;
        TAction *SaveAspectsMSpec2;
        TAction *EndsaveAspectsMSpec2;
        TAction *LoadMSpecAspects;
        TAction *SaveAspectsMSpec0;
        TAction *StopProgram;
        TAction *StartLoadPodrUSR;
        TAction *StartLoadObslOtdUSR;
        TAction *ShowForm1;
        TAction *MergeAspectsUser;
        TAction *WriteAspectsUsr;
        TAction *MergeAspectsUserQ;
        TAction *PreLoadLogins;
        TTimer *BlockServer;
        TTimer *UnBlockServer;
        TAction *UnblockAndRWD;
        TAction *ReadAspectsMSpec;
        TAction *PrepareCompareMSpecAspects;
        TAction *EndReadAspectsMSpec;
        TAction *StopProgram2;
        TAction *PrepStartMergeLoginPodr;
        TAction *PrepareReadAspectsUsr;
        TAction *MergeAspectsUser1;
        TAction *PrepWriteAspUsr;
        TAction *PrepMergeAspectsUserQ;
        TAction *ContStartReports;
        TAction *CloseMAsp;
        TAction *ContStartReports2;
        TAction *PrepWriteAspUsr_ADM;
        TAction *PrepWriteAspUsr_ADM_1;
        TAction *PrepWriteAspUsr_MSpec_1;
        TAction *PrepWriteAspUsr_MSpec;
        TAction *PostWriteAspectsUsr1;
        TAction *AspQ1;
        TAction *ReadTempAsp;
        TAction *VerIsNew;
        TAction *VerIsNew1;
        void __fastcall Timer1Timer(TObject *Sender);
        void __fastcall Image1Click(TObject *Sender);
        void __fastcall FormShow(TObject *Sender);
        void __fastcall Timer2Timer(TObject *Sender);
        void __fastcall FormDestroy(TObject *Sender);
        void __fastcall ClientSocketRead(TObject *Sender,
          TCustomWinSocket *Socket);
        void __fastcall PrepareConnectBaseExecute(TObject *Sender);
        void __fastcall ConnectBaseExecute(TObject *Sender);
        void __fastcall LoadLoginsExecute(TObject *Sender);
        void __fastcall ViewLoginsExecute(TObject *Sender);
        void __fastcall ClientSocketError(TObject *Sender,
          TCustomWinSocket *Socket, TErrorEvent ErrorEvent,
          int &ErrorCode);
        void __fastcall BeginWorkExecute(TObject *Sender);
        void __fastcall DeletePodrExecute(TObject *Sender);
        void __fastcall MergeMetodikaExecute(TObject *Sender);
        void __fastcall ReadMetodikaExecute(TObject *Sender);
        void __fastcall ReadWriteDocExecute(TObject *Sender);
        void __fastcall ReadPodrazdExecute(TObject *Sender);
        void __fastcall MergePodrazdExecute(TObject *Sender);
        void __fastcall ReadCritExecute(TObject *Sender);
        void __fastcall MergeCritExecute(TObject *Sender);
        void __fastcall ReadSitExecute(TObject *Sender);
        void __fastcall MergeSit1Execute(TObject *Sender);
        void __fastcall MergeSit2Execute(TObject *Sender);
        void __fastcall ReadVozd1Execute(TObject *Sender);
        void __fastcall ReadVozd2Execute(TObject *Sender);
        void __fastcall MergeVozdExecute(TObject *Sender);
        void __fastcall MergeVozd2Execute(TObject *Sender);
        void __fastcall ReadMeropr1Execute(TObject *Sender);
        void __fastcall MergeMeroprExecute(TObject *Sender);
        void __fastcall ReadTerr1Execute(TObject *Sender);
        void __fastcall MergeTerr1Execute(TObject *Sender);
        void __fastcall MergeTerr2Execute(TObject *Sender);
        void __fastcall ReadDeyat1Execute(TObject *Sender);
        void __fastcall ReadMeropr2Execute(TObject *Sender);
        void __fastcall ReadTerr2Execute(TObject *Sender);
        void __fastcall ReadDeyat2Execute(TObject *Sender);
        void __fastcall MergeDeyat1Execute(TObject *Sender);
        void __fastcall MergeDeyat2Execute(TObject *Sender);
        void __fastcall ReadAspect1Execute(TObject *Sender);
        void __fastcall ReadAspect2Execute(TObject *Sender);
        void __fastcall MergeAspect1Execute(TObject *Sender);
        void __fastcall MergeAspect2Execute(TObject *Sender);
        void __fastcall WriteDocExecute(TObject *Sender);
        void __fastcall WriteMetodikaExecute(TObject *Sender);
        void __fastcall WritePodrExecute(TObject *Sender);
        void __fastcall MergeServerPodrExecute(TObject *Sender);
        void __fastcall WriteCritExecute(TObject *Sender);
        void __fastcall MergeServerCritExecute(TObject *Sender);
        void __fastcall WriteSitExecute(TObject *Sender);
        void __fastcall WriteSit2Execute(TObject *Sender);
        void __fastcall WritePodr2Execute(TObject *Sender);
        void __fastcall MergeServerPodr2Execute(TObject *Sender);
        void __fastcall MergeServerSitExecute(TObject *Sender);
        void __fastcall WriteVozd1Execute(TObject *Sender);
        void __fastcall WriteVozd2Execute(TObject *Sender);
        void __fastcall MergeServerVozdExecute(TObject *Sender);
        void __fastcall EndMergeServerVozdExecute(TObject *Sender);
        void __fastcall WriteMeropr1Execute(TObject *Sender);
        void __fastcall WriteMeropr2Execute(TObject *Sender);
        void __fastcall MergeServerMeroprExecute(TObject *Sender);
        void __fastcall EndMergeServerMeroprExecute(TObject *Sender);
        void __fastcall WriteTerr1Execute(TObject *Sender);
        void __fastcall WriteTerr2Execute(TObject *Sender);
        void __fastcall MergeServerTerrExecute(TObject *Sender);
        void __fastcall EndMergeServerTerrExecute(TObject *Sender);
        void __fastcall WriteDeyat1Execute(TObject *Sender);
        void __fastcall WriteDeyat2Execute(TObject *Sender);
        void __fastcall MergeServerDeyatExecute(TObject *Sender);
        void __fastcall EndMergeServerDeyatExecute(TObject *Sender);
        void __fastcall WriteAspect1Execute(TObject *Sender);
        void __fastcall WriteAspect2Execute(TObject *Sender);
        void __fastcall MergeServerAspectExecute(TObject *Sender);
        void __fastcall EndMergeServerAspectExecute(TObject *Sender);
        void __fastcall ContSvodReportExecute(TObject *Sender);
        void __fastcall StartLoadPodrExecute(TObject *Sender);
        void __fastcall StartLoadObslOtdExecute(TObject *Sender);
        void __fastcall StartMergeLoginsPodrExecute(TObject *Sender);
        void __fastcall CompareMSpecAspectsExecute(TObject *Sender);
        void __fastcall SaveAspectsMSpecExecute(TObject *Sender);
        void __fastcall EndsaveAspectsMSpecExecute(TObject *Sender);
        void __fastcall CompareMSpecAspects2Execute(TObject *Sender);
        void __fastcall SaveAspectsMSpec2Execute(TObject *Sender);
        void __fastcall EndsaveAspectsMSpec2Execute(TObject *Sender);
        void __fastcall LoadMSpecAspectsExecute(TObject *Sender);
        void __fastcall SaveAspectsMSpec0Execute(TObject *Sender);
        void __fastcall StopProgramExecute(TObject *Sender);
        void __fastcall StartLoadPodrUSRExecute(TObject *Sender);
        void __fastcall StartLoadObslOtdUSRExecute(TObject *Sender);
        void __fastcall ShowForm1Execute(TObject *Sender);
        void __fastcall MergeAspectsUserExecute(TObject *Sender);
        void __fastcall WriteAspectsUsrExecute(TObject *Sender);
        void __fastcall MergeAspectsUserQExecute(TObject *Sender);
        void __fastcall PreLoadLoginsExecute(TObject *Sender);
        void __fastcall BlockServerTimer(TObject *Sender);
        void __fastcall UnBlockServerTimer(TObject *Sender);
        void __fastcall UnblockAndRWDExecute(TObject *Sender);
        void __fastcall ReadAspectsMSpecExecute(TObject *Sender);
        void __fastcall PrepareCompareMSpecAspectsExecute(TObject *Sender);
        void __fastcall EndReadAspectsMSpecExecute(TObject *Sender);
        void __fastcall StopProgram2Execute(TObject *Sender);
        void __fastcall PrepStartMergeLoginPodrExecute(TObject *Sender);
        void __fastcall PrepareReadAspectsUsrExecute(TObject *Sender);
        void __fastcall MergeAspectsUser1Execute(TObject *Sender);
        void __fastcall PrepWriteAspUsrExecute(TObject *Sender);
        void __fastcall PrepMergeAspectsUserQExecute(TObject *Sender);
        void __fastcall ContStartReportsExecute(TObject *Sender);
        void __fastcall CloseMAspExecute(TObject *Sender);
        void __fastcall ContStartReports2Execute(TObject *Sender);
        void __fastcall PrepWriteAspUsr_ADMExecute(TObject *Sender);
        void __fastcall PrepWriteAspUsr_ADM_1Execute(TObject *Sender);
        void __fastcall PrepWriteAspUsr_MSpec_1Execute(TObject *Sender);
        void __fastcall PrepWriteAspUsr_MSpecExecute(TObject *Sender);
        void __fastcall PostWriteAspectsUsr1Execute(TObject *Sender);
        void __fastcall AspQ1Execute(TObject *Sender);
        void __fastcall ReadTempAspExecute(TObject *Sender);
        void __fastcall VerIsNewExecute(TObject *Sender);
        void __fastcall VerIsNew1Execute(TObject *Sender);
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
void MergeLogins();
void MergeOtdels();
void MergeAspects(int NumLogin, bool Quit);
void CorrectPodrazd();
//void PrepWriteAspUsr_MSpec();


public:		// User declarations
bool Saved;
bool Stop;
bool Reg;

MDBConnector* ADOConn;
MDBConnector* ADOAspect;
MDBConnector* ADOUsrAspect;
MDBConnector* ADOSvod;

vector<String>Parameters;

int Role;
void WaitBlockServer(bool Flag);
        __fastcall TZast(TComponent* Owner);
String ServerName;
Client *MClient;

void __fastcall MyException(TObject *Sender,
                                    Exception *E);

void CompactDB(TADOConnection * Conn, String B);

int GetIDDBName(String Name);

void BlockMK(bool B);
};
//---------------------------------------------------------------------------
extern PACKAGE TZast *Zast;
//---------------------------------------------------------------------------
#endif
