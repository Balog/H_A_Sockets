//---------------------------------------------------------------------------

#ifndef MainFormH
#define MainFormH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
//#include "F_Vvedenie.h"
//#include "Metod.h"
#include <Menus.hpp>
#include <ADODB.hpp>
#include <DB.hpp>
#include <ComCtrls.hpp>
#include <Buttons.hpp>
#include <DBCtrls.hpp>
#include "Main.h"
#include <ExtCtrls.hpp>
#include <Mask.hpp>
//#include "InpDocs.h"
//#include "Settings.h"
//#include "InputFiltr.h"
//#include "Form_SendFile.h"
#include "Winbase.h"
//#include "SetZn.h"
#include <ActnList.hpp>
#include <DBActns.hpp>
//---------------------------------------------------------------------------
class TForm1 : public TForm
{
__published:	// IDE-managed Components
        TMainMenu *MainMenu1;
        TMenuItem *N4;
        TMenuItem *N5;
        TMenuItem *N6;
        TMenuItem *N7;
        TMenuItem *N1;
        TMenuItem *N8;
        TADODataSet *Branches;
        TLabel *Label2;
        TLabel *Label3;
        TLabel *Label4;
        TButton *Button1;
        TLabel *Label5;
        TButton *Button2;
        TLabel *Label6;
        TLabel *Label7;
        TButton *Button3;
        TLabel *Label8;
        TButton *Button4;
        TButton *Button5;
        TLabel *Label9;
        TDateTimePicker *DateTimePicker1;
        TLabel *Label10;
        TLabel *Label11;
        TLabel *Label12;
        TLabel *Label13;
        TLabel *Label14;
        TLabel *Label15;
        TLabel *Label16;
        TLabel *Label17;
        TLabel *Label18;
        TLabel *Label19;
        TLabel *Label20;
        TLabel *Label21;
        TLabel *Label22;
        TLabel *Label23;
        TLabel *Label24;
        TLabel *Label25;
        TStatusBar *StatusBar1;
        TLabel *Label28;
        TDateTimePicker *DateTimePicker2;
        TDateTimePicker *DateTimePicker3;
        TLabel *Label29;
        TBitBtn *BitBtn1;
        TBitBtn *BitBtn2;
        TBitBtn *BitBtn3;
        TBitBtn *BitBtn4;
        TBitBtn *BitBtn5;
        TBitBtn *BitBtn6;
        TLabel *Label30;
        TEdit *NumRecord;
        TLabel *Label31;
        TEdit *CountRecord;
        TLabel *Label32;
        TLabel *Label33;
        TBitBtn *BitBtn7;
        TLabel *Label34;
        TLabel *Label35;
        TBitBtn *BitBtn9;
        TDBMemo *DBMemo1;
        TDBMemo *DBMemo2;
        TDBMemo *DBMemo4;
        TADODataSet *Aspects;
        TDataSource *DataSource1;
        TADODataSet *Territoriya;
        TADODataSet *Deyatelnost;
        TADODataSet *Aspect;
        TADODataSet *Vozdeystvie;
        TADODataSet *Podrazdel;
        TADODataSet *Situaciya;
        TADODataSet *Posledstvie;
        TADODataSet *Tiagest;
        TADODataSet *Prioritet;
        TADODataSet *TempAspects;
        TComboBox *CPodrazdel;
        TComboBox *CSit;
        TComboBox *CProyav;
        TComboBox *CPosl;
        TComboBox *CPrior;
        TADODataSet *Znachimost;
        TEdit *Edit1;
        TEdit *Edit2;
        TEdit *Edit4;
        TEdit *Edit5;
        TCheckBox *CheckBox1;
        TCheckBox *CheckBox2;
        TCheckBox *CheckBox3;
        TCheckBox *CheckBox4;
        TCheckBox *CheckBox5;
        TCheckBox *CheckBox6;
        TCheckBox *CheckBox7;
        TEdit *Edit6;
        TDBEdit *DBEdit2;
        TButton *Button6;
        TDBMemo *DBMemo31;
        TLabel *LFiltr;
        TEdit *Edit3;
        TADODataSet *Svod;
        TADOCommand *Comm;
        TADODataSet *Report;
        TLabel *Label36;
        TLabel *Label37;
        TLabel *Label26;
        TEdit *Zmax;
        TActionList *ActionList1;
        TDataSetPost *DataSetPost1;
        TMenuItem *N2;
        TMenuItem *N3;
        TMenuItem *N9;
        TMenuItem *N10;
        TDataSetRefresh *DataSetRefresh1;
        void __fastcall FormClose(TObject *Sender, TCloseAction &Action);
        void __fastcall N8Click(TObject *Sender);
        void __fastcall CPodrazdelClick(TObject *Sender);
        void __fastcall CSitClick(TObject *Sender);
        void __fastcall CProyavClick(TObject *Sender);
        void __fastcall CPoslClick(TObject *Sender);
        void __fastcall CPriorClick(TObject *Sender);
        void __fastcall BitBtn5Click(TObject *Sender);
        void __fastcall BitBtn1Click(TObject *Sender);
        void __fastcall BitBtn4Click(TObject *Sender);
        void __fastcall BitBtn2Click(TObject *Sender);
        void __fastcall BitBtn3Click(TObject *Sender);
        void __fastcall BitBtn6Click(TObject *Sender);
        void __fastcall Button1Click(TObject *Sender);
        void __fastcall Button2Click(TObject *Sender);
        void __fastcall Button3Click(TObject *Sender);
        void __fastcall Button4Click(TObject *Sender);
        void __fastcall DBCheckBox1Click(TObject *Sender);
        void __fastcall DBCheckBox2Click(TObject *Sender);
        void __fastcall DBCheckBox3Click(TObject *Sender);
        void __fastcall DBCheckBox4Click(TObject *Sender);
        void __fastcall DBCheckBox5Click(TObject *Sender);
        void __fastcall DBCheckBox6Click(TObject *Sender);
        void __fastcall DBCheckBox7Click(TObject *Sender);
        void __fastcall CPoslChange(TObject *Sender);
        void __fastcall CheckBox1Click(TObject *Sender);
        void __fastcall CheckBox2Click(TObject *Sender);
        void __fastcall CheckBox3Click(TObject *Sender);
        void __fastcall CheckBox4Click(TObject *Sender);
        void __fastcall CheckBox5Click(TObject *Sender);
        void __fastcall CheckBox6Click(TObject *Sender);
        void __fastcall CheckBox7Click(TObject *Sender);
        void __fastcall FormShow(TObject *Sender);
        void __fastcall DateTimePicker1Change(TObject *Sender);
        void __fastcall DateTimePicker2Change(TObject *Sender);
        void __fastcall DateTimePicker3Change(TObject *Sender);
        void __fastcall Button6Click(TObject *Sender);
        void __fastcall Button5Click(TObject *Sender);
        void __fastcall BitBtn7Click(TObject *Sender);
        void __fastcall FormHide(TObject *Sender);
        void __fastcall N9Click(TObject *Sender);
        void __fastcall FormActivate(TObject *Sender);
        void __fastcall Cdjlysqhttcnhfcgtrnjd1Click(TObject *Sender);
        void __fastcall BitBtn9Click(TObject *Sender);
        void __fastcall DBMemo1Exit(TObject *Sender);
        void __fastcall DBMemo2Exit(TObject *Sender);
        void __fastcall DBMemo31Exit(TObject *Sender);
        void __fastcall DBMemo4Exit(TObject *Sender);
        void __fastcall MetodikaN19Click(TObject *Sender);
        void __fastcall N14Click(TObject *Sender);
        void __fastcall N15Click(TObject *Sender);
        void __fastcall N16Click(TObject *Sender);
        void __fastcall N17Click(TObject *Sender);
        void __fastcall CProyavKeyPress(TObject *Sender, char &Key);
        void __fastcall CPoslKeyPress(TObject *Sender, char &Key);
        void __fastcall CPriorKeyPress(TObject *Sender, char &Key);
        void __fastcall CPodrazdelKeyPress(TObject *Sender, char &Key);
        void __fastcall CSitKeyPress(TObject *Sender, char &Key);
        void __fastcall DateTimePicker1KeyPress(TObject *Sender,
          char &Key);
        void __fastcall DateTimePicker2KeyPress(TObject *Sender,
          char &Key);
        void __fastcall DateTimePicker3KeyPress(TObject *Sender,
          char &Key);
        void __fastcall N1Click(TObject *Sender);
        void __fastcall N18Click(TObject *Sender);
        void __fastcall Edit1DblClick(TObject *Sender);
        void __fastcall Edit2DblClick(TObject *Sender);
        void __fastcall Edit4DblClick(TObject *Sender);
        void __fastcall Edit5DblClick(TObject *Sender);
        void __fastcall N10Click(TObject *Sender);
        void __fastcall N3Click(TObject *Sender);
        void __fastcall N4Click(TObject *Sender);
        void __fastcall N5Click(TObject *Sender);
        void __fastcall N7Click(TObject *Sender);
        void __fastcall FormKeyUp(TObject *Sender, WORD &Key,
          TShiftState Shift);
private:	// User declarations
bool Registered;
TLocateOptions SO;
bool LoadSpravs();
void SetAspects(String Login);

bool VerifySit();
bool VerifyTerr();
bool VerifyDeyat();
bool VerifyAsp();
bool VerifyVozd();
bool VerifyCryt();
bool VerifyMeropr();

bool LoadSit();
void MergeSit();

bool LoadTerr();
void MergeTerr();

bool LoadDeyat();
void MergeDeyat();

bool LoadAsp();
void MergeAsp();

bool LoadVozd();
void MergeVozd();

bool LoadCryt();
void MergeCryt();


void RefreshLocalReference(String NameNode, String NameBranch);

bool LoadPodr();
void EnabledForm(bool B);

public:		// User declarations
        __fastcall TForm1(TComponent* Owner);

bool Demo;
int NumLogin;
String Login;
int Role;

void SynhronTerr();
void SynhronDeyat();
void SynhronAspect();
void SynhronVozd();

void NewRecord();
void InitCombo();
void SetCombo();

int Mode;
int IForm;

void InpTer();
void InpDeyat();
void InpAsp();
void InpVozd();
void InpMeropr();

void Calc();
void Initialize();
bool C1,C2,C3,C4,C5,C6,C7;
int CurrentRecord;
/*
AnsiString Address(Variant Sheet,int x,int y);


void CreateReport1(int Podr, TDateTime Date1, TDateTime Date2, AnsiString Isp);
void CreateReport2(int Podr, TDateTime Date1, TDateTime Date2, AnsiString Isp);
*/
AnsiString Filtr1, Filtr2;

bool  PrepareSaveAspects();
bool  SaveAspects();
bool LoadAspects(int NumLogin);
void PrepareMergeAspects();
void  MergeAspects(int NumLogin);

        BEGIN_MESSAGE_MAP
          VCL_MESSAGE_HANDLER(WM_SYSCOMMAND, TMessage, WMSysCommand)
        END_MESSAGE_MAP(TComponent);
void __fastcall WMSysCommand(TMessage & Msg);
};
//---------------------------------------------------------------------------
extern PACKAGE TForm1 *Form1;
//---------------------------------------------------------------------------
#endif