//---------------------------------------------------------------------------

#ifndef FMoveAspH
#define FMoveAspH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <ADODB.hpp>
#include <DB.hpp>
#include <dbcgrids.hpp>
#include <DBCtrls.hpp>
#include <Mask.hpp>
#include <ExtCtrls.hpp>
#include <ActnList.hpp>
#include <Buttons.hpp>
#include <DBActns.hpp>
#include <ImgList.hpp>
#include <StdActns.hpp>
//---------------------------------------------------------------------------
class TMAsp : public TForm
{
__published:	// IDE-managed Components
        TPanel *Panel1;
        TPanel *Panel2;
        TDataSource *DataSource4;
        TADODataSet *MoveAspects;
        TComboBox *CPodr;
        TLabel *Label1;
        TLabel *Label2;
        TLabel *Label3;
        TLabel *Label4;
        TLabel *Label5;
        TLabel *Label6;
        TDBText *DBText1;
        TDBText *DBText2;
        TDBText *DBText3;
        TDBText *DBText4;
        TDBText *DBText5;
        TActionList *ActionList1;
        TDataSetPost *DataSetPost1;
        TImageList *ImageList1;
        TBitBtn *BitBtn1;
        TBitBtn *BitBtn2;
        TBitBtn *BitBtn3;
        TBitBtn *BitBtn4;
        TFileSaveAs *FileSaveAs1;
        TBitBtn *BitBtn5;
        TFileOpen *FileOpen1;
        TBitBtn *BitBtn6;
        TPanel *Panel3;
        TRadioGroup *RadioGroup1;
        TComboBox *ComboBox1;
        TEdit *Edit1;
        TEdit *Edit2;
        TEdit *Edit3;
        TEdit *Edit4;
        TComboBox *ComboBox2;
        TComboBox *ComboBox3;
        TADODataSet *Podr;
        TDataSource *DataSource1;
        TButton *Button1;
        TButton *Button2;
        TButton *Button3;
        TButton *Button4;
        TLabel *Label7;
        TLabel *Label8;
        TLabel *Label9;
        TLabel *Label10;
        TTimer *Timer1;
        TDataSetRefresh *DataSetRefresh1;
        void __fastcall FormShow(TObject *Sender);
        void __fastcall BitBtn1Click(TObject *Sender);
        void __fastcall BitBtn2Click(TObject *Sender);
        void __fastcall BitBtn3Click(TObject *Sender);
        void __fastcall BitBtn4Click(TObject *Sender);
        void __fastcall CPodrClick(TObject *Sender);
        void __fastcall BitBtn5Click(TObject *Sender);
        void __fastcall BitBtn6Click(TObject *Sender);
        void __fastcall FormClose(TObject *Sender, TCloseAction &Action);
        void __fastcall RadioGroup1Click(TObject *Sender);
        void __fastcall Button1Click(TObject *Sender);
        void __fastcall Button2Click(TObject *Sender);
        void __fastcall Button3Click(TObject *Sender);
        void __fastcall Button4Click(TObject *Sender);
        void __fastcall ComboBox1Click(TObject *Sender);
        void __fastcall ComboBox2Change(TObject *Sender);
        void __fastcall ComboBox3Change(TObject *Sender);
        void __fastcall Edit1DblClick(TObject *Sender);
        void __fastcall Edit2DblClick(TObject *Sender);
        void __fastcall Edit3DblClick(TObject *Sender);
        void __fastcall Edit4DblClick(TObject *Sender);
        void __fastcall Timer1Timer(TObject *Sender);
        void __fastcall FormKeyUp(TObject *Sender, WORD &Key,
          TShiftState Shift);
private:	// User declarations
bool Cont;
TLocateOptions SO;

void ChangeCPodr();
String LoadAspects();
void MergeAspects();


int Index;
AnsiString CText;






public:		// User declarations
void InpTer();
void InpDeyat();
void InpAsp();
void InpVozd();
String SaveAspects();
        __fastcall TMAsp(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TMAsp *MAsp;
//---------------------------------------------------------------------------
#endif
