//---------------------------------------------------------------------------

#ifndef SetZnH
#define SetZnH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>

#include <ADODB.hpp>
#include <DB.hpp>
#include <DBGrids.hpp>
#include <Grids.hpp>
#include <Menus.hpp>
#include <dbcgrids.hpp>
#include <DBCtrls.hpp>
#include <Mask.hpp>
//---------------------------------------------------------------------------
class TSZn : public TForm
{
__published:	// IDE-managed Components
        TDBGrid *DBGrid1;
        TDataSource *DataSource1;
        TADODataSet *ADODataSet1;
        TButton *Button1;
        TPopupMenu *PopupMenu1;
        TMenuItem *N1;
        TMenuItem *N2;
        TADODataSet *ADODataSet2;
        void __fastcall ADODataSet1BeforePost(TDataSet *DataSet);
        void __fastcall ADODataSet1AfterPost(TDataSet *DataSet);
        void __fastcall DataSource1UpdateData(TObject *Sender);
        void __fastcall ADODataSet1BeforeEdit(TDataSet *DataSet);
        void __fastcall FormShow(TObject *Sender);
        void __fastcall Button1Click(TObject *Sender);
        void __fastcall N1Click(TObject *Sender);
        void __fastcall N2Click(TObject *Sender);
        void __fastcall FormClose(TObject *Sender, TCloseAction &Action);
        void __fastcall FormCloseQuery(TObject *Sender, bool &CanClose);
        void __fastcall PopupMenu1Popup(TObject *Sender);
        void __fastcall FormActivate(TObject *Sender);
        void __fastcall ADODataSet1BeforeInsert(TDataSet *DataSet);
        void __fastcall ADODataSet1AfterInsert(TDataSet *DataSet);
private:	// User declarations
public:		// User declarations
        __fastcall TSZn(TComponent* Owner);
void Crit(TDataSet *DataSet);
void Crit1(TDataSet *DataSet);
void Crit2(TDataSet *DataSet);
bool C;
bool Ins;
bool Kriteriy;
};
//---------------------------------------------------------------------------
extern PACKAGE TSZn *SZn;
//---------------------------------------------------------------------------
#endif
