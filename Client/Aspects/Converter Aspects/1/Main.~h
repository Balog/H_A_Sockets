//---------------------------------------------------------------------------

#ifndef MainH
#define MainH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <DBGrids.hpp>
#include <Grids.hpp>
#include <Menus.hpp>
#include <Dialogs.hpp>
#include <ADODB.hpp>
#include <DB.hpp>
//---------------------------------------------------------------------------
class TForm1 : public TForm
{
__published:	// IDE-managed Components
        TLabel *Label1;
        TLabel *Label2;
        TEdit *AspRefs;
        TButton *Button1;
        TButton *Button2;
        TPopupMenu *PopupMenu1;
        TMenuItem *N1;
        TMenuItem *N2;
        TDBGrid *AspPatch;
        TEdit *Edit1;
        TOpenDialog *OpenDialog1;
        TOpenDialog *OpenDialog2;
        TADOConnection *Database;
        TADODataSet *Table;
        TDataSource *DataSource1;
        TADOConnection *TempFrom;
        TADOConnection *TempTo;
        TADOConnection *AspectTo;
        void __fastcall N1Click(TObject *Sender);
        void __fastcall FormShow(TObject *Sender);
        void __fastcall N2Click(TObject *Sender);
        void __fastcall Button1Click(TObject *Sender);
        void __fastcall Button2Click(TObject *Sender);
private:	// User declarations
TLocateOptions SO;
void CopyTable(String NameNode, String NameBranch);
int EncodeBranch(String NameBranch, int NumCode);
public:		// User declarations
        __fastcall TForm1(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TForm1 *Form1;
//---------------------------------------------------------------------------
#endif
