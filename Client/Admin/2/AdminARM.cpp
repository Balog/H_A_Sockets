//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop
//---------------------------------------------------------------------------
USEFORM("Zastavka.cpp", Zast);
USEFORM("Main.cpp", Form1);
USEFORM("PassForm.cpp", Pass);
USEFORM("EditLogin.cpp", EditLogins);
//---------------------------------------------------------------------------
WINAPI WinMain(HINSTANCE, HINSTANCE, LPSTR, int)
{
        try
        {
                 Application->Initialize();
                 Application->CreateForm(__classid(TZast), &Zast);
                 Application->CreateForm(__classid(TForm1), &Form1);
                 Application->CreateForm(__classid(TPass), &Pass);
                 Application->CreateForm(__classid(TEditLogins), &EditLogins);
                 Application->Run();
        }
        catch (Exception &exception)
        {
                 Application->ShowException(&exception);
        }
        catch (...)
        {
                 try
                 {
                         throw Exception("");
                 }
                 catch (Exception &exception)
                 {
                         Application->ShowException(&exception);
                 }
        }
        return 0;
}
//---------------------------------------------------------------------------
