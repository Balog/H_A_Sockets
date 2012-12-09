//---------------------------------------------------------------------------
#include <vcl.h>
#ifndef CodeTextH
#define CodeTextH
//---------------------------------------------------------------------------
class CodeText
{
private:
String CodeText;
String DecodeText;
String Sect[256];
String CreatePassword(String ManPass);
String CombineSection(String Section, String Pass);
String DeCombineSection(String Section, String Pass);
public:
CodeText();
String Crypt(String Text,String Password);
String EnCrypt(String Text,String Password, int Dlina);
String ReCrypt(String Text,String OldPassword, String NewPassword,int Dlina);
};
#endif
