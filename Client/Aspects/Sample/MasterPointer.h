//---------------------------------------------------------------------------
#include <vcl.h>
#ifndef MasterPointerH
#define MasterPointerH
//---------------------------------------------------------------------------
template <class Type>
class MP {
private:
Type* t;
public:
MP(); // Создает указываемый объект
MP(const MP<Type>&); // Копирует указываемый объект
MP(TControl*);
MP(String);
~MP(); // Удаляет указываемый объект
MP<Type>& operator=(const MP<Type>&);
Type* operator->() const;
operator Type*();
};

//----------------------

template <class Type>
inline MP<Type>::MP(String F) : t(new Type(F))
{}

//----------------------
template <class Type>
inline MP<Type>::MP() : t(new Type())
{}
//--------------------------

template <class Type>
inline MP<Type>::MP(TControl *F): t(new Type(F))
{}
//------------------------------
template <class Type>
inline MP<Type>::MP(const MP<Type>& mp) : t(new Type(*(mp.t)))
{}
//------------------------------
template <class Type>
inline MP<Type>::~MP()
{
delete t;
}
//------------------------------
template <class Type>
inline MP<Type>& MP<Type>::operator=(const MP<Type>& mp)
{
if (this != &mp) {
delete t;
t = new Type(*(mp.t));
}
return *this;
}
//-----------------------------
template <class Type>
inline Type* MP<Type>::operator->() const
{
return t;
}
//-----------------------------
template <class Type>
MP<Type>::operator Type*()
{
 return t;
}
#endif
 