using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1
{
   
        internal class Pers
        {//поля
            private int _nomer;
            private int _filial;
            private int _code;
            private string _fio;
            private int _age;
            private int _money;
            private int _stag;

            public Pers(int nomer, int filial, int code, string fio, int age, int money, int stag)
            {
                this._nomer = nomer;
                this._filial = filial;
                this._code = code;
                this._fio = fio;
                this._age = age;
                this._money = money;
                this._stag = stag;
            }
            public int nomer
            {
                get { return _nomer; }
                set { _nomer = value; }
            }
            public int filial
            {
                get { return _filial; }
                set { _filial = value; }
            }
            public int code
            {
                get { return _code; }
                set { _code = value; }
            }
            public string fio
            {
                //свойства
                get { return _fio; }
                set { _fio = value; }
            }

            public int age
            {
                get { return _age; }
                set { _age = value; }
            }
            public int money
            {
                get { return _money; }
                set { _money = value; }
            }
            public int stag
            {
                get { return _stag; }
                set { _stag = value; }
            }


        }
    }

