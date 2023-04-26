using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1
{
    
        internal class PeopleCollect
        {
            public static List<Pers> lst;
            public static string filepath = "data.txt";
           
            bool deletes = false;
            public PeopleCollect()
            {
               
                lst = new List<Pers>();
            
                FileStream f1 = new FileStream(filepath, FileMode.Open);
                StreamReader rdr = new StreamReader(f1, Encoding.UTF8);
                string line;
                string[] split;
                while ((line = rdr.ReadLine()) != null)
                {
                   
                    split = line.Split('+');
                    Pers obj = new Pers(Int32.Parse(split[0]), Int32.Parse(split[1]), Int32.Parse(split[2]), split[3], Int32.Parse(split[4]),
                    Int32.Parse(split[5]), Int32.Parse(split[6]));
                    lst.Add(obj);

                }




            }
            public void Add(Pers obj)
            {
                lst.Add(obj);
            }
            public void Remove(int nom)
            {
                foreach (Pers v in lst)
                {
                    if (v.nomer == nom)
                    {
                        lst.Remove(v);
                    }
                }
            }
            public int Count()
            {
                return lst.Count;
            }
            public Pers Item(int nom)
            {
                Pers x = null;
                foreach (Pers v in lst)
                {
                    if (v.nomer == nom)
                    {
                    x = v;
                    }
                }
                return x;

            }
            public IEnumerator<Pers> GetEnumerator()
            {
                return lst.GetEnumerator();
            }

        }
    

    
}
