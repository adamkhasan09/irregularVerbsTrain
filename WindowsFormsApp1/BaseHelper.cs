using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1
{
    class BaseHelper
    {
        public string[] explode(string delimiter, string args)
        {
            //♕ ♖ ♗ ♘ ♙ ♚ ♛ ♜ ♝ ♞ ♟ ♠ ♡ ♢ ♣ ♤ ♥ ♦ ♧ ♩ ♪ ♫ ♬ ♭ ♮ ♯
            // (char)9820♜ (char)9822♞
            string sym = (char)9820 + "";
            string repArgs = args.Replace(delimiter, sym);
            string[] arrArgs = repArgs.Split((char)9820);
            return arrArgs;
        }
        public string inplode(string[] Array, string delimiter)
        {
            string res = "";
            for (int i = 0; i < Array.Count(); i++)
            {
                string del = i < Array.Count() - 1 ? delimiter : "";
                res += Array[i] + del;
            }
            return res;
        }
       
       
    }
}
