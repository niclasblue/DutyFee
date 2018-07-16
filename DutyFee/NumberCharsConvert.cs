using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DutyFee
{
    public static class NumberCharsConvert
    {
        /// 
        /// 把1,2,3,...,26转换成A,B,C,...,Y,Z
        /// 
        /// 要转换成字母的数字（数字范围在闭区间[1,26]）
        /// 
        /// 
        public static string NumberToChar(int number)
        {
            try
            {
                if ((number < 1) || (number > 26))
                    throw new ArgumentOutOfRangeException();

                    int num = number + 64;
                    System.Text.ASCIIEncoding asciiEncoding = new System.Text.ASCIIEncoding();
                    byte[] btNumber = new byte[] { (byte)num };
                    return asciiEncoding.GetString(btNumber);
            }
            catch (ArgumentOutOfRangeException)
            {
                Console.WriteLine("Number out of Range(1-26).");
                throw;
            }

        }


        /// 
        /// 把1,2,3,...转换成A,B,C,...,Y,Z,AA,AB,...,BA...
        /// 
        /// 本函数主要用在excel表格中，将数字序列转换为excel列标。
        /// 
        public static string NumberToChars(int number)
        {
            try
            {
                string result = "";
                if (number < 1)
                    throw new ArgumentOutOfRangeException();
                int tempNum = number;
                while (tempNum > 0)
                {
                    int m = ((tempNum % 26) == 0 ? 26 : (tempNum % 26));
                    result = NumberToChar(m) + result;
                    tempNum = (tempNum - m)/26;
                }
                return result;
                
            }
            catch(ArgumentOutOfRangeException)
            {
                Console.WriteLine("Number out of range");
                throw;
            }
        }

    }
}
