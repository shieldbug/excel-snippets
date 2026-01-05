using System;
using System.Collections.Generic;
using System.Text;

namespace Excel.Snippets
{
    public static partial class Demo
    {
        public static string IncrementColumn(string ColumnName)
        {
            /* Excel columns are essentially a base‑26 number system using letters A–Z */

            char[] Chars = ColumnName.ToUpper().ToCharArray();

            int index = Chars.Length - 1;

            while (index >= 0)
            {
                if (Chars[index] < 'Z')
                {
                    Chars[index]++;

                    return new string(Chars);
                }

                /* If it's 'Z', wrap to 'A' and carry to the next position. */

                Chars[index] = 'A';

                index--;
            }

            /* If we carried past the first character, prepend 'A' */

            /*
                A	=> B
                Z	=> AA
                AA	=> AB
                AZ	=> BA
                ZZ	=> AAA
                XFD	=> XFE

             */

            return "A" + new string(Chars);
        }
    }
}
