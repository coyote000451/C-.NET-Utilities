using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Utilities
{
    public class CompareList
    {
        //Method to compare two list of string
        ////public List<string> Contains(List<string> list1, List<string> list2)
        ////{
        ////    List<string> result = new List<string>();

        ////    result.AddRange(list1.Except(list2, StringComparer.OrdinalIgnoreCase));
        ////    //result.AddRange(list2.Except(list1, StringComparer.OrdinalIgnoreCase));

        ////    return result;
        ////}

        public IEnumerable<string> Contains(List<string> list1, List<string> list2)
        {
            var newData = list1.Select(i => i.ToString()).Intersect(list2);

            return newData;
        }

    }

}
