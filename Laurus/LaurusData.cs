using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Laurus {
    internal class LaurusData {
        public static readonly IReadOnlyDictionary<short, string> laurusDataSet = new Dictionary<short, string>(){
            {12,"hi"},{13,"hello"}
        };
        public void Hha() {
            //string hi;
            laurusDataSet.TryGetValue(12,out string hi);
        }

    }

}
