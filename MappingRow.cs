using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace evaluacion_ceapsi
{
    class MappingRow
    {

        public int SourcePosition { get; set; }
        public int TargetSheet { get; set; }
        public int TargetRow { get; set; }
        public int TargetColumn { get; set; }
        public string Formula { get; set; }


    }
}
