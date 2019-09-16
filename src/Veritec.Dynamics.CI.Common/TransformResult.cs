using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Veritec.Dynamics.CI.Common
{
    internal class TransformResult
    {
        public string ReplacementAttributeName { get; set; }
        public Object ReplacementValue { get; set; }
        public Type ReplacementValueType { get; set; }
    }
}
