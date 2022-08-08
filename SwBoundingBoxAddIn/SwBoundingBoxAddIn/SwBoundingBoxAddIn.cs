using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swcommands;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;

namespace SwBoundingBoxAddIn
{
    [ComVisible(true)]
    [Guid("7527A62E-21A5-4F35-B9A6-3043DBE1E07C")]
    public class SwBoundingBoxAddIn:ISwAddin
    {
        public bool ConnectToSW(object ThisSW, int Cookie)
        {
            throw new NotImplementedException();
        }

        public bool DisconnectFromSW()
        {
            throw new NotImplementedException();
        }
    }
}
