using Microsoft.Office.Interop.PowerPoint;
using System.Collections.Generic;

namespace VSTO_PPT_AddIn_Clean_SlideMaster.Cleanup.Model
{
    internal class SlideMasterCleanup : ISlideMasterCleanup
    {
        public IReadOnlyList<CustomLayout> UnUsedlayouts { get; set; }
        public IReadOnlyList<Master> UnUsedMaster { get; set; }
    }
}