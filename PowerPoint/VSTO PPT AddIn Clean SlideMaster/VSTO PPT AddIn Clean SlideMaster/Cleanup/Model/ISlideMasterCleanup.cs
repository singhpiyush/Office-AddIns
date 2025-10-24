using Microsoft.Office.Interop.PowerPoint;
using System.Collections.Generic;

namespace VSTO_PPT_AddIn_Clean_SlideMaster.Cleanup.Model
{
    internal interface ISlideMasterCleanup
    {
        IReadOnlyList<CustomLayout> UnUsedlayouts { get; set; }
        IReadOnlyList<Master> UnUsedMaster { get; set; }
    }
}