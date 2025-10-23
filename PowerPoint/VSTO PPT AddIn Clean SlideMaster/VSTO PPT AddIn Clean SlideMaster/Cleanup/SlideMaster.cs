using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using VSTO_PPT_AddIn_Clean_SlideMaster.Cleanup.Model;

namespace VSTO_PPT_AddIn_Clean_SlideMaster.Cleanup
{
    internal class SlideMaster
    {
        public async Task Delete()
        {

            ISlideMasterCleanup slideMasterCleanup = await GetSlidesToBeDeleted();

            DeleteSlides(slideMasterCleanup);
        }

        private async Task<ISlideMasterCleanup> GetSlidesToBeDeleted()
        {
            var activePresentations = Globals.ThisAddIn.Application.ActivePresentation;

            var allSlides = activePresentations.Slides;
            var slideMasterSlides = activePresentations.SlideMaster;

            List<Master> lstAvailableMaster = new List<Master>();
            List<CustomLayout> lstAvailableCustomLayout = new List<CustomLayout>();


            IReadOnlyList<CustomLayout> lstUsedCustomLayouts = await Task.FromResult((from Slide slide in allSlides
                                                                                      select slide.CustomLayout).ToList()).ConfigureAwait(false);

            foreach (Design design in activePresentations.Designs)
            {
                var slideMaster = design.SlideMaster;

                lstAvailableMaster.Add(slideMaster);
                IReadOnlyList<CustomLayout> lstDesignCustomLayouts = await Task.FromResult((from CustomLayout layout in slideMaster.CustomLayouts
                                                                                            select layout).ToList()).ConfigureAwait(false);

                lstAvailableCustomLayout.AddRange(lstDesignCustomLayouts);
            }

            IReadOnlyList<CustomLayout> unUsedlayouts = lstAvailableCustomLayout.Where(p => !lstUsedCustomLayouts.Contains(p)).ToList();

            return new SlideMasterCleanup
            {
                UnUsedlayouts = unUsedlayouts,
                UnUsedMaster = lstAvailableMaster.Where(p => !(from CustomLayout layout in p.CustomLayouts select layout).ToList().Except(unUsedlayouts).Any()).ToList()
            };
        }

        private void DeleteSlides(ISlideMasterCleanup slideMasterCleanup)
        {
            double initialSize = GetFileSize();
            int deleteSlideCount = 0;

            IReadOnlyList<CustomLayout> unUsedlayouts = slideMasterCleanup.UnUsedlayouts;
            IReadOnlyList<Master> unUsedMaster = slideMasterCleanup.UnUsedMaster;

            foreach (CustomLayout layout in unUsedlayouts)
            {
                deleteSlideCount++;
                layout.Delete();
            }

            foreach (Master master in unUsedMaster)
            {
                master.Delete();
            }

            Globals.ThisAddIn.Application.ActivePresentation.Save();

            double finalSize = GetFileSize();

            MessageBox.Show($"Done {Environment.NewLine}Previous size: {initialSize}MB{Environment.NewLine}New Size: {finalSize}MB{Environment.NewLine}Total slides deleted: {deleteSlideCount}", "SLide Master Delete Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private double GetFileSize()
        {
            string filePath = Globals.ThisAddIn.Application.ActivePresentation.FullName;

            if (File.Exists(filePath))
            {
                FileInfo fileInfo = new FileInfo(filePath);
                return fileInfo.Length / (1024 * 1024);
            }

            return 0;
        }
    }
}
