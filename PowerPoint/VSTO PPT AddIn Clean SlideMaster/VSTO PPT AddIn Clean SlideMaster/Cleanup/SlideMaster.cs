using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using VSTO_PPT_AddIn_Clean_SlideMaster.Cleanup.Model;

namespace VSTO_PPT_AddIn_Clean_SlideMaster.Cleanup
{
    internal class SlideMaster
    {
        /// <summary>
        /// Entry point kept as Task to match prior signature.
        /// </summary>
        public void Delete()
        {
            ISlideMasterCleanup slideMasterCleanup = GetSlidesToBeDeleted();

            DeleteSlides(slideMasterCleanup);
        }

        /// <summary>
        /// Returns a collection of unused layouts and master slides from the Slide Master
        /// </summary>
        /// <returns>Returns an object of class SlideMasterCleanup with the unused custom & master slides</returns>
        /// <exception cref="InvalidOperationException">Active presentation can sometimes be NULL when the add-in is not loaded successfully or due to COMS error. In those edge cases, this exception will be thrown.</exception>
        private ISlideMasterCleanup GetSlidesToBeDeleted()
        {
            var activePresentations = Globals.ThisAddIn.Application.ActivePresentation ?? throw new InvalidOperationException("No active presentation.");

            var allSlides = activePresentations.Slides;

            // collect all masters and all custom layouts across designs
            IReadOnlyList<Master> lstAvailableMaster = activePresentations.Designs
                .Cast<Design>()
                .Select(d => d.SlideMaster)
                .Where(m => m != null)
                .ToList();

            IReadOnlyList<CustomLayout> lstAvailableCustomLayout = lstAvailableMaster
                .SelectMany(m => m.CustomLayouts.Cast<CustomLayout>())
                .Where(cl => cl != null)
                .ToList();

            // collect custom layouts currently used by slides
            IReadOnlyList<CustomLayout> lstUsedCustomLayouts = allSlides.Cast<Slide>()
                                                                        .Select(s => s.CustomLayout)
                                                                        .Where(cl => cl != null)
                                                                        .ToList();

            // layouts not referenced by any slide
            IReadOnlyList<CustomLayout> unUsedlayouts = lstAvailableCustomLayout
                .Where(p => !lstUsedCustomLayouts.Contains(p))
                .ToList();

            // masters where ALL of their custom layouts are unused (i.e., master can be removed)
            IReadOnlyList<Master> unUsedMaster = lstAvailableMaster
                .Where(m => !m.CustomLayouts.Cast<CustomLayout>()
                .Except(unUsedlayouts)
                .Any())
                .ToList();

            return new SlideMasterCleanup
            {
                UnUsedlayouts = unUsedlayouts,
                UnUsedMaster = unUsedMaster
            };
        }

        /// <summary>
        /// Deletes the unused slides and saves the changes. The changes are permanent and cannot be reverted as initiating the file save operation in this method.
        /// </summary>
        /// <param name="slideMasterCleanup">Details of all the slides to be deleted.</param>
        private void DeleteSlides(ISlideMasterCleanup slideMasterCleanup)
        {
            IReadOnlyList<CustomLayout> unUsedlayouts = slideMasterCleanup.UnUsedlayouts ?? new List<CustomLayout>();
            IReadOnlyList<Master> unUsedMaster = slideMasterCleanup.UnUsedMaster ?? new List<Master>();

            if (IsNothingToDelete(unUsedlayouts, unUsedMaster))
            {
                return;
            }

            double initialSize = GetFileSize();
            int deleteSlideCount = 0;

            try
            {
                // Delete unused custom layouts
                foreach (CustomLayout layout in unUsedlayouts)
                {
                    try
                    {
                        deleteSlideCount++;
                        layout.Delete();
                    }
                    catch (COMException) { }
                    finally
                    {
                        ReleaseComObjectSafe(layout);
                    }
                }

                // Delete unused masters
                foreach (Master master in unUsedMaster)
                {
                    try
                    {
                        master.Delete();
                    }
                    catch (COMException) { }
                    finally
                    {
                        ReleaseComObjectSafe(master);
                    }
                }


                // Save the presentation after modifications    
                Globals.ThisAddIn.Application.ActivePresentation.Save();
            }
            finally
            {
                // attempt to release many common COM objects to reduce memory and COM reference leaks
                ReleaseComObjectSafe(Globals.ThisAddIn.Application.ActivePresentation.Slides);
                ReleaseComObjectSafe(Globals.ThisAddIn.Application.ActivePresentation.Designs);
                ReleaseComObjectSafe(Globals.ThisAddIn.Application.ActivePresentation.SlideMaster);
            }

            double finalSize = GetFileSize();

            MessageBox.Show($"Done {Environment.NewLine}Previous size: {initialSize}MB{Environment.NewLine}New Size: {finalSize}MB{Environment.NewLine}Total slides deleted: {deleteSlideCount}", "Slide Master Delete Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// Checks if there are no slides to be deleted and all the slide master slides are indeed in use. Displays a warning message. Can be removed before going to PROD
        /// </summary>
        /// <param name="unUsedlayouts">A collection of all the unused custom layouts.</param>
        /// <param name="unUsedMaster">A collection of all the unused master layouts</param>
        /// <returns></returns>
        private bool IsNothingToDelete(IReadOnlyList<CustomLayout> unUsedlayouts, IReadOnlyList<Master> unUsedMaster)
        {
            if (!unUsedlayouts.Any() && !unUsedMaster.Any())
            {
                MessageBox.Show("No unused slide master layouts or masters found to delete.", "Slide Master Cleanup", MessageBoxButtons.OK, MessageBoxIcon.Information);

                return true;
            }

            return false;
        }

        /// <summary>
        /// Optional method to gauge file size before and after the cleanup. Can be removed before going to PROD
        /// </summary>
        /// <returns>The current file size in MB at the time of the method invocation.</returns>
        private double GetFileSize()
        {
            string filePath = Globals.ThisAddIn.Application.ActivePresentation?.FullName;

            if (File.Exists(filePath))
            {
                FileInfo fileInfo = new FileInfo(filePath);
                return fileInfo.Length / (1024 * 1024);
            }

            return 0;
        }

        /// <summary>
        /// Releases a COM object safely. Swallows exceptions - best effort to avoid leaking COM refs.
        /// </summary>
        /// <param name="comObj">COM object required to be released</param>
        private void ReleaseComObjectSafe(object comObj)
        {
            if (comObj == null) return;

            try
            {
                while (Marshal.ReleaseComObject(comObj) > 0) { }
            }
            catch
            {
                // ignore - we tried our best
            }
        }
    }
}
