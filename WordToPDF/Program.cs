using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace WordToPDF
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length != 2)
            {
                Console.Out.WriteLine("Incorrect number opf parameters.  Expected: <source> <result>, was " + args.Length);
                return;
            }

            Application wordApplication = new Application();
            Document wordDocument = null;

            object paramSourceDocPath = Path.GetFullPath(args[0]);
            string paramExportFilePath = Path.GetFullPath(args[1]);

            WdExportFormat paramExportFormat = WdExportFormat.wdExportFormatPDF;
            bool paramOpenAfterExport = false;
            WdExportOptimizeFor paramExportOptimizeFor = WdExportOptimizeFor.wdExportOptimizeForPrint;
            WdExportRange paramExportRange = WdExportRange.wdExportAllDocument;

            int paramStartPage = 0;
            int paramEndPage = 0;

            WdExportItem paramExportItem = WdExportItem.wdExportDocumentContent;

            bool paramIncudeDocProps = true;
            bool paramKeepIRM = true;
            WdExportCreateBookmarks paramCreateBookmarks = WdExportCreateBookmarks.wdExportCreateWordBookmarks;

            bool paramDocStructureTags = true;
            bool paramBitmapMissingFonts = true;
            bool paramUseISO19005_1 = false;

            object paramMissing = Type.Missing;

            try
            {
                wordDocument = wordApplication.Documents.Open(
                    ref paramSourceDocPath, ref paramMissing, ref paramMissing,
                    ref paramMissing, ref paramMissing, ref paramMissing,
                    ref paramMissing, ref paramMissing, ref paramMissing,
                    ref paramMissing, ref paramMissing, ref paramMissing,
                    ref paramMissing, ref paramMissing, ref paramMissing,
                    ref paramMissing);

                if (wordDocument != null)
                {
                    wordDocument.ExportAsFixedFormat(paramExportFilePath, paramExportFormat,
                        paramOpenAfterExport, paramExportOptimizeFor, paramExportRange,
                        paramStartPage, paramEndPage, paramExportItem, paramIncudeDocProps,
                        paramKeepIRM, paramCreateBookmarks, paramDocStructureTags,
                        paramBitmapMissingFonts, paramUseISO19005_1, ref paramMissing);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("The Error recieved is {0}: ", ex.Message);
            }
            finally
            {
                if (wordDocument != null)
                {
                    wordDocument.Close(ref paramMissing, ref paramMissing, ref paramMissing);
                    wordDocument = null;
                }

                if (wordApplication != null)
                {
                    wordApplication.Quit(ref paramMissing, ref paramMissing, ref paramMissing);
                    wordApplication = null;

                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

        }
    }
}
