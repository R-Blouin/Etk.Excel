using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Etk.Excel;

namespace Etk.Tests.Templates.ExcelDna1.BasicExcelComTests
{
    class TestClassWithARange : IDisposable
    {
        private ExcelInterop.Range testRange;

        public void AffectAndUSeRange(ExcelInterop.Range range)
        {
            testRange = range;
            testRange.Value2 = "ClassPropery range";
        }

        public void Dispose()
        {
            if(testRange != null)
                BasicExcelComTestsManager.ReleaseComObject(testRange);
        }
    }

    class BasicExcelComTestsManager
    {
        const string TEST_SHEET_NAME = "ComTests";

        static public int ReleaseComObject(object obj)
        {
            int refCpt = Marshal.ReleaseComObject(obj);
            Trace.WriteLine($"Marshal cpt: {refCpt}");
            if (refCpt < 0)
                throw new Exception("ReleaseComObject returns counter < 1");
            return refCpt;
        }

        public void Execute()
        {
            ExcelInterop.Worksheet sheet = ETKExcel.ExcelApplication.GetWorkSheetFromName(null, TEST_SHEET_NAME);
            sheet.Activate();
            ReleaseComObject(sheet);


            PassARangeAsParameterUseItAndReleaseIt();
            GetARangeFromARangeWithResizingUseItThenReleaseIt();
            RetrieveTheFirstRangeOfARangeUseItAndReleaseIt();
            PassARangeAsParameterAffectItUseBothAndReleaseBoth();
            GetARangeFromAsheetReaffectItFromItselfThenReleaseIt();
            PassARangeAsParameterAffectItResizeItUseItAndReleaseIt();
            GetARangeFromARangeUseItAndReleaseIt();
            MergeRangeUseItUnmergedItReleaseIt();
            InitialyzeAClassProperyWithARangeUseThemReleaseThem();
            AddCommentToRangeThenDeleteIt();
            SelectARangeRetrieveSelectedRangeUseItThenReleaseIt();

            RetrieveWorksheetFromARangeThenReleaseIt();
            GetAWorkbookFromASheetUseItThenReleaseIt();
            RetrieveBordersFromARangeThenReleaseIt();
            Test2();
        }

        #region PassARangeAsParameterUseItAndReleaseIt
        void PassARangeAsParameterUseItAndReleaseIt()
        {
            try
            {
                ExcelInterop.Worksheet sheet = ETKExcel.ExcelApplication.GetWorkSheetFromName(null, TEST_SHEET_NAME);
                ExcelInterop.Range range = sheet.Range["A1"];

                PassARangeAsParameterUseItAndReleaseIt_(range);
                int cpt = ReleaseComObject(range);
                if (cpt != 0)
                    throw new Exception("range com counter != 0");

                ReleaseComObject(sheet);
            }
            catch (Exception ex)
            {
                throw new Exception("'PassARangeAsParameterUseItAndReleaseIt' Failed: {ex.Message}", ex);
            }
        }
        void PassARangeAsParameterUseItAndReleaseIt_(ExcelInterop.Range range)
        {
            range.Value2 = "Yo !";
        }
        #endregion

        void GetARangeFromARangeWithResizingUseItThenReleaseIt()
        {
            try
            {
                ExcelInterop.Worksheet sheet = ETKExcel.ExcelApplication.GetWorkSheetFromName(null, TEST_SHEET_NAME);
                ExcelInterop.Range range = sheet.Range["A1"];
                range = range[1, 1];

                ExcelInterop.Range range2 = range.Resize[2, 2];
                range2.Value2 = "Yo";

                if (ReleaseComObject(range2) != 0)
                    throw new Exception("range2 com counter != 0");

                if (ReleaseComObject(range) != 0)
                    throw new Exception("range com counter != 0");

                ReleaseComObject(sheet);
            }
            catch (Exception ex)
            {
                throw new Exception("'GetARangeFromARangeWithResizingUseItThenReleaseIt' Failed: {ex.Message}", ex);
            }
        }

        void RetrieveTheFirstRangeOfARangeUseItAndReleaseIt()
        {
            try
            {
                ExcelInterop.Worksheet sheet = ETKExcel.ExcelApplication.GetWorkSheetFromName(null, TEST_SHEET_NAME);
                ExcelInterop.Range range = sheet.Range["A1"];

                ExcelInterop.Range range2 = range[1, 1];
                range2.Value2 = "Yo";

                if (ReleaseComObject(range2) != 0)
                    throw new Exception("range2 com counter != 0");

                if (ReleaseComObject(range) != 0)
                    throw new Exception("range com counter != 0");

                ReleaseComObject(sheet);
            }
            catch (Exception ex)
            {
                throw new Exception("'PassARangeAsParameterUseItAndReleaseIt' Failed: {ex.Message}", ex);
            }
        }

        #region PassARangeAsParameterAffectItUseBothItAndReleaseBoth
        void PassARangeAsParameterAffectItUseBothAndReleaseBoth()
        {
            try
            {
                ExcelInterop.Worksheet sheet = ETKExcel.ExcelApplication.GetWorkSheetFromName(null, TEST_SHEET_NAME);
                ExcelInterop.Range range = sheet.Range["A1"];

                PassARangeAsParameterAffectItUseBothItAndReleaseBoth_(range);
                try
                {
                    ReleaseComObject(range);
                    throw new Exception("'PassARangeAsParameterAffectItUseBothItAndReleaseBoth' Failed: 'range' should have been already released");
                }
                catch
                { }

                ReleaseComObject(sheet);
            }
            catch (Exception ex)
            {
                throw new Exception("'PassARangeAsParameterAffectItUseBothItAndReleaseBoth' Failed: {ex.Message}", ex);
            }
        }

        void PassARangeAsParameterAffectItUseBothItAndReleaseBoth_(ExcelInterop.Range range)
        {
            ExcelInterop.Range refRange = range;
            range.Value2 = "Yo !";
            refRange.Value2 = "Yo ! Yo !";

            if(ReleaseComObject(refRange) != 0)
                throw new Exception("refRange com counter != 0");
        }
        #endregion

        void RetrieveWorksheetFromARangeThenReleaseIt()
        {
            try
            {
                ExcelInterop.Worksheet sheet = ETKExcel.ExcelApplication.GetWorkSheetFromName(null, TEST_SHEET_NAME);
                ExcelInterop.Range range = sheet.Range["A1"];

                ExcelInterop.Worksheet worksheet = range.Worksheet;

                ReleaseComObject(range);
                int cpt = ReleaseComObject(worksheet);
                cpt = ReleaseComObject(sheet);
            }
            catch(Exception ex)
            {
                throw new Exception("'RetrieveWorksheetFromARangeThenReleaseIt' Failed: {ex.Message}", ex);
            }
        }

        void RetrieveBordersFromARangeThenReleaseIt()
        {
            try
            {
                ExcelInterop.Worksheet sheet = ETKExcel.ExcelApplication.GetWorkSheetFromName(null, TEST_SHEET_NAME);
                ExcelInterop.Range range = sheet.Range["A1"];

                ExcelInterop.Borders borders = range.Borders;
                int cpt = ReleaseComObject(borders);
                if (cpt != 0)
                    throw new Exception("borders com counter != 0");
                cpt = ReleaseComObject(range);
                if (cpt != 0)
                    throw new Exception("range com counter != 0");
                ReleaseComObject(sheet);
            }
            catch (Exception ex)
            {
                throw new Exception("'RetrieveWorksheetFromARangeThenReleaseIt' Failed: {ex.Message}", ex);
            }
        }

        void GetARangeFromAsheetReaffectItFromItselfThenReleaseIt()
        {
            try
            {
                ExcelInterop.Worksheet sheet = ETKExcel.ExcelApplication.GetWorkSheetFromName(null, TEST_SHEET_NAME);
                ExcelInterop.Range range = sheet.Range["A1"];

                range = range.Offset[Type.Missing, 1];
                range.Value2 = "Yo !";
                range = range.Offset[Type.Missing, 2];
                range.Value2 = "Yo !";

                int cpt = ReleaseComObject(range);
                if (cpt != 0)
                    throw new Exception("'range' com counter != 0");
                ReleaseComObject(sheet);
            }
            catch (Exception ex)
            {
                throw new Exception("'GetARAngeFromAsheetReaffectItFromItselfThenReleaseIt' Failed: {ex.Message}", ex);
            }
        }

        #region PassARangeAsParameterAffectItResizeItUseItAndReleaseIt
        void PassARangeAsParameterAffectItResizeItUseItAndReleaseIt()
        {
            try
            {
                ExcelInterop.Worksheet sheet = ETKExcel.ExcelApplication.GetWorkSheetFromName(null, TEST_SHEET_NAME);
                ExcelInterop.Range range = sheet.Range["A1"];

                PassARangeAsParameterAffectItResizeItUseItAndReleaseIt_(range, 1);
                PassARangeAsParameterAffectItResizeItUseItAndReleaseIt_(range, 2);

                int cpt = ReleaseComObject(range);
                if (cpt != 0)
                    throw new Exception("'range' com counter != 0");
                ReleaseComObject(sheet);
            }
            catch (Exception ex)
            {
                throw new Exception("'PassARangeAsParameterAffectItResizeItUseItAndReleaseIt' Failed: {ex.Message}", ex);
            }
        }

        void PassARangeAsParameterAffectItResizeItUseItAndReleaseIt_(ExcelInterop.Range range, int numberOfColumns)
        {
            ExcelInterop.Range workingRange;
            if (numberOfColumns < 0)
                workingRange = range.Offset[Type.Missing, numberOfColumns];
            else
                workingRange = range.Offset[Type.Missing, 1];

            workingRange = workingRange.Resize[Type.Missing, 2];

            workingRange.Value2 = "Yo !";

            int cpt = ReleaseComObject(workingRange);
            if (cpt != 0)
                throw new Exception("'workingRange' com counter != 0");
        }
        #endregion

        void GetARangeFromARangeUseItAndReleaseIt()
        {
            try
            {
                ExcelInterop.Worksheet sheet = ETKExcel.ExcelApplication.GetWorkSheetFromName(null, TEST_SHEET_NAME);
                ExcelInterop.Range range = sheet.Range["A1"];

                ExcelInterop.Range columns = range.EntireColumn;
                object obj = columns.Value2;

                ExcelInterop.Range cell = range[1, 1];
                cell.Value2 = "Cell";

                if (ReleaseComObject(columns) != 0)
                    throw new Exception("'columns' com counter != 0");

                if (ReleaseComObject(cell) != 0)
                    throw new Exception("'cell' com counter != 0");

                if (ReleaseComObject(range) != 0)
                    throw new Exception("'range' com counter != 0");

                ReleaseComObject(sheet);
            }
            catch (Exception ex)
            {
                throw new Exception("'GetARangeFromARAngeUseItAndReleaseIt' Failed: {ex.Message}", ex);
            }
        }

        void MergeRangeUseItUnmergedItReleaseIt()
        {
            try
            {
                ExcelInterop.Worksheet sheet = ETKExcel.ExcelApplication.GetWorkSheetFromName(null, TEST_SHEET_NAME);
                ExcelInterop.Range range = sheet.Range["E1"];

                range.Value2 = "Yo !";
                ExcelInterop.Range toMerge = range.Resize[2, 2];
                toMerge.Merge();
                toMerge.UnMerge();

                int cpt = ReleaseComObject(toMerge);
                if (cpt != 0)
                    throw new Exception("'toMerge' com counter != 0");

                cpt = ReleaseComObject(range);
                if (cpt != 0)
                    throw new Exception("'range' com counter != 0");

                ReleaseComObject(sheet);
            }
            catch (Exception ex)
            {
                throw new Exception("'MergeRangeUseItUnmergedItReleaseIt' Failed: {ex.Message}", ex);
            }
        }

        void InitialyzeAClassProperyWithARangeUseThemReleaseThem()
        {
            try
            {
                ExcelInterop.Worksheet sheet = ETKExcel.ExcelApplication.GetWorkSheetFromName(null, TEST_SHEET_NAME);
                ExcelInterop.Range range = sheet.Range["A1"];

                TestClassWithARange testClass = new TestClassWithARange();
                testClass.AffectAndUSeRange(range);
                range.Value2 = "Yo !";

                testClass.Dispose();
                try
                {
                    int cpt = ReleaseComObject(range);
                    throw new Exception("'InitialyzeAClassProperyWithARangeUseThemReleaseThem' Failed: 'range' should have been already released");
                }
                catch
                { }

                ReleaseComObject(sheet);
            }
            catch (Exception ex)
            {
                throw new Exception("'InitialyzeAClassProperyWithARangeUseThemReleaseThem' Failed: {ex.Message}", ex);
            }
        }

        void AddCommentToRangeThenDeleteIt()
        {
            try
            {
                ExcelInterop.Worksheet sheet = ETKExcel.ExcelApplication.GetWorkSheetFromName(null, TEST_SHEET_NAME);
                ExcelInterop.Range range = sheet.Range["A1"];

                range.AddComment("It's a comment !");
                ExcelInterop.Comment addedComment = range.Comment;
                addedComment.Visible = true;
                ExcelInterop.Shape shape = addedComment.Shape;
                ExcelInterop.TextFrame textFrame = shape.TextFrame;
                textFrame.AutoSize = true;

                if (ReleaseComObject(textFrame) != 0)
                    throw new Exception("'textFrame' com counter != 0");
                if (ReleaseComObject(shape) != 0)
                    throw new Exception("'shape' com counter != 0");
                if (ReleaseComObject(addedComment) != 0)
                    throw new Exception("'addedComment' com counter != 0");

                ExcelInterop.Comment comment = range.Comment;
                comment.Delete();
                if (ReleaseComObject(comment) != 0)
                    throw new Exception("'comment' com counter != 0");

                if (ReleaseComObject(range) != 0)
                    throw new Exception("'range' com counter != 0");
                ReleaseComObject(sheet);
            }
            catch (Exception ex)
            {
                throw new Exception("'AddCommentToRangeThenDeleteIt' Failed: {ex.Message}", ex);
            }
        }

        void SelectARangeRetrieveSelectedRangeUseItThenReleaseIt()
        {
            try
            {
                ExcelInterop.Worksheet sheet = ETKExcel.ExcelApplication.GetWorkSheetFromName(null, TEST_SHEET_NAME);
                ExcelInterop.Range range = sheet.Range["A1"];

                range.Select();

                ExcelInterop.Range selectedRange = ETKExcel.ExcelApplication.Application.Selection as ExcelInterop.Range;
                selectedRange.Value2 = "Selected !!";

                if (ReleaseComObject(selectedRange) != 0)
                    throw new Exception("'selectedRange' com counter != 0");

                if (ReleaseComObject(range) != 0)
                    throw new Exception("'range' com counter != 0");
                ReleaseComObject(sheet);
            }
            catch (Exception ex)
            {
                throw new Exception("'SelectARangeRetrieveSelectedRangeUseItThenReleaseIt' Failed: {ex.Message}", ex);
            }
        }

        #region GetAWorkbookFromASheetUseItThenReleaseIt
        void GetAWorkbookFromASheetUseItThenReleaseIt()
        {
            try
            {
                ExcelInterop.Worksheet sheet = ETKExcel.ExcelApplication.GetWorkSheetFromName(null, TEST_SHEET_NAME);

                ExcelInterop.Workbook book = sheet.Parent as ExcelInterop.Workbook;
                if (book != null)
                {
                    book.SheetCalculate += GetAWorkbookFromASheetUseItThenReleaseItOnSheetActivation;
                    book.SheetCalculate -= GetAWorkbookFromASheetUseItThenReleaseItOnSheetActivation;
                    if (ReleaseComObject(book) < 0)
                        throw new Exception("'book' com counter < 0");
                }
                ReleaseComObject(sheet);
            }
            catch (Exception ex)
            {
                throw new Exception("'GetAWorkbookFromASheetUseItThenReleaseIt' Failed: {ex.Message}", ex);
            }
        }

        void GetAWorkbookFromASheetUseItThenReleaseItOnSheetActivation(object sheet)
        { }
        #endregion

        void Test2()
        {
            ExcelInterop.Workbook workbook = null;
            ExcelInterop.Sheets sheets = null;
            ExcelInterop.Worksheet lastSheet = null;
            ExcelInterop.Worksheet firstSheet = null;
            try
            {
                workbook = ETKExcel.ExcelApplication.Application.ActiveWorkbook;
                sheets = workbook.Sheets;
                firstSheet = workbook.Sheets[1];
                lastSheet = workbook.Sheets[sheets.Count];
            }
            finally
            {
                if (firstSheet != null)
                    ReleaseComObject(firstSheet);
                if (lastSheet != null)
                    ReleaseComObject(lastSheet);
                if (sheets != null)
                    ReleaseComObject(sheets);
                if (workbook != null)
                    ReleaseComObject(workbook);
            }
        }
    }
}
