namespace Etk.Excel.BindingTemplates.Views
{
    using Microsoft.Office.Interop.Excel;

    public class RenderingArea
    {
        public int XFirstCell
        { get; private set;}

        public int YFirstCell
        { get; private set; }

        public int Width
        { get; private set; }

        public int Height
        { get; private set; }

        private RenderingArea(int xFirstCell, int yFirstCell, int width, int height)
        {
            XFirstCell = xFirstCell;
            YFirstCell = yFirstCell;
            Width = width;
            Height = height;
        }

        public static RenderingArea CreateInstance(Range range)
        {
            RenderingArea ret = null;
            if (range != null)
                ret = new RenderingArea(range.Column, range.Row, range.Columns.Count, range.Rows.Count);
            return ret;
        }
    }
}
