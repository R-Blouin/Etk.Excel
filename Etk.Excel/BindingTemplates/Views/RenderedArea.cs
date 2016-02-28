namespace Etk.Excel.BindingTemplates.Views
{
    /// <summary>Rendered aread of a view</summary>
    public class RenderedArea
    {
        /// <summary>First rendered column</summary>
        public int XPos
        { get; private set; }

        /// <summary>First rendered col</summary>
        public int YPos
        { get; private set; }

        /// <summary>Number of columns rendered</summary>
        public int Width
        { get; private set; }

        /// <summary>Number of rows rendered</summary>
        public int Height
        { get; private set; }

        public RenderedArea(int xPos, int yPos, int width, int height)
        {
            XPos = xPos;
            YPos = yPos;
            Width = width;
            Height = height;
        }
    }
}
