namespace Etk.BindingTemplates.Definitions.Decorators
{
    /// <summary>The result of an invocation of a decorator resolver method</summary>
    public class DecoratorResult
    {
        /// <summary>The decorator item index to use to change the object to modify (for example the range in the Excel case)</summary>
        public int? Item
        { get; set; }

        /// <summary>The comment to apply to the object to modify (for example the range in the Excel case)</summary>
        public string Comment
        { get; set; }

        /// <summary>Construct a Decorator result</summary>
        public DecoratorResult(int? item, string comment)
        {
            this.Item = item;
            this.Comment = comment;
        }
    }
}
