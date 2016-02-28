namespace Etk.Excel.UI.Emit
{
    using System;

    public class EmitProperty
    {
        public string PropertyName 
        { get; private set; }
        
        public Type PropertyType
        { get; private set; }

        public EmitProperty(string propertyName, Type type)
        {
            PropertyName = propertyName;
            PropertyType = type;
        }
    }
}
