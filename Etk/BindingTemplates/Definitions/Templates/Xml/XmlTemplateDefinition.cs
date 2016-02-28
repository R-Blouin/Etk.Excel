namespace Etk.BindingTemplates.Definitions.Templates.Xml
{
    using System;
    using System.Xml.Serialization;

    public class XmlTemplateOption
    {
        [XmlAttribute]
        public string Name
        { get; set; }

        [XmlAttribute]
        public string Description
        { get; set; }

        [XmlAttribute]
        public string BindingWith
        { get; set; }

        /// <summary> Only valid if 'DataAccessor' is settled. Can be 'Static, Singleton'</summary>
        [XmlAttribute]
        public string InstanceType
        { get; set; }

        /// <summary> For singleton only </summary>
        [XmlAttribute]
        public string InstanceName
        { get; set; }

        [XmlAttribute]
        public bool CanSort
        { get; set; }

        [XmlAttribute]
        public bool AddBorder
        { get; set; }

        [XmlAttribute("Dec")]
        public string Decorator
        { get; set; }

        [XmlIgnore]
        public Orientation Orientation
        { get; set; }

        private string orientationStr;
        [XmlAttribute("Orientation")]
        public string OrientationStr
        {
            get { return orientationStr; }
            set
            {
                if (!string.IsNullOrEmpty(value))
                {
                    orientationStr = value.Trim().ToUpper();
                    if (orientationStr.Equals("H") || orientationStr.Equals("HORIZONTAL"))
                        Orientation = Orientation.Horizontal;
                    else if (orientationStr.Equals("V") || orientationStr.Equals("VERTICAL"))
                        Orientation = Orientation.Vertical;
                    else
                        throw new ArgumentException(string.Format("The attribut 'Orientation' '{0}' is invalid. Value must be 'Vertical' (or 'V') or 'Horizontal' (or 'V') (no case sentitive)", value));
                }
            }
        }

        [XmlAttribute]
        public bool HeaderAsExpander
        { get; set; }

        [XmlAttribute]
        public string Expander
        { get; set; }

        [XmlIgnore]
        public ExpanderMode ExpanderMode
        { get; set; }

        private string expanderModeStr;
        [XmlAttribute("ExpanderMode")]
        public string ExpanderModeStr
        {
            get { return expanderModeStr; }
            set
            {
                if (!string.IsNullOrEmpty(value))
                {
                    expanderModeStr = value.Trim().ToUpper();
                    if (expanderModeStr.Equals("DONTRENDER") || expanderModeStr.Equals("DR"))
                        ExpanderMode = ExpanderMode.DontRender;
                    if (expanderModeStr.Equals("HIDE") || expanderModeStr.Equals("H"))
                        ExpanderMode = ExpanderMode.Hide;
                    else
                        throw new ArgumentException(string.Format("The attribut 'ExpanderMode' '{0}' is invalid. Value must be 'DontRender' (or 'DR') or 'Hide' (or 'H') (no case sentitive)", value));
                }
            }
        }

        [XmlAttribute]
        public string SelectionChanged
        { get; set; }

        [XmlAttribute]
        public string ContextMenu
        { get; set; }

        public XmlTemplateOption()
        {
            Orientation = Orientation.Vertical;
            ExpanderMode = ExpanderMode.DontRender;
            CanSort = true;
            HeaderAsExpander = true;
        }
    }
}
