using System;
using System.Xml.Serialization;

namespace Etk.BindingTemplates.Definitions.Templates.Xml
{
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

        [XmlIgnore]
        public bool? CanSort
        { get; private set; }

        private bool canSortSet;
        [XmlAttribute("CanSort")]
        public bool CanSortSet
        {
            get { return canSortSet; }
            set { CanSort = canSortSet = value; }
        }

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
                        throw new ArgumentException($"The attribut 'Orientation' '{value}' is invalid. Value must be 'Vertical' (or 'V') or 'Horizontal' (or 'V') (no case sentitive)");
                }
            }
        }

        [XmlIgnore]
        public ExpanderType ExpanderType
        { get; set; }

        private string expanderTypeStr;
        [XmlAttribute("ExpanderType")]
        public string ExpanderTypeStr
        {
            get { return expanderTypeStr; }
            set
            {
                if (!string.IsNullOrEmpty(value))
                {
                    expanderTypeStr = value.Trim().ToUpper();
                    if (expanderTypeStr.Equals("DONTRENDER") || expanderTypeStr.Equals("DR"))
                        ExpanderType = ExpanderType.DontRender;
                    if (expanderTypeStr.Equals("HIDE") || expanderTypeStr.Equals("H"))
                        ExpanderType = ExpanderType.Hide;
                    else
                        throw new ArgumentException($"The attribut 'ExpanderType' '{value}' is invalid. Value must be 'DontRender' (or 'DR') or 'Hide' (or 'H') (no case sentitive)");
                }
            }
        }

        [XmlIgnore]
        public HeaderAsExpander HeaderAsExpander
        { get; set; }

        private string headerAsExpanderStr;
        [XmlAttribute("HeaderAsExpander")]
        public string HeaderAsExpanderStr
        {
            get { return headerAsExpanderStr; }
            set
            {
                if (!string.IsNullOrEmpty(value))
                {
                    headerAsExpanderStr = value.Trim().ToUpper();
                    if (headerAsExpanderStr.Equals("STARTEXPANDED") || headerAsExpanderStr.Equals("SE"))
                        HeaderAsExpander = HeaderAsExpander.StartExpanded;
                    else if (headerAsExpanderStr.Equals("STARTCLOSED") || headerAsExpanderStr.Equals("SC"))
                        HeaderAsExpander = HeaderAsExpander.StartClosed;
                    else
                        throw new ArgumentException($"The attribut 'HeaderAsExpander' '{value}' is invalid. Value must be 'StartExpanded' (or 'SE') or 'StartClosed' (or 'SC') (no case sentitive)");
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
            ExpanderType = ExpanderType.Hide;
            HeaderAsExpander = HeaderAsExpander.None;
            //AddBorder = true;
        }
    }
}
