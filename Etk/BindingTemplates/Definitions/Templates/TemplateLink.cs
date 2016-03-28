using Etk.BindingTemplates.Definitions.Templates.Xml;
using Etk.Tools.Extensions;

namespace Etk.BindingTemplates.Definitions.Templates
{
    public class TemplateLink
    {
        public string Name
        { get; private set; }
        
        public string Description
        { get; private set; }
        
        public string To
        { get; private set; }

        public string With
        { get; private set; }

        public LinkedTemplatePositioning Positioning
        { get; private set; }

        /// <summary> Method info to invoke to determinate the min nomber of occurences the link templates must occupied</summary>
        public string MinOccurencesMethod
        { get; private set; }

        #region .ctors and factories
        private TemplateLink()
        {}

        public static TemplateLink CreateInstance(XmlTemplateLink xmlTemplateLink)
        {
            if (xmlTemplateLink == null)
                return null;

            TemplateLink templateLink = new TemplateLink();

            templateLink.Name = xmlTemplateLink.Name;
            templateLink.Description = xmlTemplateLink.Description;

            templateLink.To = xmlTemplateLink.To.EmptyIfNull().Trim();
            templateLink.With = xmlTemplateLink.With.EmptyIfNull().Trim();
            templateLink.Positioning = xmlTemplateLink.Position;
            templateLink.MinOccurencesMethod = xmlTemplateLink.MinOccurences;

            if (string.IsNullOrEmpty(templateLink.To))
                throw new EtkException("Attribut 'To' cannot be null or empty", false);

            return templateLink;
        }
        #endregion
    }
}
