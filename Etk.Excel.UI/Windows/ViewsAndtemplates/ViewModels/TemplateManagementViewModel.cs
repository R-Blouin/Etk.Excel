using System.ComponentModel;
using Etk.BindingTemplates.Definitions.Templates;
using Etk.Excel.BindingTemplates.Definitions;
using Etk.Excel.MvvmBase;

namespace Etk.Excel.UI.Windows.ViewsAndtemplates.ViewModels
{
    public class TemplateManagementViewModel : ViewModelBase
    {
        #region attributes and properties
        private ExcelTemplateDefinition templateDefinition;

        /// <summary>Template Name</summary>
        public string Name
        {
            get {return templateDefinition.TemplateOption.Name;} 
            set 
            {
                templateDefinition.TemplateOption.Name = value;
                OnPropertyChanged("Name");
            }
        }

        /// <summary>Template Description</summary>
        public string Description
        {
            get { return templateDefinition.TemplateOption.Description; }
            set
            {
                templateDefinition.TemplateOption.Description = value;
                OnPropertyChanged("Description");
            }
        }

        /// <summary>Template Orientation</summary>
        public Orientation Orientation
        {
            get { return templateDefinition.TemplateOption.Orientation; }
            set
            {
                templateDefinition.TemplateOption.Orientation = value;
                OnPropertyChanged("Orientation");
            }
        }

        private bool mainBindingDefinitionFromDataAccessor;
        public bool MainBindingDefinitionFromDataAccessor
        {
            get { return mainBindingDefinitionFromDataAccessor; } 
            set 
            {
                mainBindingDefinitionFromDataAccessor = value;
                OnPropertyChanged("MainBindingDefinitionFromDataAccessor");
            } 
        }

        private string dataAccessorString;
        /// <summary>Template Description</summary>
        public string DataAccessorString
        {
            get { return dataAccessorString; }
            set
            {
                //partToRenderDefinition.TemplateOption.DataAccessor = value;
                OnPropertyChanged("DataAccessorString");
            }
        }

        private string mainBindingDefinitionString;
        /// <summary>Template Description</summary>
        public string MainBindingDefinitionString
        {
            get { return mainBindingDefinitionString; }
            set
            {
                //partToRenderDefinition.TemplateOption.DataAccessor = value;
                OnPropertyChanged("MainBindingDefinitionString");
            }
        }

        /// <summary> Determine if a header is an expander</summary>
        public HeaderAsExpander HeaderAsExpander
        {
            get { return templateDefinition.TemplateOption.HeaderAsExpander; }
            set
            {
                templateDefinition.TemplateOption.HeaderAsExpander = value;
                OnPropertyChanged("HeaderAsExpander");
            }
        }

        /// <summary>Template Header Expadner mode</summary>
        public ExpanderType ExpanderMode
        {
            get { return templateDefinition.TemplateOption.ExpanderType; }
            set
            {
                templateDefinition.TemplateOption.ExpanderType = value;
                OnPropertyChanged("ExpanderMode");
            }
        }

        private string expanderBindingDefinitionString;
        /// <summary>Template Header Expadner mode</summary>
        public string ExpanderBindingDefinitionString
        {
            get { return expanderBindingDefinitionString; }
            set
            {
                expanderBindingDefinitionString = value;
                OnPropertyChanged("ExpanderBindingDefinitionString");
            }
        }
        #endregion

        #region .ctors
        public TemplateManagementViewModel(ExcelTemplateDefinition templateDefinition)
        {
            this.templateDefinition = null;
            dataAccessorString = null;
            mainBindingDefinitionString = null;

            //if (templateDefinition == null)
            //    templateDefinition = new ExcelTemplateDefinition(new TemplateOption());

            ////if (partToRenderDefinition.TemplateOption.DataAccessor == null && partToRenderDefinition.TemplateOption.MainBindingDefinition == null)
            ////    TypeToBindWithType = TypeToBindWithMode.None;
            ////else
            ////    TypeToBindWithType = partToRenderDefinition.TemplateOption.DataAccessor == null ? TypeToBindWithMode.Type : TypeToBindWithMode.Accessor;

            //this.templateDefinition = templateDefinition;
            //this.PropertyChanged += OnPropertyChanged;
        }
        #endregion

        #region private methods
        private void OnPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (sender != this)
                return;

            switch(e.PropertyName)
            {
                case "MainBindingDefinitionFromDataAccessor":
                    ManageMainBindingDefinitionFromDataAccessorChanged();
                break;
                case "TypeToBindWithMode":
                    ManageTypeToBindWithModeChanged();
                break;
                case "ExpanderBindingDefinitionString":
                case "HeaderAsExpander":
                    ManageExpanderBindingDefinitionChanged();
                break;
            }
        }

        private void ManageMainBindingDefinitionFromDataAccessorChanged()
        { 
            if (! mainBindingDefinitionFromDataAccessor)
                templateDefinition.TemplateOption.DataAccessor = null;
        }

        private void ManageTypeToBindWithModeChanged()
        {
            if (!mainBindingDefinitionFromDataAccessor)
                templateDefinition.TemplateOption.DataAccessor = null;
        }

        private void ManageExpanderBindingDefinitionChanged()
        { 
        }
        #endregion
    }
}
