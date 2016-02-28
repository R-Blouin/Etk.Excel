//namespace Etk.BindingTemplates.Definitions
//{
//    using Etk.ModelManagement;

//    public class ModelViewTemplateDefinition : FilterOwner
//    {
//        #region .ctors and factories
//        private ModelViewTemplateDefinition(ModelView view, TemplateOption templateOption) : base(templateOption)
//        {
//            if(view.Parts != null)
//            {
//                foreach(IModelViewPart part in view.Parts)    
//                {
//                    if (part is IModelView)
//                    {
//                        this.AddLinkedTemplate(null);
//                    }
//                    else
//                    {
//                        this.AddBindingDefinition(null);
//                    }
//                }        
//            }
//        }

//        public ITemplateDefinition CreateInstances(IModelView view, TemplateOption templateOption)
//        {
//            ModelViewTemplateDefinition ret = null;
//            if (view != null)
//            { 
                
//            }
//            return ret;
//        }
//        #endregion
//    }
//}
