//namespace Etk.BindingTemplates.Definitions.Binding
//{
//    using System.Collections.Generic;
//    using System.Collections.ObjectModel;
//    using Etk.BindingTemplates.Context;

//    class BindingDefinitionKeyWord : BindingDefinition
//    { 
//        public static readonly ReadOnlyCollection<string> KeyWords = new ReadOnlyCollection<string>(new List<string>(new string[] { "_Seq_" }));

//        override public object UpdateDataSource(IBindingContextItem contextItem, object dataSource, object data)
//        {           
//            return null; 
//        }

//        override public object ResolveBinding(IBindingContextItem contextItem, object dataSource)
//        { 
//            string ret;
//            switch(BindingExpression)
//            {
//                case "_Seq_":
//                    ret = contextItem == null ? string.Empty : contextItem.ParentElement.Index.ToString();
//                break;
//                default:
//                    ret = string.Empty;
//                break;
//            }
//            return ret; 
//        }

//        #region static public methods
//        static public BindingDefinitionKeyWord CreateInstances(string keyWord)
//        {
//            if(!KeyWords.Contains(keyWord))
//                throw new BindingTemplateException(string.Format("'{0}' is not a keyword.", keyWord ?? string.Empty));

//            BindingDefinitionDescription definitionDescription = new BindingDefinitionDescription() { BindingExpression = keyWord }; 
//            BindingDefinitionKeyWord bindingDefinition = new BindingDefinitionKeyWord() { DefinitionDescription = definitionDescription,
//                                                                                          BindingType = typeof(string),
//                                                                                          IsBoundWithData = false};
//            return bindingDefinition;
//        }
//        #endregion
//    }
//}
