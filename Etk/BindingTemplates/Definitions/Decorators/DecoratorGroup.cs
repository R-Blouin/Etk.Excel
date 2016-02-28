//namespace Etk.BindingTemplates.Definitions.Decorators
//{
//    using System.Collections.Generic;
//    using System.Linq;

//    class DecoratorGroup : Decorator
//    {
//        #region attributes and properties
//        public IEnumerable<Decorator> Decorators
//        { get; private set; }
//        #endregion

//        #region attributes and properties
//        public DecoratorGroup(string ident, string description, IEnumerable<Decorator> decorators) : base(ident, description, null)
//        { 
//            Decorators = decorators;
//            if (Decorators != null)
//                Decorators = Decorators.Where(d => d != null);
//        }
//        #endregion


//        public override bool Resolve(object sender, Context.IBindingContextItem contextItem)
//        {
//            bool ret = false;
//            if (Decorators != null)
//            {
//                foreach (Decorator decorator in Decorators)
//                {
//                    if (decorator.Resolve(sender, contextItem))
//                    {
//                        ret = true;
//                        break;
//                    }
//                }
//            }
//            return ret;
//        }
//    }
//}
