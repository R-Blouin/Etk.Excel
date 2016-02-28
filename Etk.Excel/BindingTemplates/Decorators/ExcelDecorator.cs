//namespace Etk.Excel.BindingTemplates.Decorators
//{
//    using System.Reflection;
//    using Etk.BindingTemplates.Definitions.Decorators;
//    using Microsoft.Office.Interop.Excel;
//    using Etk.BindingTemplates.Context;

//    class ExcelDecorator : Decorator
//    {
//        private bool addConcernedRangeParameter;

//        public ExcelDecorator(string ident, string description, MethodInfo toInvoke)
//                             : base(ident, description, toInvoke)    
//        {
        

//        }

//        /// <summary> Invoke the decorator</summary>
//        /// <param name="sender">Range that ask for a decoration</param>
//        /// <param name="contextItem">Binding bindingContextPart of the decoration request</param>
//        /// <param name="canDeleteComment">True if the previous Comment can be deleted</param>
//        /// <returns>True if  the décorator result</returns>
//        public override bool Invoke(object sender, IBindingContextItem contextItem)
//        {
//            Range concernedRange = sender as Range;
//            if (concernedRange == null)
//                return false;

//            // We delete the previous concernedRange comment 
//            Comment comment = concernedRange.Comment;
//            if (comment != null)
//                comment.Delete();

//            object[] parameters;

//            if (addConcernedRangeParameter)
//                parameters = new object[] { concernedRange, contextItem.DataSource, contextItem.BindingDefinition.Name };
//            else
//                parameters = new object[] { contextItem.DataSource, contextItem.BindingDefinition.Name };
//            object result = Callback.Invoke(Callback.IsStatic ? null : contextItem.DataSource, parameters);

//            if (result != null)
//            {
//                return true;
//            }
//            return false;
//        }
//    }
//}
