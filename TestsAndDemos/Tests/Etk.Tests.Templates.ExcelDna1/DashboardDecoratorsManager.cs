using System.Collections.Generic;
using Etk.BindingTemplates.Definitions.Decorators;
using Etk.Tests.Templates.ExcelDna1.Tests;

namespace Etk.Tests.Templates.ExcelDna1
{
    /// <summary>Manage the decorators declared in the templates used by the Dashboard </summary>
    static class DashboardDecoratorsManager
    {
        /// <summary>Decore the 'Status' property of the 'ExcelTestsManager' instance.
        /// The style of the result are defined by the named range 'DEC_STATUS'</summary>
        /// <param name="testsManager">The object that requests a decoration</param>
        /// <param name="bindingName">The name of the binding definition that requests a decoration</param>
        /// <returns>The decorator result</returns>
        static public DecoratorResult DecorateStatus(ExcelTestsManager testsManager, string bindingName)
        {
            if(string.IsNullOrEmpty(testsManager.Status))
                return new DecoratorResult(0, null);
            return new DecoratorResult(1, null);
        }

        /// <summary>Decore the 'InitSuccessful' property of a 'IExcelTestTopic' instance.
        /// The style of the result are defined by the range "'Dashboard Templates'!B18:B19"</summary>
        /// <param name="concernedTestTopic">The test topic that requests a decoration</param>
        /// <param name="bindingName">The name of the binding definition that requests a decoration</param>
        /// <returns>The decorator result</returns>
        static public DecoratorResult DecorateInitSuccessful(IExcelTestTopic concernedTestTopic, string bindingName)
        {
            if (concernedTestTopic.InitSuccessful)
                return new DecoratorResult(0, null);
            return new DecoratorResult(1, null);
        }

        /// <summary>Decore the 'Done' property of a 'IExcelTest' instance.
        /// The style of the result are defined by the named range 'DEC_DONE'</summary>
        /// <param name="concernedTest">The test that requests a decoration</param>
        /// <param name="bindingName">The name of the binding definition that requests a decoration</param>
        /// <returns>The decorator result</returns>
        static public DecoratorResult DecorateDone(IExcelTest concernedTest, string bindingName)
        {
            if (concernedTest.Done)
                return new DecoratorResult(0, null);
            return new DecoratorResult(1, null);
        }

        /// <summary>Decore the 'Success' property of a 'IExcelTest' instance.
        /// The style of the result are defined by the named range 'DEC_EXEC'</summary>
        /// <param name="concernedTest">The test that requests a decoration</param>
        /// <param name="bindingName">The name of the binding definition that requests a decoration</param>
        /// <returns>The decorator result</returns>
        static public DecoratorResult DecorateSuccess(IExcelTest concernedTest, string bindingName)
        {
            if (concernedTest.Success)
                return new DecoratorResult(0, null);
            return new DecoratorResult(1, null);
        }
    }
}
