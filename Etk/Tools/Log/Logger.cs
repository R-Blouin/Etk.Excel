using System.ComponentModel.Composition;
using System.ComponentModel.Composition.Hosting;
using System.IO;
using System.Reflection;

namespace Etk.Tools.Log
{
    public sealed class Logger
    {
        #region attributes and properties

        [Import]
        private LoggerManager LoggerManager = null;

        private static volatile Logger instance;
        private static object syncRoot = new object();

        public static ILogger Instance
        {
            get
            {
                if (instance == null)
                {
                    lock (syncRoot)
                    {
                        if (instance == null)
                        {
                            instance = new Logger();
                            string currentAssemblyDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                            DirectoryCatalog catalog = new DirectoryCatalog(currentAssemblyDirectory);
                            CompositionContainer container = new CompositionContainer(catalog);
                            container.ComposeParts(instance);
                        } 
                   }
                }
                return instance.LoggerManager.Instance;
            }
        }
        #endregion

        private Logger()
        {}
    }
}
