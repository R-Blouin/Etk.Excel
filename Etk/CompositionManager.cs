using System;
using System.ComponentModel.Composition;
using System.ComponentModel.Composition.Hosting;
using System.Reflection;

namespace Etk
{
    public class CompositionManager
    {
        #region singleton
        private static readonly Lazy<CompositionManager> instance = new Lazy<CompositionManager>(() => 
                                                                        { 
                                                                            CompositionManager manager = Activator.CreateInstance(typeof(CompositionManager), true) as CompositionManager;
                                                                            manager.Init();
                                                                            return manager;
                                                                        });

        public static CompositionManager Instance { get { return instance.Value; } }
        #endregion

        #region attributes and properties
        private CompositionContainer container;   
        #endregion

        #region .ctors
        private CompositionManager()
        { }
        #endregion

        #region public methods
        public void ComposeParts(object o)
        {
            container.ComposeParts(o);
        }

        public void ComposeExportedValue<T>(T o)
        {
            container.ComposeExportedValue<T>(o);
        }

        public T GetExportedValue<T>()
        {
            return container.GetExportedValue<T>();
        }
        #endregion


        #region private methods
        private void Init()
        {
            AggregateCatalog aggregateCatalog = new AggregateCatalog();
            // Add a catalog from the 'Etk.Excel' assembly
            aggregateCatalog.Catalogs.Add(new AssemblyCatalog(Assembly.Load("Etk.Excel")));
            // Add a catalog from the 'Etk.Excel' assembly
            aggregateCatalog.Catalogs.Add(new AssemblyCatalog(Assembly.Load("Etk")));

            // Creation container
            container = new CompositionContainer(aggregateCatalog);
        }
        #endregion

    }
}
