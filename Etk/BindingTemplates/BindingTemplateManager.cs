﻿using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using Etk.BindingTemplates.Definitions.Templates;
using Etk.BindingTemplates.Views;
using Etk.Tools.Log;

namespace Etk.BindingTemplates
{
    [Export]
    [PartCreationPolicy(CreationPolicy.Shared)]
    public sealed class BindingTemplateManager
    {
        private static readonly object syncRoot = new object();

        #region .ctors
        private BindingTemplateManager()
        {}
        #endregion

        #region attributes and properties
        private ILogger log = Logger.Instance;

        private readonly Dictionary<string, TemplateDefinition> templateDescriptionById = new Dictionary<string, TemplateDefinition>();
        private readonly Dictionary<Guid, ITemplateView> viewById = new Dictionary<Guid, ITemplateView>();
        private readonly Dictionary<string, List<ITemplateView>> viewsByTemplateDefinition = new Dictionary<string, List<ITemplateView>>();
        #endregion

        #region public methods
        public void RegisterTemplateDefinition(TemplateDefinition definition)
        {
            if(definition != null)
            {
                lock (syncRoot)
                {
                    templateDescriptionById[definition.Name] = definition;
                    viewsByTemplateDefinition[definition.Name] = new List<ITemplateView>();
                }
            }
        }

        public TemplateDefinition GetTemplateDefinition(string name)
        {
            TemplateDefinition definition = null;
            if (! string.IsNullOrEmpty(name))
            {
                lock (syncRoot)
                {
                    templateDescriptionById.TryGetValue(name, out definition);
                }
            }
            return definition;
        }

        public void AddView(ITemplateView view)
        {
            if (view != null)
            {
                lock (syncRoot)
                {
                    try
                    {
                        if (view.TemplateDefinition == null)
                            throw new BindingTemplateException("the template dataAccessor cannot be null");
                        if (GetTemplateDefinition(view.TemplateDefinition.Name) != null)

                        viewsByTemplateDefinition[view.TemplateDefinition.Name].Add(view);
                        viewById[view.Ident] = view;
                    }
                    catch (Exception ex)
                    {
                        throw new BindingTemplateException($"Cannot add view '{view.Ident}'.{ex.Message}", ex);
                    }
                }
            }
        }

        public ITemplateView GetView(Guid ident)
        {
            ITemplateView view = null;
            if (ident != null)
            {
                lock (syncRoot)
                {
                    viewById.TryGetValue(ident, out view);
                }
            }
            return view;
        }

        public IEnumerable<ITemplateView> GetAllViews()
        {
            return viewById.Values;
        }

        public void RemoveView(ITemplateView view)
        {
            if (view != null && view.Ident != null )
            {
                lock (syncRoot)
                {
                    try
                    {
                        viewById.Remove(view.Ident);
                        if (view.TemplateDefinition != null)
                        {
                            if (viewsByTemplateDefinition.ContainsKey(view.TemplateDefinition.Name))
                                viewsByTemplateDefinition[view.TemplateDefinition.Name].Remove(view);
                        }
                        view.Dispose();
                    }
                    catch (Exception ex)
                    {
                        string message = $"Remove view '{view.Ident}' failed. {ex.Message}";
                        throw new BindingTemplateException(message);
                    }
                }
            }
        }

        public void RemoveViews(IEnumerable<ITemplateView> views)
        {
            if (views != null)
            {
                try
                {
                    lock (syncRoot)
                    {
                        bool success = true;
                        foreach (ITemplateView view in views)
                        {
                            try { RemoveView(view); }
                            catch { success = false; }
                        }
                        if (!success)
                            throw new BindingTemplateException("No all views have been removed. Please check the logs.");
                    }
                }
                catch (Exception ex)
                {
                    string message = $"Remove views failed. {ex.Message}";
                    throw new BindingTemplateException(message);
                }
            }
        } 
        #endregion
    }
}
