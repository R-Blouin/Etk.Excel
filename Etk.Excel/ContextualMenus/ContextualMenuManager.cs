using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using Etk.Tools.Extensions;
using Etk.Tools.Log;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

namespace Etk.Excel.ContextualMenus
{
    [Export]
    [PartCreationPolicy(CreationPolicy.Shared)]
    class ContextualMenuManager : IContextualMenuManager, IDisposable
    {
        #region attributes and properties
        private ILogger log = Logger.Instance;

        private readonly object syncObj = new object();
        private bool isDisposed = false;

        private List<ContextualMenusRequestedHandler> contextualMenusManagers = new List<ContextualMenusRequestedHandler>();
        private List<Workbook> manageWorkbooks = new List<Workbook>();

        private Dictionary<string, IContextualMenu> contextualMenuByIdent = new Dictionary<string, IContextualMenu>();

        public event ContextualMenusRequestedHandler OnContextualMenusRequested
        {
            add
            {
                lock (syncObj)
                {
                    if (value != null && ! isDisposed)
                    {
                        if (!contextualMenusManagers.Contains(value))
                            contextualMenusManagers.Add(value);
                    }
                }
            }
            remove
            {
                lock (syncObj)
                {
                    if (value != null && ! isDisposed)
                    {
                        if (contextualMenusManagers.Contains(value))
                            contextualMenusManagers.Remove(value);
                    }
                }
            }
        }
        #endregion

        #region .ctors
        public ContextualMenuManager()
        {}
        #endregion

        #region public methods
        public void RegisterWorkbook(Workbook workbook)
        {
            if(workbook == null || isDisposed)
                return;

            lock (syncObj)
            {
                if(! manageWorkbooks.Contains(workbook))
                    workbook.SheetBeforeRightClick += OnSheetBeforeRightClickViewsManagement;
            }
        }

        public void UnRegisterWorkbook(Workbook workbook)
        {
            if (workbook == null)
                return;

            lock (syncObj)
            {
                if (manageWorkbooks.Contains(workbook))
                    workbook.SheetBeforeRightClick -= OnSheetBeforeRightClickViewsManagement;
            }
        }

        public void RegisterMenuDefinitionsFromXml(string xml)
        {
            try
            {
                IEnumerable<IContextualMenu> menus = ContextualMenuFactory.CreateInstances(xml);
                if (menus != null)
                {
                    lock ((contextualMenuByIdent as ICollection).SyncRoot)
                    {
                        foreach (IContextualMenu menu in menus)
                        {
                            if (menu != null)
                            {
                                if (contextualMenuByIdent.ContainsKey(menu.Name))
                                    log.LogFormat(LogType.Warn, "Menu {0} already registred.", menu.Name ?? string.Empty);
                                contextualMenuByIdent[menu.Name] = menu;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string message = xml.Length > 350 ? xml.Substring(0, 350) + "..." : xml;
                throw new EtkException(string.Format("Cannot create contextual menus from xml '{0}':{1}", message, ex.Message));
            }
        }

        public IContextualMenu GetContextualMenu(string name)
        {
            IContextualMenu ret = null;
            if (!String.IsNullOrEmpty(name))
            {
                lock ((contextualMenuByIdent as ICollection).SyncRoot)
                {
                    if(! contextualMenuByIdent.TryGetValue(name, out ret))
                        throw new Exception(string.Format("Cannot find contextual menu '{0}'", name ?? string.Empty));
                }
            }
            return ret;
        }

        public void Dispose()
        {
            lock (syncObj)
            {
                if (!isDisposed)
                {
                    isDisposed = true;

                    foreach (Workbook workbook in manageWorkbooks)
                        UnRegisterWorkbook(workbook);

                    contextualMenusManagers.Clear();
                    manageWorkbooks.Clear();
                }
            }
        }
        #endregion

        #region private methods
        /// <summary> Manage the user's right click: manage the menus creation</summary>
        /// <param name="sheet">The sheet where the right click is done</param>
        /// <param name="concernedRange">Concerned Range</param>
        /// <param name="cancel"></param>
        private void OnSheetBeforeRightClickViewsManagement(object sheet, Range range, ref bool cancel)
        {
            Microsoft.Office.Interop.Excel.Application application = range.Application;
            CommandBar commandBar = application.CommandBars["Cell"];
            commandBar.Reset();

            Range realRange = range.Cells[1, 1];
            foreach (ContextualMenusRequestedHandler manager in contextualMenusManagers)
            {
                try
                {
                    IEnumerable<IContextualMenu> menus = manager(sheet as Worksheet, realRange);
                    if (menus != null)
                    {
                        foreach (IContextualMenu menu in menus)
                            CreateMenus(commandBar.Controls, menu);
                    }
                }
                catch (Exception ex)
                { 
                    string methodName = manager.Method == null ? string.Empty : manager.Method.Name.EmptyIfNull();
                    throw new EtkException(string.Format("Contextual menu manager: '{0}' invocation failed: {1}.", methodName, ex.Message));
                }
            }
            realRange = null;
        }

        /// <summary> Create a menu and its submenus and manage their actions</summary>
        /// <param name="parentControls">Control where the menu must be inserted</param>
        /// <param name="contextualMenu">Menu to insert</param>
        private void CreateMenus(CommandBarControls parentControls, IContextualMenu contextualMenu)
        {
            if (contextualMenu != null && contextualMenu.Items != null)
            {
                foreach (IContextualPart part in contextualMenu.Items)
                {
                    if (part is IContextualMenu)
                    {
                        IContextualMenu contextualSubMenu = part as IContextualMenu;
                        CommandBarPopup subMenu = (CommandBarPopup) parentControls.Add(MsoControlType.msoControlPopup, Type.Missing, Type.Missing, Type.Missing, true);
                        subMenu.Caption = contextualSubMenu.Caption;
                        subMenu.BeginGroup = contextualSubMenu.BeginGroup;
                        CreateMenus(subMenu.Controls, contextualSubMenu);
                    }
                    else
                    {
                        ContextualMenuItem contextualMenuItem = part as ContextualMenuItem;
                        if (contextualMenuItem != null)
                        {
                            MsoControlType menuItem = MsoControlType.msoControlButton;
                            CommandBarButton commandBarButton = (CommandBarButton) parentControls.Add(menuItem, Type.Missing, Type.Missing, Type.Missing, true);
                            commandBarButton.Style = MsoButtonStyle.msoButtonIconAndCaption;
                            commandBarButton.Caption = contextualMenuItem.Caption;
                            commandBarButton.BeginGroup = contextualMenuItem.BeginGroup;
                            if (contextualMenuItem.FaceId != 0)
                                commandBarButton.FaceId = contextualMenuItem.FaceId;

                            commandBarButton.Click += (CommandBarButton ctrl, ref bool cancel1) =>
                                                      {
                                                          try
                                                          {
                                                              if(contextualMenuItem.Action != null)
                                                                  contextualMenuItem.Action(); 
                                                          }
                                                          catch (Exception ex)
                                                          {
                                                            string methodName = contextualMenuItem.MethodInfo == null ? string.Empty : contextualMenuItem.MethodInfo.Name; 
                                                            throw new EtkException(string.Format("Contextual menu: '{0}' invocation failed: {1}.", methodName, ex.Message));
                                                          }
                                                      };
                        }
                    }
                }
            }
        }
        #endregion
    }
}
