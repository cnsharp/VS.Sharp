using System;
using System.Collections.Generic;
using System.Linq;

namespace CnSharp.VisualStudio.Extensions.Commands
{
    public class CommandManager
    {
        private readonly ICommandBarAccessor _commandBarAccessor;
        private readonly Plugin _plugin;


        public CommandManager(Plugin plugin)
        {
            if (Host.Plugins.Any(m => m.Id == plugin.Id))
            {
                throw new InvalidOperationException($"plugin {plugin.Id} already exists.");
            }

            CommandConfig = plugin.CommandConfig;
            _plugin = plugin;
            plugin.CommandManager = this;

            _commandBarAccessor = new CommandBarAccessor();

            Host.Plugins.Add(plugin);
        }

        public CommandConfig CommandConfig { get; }

        public virtual void Load()
        {
            CommandConfig.Buttons.ForEach(m =>
            {
                m.Plugin = _plugin;
                _commandBarAccessor.AddControl(m);
            });
            CommandConfig.Menus.ForEach(m =>
            {
                m.Plugin = _plugin;
                _commandBarAccessor.AddControl(m);
            });
            CommandConfig.ContextMenus.ForEach(m =>
            {
                m.Plugin = _plugin;
                _commandBarAccessor.AddControl(m);
            });
        }

        public void Execute(string commandName)
        {
            CommandControl control = FindCommandControl(commandName);
            if (control != null)
            {
                control.Plugin = _plugin;
                control.Execute();
            }
        }

        private CommandControl FindCommandControl(string commandName)
        {
            var menu = FindCommandMenu(CommandConfig.Menus, commandName);
            if (menu != null)
            {
                return menu;
            }

            var button = FindCommandMenu(CommandConfig.Buttons, commandName);
            if (button != null)
            {
                return button;
            }

            menu = FindCommandMenu(CommandConfig.ContextMenus, commandName);

            return menu;
        }

        private CommandMenu FindCommandMenu(IEnumerable<CommandMenu> menus, string commandName)
        {
            foreach (var m in menus)
            {
                if (string.Compare(m.Id, commandName, StringComparison.OrdinalIgnoreCase) == 0)
                    return m;
                if (m.SubMenus.Count > 0)
                {
                    var subMenu = FindCommandMenu(m.SubMenus, commandName);
                    if (subMenu != null)
                        return subMenu;
                }
            }
            return null;
        }

        public void Reset(Type type)
        {
            var control = FindCommandControl(type);
            if (control == null)
                return;
            _commandBarAccessor.ResetControl(control);
        }

        public void Reset(string commandName)
        {
            var control = FindCommandControl(commandName);
            if (control == null)
                return;
            _commandBarAccessor.ResetControl(control);
        }


        public void ApplyDependencies(DependentItems items, bool enabled)
        {
            var ids = new List<string>();

            ids.AddRange(GetMatchedMenus(CommandConfig.Menus, items));
            ids.AddRange(GetMatchedMenus(CommandConfig.Buttons, items));
            ids.AddRange(GetMatchedMenus(CommandConfig.ContextMenus, items));

            _commandBarAccessor.EnableControls(ids, enabled);
        }


        private IEnumerable<string> GetMatchedMenus(IEnumerable<CommandMenu> menus, DependentItems items)
        {
            var menuIds = new List<string>();
            foreach (var menu in menus)
            {
                if (menu.SubMenus?.Any() == true)
                {
                    menuIds.AddRange(GetMatchedMenus(menu.SubMenus, items));
                }
                if (menu.DependentItems.HasFlag(items))
                    menuIds.Add(menu.Id);
            }
            return menuIds;
        }

        public CommandControl FindCommandControl(Type type)
        {
            string typeName = type.FullName;
            var menu = FindCommandMenuByClassName(CommandConfig.Menus, typeName);
            if (menu != null)
            {
                return menu;
            }

            var button = FindCommandMenuByClassName(CommandConfig.Buttons, typeName);
            if (button != null)
            {
                return button;
            }

            menu = FindCommandMenuByClassName(CommandConfig.ContextMenus, typeName);

            return menu;
        }


        private CommandMenu FindCommandMenuByClassName(IEnumerable<CommandMenu> menus, string typeName)
        {
            foreach (var m in menus)
            {
                if (string.Compare(m.ClassName, typeName, StringComparison.OrdinalIgnoreCase) == 0)
                    return m;
                if (m.SubMenus.Count > 0)
                    return FindCommandMenu(m.SubMenus, typeName);
            }
            return null;
        }


        public virtual void Disconnect()
        {
            _commandBarAccessor.Delete();
        }

    }
}