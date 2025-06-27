using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using EnvDTE;
using Microsoft.VisualStudio.CommandBars;

namespace CnSharp.VisualStudio.Extensions.Commands
{
    public class CommandBarAccessor : ICommandBarAccessor
    {
        private static readonly Host Host = Host.Instance;
        private readonly List<Command> _commands = new List<Command>();
        private readonly Dictionary<string, CommandBarControl> _commandBarControlDict =
            new Dictionary<string, CommandBarControl>();
        private readonly List<CommandControl> _commandControls = new List<CommandControl>();

        //public Plugin Plugin { get; set; }

        public void AddControl(CommandControl control)
        {
            //control.Plugin = Plugin;
            AddControl(control, false);
            RegisterCommandControl(control);
        }

        private void RegisterCommandControl(CommandControl control)
        {
            if (!_commandControls.Contains(control))
                _commandControls.Add(control);
            if (control is CommandMenu commandMenu)
            {
                if (commandMenu.SubMenus?.Any() == true)
                {
                    commandMenu.SubMenus.ForEach(RegisterCommandControl);
                }
            }
        }


        private CommandBarControl AddControl(CommandControl control, bool keepPosition)
        {
            if (control is CommandButton)
            {
                var btn = control as CommandButton;
                return AddButton(btn, keepPosition);
            }

            if (control is CommandMenu)
            {
                var menu = control as CommandMenu;
                return AddMenu(menu, keepPosition);
            }
            return null;
        }

        public void ResetControl(CommandControl control)
        {
            AddControl(control, true);
        }

        public void EnableControls(IEnumerable<string> ids, bool enabled)
        {
            foreach (var id in ids)
            {
                CommandBarControl control;
                var cc = _commandControls.Find(m => m.Id == id);
                if (IsDisposed(_commandBarControlDict[id]))
                {
                    _commandBarControlDict.Remove(id);
                    control = AddControl(cc, true);
                }
                else
                {
                    control = _commandBarControlDict[id];
                }

                var controlEnabled = enabled;
                if (cc.EnabledFunc != null)
                    controlEnabled = controlEnabled && cc.EnabledFunc();
                control.Enabled = controlEnabled;
                control.Visible = controlEnabled || cc.UnavailableState != ControlUnavailableState.Invisbile;
            }
        }

        private bool IsDisposed(CommandBarControl control)
        {
            try
            {
                var temp = control.Enabled;
                return false;
            }
            catch
            {
                return true;
            }
        }


        public void Delete()
        {
            foreach (var command in _commands)
            {
                command.Delete();
            }
            foreach (var c in _commandBarControlDict.Values)
            {
                try
                {
                    c.Delete(true);
                }
                catch
                {

                }
            }
        }


        private CommandBarButton AddCommandBarButton(CommandControl menu, bool keepPosition)
        {
            var bar = ResetCommandBar(menu, MsoControlType.msoControlButton, keepPosition);
            var btn = (CommandBarButton)bar;
                                                                                                                                                
            FormatCommandBarButton(menu, btn);
            if (!string.IsNullOrWhiteSpace(menu.ShortcutText))
                btn.ShortcutText = menu.ShortcutText;

            _commandBarControlDict.Add(menu.Id, btn);
            return btn;
        }

        private CommandBarPopup AddCommandBarButtonPop(CommandButton button, bool keepPosition)
        {

            var bar = ResetCommandBar(button, MsoControlType.msoControlPopup, keepPosition);
            var popup = (CommandBarPopup)bar;
            popup.Tag = button.Id;

            popup.Caption = button.Text;
            //StdPicture pic = button.StdPicture;
            //if (pic != null)
            //    popup.Picture = pic;
            popup.Visible = true;
            popup.BeginGroup = button.BeginGroup;
            _commandBarControlDict.Add(button.Id, popup);

            return popup;
        }


        private CommandBarControl AddMenu(CommandMenu menu, bool keepPosition)
        {
            if (keepPosition)
                menu.LoadSubMenus();
            if (menu.SubMenus.Count > 0)
            {
                CommandBarControl mainMenu = AddMainMenu(menu, keepPosition);
                var toolsPopup = (CommandBarPopup)mainMenu;
                menu.SubMenus.ForEach(m =>
                {
                    m.Plugin = menu.Plugin;
                    AddSubMenu(toolsPopup, m);
                }
                    );
                return mainMenu;
            }
            return AddCommandBarButton(menu, keepPosition);
        }

        private CommandBarControl ResetCommandBar(CommandControl control, MsoControlType controlType, bool keepPosition)
        {
            var cmdBars = (CommandBars)Host.DTE.CommandBars;
            CommandBar bar = cmdBars[control.AttachTo];
            IEnumerator controls = bar.Controls.GetEnumerator();

            var i = 0;
            while (controls.MoveNext())
            {
                i++;
                var c = (CommandBarControl)controls.Current;
                if (c != null && c.Tag == control.Id)
                {
                    RemoveCache(control);
                    c.Delete(false);
                    break;
                }
            }
            return bar.Controls.Add(controlType, Type.Missing,
             Type.Missing,
             keepPosition ? i : bar.Controls.Count + control.Position,
             true)
             ;
        }

        private void RemoveCache(CommandControl control)
        {
            var menu = control as CommandMenu;
            if (menu != null && menu.SubMenus.Count > 0)
            {
                foreach (var subMenu in menu.SubMenus)
                {
                    RemoveCache(subMenu);
                }
            }
            _commandBarControlDict.Remove(menu.Id);
        }

        public CommandBarControl AddMainMenu(CommandMenu menu, bool keepPosition)
        {
            const string menuBar = "MenuBar";

            if (string.IsNullOrWhiteSpace(menu.AttachTo))
            {
                menu.AttachTo = menuBar;
            }

            var btn = ResetCommandBar(menu, MsoControlType.msoControlPopup, keepPosition);
            btn.Tag = menu.Id;
            btn.Caption = menu.Text;
            SetState(btn,menu);
            if (menu.AttachTo != menuBar)
            {
                btn.BeginGroup = menu.BeginGroup;
            }
            return btn;
        }

        private void SetState(CommandBarControl bar,CommandControl control)
        {
            var ok = Host.IsDependencySatisfied(control.DependentItems);
            if (ok)
            {
                bar.Visible = bar.Enabled = true;
                return;
            }
            bar.Enabled = false;
            bar.Visible = control.UnavailableState != ControlUnavailableState.Invisbile;
        }

        private void AddSubMenu(CommandBarPopup parentMenu, CommandMenu menu, Command command = null)
        {
            var bar = (CommandBarButton)parentMenu.Controls.Add(
                MsoControlType.msoControlButton, Missing.Value, Missing.Value, Missing.Value, true);
            FormatCommandBarButton(menu, bar);
            bar.Tag = menu.Id;
            _commandBarControlDict.Add(menu.Id, bar);

            if (menu.SubMenus != null && menu.SubMenus.Count > 0)
            {
                menu.SubMenus.ForEach(sub => AddSubMenu((CommandBarPopup)bar, sub));
            }
        }


        private CommandBarControl AddButton(CommandButton button, bool keepPosition)
        {
            //Command cmd = AddCommond(button);
            if (keepPosition)
                button.LoadSubMenus();
            CommandBarControl cmdControl;
            if (button.SubMenus != null && button.SubMenus.Count > 0)
            {
                var popup = AddCommandBarButtonPop(button, keepPosition);
                button.SubMenus.ForEach(sub => AddSubMenu(popup, sub));
                cmdControl = popup;
            }
            else
            {
                CommandBarButton btn = AddCommandBarButton(button, keepPosition);
                btn.Style = MsoButtonStyle.msoButtonIcon;
                cmdControl = btn;

            }
            SetState(cmdControl,button);
            return cmdControl;
        }

        private void FormatCommandBarButton(CommandControl control, CommandBarButton btn)
        {
            btn.Caption = control.Text;
            var pic = control.StdPicture;
            if (pic != null)
                btn.Picture = pic;
            btn.Visible = true;
            btn.BeginGroup = control.BeginGroup;
            btn.Tag = control.Id;
            btn.Click += (CommandBarButton ctrl, ref bool @default) =>
            {
                if(control.Action != null)
                    control.Action.Invoke();
                else 
                    control.Execute();
            };
            SetState(btn,control);
        }
    }
}