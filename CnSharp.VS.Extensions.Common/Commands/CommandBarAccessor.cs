using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using CnSharp.VisualStudio.Extensions;
using CnSharp.VisualStudio.Extensions.Commands;
using EnvDTE;
using Microsoft.VisualStudio.CommandBars;
using stdole;

public class CommandBarAccessor : ICommandBarAccessor
{
    // Fields
    private static readonly Host Host = Host.Instance;
    private readonly List<CommandBarControl> _commandBarControls = new List<CommandBarControl>();
    private readonly IList<Command> _commands = new List<Command>();

    public void AddControl(CommandControl control)
    {
        AddControl(control, false);
    }

    public void Delete()
    {
        foreach (Command command in _commands)
        {
            command.Delete();
        }
        foreach (CommandBarControl control in _commandBarControls)
        {
            try
            {
                control.Delete(true);
            }
            catch
            {
            }
        }
    }

    public void EnableControls(IEnumerable<string> ids, bool enabled)
    {
        foreach (CommandBarControl control in from c in _commandBarControls
            where ids.Contains(c.Tag)
            select c)
        {
            control.Enabled = enabled;
        }
    }

    public void ResetControl(CommandControl control)
    {
        AddControl(control, true);
    }

    // Methods
    private CommandBarControl AddButton(CommandButton button, bool keepPosition)
    {
        if (keepPosition)
        {
            button.LoadSubMenus();
        }
        CommandBarControl control = null;
        if ((button.SubMenus != null) && (button.SubMenus.Count > 0))
        {
            CommandBarPopup popup = AddCommandBarButtonPop(button, keepPosition);
            button.SubMenus.ForEach(delegate(CommandMenu sub) { AddSubMenu(popup, sub, null); });
            control = popup;
        }
        else
        {
            CommandBarButton button2 = AddCommandBarButton(button, keepPosition);
            button2.Style = MsoButtonStyle.msoButtonIcon;
            control = button2;
        }
        control.Enabled = Host.IsDependencySatisfied(button.DependentItems);
        return control;
    }

    private CommandBarButton AddCommandBarButton(CommandControl menu, bool keepPosition)
    {
        var btn = (CommandBarButton) ResetCommandBar(menu, keepPosition);
        FormatCommandBarButton(menu, btn);
        _commandBarControls.Add(btn);
        return btn;
    }

    private CommandBarPopup AddCommandBarButtonPop(CommandButton button, bool keepPosition)
    {
        var item = (CommandBarPopup) ResetCommandBar(button, MsoControlType.msoControlPopup, keepPosition);
        item.Tag = button.Id;
        item.Caption = button.Text;
        item.Visible = true;
        item.BeginGroup = button.BeginGroup;
        _commandBarControls.Add(item);
        return item;
    }

    private CommandBarButton AddCommandMenu(CommandMenu menu, bool keepPosition)
    {
        CommandBarControl item = ResetCommandBar(menu, keepPosition);
        _commandBarControls.Add(item);
        var btn = (CommandBarButton) item;
        FormatCommandBarButton(menu, btn);
        return btn;
    }

    private Command AddCommond(CommandControl control)
    {
        Command item = GetCommand(control);
        var contextUIGUIDs = new object[0];
        if (item == null)
        {
            item = Host.DTE.Commands.AddNamedCommand(Host.AddIn, control.Id, control.Text, control.Tooltip, true,
                control.FaceId, ref contextUIGUIDs, 3);
        }
        _commands.Add(item);
        return item;
    }

    private CommandBarControl AddControl(CommandControl control, bool keepPosition)
    {
        if (control is CommandButton)
        {
            var button = control as CommandButton;
            return AddButton(button, keepPosition);
        }
        if (control is CommandMenu)
        {
            var menu = control as CommandMenu;
            return AddMenu(menu, keepPosition);
        }
        return null;
    }

    public CommandBarControl AddMainMenu(CommandMenu menu, bool keepPosition)
    {
        if (string.IsNullOrEmpty(menu.AttachTo) || (menu.AttachTo.Trim().Length == 0))
        {
            menu.AttachTo = "MenuBar";
        }
        CommandBarControl control = ResetCommandBar(menu, MsoControlType.msoControlPopup, keepPosition);
        control.Tag = menu.Id;
        control.Caption = menu.Text;
        control.Visible = true;
        control.Enabled = Host.IsDependencySatisfied(menu.DependentItems);
        if (menu.AttachTo != "MenuBar")
        {
            control.BeginGroup = menu.BeginGroup;
        }
        return control;
    }

    private CommandBarControl AddMenu(CommandMenu menu, bool keepPosition)
    {
        if (keepPosition)
        {
            menu.LoadSubMenus();
        }
        if (menu.SubMenus.Count > 0)
        {
            CommandBarControl control = AddMainMenu(menu, keepPosition);
            var toolsPopup = (CommandBarPopup) control;
            menu.SubMenus.ForEach(delegate(CommandMenu m)
            {
                m.Plugin = menu.Plugin;
                AddSubMenu(toolsPopup, m, null);
            });
            return control;
        }
        return AddCommandMenu(menu, keepPosition);
    }

    private void AddSubMenu(CommandBarPopup parentMenu, CommandMenu menu, Command command = null)
    {
        Action<CommandMenu> action = null;
        Command command2 = command ?? AddCommond(menu);
        var bar =
            (CommandBarButton) command2.AddControl(parentMenu.CommandBar, parentMenu.Controls.Count + menu.Position);
        FormatCommandBarButton(menu, bar);
        _commandBarControls.Add(bar);
        if ((menu.SubMenus != null) && (menu.SubMenus.Count > 0))
        {
            if (action == null)
            {
                action = sub => AddSubMenu((CommandBarPopup) bar, sub, null);
            }
            menu.SubMenus.ForEach(action);
        }
    }

    private void FormatCommandBarButton(CommandControl control, CommandBarButton btn)
    {
        btn.Caption = control.Text;
        StdPicture stdPicture = control.StdPicture;
        if (stdPicture != null)
        {
            btn.Picture = stdPicture;
        }
        btn.Visible = true;
        btn.BeginGroup = control.BeginGroup;
        btn.Tag = control.Id;
        btn.Enabled = Host.IsDependencySatisfied(control.DependentItems);
    }

    private Command GetCommand(CommandControl control)
    {
        Command command = null;
        _DTE dTE = Host.DTE;
        try
        {
            command = dTE.Commands.Item(Host.AddIn.ProgID + "." + control.Id, -1);
        }
        catch
        {
        }
        return command;
    }

    private void RemoveCache(CommandControl control)
    {
        var menu = control as CommandMenu;
        if ((menu != null) && (menu.SubMenus.Count > 0))
        {
            foreach (CommandMenu menu2 in menu.SubMenus)
            {
                RemoveCache(menu2);
            }
        }
        _commandBarControls.RemoveAll(c => c.Tag == control.Id);
    }

    private CommandBarControl ResetCommandBar(CommandControl control, bool keepPosition)
    {
        var commandBars = (CommandBars) Host.DTE.CommandBars;
        CommandBar owner = commandBars[control.AttachTo];
        IEnumerator enumerator = owner.Controls.GetEnumerator();
        int num = 0;
        while (enumerator.MoveNext())
        {
            num++;
            var current = (CommandBarControl) enumerator.Current;
            if ((current != null) && (current.Tag == control.Id))
            {
                RemoveCache(control);
                current.Delete(false);
                break;
            }
        }
        return
            (CommandBarControl)
                AddCommond(control).AddControl(owner, keepPosition ? num : (owner.Controls.Count + control.Position));
    }

    private CommandBarControl ResetCommandBar(CommandControl control, MsoControlType controlType, bool keepPosition)
    {
        var commandBars = (CommandBars) Host.DTE.CommandBars;
        CommandBar bar = commandBars[control.AttachTo];
        IEnumerator enumerator = bar.Controls.GetEnumerator();
        int num = 0;
        while (enumerator.MoveNext())
        {
            num++;
            var current = (CommandBarControl) enumerator.Current;
            if ((current != null) && (current.Tag == control.Id))
            {
                RemoveCache(control);
                current.Delete(false);
                break;
            }
        }
        return bar.Controls.Add(controlType, Type.Missing, Type.Missing,
            keepPosition ? num : (bar.Controls.Count + control.Position), true);
    }
}