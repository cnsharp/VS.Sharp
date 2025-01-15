using System.Collections.Generic;

namespace CnSharp.VisualStudio.Extensions.Commands
{
    public interface ICommandBarAccessor
    {
        void AddControl(CommandControl control);
        void ResetControl(CommandControl control);
        void EnableControls(IEnumerable<string> ids ,bool enabled);
        void Delete();
    }
}
