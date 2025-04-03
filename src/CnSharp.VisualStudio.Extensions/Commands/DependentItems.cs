using System;

namespace CnSharp.VisualStudio.Extensions.Commands
{
    [Flags]
    public enum DependentItems
    {
        None,
        Document,
        SolutionProject
    }
}
