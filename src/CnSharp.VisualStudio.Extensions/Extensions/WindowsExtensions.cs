using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;

namespace CnSharp.VisualStudio.Extensions
{
    public static class WindowsExtensions
    {
        public static IVsWindowFrame Dock(this IVsWindowFrame windowFrame, ToolWindowOrientation orientation)
        {
            if (orientation == ToolWindowOrientation.none)
                return windowFrame;
            windowFrame.SetProperty(VsIds.VSFPROPID_DockOwner, ToolWindowGuids.GetGuid(orientation));
            return windowFrame;
        }

        public static IVsWindowFrame DockAsTabbed(this IVsWindowFrame windowFrame)
        {
            windowFrame.SetProperty(VsIds.VSFPROPID_DockStyle, VsIds.VSIDS_DockStyleAlwaysTabbed);
            return windowFrame;
        }

        public static IVsWindowFrame Relayout(this IVsWindowFrame windowFrame)
        {
            windowFrame.SetProperty((int)__VSFPROPID.VSFPROPID_FrameMode, (int)VSFRAMEMODE.VSFM_Dock);
            return windowFrame;
        }
    }

    public static class VsIds
    {
        // From Microsoft.VisualStudio.Shell.Interop.__VSFPROPID
        public const int VSFPROPID_DockOwner = -8027;
        public const int VSFPROPID_DockStyle = -8028;
        public const int VSFPROPID_FrameMode = -8029;

        // From Microsoft.VisualStudio.Shell.Interop.VSIDS
        public const int VSIDS_DockStyleAlwaysTabbed = 1;

        // From Microsoft.VisualStudio.Shell.Interop.VSFRAMEMODE
        public const int VSFM_Dock = 1;
    }


    public static class ToolWindowGuids
    {
        public static readonly Guid BottomPane = new Guid("34E76E81-EE4A-11D0-AE2E-00A0C90FFFC3");
        public static readonly Guid RightPane = new Guid("D5090FBD-2F05-4DBA-ABFA-35FD0A325348");
        public static readonly Guid LeftPane = new Guid("D5090FBD-2F05-4DBA-ABFA-35FD0A325347");
        public static readonly Guid TopPane = new Guid("574B6C9A-D0F8-45B4-BAB9-DFF1D2B073E5");
        public static readonly Guid DocumentPane = new Guid("CD1A2C6E-2F8E-4BBD-9CE3-FF0085B0F4EF");

        public static Guid? GetGuid(ToolWindowOrientation orientation)
        {
            switch (orientation)
            {
                case ToolWindowOrientation.Top:
                    return TopPane;
                case ToolWindowOrientation.Left:
                    return LeftPane;
                case ToolWindowOrientation.Right:
                    return RightPane;
                case ToolWindowOrientation.Bottom:
                    return BottomPane;
            }

            return null;
        }
    }
}