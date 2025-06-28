using System.Windows.Controls;

namespace CnSharp.VisualStudio.Extensions.Controls
{
    /// <summary>
    /// WPF UserControl to hold a windows forms UserControl in it.
    /// </summary>
    public partial class WinFormsUserControlHolderControl : UserControl
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="WinFormsUserControlHolderControl"/> class.
        /// </summary>
        public WinFormsUserControlHolderControl(System.Windows.Forms.UserControl winFormsControl)
        {
            this.InitializeComponent();

            this.winFormsHost.Child = winFormsControl;
        }
    }
}