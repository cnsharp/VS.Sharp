using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Serialization;
using CnSharp.VisualStudio.Extensions.Resources;
using stdole;

namespace CnSharp.VisualStudio.Extensions.Commands
{
    [XmlInclude(typeof(CommandMenu))]
    [XmlInclude(typeof(CommandButton))]
    public abstract class CommandControl
    {
        private Form _form;
        private int _position;
        private Image _image;
        private string _arg;
        private ICommand _command;

        protected CommandControl()
        {
            CommandActionType = CommandActionType.Menu;
            Position = 1;
        }

        [XmlAttribute("id")]
        public string Id { get; set; }

        [XmlAttribute("text")]
        public string Text { get; set; }

        [XmlAttribute("tooltip")]
        public string Tooltip { get; set; }

        [XmlAttribute("faceId")]
        public int FaceId { get; set; }

        /// <summary>
        ///     控件在父控件上的相对位置，相对于控件总数n而言，大于等于0则放在末尾n+1的位置，为负数则放在倒数第n-Position的位置
        /// </summary>
        [XmlAttribute("position")]
        public int Position
        {
            get { return _position; }
            set
            {
                if (value >= 0)
                    value = 1;
                _position = value;
            }
        }

        [XmlAttribute("picture")]
        public string Picture { get; set; }

        [XmlIgnore]
        public StdPicture StdPicture
        {
            get
            {
                if (!String.IsNullOrEmpty(Picture) && Plugin != null && Plugin.ResourceManager != null)
                {
                    return Plugin.ResourceManager.LoadPicture(Picture);
                }
                return null;
            }
        }

        [XmlIgnore]
        public Image Image
        {
            get
            {
                if (_image == null && !string.IsNullOrEmpty(Picture) && Picture.Trim().Length > 0 && Plugin != null && Plugin.ResourceManager != null)
                {
                    _image = Plugin.ResourceManager.LoadBitmap(Picture);
                }
                return _image;
            }
            set
            {
                _image = value;
            }
        }



        [XmlAttribute("class")]
        public string ClassName { get; set; }

        [XmlAttribute("type")]
        public CommandActionType CommandActionType { get; set; }

        [XmlAttribute("attachTo")]
        public string AttachTo { get; set; }

        //[XmlAttribute("hotKey")]
        //public string HotKey { get; set; }

        [XmlAttribute("beginGroup")]
        public bool BeginGroup { get; set; }

        [XmlIgnore]
        public ICommand Command
        {
            get { return _command ?? (_command = LoadInstance(ClassName) as ICommand); }
            set { _command = value; }
        }

        [XmlIgnore]
        public Plugin Plugin { get; set; }


        [XmlAttribute("arg")]
        public string Arg
        {
            get { return _arg; }
            set
            {
                _arg = value;
                Tag = _arg;
              
            }
        }


        [XmlAttribute("dependOn")]
        public string DependOn { get; set; }

          [XmlIgnore]
        public DependentItems DependentItems {
              get
              {
                  return string.IsNullOrEmpty(DependOn) || DependOn.Trim().Length == 0
                      ? DependentItems.None
                      : (DependentItems)Enum.Parse(typeof(DependentItems), DependOn);
              } }

        [XmlIgnore]
        public object Tag { get; set; }

        public override string ToString()
        {
            return Text;
        }

        public virtual void Execute()
        {
            var arg = Arg ?? Tag;
            switch (CommandActionType)
            {
                case CommandActionType.Program:
                    if (Command != null)
                    {
                        Command.Execute(arg);
                    }
                    break;
                case CommandActionType.Window:
                    var window = GetForm();
                    window.Show();
                    break;
                case CommandActionType.Dialog:
                    var dialog = GetForm();
                    dialog.ShowDialog();
                    break;
            }
        }


        public object LoadInstance(string typeName)
        {
            if (typeName.Contains(","))
            {
                var arr = typeName.Split(',');
                if (arr.Length < 2)
                    return null;
                var assemblyName = arr[1];
                try
                {
                    var assembly = Assembly.Load(assemblyName);
                    return assembly.CreateInstance(arr[0]);
                }
                catch
                {

                    var file = Path.Combine(Plugin.Location, assemblyName + ".dll");
                    if (File.Exists(file))
                    {
                        var assembly = Assembly.LoadFile(file);
                        return assembly.CreateInstance(arr[0]);
                    }
                }
            }


            return Plugin.Assembly.CreateInstance(typeName);

        }

        private Form GetForm()
        {
            if (_form != null && !_form.IsDisposed)
                return _form;
            _form = (Form)LoadInstance(ClassName);
            return _form;
        }
    }
}